"""Collect case image assets from local background files or explicit image URLs.

The script is intentionally small and dependency-light:

- PDF image extraction uses PyMuPDF, already required by the repository.
- DOCX image extraction reads the Office zip package directly.
- HTML image discovery uses Python's standard library.
- URL downloads use urllib from the standard library.

For web-sourced images, first use Codex web/image search to identify suitable
public image URLs and source pages, then pass the image URLs to this script with
matching --source-note values so the generated manifest preserves traceability.
The workflow must not generate its own diagrams or synthetic images.
"""

from __future__ import annotations

import argparse
import hashlib
import html.parser
import json
import mimetypes
import shutil
import sys
import urllib.parse
import urllib.request
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path


IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff"}
DEFAULT_USER_AGENT = "case-study-writer/1.0"


@dataclass
class ImageRecord:
    path: str
    source: str
    source_note: str
    kind: str
    bytes: int
    sha256: str


class ImageHTMLParser(html.parser.HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.images: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "img":
            return
        attr_map = {name.lower(): value for name, value in attrs if value}
        src = attr_map.get("src")
        if src:
            self.images.append(src)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def unique_path(output_dir: Path, stem: str, suffix: str) -> Path:
    clean_stem = "".join(char if char.isalnum() or char in "-_" else "_" for char in stem).strip("_")
    if not clean_stem:
        clean_stem = "image"
    candidate = output_dir / f"{clean_stem}{suffix.lower()}"
    counter = 2
    while candidate.exists():
        candidate = output_dir / f"{clean_stem}_{counter}{suffix.lower()}"
        counter += 1
    return candidate


def write_record(path: Path, source: str, source_note: str, kind: str) -> ImageRecord:
    return ImageRecord(
        path=str(path),
        source=source,
        source_note=source_note,
        kind=kind,
        bytes=path.stat().st_size,
        sha256=sha256(path),
    )


def copy_image(path: Path, output_dir: Path, source_note: str) -> ImageRecord | None:
    if path.suffix.lower() not in IMAGE_SUFFIXES:
        return None
    destination = unique_path(output_dir, path.stem, path.suffix)
    shutil.copy2(path, destination)
    return write_record(destination, str(path), source_note or path.name, "file")


def extract_pdf_images(path: Path, output_dir: Path, source_note: str, min_bytes: int) -> list[ImageRecord]:
    import fitz

    records: list[ImageRecord] = []
    document = fitz.open(path)
    try:
        for page_index in range(document.page_count):
            page = document[page_index]
            for image_index, image in enumerate(page.get_images(full=True), start=1):
                xref = image[0]
                info = document.extract_image(xref)
                data = info.get("image")
                if not data or len(data) < min_bytes:
                    continue
                extension = f".{info.get('ext') or 'png'}"
                destination = unique_path(output_dir, f"{path.stem}_p{page_index + 1}_img{image_index}", extension)
                destination.write_bytes(data)
                note = source_note or f"{path.name}, page {page_index + 1}, image {image_index}"
                records.append(write_record(destination, str(path), note, "pdf"))
    finally:
        document.close()
    return records


def extract_docx_images(path: Path, output_dir: Path, source_note: str, min_bytes: int) -> list[ImageRecord]:
    records: list[ImageRecord] = []
    with zipfile.ZipFile(path) as archive:
        for member in archive.namelist():
            if not member.startswith("word/media/"):
                continue
            data = archive.read(member)
            if len(data) < min_bytes:
                continue
            suffix = Path(member).suffix or ".png"
            destination = unique_path(output_dir, f"{path.stem}_{Path(member).stem}", suffix)
            destination.write_bytes(data)
            note = source_note or f"{path.name}, embedded image {Path(member).name}"
            records.append(write_record(destination, str(path), note, "docx"))
    return records


def discover_html_images(path: Path, output_dir: Path, source_note: str, min_bytes: int) -> list[ImageRecord]:
    parser = ImageHTMLParser()
    parser.feed(path.read_text(encoding="utf-8", errors="ignore"))

    records: list[ImageRecord] = []
    base = path.parent
    for index, src in enumerate(parser.images, start=1):
        parsed = urllib.parse.urlparse(src)
        if parsed.scheme in {"http", "https"}:
            continue
        image_path = (base / urllib.parse.unquote(parsed.path)).resolve()
        if not image_path.exists() or image_path.stat().st_size < min_bytes:
            continue
        record = copy_image(image_path, output_dir, source_note or f"{path.name}, img {index}: {src}")
        if record:
            record.kind = "html"
            records.append(record)
    return records


def collect_local(source: Path, output_dir: Path, source_note: str, min_bytes: int) -> list[ImageRecord]:
    if source.is_dir():
        records: list[ImageRecord] = []
        for path in sorted(item for item in source.rglob("*") if item.is_file()):
            records.extend(collect_local(path, output_dir, "", min_bytes))
        return records

    suffix = source.suffix.lower()
    if suffix in IMAGE_SUFFIXES:
        record = copy_image(source, output_dir, source_note)
        return [record] if record and record.bytes >= min_bytes else []
    if suffix == ".pdf":
        return extract_pdf_images(source, output_dir, source_note, min_bytes)
    if suffix == ".docx":
        return extract_docx_images(source, output_dir, source_note, min_bytes)
    if suffix in {".html", ".htm"}:
        return discover_html_images(source, output_dir, source_note, min_bytes)
    return []


def extension_from_response(url: str, content_type: str | None) -> str:
    url_suffix = Path(urllib.parse.urlparse(url).path).suffix.lower()
    if url_suffix in IMAGE_SUFFIXES:
        return url_suffix
    guessed = mimetypes.guess_extension((content_type or "").split(";")[0].strip())
    if guessed and guessed.lower() in IMAGE_SUFFIXES:
        return guessed.lower()
    return ".jpg"


def download_url(url: str, output_dir: Path, source_note: str, min_bytes: int) -> ImageRecord | None:
    request = urllib.request.Request(url, headers={"User-Agent": DEFAULT_USER_AGENT})
    with urllib.request.urlopen(request, timeout=30) as response:
        data = response.read()
        content_type = response.headers.get("Content-Type")
    if len(data) < min_bytes:
        return None
    suffix = extension_from_response(url, content_type)
    stem = Path(urllib.parse.urlparse(url).path).stem or "web_image"
    destination = unique_path(output_dir, stem, suffix)
    destination.write_bytes(data)
    return write_record(destination, url, source_note or url, "url")


def write_manifest(records: list[ImageRecord], output_dir: Path) -> None:
    manifest_json = output_dir / "manifest.json"
    manifest_md = output_dir / "manifest.md"
    manifest_json.write_text(json.dumps([asdict(record) for record in records], ensure_ascii=False, indent=2), encoding="utf-8")

    lines = ["# Image Asset Manifest", ""]
    for index, record in enumerate(records, start=1):
        lines.extend(
            [
                f"## {index}. {Path(record.path).name}",
                f"- Path: `{record.path}`",
                f"- Kind: {record.kind}",
                f"- Source: {record.source}",
                f"- Source note: {record.source_note}",
                f"- Bytes: {record.bytes}",
                f"- SHA256: `{record.sha256}`",
                "",
            ]
        )
    manifest_md.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Collect image assets for case drafts.")
    parser.add_argument("sources", nargs="*", type=Path, help="Local background files or folders to scan.")
    parser.add_argument("--url", action="append", default=[], help="Direct image URL discovered by web/image search. Repeatable.")
    parser.add_argument(
        "--source-note",
        action="append",
        default=[],
        help="Source note for matching --url entries, such as title, publisher, date, and source page.",
    )
    parser.add_argument("--output", type=Path, default=Path("outputs/drafts/images"), help="Output image directory.")
    parser.add_argument("--min-bytes", type=int, default=4096, help="Skip tiny decorative images below this size.")
    args = parser.parse_args()

    if args.source_note and len(args.source_note) != len(args.url):
        parser.error("--source-note must be provided the same number of times as --url, or omitted entirely.")

    args.output.mkdir(parents=True, exist_ok=True)
    records: list[ImageRecord] = []

    for source in args.sources:
        if not source.exists():
            print(f"Skipping missing source: {source}", file=sys.stderr)
            continue
        records.extend(collect_local(source, args.output, "", args.min_bytes))

    for index, url in enumerate(args.url):
        note = args.source_note[index] if args.source_note else url
        record = download_url(url, args.output, note, args.min_bytes)
        if record:
            records.append(record)

    write_manifest(records, args.output)
    print(f"Wrote {len(records)} image asset(s) to {args.output}")
    print(f"Wrote manifest files to {args.output / 'manifest.json'} and {args.output / 'manifest.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
