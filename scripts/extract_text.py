"""Extract text from source documents into mirrored .txt files.

Supported inputs: PDF, DOCX, TXT, MD, HTML, HTM.
"""

from __future__ import annotations

import argparse
import html
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree


SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".txt", ".md", ".html", ".htm"}


def extract_pdf(path: Path) -> str:
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError("PyMuPDF is required for PDF extraction. Run scripts/check_dependencies.py.") from exc

    chunks: list[str] = []
    with fitz.open(path) as document:
        for index, page in enumerate(document, start=1):
            text = page.get_text("text").strip()
            if text:
                chunks.append(f"\n\n--- Page {index} ---\n{text}")
    return "\n".join(chunks).strip()


def extract_docx(path: Path) -> str:
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(path) as archive:
        with archive.open("word/document.xml") as document_xml:
            root = ElementTree.parse(document_xml).getroot()

    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:p", namespace):
        text_parts = [node.text or "" for node in paragraph.findall(".//w:t", namespace)]
        text = "".join(text_parts).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def extract_plain(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def extract_html(path: Path) -> str:
    raw = extract_plain(path)
    raw = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", raw)
    raw = re.sub(r"(?s)<[^>]+>", " ", raw)
    raw = html.unescape(raw)
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]+", " ", raw)).strip()


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".docx":
        return extract_docx(path)
    if suffix in {".txt", ".md"}:
        return extract_plain(path)
    if suffix in {".html", ".htm"}:
        return extract_html(path)
    raise ValueError(f"Unsupported file type: {path}")


def iter_input_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() in SUPPORTED_EXTENSIONS else []
    return sorted(
        path
        for path in input_path.rglob("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def output_path_for(source: Path, input_root: Path, output_root: Path) -> Path:
    base = input_root if input_root.is_dir() else input_root.parent
    relative = source.relative_to(base)
    return output_root / relative.with_suffix(relative.suffix + ".txt")


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract text from documents into intermediate .txt files.")
    parser.add_argument("input", type=Path, help="Input file or directory.")
    parser.add_argument("--output", type=Path, default=Path("outputs/text"), help="Output directory.")
    args = parser.parse_args()

    input_path = args.input.resolve()
    output_root = args.output.resolve()

    if not input_path.exists():
        print(f"Input path does not exist: {input_path}", file=sys.stderr)
        return 2

    files = iter_input_files(input_path)
    if not files:
        print(f"No supported files found in {input_path}")
        return 0

    had_errors = False

    for source in files:
        destination = output_path_for(source, input_path, output_root)
        destination.parent.mkdir(parents=True, exist_ok=True)
        try:
            text = extract_text(source)
        except Exception as exc:
            print(f"ERROR: {source}: {exc}", file=sys.stderr)
            had_errors = True
            continue
        destination.write_text(text, encoding="utf-8")
        print(f"Wrote {destination}")

    return 1 if had_errors else 0


if __name__ == "__main__":
    sys.exit(main())
