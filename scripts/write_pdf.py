"""Render a DOCX, UTF-8 text, or Markdown-like draft to PDF.

For style fidelity, final deliverables should use DOCX input and the Word or
LibreOffice backend. For DOCX inputs, the default backend first tries Microsoft
Word COM on Windows, then LibreOffice, then a lightweight PyMuPDF text renderer.
Text input is an emergency fallback and will not fully match DOCX layout.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree


PAGE_WIDTH = 595
PAGE_HEIGHT = 842
MARGIN_X = 90
MARGIN_Y = 76
HEADER_Y = 51
FOOTER_Y = 790
FONT_SIZE = 12
HEADER_FONT_SIZE = 9
LINE_HEIGHT = 20
IMAGE_PATTERN = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)]+)\)$")
HEADER_LEFT_TEXT = "商学院教学案例库"

SOFFICE_CANDIDATES = [
    "soffice",
    "libreoffice",
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
]

WORD_PDF_FORMAT = 17


def find_default_font() -> str | None:
    candidates = [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simsun.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("/System/Library/Fonts/PingFang.ttc"),
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def normalize_markdown(text: str) -> str:
    text = re.sub(r"^#{1,6}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.replace("\r\n", "\n").replace("\r", "\n")


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


def split_wrapped_lines(text: str, max_units: int = 68) -> list[str]:
    lines: list[str] = []
    for paragraph in text.split("\n"):
        if not paragraph.strip():
            lines.append("")
            continue

        current = ""
        current_units = 0
        for char in paragraph.strip():
            units = 2 if ord(char) > 127 else 1
            if current and current_units + units > max_units:
                lines.append(current)
                current = char
                current_units = units
            else:
                current += char
                current_units += units
        if current:
            lines.append(current)
    return lines


def find_soffice() -> str | None:
    for candidate in SOFFICE_CANDIDATES:
        resolved = shutil.which(candidate) if not Path(candidate).is_absolute() else candidate
        if resolved and Path(resolved).exists():
            return str(resolved)
    return None


def write_pdf_with_libreoffice(source: Path, destination: Path) -> None:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice soffice executable was not found.")

    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        command = [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            temp_dir,
            str(source.resolve()),
        ]
        completed = subprocess.run(command, capture_output=True, text=True, timeout=120)
        if completed.returncode != 0:
            message = completed.stderr.strip() or completed.stdout.strip()
            raise RuntimeError(f"LibreOffice PDF conversion failed: {message}")

        converted = Path(temp_dir) / f"{source.stem}.pdf"
        if not converted.exists():
            raise RuntimeError("LibreOffice did not produce the expected PDF output.")
        shutil.move(str(converted), destination)


def write_pdf_with_word(source: Path, destination: Path) -> None:
    if source.suffix.lower() != ".docx":
        raise ValueError("The word backend only supports .docx input.")

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for Microsoft Word PDF export.") from exc

    destination.parent.mkdir(parents=True, exist_ok=True)
    source_path = str(source.resolve())
    destination_path = str(destination.resolve())

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(source_path, ReadOnly=True)
        document.ExportAsFixedFormat(destination_path, WORD_PDF_FORMAT)
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def insert_text_line(page, line: str, y: float, font_path: str | None) -> None:
    page.insert_text(
        (MARGIN_X, y),
        line,
        fontsize=FONT_SIZE,
        fontname="casefont" if font_path else "helv",
        fontfile=font_path,
    )


def insert_centered_text(page, text: str, y: float, fontsize: int, font_path: str | None) -> None:
    text_width = fitz_text_width(text, fontsize)
    page.insert_text(
        ((PAGE_WIDTH - text_width) / 2, y),
        text,
        fontsize=fontsize,
        fontname="casefont" if font_path else "helv",
        fontfile=font_path,
    )


def fitz_text_width(text: str, fontsize: int) -> float:
    units = sum(2 if ord(char) > 127 else 1 for char in text)
    return units * fontsize * 0.5


def infer_document_kind(source: Path, text: str) -> str:
    joined = f"{source.name} {text[:80]}"
    if "使用说明" in joined or "teaching" in joined.lower() or "note" in joined.lower():
        return "案例使用说明"
    return "案例正文"


def infer_company_suffix(*paths: Path) -> str:
    for path in paths:
        stem = path.stem
        for prefix in ("案例正文_", "案例使用说明_"):
            if stem.startswith(prefix):
                suffix = stem[len(prefix) :].strip()
                if suffix:
                    return suffix
    return ""


def heading_text(raw_line: str) -> str | None:
    match = re.match(r"^#{1,6}\s+(.+)$", raw_line.strip())
    if match:
        return normalize_markdown(match.group(1)).strip()
    return None


def insert_template_title(page, title: str, kind: str, font_path: str | None) -> float:
    page.insert_text(
        (MARGIN_X, 99),
        f"{kind}：",
        fontsize=15,
        fontname="casefont" if font_path else "helv",
        fontfile=font_path,
    )
    insert_centered_text(page, title, 134, 16, font_path)
    return 194


def new_template_page(document, font_path: str | None, company_suffix: str = ""):
    page = document.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
    page_number = document.page_count
    page.insert_text(
        (MARGIN_X, HEADER_Y),
        HEADER_LEFT_TEXT,
        fontsize=HEADER_FONT_SIZE,
        fontname="casefont" if font_path else "helv",
        fontfile=font_path,
    )
    if company_suffix:
        right_width = fitz_text_width(company_suffix, HEADER_FONT_SIZE)
        page.insert_text(
            (PAGE_WIDTH - MARGIN_X - right_width, HEADER_Y),
            company_suffix,
            fontsize=HEADER_FONT_SIZE,
            fontname="casefont" if font_path else "helv",
            fontfile=font_path,
        )
    footer_text = str(page_number)
    footer_width = len(footer_text) * HEADER_FONT_SIZE * 0.5
    page.insert_text(
        ((PAGE_WIDTH - footer_width) / 2, FOOTER_Y),
        footer_text,
        fontsize=HEADER_FONT_SIZE,
        fontname="casefont" if font_path else "helv",
        fontfile=font_path,
    )
    return page


def write_pdf_with_pymupdf(source: Path, destination: Path, font_path: str | None) -> None:
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError("PyMuPDF is required. Run scripts/check_dependencies.py.") from exc

    if source.suffix.lower() == ".docx":
        text = extract_docx(source)
    else:
        text = source.read_text(encoding="utf-8-sig")
    
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    kind = infer_document_kind(source, text)
    company_suffix = infer_company_suffix(destination, source)
    title_written = False
    destination.parent.mkdir(parents=True, exist_ok=True)

    document = fitz.open()
    page = new_template_page(document, font_path, company_suffix)
    y = MARGIN_Y

    for raw_line in text.split("\n"):
        first_heading = heading_text(raw_line)
        if first_heading and not title_written:
            y = insert_template_title(page, first_heading, kind, font_path)
            title_written = True
            continue

        image_match = IMAGE_PATTERN.match(raw_line.strip())
        if image_match and source.suffix.lower() != ".docx":
            image_path = Path(image_match.group("path"))
            if not image_path.is_absolute():
                image_path = (source.parent / image_path).resolve()
            if image_path.exists():
                image_doc = fitz.open(image_path)
                image_page = image_doc[0]
                width = PAGE_WIDTH - 2 * MARGIN_X
                height = width * image_page.rect.height / image_page.rect.width
                if y + height > PAGE_HEIGHT - MARGIN_Y:
                    page = new_template_page(document, font_path, company_suffix)
                    y = MARGIN_Y
                rect = fitz.Rect(MARGIN_X, y, MARGIN_X + width, y + height)
                page.insert_image(rect, filename=str(image_path))
                y += height + LINE_HEIGHT
                alt = image_match.group("alt").strip()
                if alt:
                    for caption_line in split_wrapped_lines(alt, max_units=46):
                        if y > PAGE_HEIGHT - MARGIN_Y:
                            page = new_template_page(document, font_path, company_suffix)
                            y = MARGIN_Y
                        insert_text_line(page, caption_line, y, font_path)
                        y += LINE_HEIGHT
                image_doc.close()
                continue

        for line in split_wrapped_lines(normalize_markdown(raw_line), max_units=46):
            if y > PAGE_HEIGHT - MARGIN_Y:
                page = new_template_page(document, font_path, company_suffix)
                y = MARGIN_Y

            if line:
                insert_text_line(page, line, y, font_path)
            y += LINE_HEIGHT

        if raw_line.strip():
            continue
        if y > PAGE_HEIGHT - MARGIN_Y:
            page = new_template_page(document, font_path, company_suffix)
            y = MARGIN_Y
        y += LINE_HEIGHT

    document.save(destination)
    document.close()


def write_pdf(source: Path, destination: Path, font_path: str | None, backend: str) -> str:
    if backend not in {"auto", "word", "libreoffice", "pymupdf"}:
        raise ValueError(f"Unsupported backend: {backend}")

    if backend == "word" and source.suffix.lower() != ".docx":
        raise ValueError("The word backend only supports .docx input.")
    if backend == "libreoffice" and source.suffix.lower() != ".docx":
        raise ValueError("The libreoffice backend only supports .docx input.")

    if source.suffix.lower() != ".docx":
        print(
            "Warning: rendering PDF directly from text/Markdown will not fully match DOCX layout. "
            "Use DOCX input with --backend word for final deliverables.",
            file=sys.stderr,
        )

    if backend in {"auto", "word"} and source.suffix.lower() == ".docx":
        try:
            write_pdf_with_word(source, destination)
            return "word"
        except Exception as exc:
            if backend == "word":
                raise
            print(f"Microsoft Word conversion unavailable, falling back: {exc}", file=sys.stderr)

    if backend in {"auto", "libreoffice"} and source.suffix.lower() == ".docx":
        try:
            write_pdf_with_libreoffice(source, destination)
            return "libreoffice"
        except Exception as exc:
            if backend == "libreoffice":
                raise
            print(f"LibreOffice conversion unavailable, falling back to PyMuPDF: {exc}", file=sys.stderr)

    write_pdf_with_pymupdf(source, destination, font_path)
    return "pymupdf"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Render a DOCX, UTF-8 text, or Markdown-like file to PDF. Use DOCX input with --backend word for best style fidelity."
    )
    parser.add_argument("source", type=Path, help="Source .docx, .txt, or .md file.")
    parser.add_argument("destination", type=Path, help="Destination .pdf file.")
    parser.add_argument(
        "--backend",
        choices=["auto", "word", "libreoffice", "pymupdf"],
        default="auto",
        help="PDF backend. auto tries Word COM, then LibreOffice for DOCX inputs, then falls back to PyMuPDF.",
    )
    parser.add_argument("--font", type=Path, help="Optional font file path, recommended for Chinese text.")
    args = parser.parse_args()

    font_path = str(args.font) if args.font else find_default_font()
    backend = write_pdf(args.source, args.destination, font_path, args.backend)

    print(f"Wrote {args.destination} with {backend}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
