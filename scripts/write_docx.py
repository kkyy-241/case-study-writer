"""Render a UTF-8 text or Markdown-like draft to DOCX."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


DEFAULT_FONT = "Microsoft YaHei"
HEADING_FONT = "Microsoft YaHei"
IMAGE_PATTERN = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<path>[^)]+)\)$")


def clean_inline_markdown(text: str) -> str:
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.strip()


def parse_line(line: str) -> tuple[str, str]:
    stripped = line.strip()
    heading = re.match(r"^(#{1,3})\s+(.+)$", stripped)
    if heading:
        level = str(min(len(heading.group(1)), 3))
        return f"Heading {level}", clean_inline_markdown(heading.group(2))
    return "Normal", clean_inline_markdown(stripped)


def set_east_asia_font(style, font_name: str) -> None:
    from docx.oxml.ns import qn

    style.font.name = font_name
    r_pr = style._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:eastAsia"), font_name)


def configure_document(document) -> None:
    from docx.shared import Cm, Pt

    section = document.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    styles = document.styles
    normal = styles["Normal"]
    set_east_asia_font(normal, DEFAULT_FONT)
    normal.font.size = Pt(11)
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(6)

    for level, size in ((1, 16), (2, 14), (3, 12)):
        style = styles[f"Heading {level}"]
        set_east_asia_font(style, HEADING_FONT)
        style.font.size = Pt(size)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)


def write_docx(source: Path, destination: Path) -> None:
    try:
        from docx import Document
        from docx.shared import Cm
    except ImportError as exc:
        raise RuntimeError("python-docx is required. Run scripts/check_dependencies.py.") from exc

    text = source.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\r", "\n")
    document = Document()
    configure_document(document)

    for raw_line in text.split("\n"):
        if not raw_line.strip():
            continue

        image_match = IMAGE_PATTERN.match(raw_line.strip())
        if image_match:
            image_path = Path(image_match.group("path"))
            if not image_path.is_absolute():
                image_path = (source.parent / image_path).resolve()
            document.add_picture(str(image_path), width=Cm(14.5))
            alt_text = clean_inline_markdown(image_match.group("alt"))
            if alt_text:
                caption = document.add_paragraph(style="Normal")
                caption.add_run(alt_text)
            continue

        style, content = parse_line(raw_line)
        if not content:
            continue

        paragraph = document.add_paragraph(style=style)
        paragraph.add_run(content)

    destination.parent.mkdir(parents=True, exist_ok=True)
    document.save(destination)


def main() -> int:
    parser = argparse.ArgumentParser(description="Render a UTF-8 text or Markdown-like draft to DOCX.")
    parser.add_argument("source", type=Path, help="Source .txt or .md file.")
    parser.add_argument("destination", type=Path, help="Destination .docx file.")
    args = parser.parse_args()

    write_docx(args.source, args.destination)
    print(f"Wrote {args.destination}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
