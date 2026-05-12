"""Archive a completed case run's background and output files.

The workflow should ask the user before running this script. By default it
copies files into outputs/archive/<company>. Use --mode move only when the user
explicitly wants to clear the active background/output folders for a new case.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BACKGROUND = ROOT / "supporting_documents" / "background"
OUTPUTS = ROOT / "outputs"
ARCHIVE = OUTPUTS / "archive"


def slug_company(name: str) -> str:
    slug = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().strip(".")
    slug = re.sub(r"\s+", "_", slug)
    if not slug:
        raise ValueError("Company name cannot be empty after sanitization.")
    return slug


def ensure_inside(path: Path, parent: Path) -> None:
    resolved = path.resolve()
    parent_resolved = parent.resolve()
    if parent_resolved != resolved and parent_resolved not in resolved.parents:
        raise ValueError(f"Refusing to operate outside {parent_resolved}: {resolved}")


def should_skip(path: Path, archive_target: Path) -> bool:
    if path.name == ".gitkeep":
        return True
    resolved = path.resolve()
    archive_resolved = archive_target.resolve()
    if archive_resolved == resolved or archive_resolved in resolved.parents:
        return True
    if ARCHIVE.resolve() == resolved or ARCHIVE.resolve() in resolved.parents:
        return True
    return False


def copy_or_move_file(source: Path, destination: Path, mode: str) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if mode == "copy":
        shutil.copy2(source, destination)
    elif mode == "move":
        shutil.move(str(source), str(destination))
    else:
        raise ValueError(f"Unsupported mode: {mode}")


def archive_tree(source_root: Path, destination_root: Path, archive_target: Path, mode: str) -> int:
    if not source_root.exists():
        return 0
    ensure_inside(source_root, ROOT)
    ensure_inside(destination_root, archive_target)
    count = 0
    for source in sorted(path for path in source_root.rglob("*") if path.is_file()):
        if should_skip(source, archive_target):
            continue
        relative = source.relative_to(source_root)
        copy_or_move_file(source, destination_root / relative, mode)
        count += 1
    return count


def main() -> int:
    parser = argparse.ArgumentParser(description="Archive background and output files for a completed case run.")
    parser.add_argument("company", help="Company suffix used for case files, e.g. ExampleCo.")
    parser.add_argument("--mode", choices=["copy", "move"], default="copy", help="Copy by default; move only after user approval.")
    args = parser.parse_args()

    company = slug_company(args.company)
    target = ARCHIVE / company
    ensure_inside(target, ARCHIVE)
    target.mkdir(parents=True, exist_ok=True)

    background_count = archive_tree(BACKGROUND, target / "background", target, args.mode)
    output_count = archive_tree(OUTPUTS, target / "outputs", target, args.mode)

    print(f"Archived company: {company}")
    print(f"Mode: {args.mode}")
    print(f"Background files: {background_count}")
    print(f"Output files: {output_count}")
    print(f"Archive path: {target}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
