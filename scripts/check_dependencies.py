"""Check Python dependencies required by the repository."""

from __future__ import annotations

import importlib.util
import platform
import sys


REQUIRED_MODULES = {
    "docx": "python-docx",
    "fitz": "PyMuPDF",
    "yaml": "PyYAML",
}

OPTIONAL_MODULES = {
    "win32com.client": ("pywin32", "Microsoft Word COM PDF export on Windows"),
}


def module_available(module_name: str) -> bool:
    try:
        return importlib.util.find_spec(module_name) is not None
    except ModuleNotFoundError:
        return False


def main() -> int:
    missing: list[str] = []

    print(f"Python executable: {sys.executable}")
    print(f"Python version: {sys.version.split()[0]}\n")

    for module_name, package_name in REQUIRED_MODULES.items():
        if not module_available(module_name):
            missing.append(f"{package_name} (import name: {module_name})")
        else:
            print(f"OK: {package_name} is available as {module_name}")

    if platform.system() == "Windows":
        for module_name, (package_name, purpose) in OPTIONAL_MODULES.items():
            if not module_available(module_name):
                print(f"Optional missing: {package_name} (import name: {module_name}) for {purpose}")
            else:
                print(f"OK: optional {package_name} is available as {module_name}")

    if missing:
        print("\nMissing required dependencies:")
        for item in missing:
            print(f"- {item}")
        print("\nInstall them with: pip install -r requirements.txt")
        return 1

    print("\nAll required dependencies are available.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
