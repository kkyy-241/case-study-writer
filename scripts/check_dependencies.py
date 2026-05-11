"""Check Python dependencies required by the repository."""

from __future__ import annotations

import importlib.util
import sys


REQUIRED_MODULES = {
    "fitz": "PyMuPDF",
    "yaml": "PyYAML",
}


def main() -> int:
    missing: list[str] = []

    print(f"Python executable: {sys.executable}")
    print(f"Python version: {sys.version.split()[0]}\n")

    for module_name, package_name in REQUIRED_MODULES.items():
        if importlib.util.find_spec(module_name) is None:
            missing.append(f"{package_name} (import name: {module_name})")
        else:
            print(f"OK: {package_name} is available as {module_name}")

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
