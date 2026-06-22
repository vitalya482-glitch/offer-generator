from __future__ import annotations

import sys
from pathlib import Path

DEFAULT_TEMPLATE_NAME = "HVAC_offer_template_TAGS.docx"


def _candidate_roots() -> list[Path]:
    """Return roots where the bundled template may exist.

    Works both from source tree and from PyInstaller bundle.
    """
    roots: list[Path] = []

    # Source layout: brands/hvac/template_finder.py -> brands/hvac/templates/...
    roots.append(Path(__file__).resolve().parent)

    # PyInstaller --onefile/--onedir layout.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(Path(meipass) / "brands" / "hvac")

    # App root fallback: cwd/brands/hvac/templates/...
    roots.append(Path.cwd() / "brands" / "hvac")

    # Remove duplicates while preserving order.
    unique: list[Path] = []
    for root in roots:
        if root not in unique:
            unique.append(root)
    return unique


def find_default_hvac_template() -> str:
    """Find bundled HVAC DOCX template.

    Returns an absolute path if the template exists. Returns an empty string if not found,
    so the UI can leave the template field blank and let the user choose it manually.
    """
    for root in _candidate_roots():
        path = root / "templates" / DEFAULT_TEMPLATE_NAME
        if path.exists():
            return str(path)
    return ""
