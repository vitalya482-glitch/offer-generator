from __future__ import annotations

import sys
from pathlib import Path

DEFAULT_TEMPLATE_NAME = "Offer_Company_22-05-26_TAGGED_HVAC.docx"
DEFAULT_TEMPLATE_RELATIVE_PATH = Path("templates") / "HVAC" / DEFAULT_TEMPLATE_NAME


def _candidate_app_roots() -> list[Path]:
    """Return possible application roots for source and PyInstaller builds."""
    roots: list[Path] = []

    # Source tree:
    # <project>/brands/hvac/template_finder.py -> <project>
    roots.append(Path(__file__).resolve().parents[2])

    # PyInstaller extraction/bundle directory.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(Path(meipass))

    # Portable onedir build: templates may be next to the executable.
    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).resolve().parent)

    # Useful when the program is launched from the project/application folder.
    roots.append(Path.cwd())

    unique: list[Path] = []
    for root in roots:
        resolved = root.resolve()
        if resolved not in unique:
            unique.append(resolved)
    return unique


def find_default_hvac_template() -> str:
    """Find the standard HVAC DOCX template.

    Returns an absolute path when found. If the template is missing, returns an
    empty string so the HVAC page leaves the field blank for manual selection.
    """
    for root in _candidate_app_roots():
        template_path = root / DEFAULT_TEMPLATE_RELATIVE_PATH
        if template_path.is_file():
            return str(template_path)

    return ""
