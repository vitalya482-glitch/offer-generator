from __future__ import annotations

import sys
from pathlib import Path


DEFAULT_TEMPLATE_NAME = "Offer_Company_22-05-26_TAGGED_HVAC.docx"
DEFAULT_TEMPLATE_RELATIVE_PATH = Path("templates") / "HVAC" / DEFAULT_TEMPLATE_NAME


def _candidate_app_roots() -> list[Path]:
    """Return possible application roots for source and PyInstaller builds."""
    roots: list[Path] = [Path(__file__).resolve().parents[2]]

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(Path(meipass))

    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).resolve().parent)

    roots.append(Path.cwd())

    unique: list[Path] = []
    for root in roots:
        try:
            resolved = root.resolve()
        except OSError:
            resolved = root
        if resolved not in unique:
            unique.append(resolved)
    return unique


def find_default_hvac_template() -> str:
    """Return the standard HVAC template path or an empty string."""
    for root in _candidate_app_roots():
        template_path = root / DEFAULT_TEMPLATE_RELATIVE_PATH
        if template_path.is_file():
            return str(template_path)
    return ""
