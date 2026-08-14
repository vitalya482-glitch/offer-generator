from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

from core.pdf_parsers.stulz_calc_pdf import parse_stulz_calc_totals


@dataclass(frozen=True)
class StulzSpecEntry:
    """One physical STULZ specification set discovered from a Calc PDF.

    Important: entries are intentionally NOT grouped by model. The same model can
    legitimately appear several times in one project (for example for different
    cities or different configurations). The parent directory of Calc.pdf is the
    boundary of one specification set.
    """

    key: str
    model: str
    quantity: float
    calc_pdf: Path
    source_dir: Path
    source_label: str


def _format_model(value: str) -> str:
    return (value or "").replace(" ", "").strip()


def _source_label(root: Path, source_dir: Path) -> str:
    try:
        relative = source_dir.relative_to(root)
        text = str(relative)
        if text and text != ".":
            return text
    except Exception:
        pass
    return source_dir.name or str(source_dir)


def load_stulz_spec_entry(
    calc_pdf: str | Path | None,
    root: str | Path | None = None,
) -> StulzSpecEntry | None:
    """Parse one explicitly selected STULZ Calc PDF into a catalog entry.

    Unlike automatic discovery this function does not require "calc" in the
    filename. It is therefore suitable for the GUI's manual file picker.
    """

    if not calc_pdf:
        return None

    path = Path(calc_pdf)
    if not path.is_file() or path.suffix.lower() != ".pdf":
        return None

    try:
        totals = parse_stulz_calc_totals(path)
    except Exception:
        return None

    model = _format_model(getattr(totals, "model", ""))
    if not model:
        return None

    raw_qty = getattr(totals, "quantity", None)
    try:
        quantity = float(raw_qty) if raw_qty not in (None, "") else 1.0
    except Exception:
        quantity = 1.0
    if quantity <= 0:
        quantity = 1.0

    source_dir = path.parent
    label_root = Path(root) if root else source_dir.parent
    try:
        key = str(path.resolve()).lower()
    except Exception:
        key = str(path).lower()

    return StulzSpecEntry(
        key=key,
        model=model,
        quantity=quantity,
        calc_pdf=path,
        source_dir=source_dir,
        source_label=_source_label(label_root, source_dir),
    )


def discover_stulz_spec_entries(spec_dir: str | Path | None) -> list[StulzSpecEntry]:
    """Discover STULZ specification sets below *spec_dir*.

    Every successfully parsed Calc PDF becomes a separate entry. This deliberately
    preserves duplicate model names instead of summing their quantities.
    """

    if not spec_dir:
        return []

    root = Path(spec_dir)
    if not root.exists():
        return []

    entries: list[StulzSpecEntry] = []
    for calc_pdf in sorted(root.rglob("*.pdf")):
        if not calc_pdf.is_file() or "calc" not in calc_pdf.stem.lower():
            continue

        entry = load_stulz_spec_entry(calc_pdf, root=root)
        if entry is not None:
            entries.append(entry)

    return entries


def _patch_loaded_stulz_page() -> None:
    """Upgrade an already loaded STULZ page without touching MainWindow.

    brands.stulz_runtime imports this catalog while the STULZ page is already
    alive. Replacing methods on the original class also updates that existing
    page instance. The current page is then refreshed once so new controls become
    visible immediately. CLI use is unaffected because the GUI module is not
    loaded there.
    """

    page_module = sys.modules.get("gui.pages.stulz_page")
    if page_module is None:
        return

    try:
        from gui.pages.stulz_page_runtime import StulzPage as RuntimeStulzPage
    except Exception:
        return

    page_class = getattr(page_module, "StulzPage", None)
    if page_class is None or page_class is RuntimeStulzPage:
        return

    for method_name in (
        "_row_metadata",
        "_row_key",
        "current_spec_model_state",
        "selected_spec_models",
        "_scan_calc_pdf_models",
        "_manual_spec_files",
        "_manual_spec_models",
        "_ensure_manual_spec_controls",
        "_update_manual_mode_label",
        "browse_manual_spec_files",
        "clear_manual_spec_files",
        "_short_source_label",
        "refresh_spec_models",
    ):
        setattr(page_class, method_name, getattr(RuntimeStulzPage, method_name))

    # The STULZ runtime module is first imported from refresh_preview(), after the
    # original page object has already been constructed. Refresh existing page
    # instances once so the manual-selection buttons appear without reopening it.
    try:
        from PySide6.QtWidgets import QApplication

        for widget in QApplication.allWidgets():
            if isinstance(widget, page_class):
                widget._ensure_manual_spec_controls()
                widget.refresh_spec_models()
    except Exception:
        pass


_patch_loaded_stulz_page()
