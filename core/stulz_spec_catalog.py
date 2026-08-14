from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import datetime
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


def _restore_common_spec_dir(widget: object) -> None:
    """Repair old saved state that points to one physical specification folder.

    Earlier STULZ builds could remember a single leaf folder such as the Almaty
    ASR552AS folder. With physical-spec matching enabled that would hide sibling
    Astana/Karaganda Calc PDFs. If the project-level inferred folder contains
    more valid Calc PDFs than the saved leaf, automatically move the GUI back to
    that common parent. Explicit manual PDF selection is handled separately and
    is never changed here.
    """

    try:
        from gui.path_helpers import infer_specifications_dir

        project_text = str(widget.project_path_text() or "").strip()  # type: ignore[attr-defined]
        current_text = str(widget.spec_path_text() or "").strip()  # type: ignore[attr-defined]
        if not project_text:
            return

        inferred_text = str(infer_specifications_dir(project_text) or "").strip()
        if not inferred_text or inferred_text == current_text:
            return

        inferred = Path(inferred_text)
        if not inferred.exists():
            return

        current_entries = discover_stulz_spec_entries(current_text) if current_text else []
        inferred_entries = discover_stulz_spec_entries(inferred)
        if len(inferred_entries) <= len(current_entries):
            return

        # Only auto-expand a saved folder when it is actually inside the inferred
        # project specification root. This preserves unrelated user-selected dirs.
        if current_text:
            try:
                Path(current_text).resolve().relative_to(inferred.resolve())
            except Exception:
                return

        widget._set_spec_dir_path(str(inferred))  # type: ignore[attr-defined]
        try:
            widget.settings.setValue("spec_dir", str(inferred))  # type: ignore[attr-defined]
            widget.settings.sync()  # type: ignore[attr-defined]
        except Exception:
            pass
    except Exception:
        return


def _install_data_refresh_controls(widget: object) -> None:
    """Add explicit disk refresh controls to the STULZ preview card.

    The refresh button and F5 intentionally re-read the already selected Excel
    workbook and specification PDFs without forcing the user to reselect the
    project. This is useful while the calculation is being edited in parallel.
    """

    if getattr(widget, "_stulz_data_refresh_installed", False):
        return

    try:
        from PySide6.QtCore import Qt
        from PySide6.QtGui import QKeySequence, QShortcut
        from PySide6.QtWidgets import QApplication, QHBoxLayout, QLabel, QPushButton

        preview = getattr(widget, "preview", None)
        if preview is None:
            return

        card = preview.parentWidget()
        layout = card.layout() if card is not None else None
        if layout is None:
            return

        title_item = layout.takeAt(0)
        title_widget = title_item.widget() if title_item is not None else None

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(10)
        if title_widget is not None:
            title_row.addWidget(title_widget)
        title_row.addStretch(1)

        refreshed_label = QLabel()
        refreshed_label.setObjectName("Hint")
        title_row.addWidget(refreshed_label)

        refresh_button = QPushButton("↻ Обновить")
        refresh_button.setObjectName("GhostButton")
        refresh_button.setToolTip("Перечитать Excel и спецификации с диска (F5)")
        refresh_button.setMinimumWidth(120)
        title_row.addWidget(refresh_button)

        layout.insertLayout(0, title_row)

        def set_timestamp() -> str:
            stamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            refreshed_label.setText(f"Обновлено: {stamp}")
            return stamp

        def refresh_from_disk() -> None:
            if getattr(widget, "_stulz_data_refresh_busy", False):
                return

            setattr(widget, "_stulz_data_refresh_busy", True)
            refresh_button.setEnabled(False)
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                # Re-read the current workbook from disk while keeping the selected
                # sheet when it still exists. refresh_preview() also re-scans the
                # physical Calc.pdf specification rows through the runtime patch.
                widget.load_sheets(refresh=False)  # type: ignore[attr-defined]
                widget.refresh_preview()  # type: ignore[attr-defined]
                set_timestamp()
            finally:
                QApplication.restoreOverrideCursor()
                refresh_button.setEnabled(True)
                setattr(widget, "_stulz_data_refresh_busy", False)

        refresh_button.clicked.connect(refresh_from_disk)

        refresh_shortcut = QShortcut(QKeySequence("F5"), widget)
        refresh_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        refresh_shortcut.activated.connect(refresh_from_disk)

        # Keep Python/Qt references alive for the lifetime of the STULZ page.
        setattr(widget, "_stulz_data_refresh_button", refresh_button)
        setattr(widget, "_stulz_data_refresh_label", refreshed_label)
        setattr(widget, "_stulz_data_refresh_shortcut", refresh_shortcut)
        setattr(widget, "_stulz_data_refresh_callback", refresh_from_disk)
        setattr(widget, "_stulz_data_refresh_installed", True)

        # The preview has already been loaded once by the page constructor when
        # these runtime controls are injected, so show that initial read time too.
        set_timestamp()
    except Exception:
        return


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
    # instances once so runtime controls appear without reopening the application.
    try:
        from PySide6.QtWidgets import QApplication

        for widget in QApplication.allWidgets():
            if isinstance(widget, page_class):
                _restore_common_spec_dir(widget)
                widget._ensure_manual_spec_controls()
                widget.refresh_spec_models()
                _install_data_refresh_controls(widget)
    except Exception:
        pass


_patch_loaded_stulz_page()
