from __future__ import annotations

import sys

from PySide6.QtCore import QEvent, QObject, QTimer, Qt
from PySide6.QtWidgets import QApplication, QSizePolicy, QTableWidget, QTextEdit


class _PreviewResizeFilter(QObject):
    """Keep the read-only calculation preview tall enough after width changes."""

    def eventFilter(self, watched, event):  # noqa: N802 - Qt API name
        if event.type() in (QEvent.Resize, QEvent.Show, QEvent.FontChange):
            QTimer.singleShot(0, lambda obj=watched: _fit_text_edit(obj))
        return False


def _fit_text_edit(widget: QTextEdit) -> None:
    """Show the whole document and let the outer page own vertical scrolling."""

    try:
        widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setLineWrapMode(QTextEdit.WidgetWidth)
        widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # documentSize() reflects wrapping for the current viewport width.
        document = widget.document()
        document.adjustSize()
        document_height = float(document.documentLayout().documentSize().height())
        margins = widget.contentsMargins()
        extra = (
            widget.frameWidth() * 2
            + margins.top()
            + margins.bottom()
            + 12
        )
        wanted = max(90, int(document_height + extra + 0.999))
        if abs(widget.height() - wanted) > 2:
            widget.setFixedHeight(wanted)
    except Exception:
        pass


def _fit_table(table: QTableWidget, minimum: int = 72) -> None:
    """Expand a main-page table to every row; page scrollbar handles overflow."""

    try:
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        table.resizeRowsToContents()

        height = table.horizontalHeader().height()
        for row in range(table.rowCount()):
            height += table.rowHeight(row)
        height += table.frameWidth() * 2 + 6

        wanted = max(minimum, height)
        table.setMinimumHeight(wanted)
        table.setMaximumHeight(wanted)
    except Exception:
        pass


def _ensure_full_content_layout(self) -> None:
    if getattr(self, "_stulz_full_content_layout_installed", False):
        return

    preview = getattr(self, "preview", None)
    if isinstance(preview, QTextEdit):
        # Remove limits introduced by earlier compact-layout versions.
        preview.setMinimumHeight(0)
        preview.setMaximumHeight(16_777_215)
        preview.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setLineWrapMode(QTextEdit.WidgetWidth)

        resize_filter = _PreviewResizeFilter(preview)
        preview.installEventFilter(resize_filter)
        preview.document().contentsChanged.connect(lambda: QTimer.singleShot(0, lambda: _fit_text_edit(preview)))
        self._stulz_preview_resize_filter = resize_filter

    position_table = getattr(self, "_stulz_calc_position_table", None)
    if isinstance(position_table, QTableWidget):
        position_table.setMaximumHeight(16_777_215)
        position_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

    spec_table = getattr(self, "spec_models_table", None)
    if isinstance(spec_table, QTableWidget):
        spec_table.setMaximumHeight(16_777_215)
        spec_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

    self._stulz_full_content_layout_installed = True
    QTimer.singleShot(0, lambda: self._fit_all_main_content())


def _fit_all_main_content(self) -> None:
    preview = getattr(self, "preview", None)
    if isinstance(preview, QTextEdit):
        _fit_text_edit(preview)

    position_table = getattr(self, "_stulz_calc_position_table", None)
    if isinstance(position_table, QTableWidget):
        _fit_table(position_table, minimum=72)

    spec_table = getattr(self, "spec_models_table", None)
    if isinstance(spec_table, QTableWidget):
        _fit_table(spec_table, minimum=90)


def refresh_preview(self) -> None:
    self._stulz_original_refresh_preview_for_full_content()
    self._ensure_full_content_layout()
    QTimer.singleShot(0, lambda: self._fit_all_main_content())


def refresh_spec_models(self, context=None) -> None:
    self._stulz_original_refresh_spec_models_for_full_content(context)
    self._ensure_full_content_layout()
    QTimer.singleShot(0, lambda: self._fit_all_main_content())


def _refresh_calc_position_controls(self) -> None:
    self._stulz_original_refresh_calc_positions_for_full_content()
    self._ensure_full_content_layout()
    QTimer.singleShot(0, lambda: self._fit_all_main_content())


def _install() -> None:
    page_module = sys.modules.get("gui.pages.stulz_page")
    page_class = getattr(page_module, "StulzPage", None) if page_module is not None else None
    if page_class is None:
        return

    if not hasattr(page_class, "_stulz_original_refresh_preview_for_full_content"):
        page_class._stulz_original_refresh_preview_for_full_content = page_class.refresh_preview
    if not hasattr(page_class, "_stulz_original_refresh_spec_models_for_full_content"):
        page_class._stulz_original_refresh_spec_models_for_full_content = page_class.refresh_spec_models
    if not hasattr(page_class, "_stulz_original_refresh_calc_positions_for_full_content"):
        page_class._stulz_original_refresh_calc_positions_for_full_content = page_class._refresh_calc_position_controls

    for name, function in {
        "_ensure_full_content_layout": _ensure_full_content_layout,
        "_fit_all_main_content": _fit_all_main_content,
        "refresh_preview": refresh_preview,
        "refresh_spec_models": refresh_spec_models,
        "_refresh_calc_position_controls": _refresh_calc_position_controls,
    }.items():
        setattr(page_class, name, function)

    app = QApplication.instance()
    if app is None:
        return

    for widget in app.allWidgets():
        if isinstance(widget, page_class):
            widget._ensure_full_content_layout()
            QTimer.singleShot(0, widget._fit_all_main_content)


_install()
