from __future__ import annotations

import sys
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
)

from brands.stulz_legend_runtime import load_calc as load_full_calc



def _plain(value: object) -> str:
    return str(value or "").replace("\u00a0", " ").strip()


def _calc_key(index: int, item: object) -> str:
    model = _plain(getattr(item, "name", ""))
    legend = _plain(getattr(item, "legend", ""))
    return f"{index}|{legend}|{model}"


def _format_qty(value: object) -> str:
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
    except Exception:
        return _plain(value)


def _format_money(value: object) -> str:
    try:
        s = f"{float(value):,.2f}".replace(",", "TMP").replace(".", ",").replace("TMP", "\u00a0")
        return s
    except Exception:
        return _plain(value)


def _ensure_calc_position_controls(self) -> None:
    if getattr(self, "_stulz_calc_position_controls_installed", False):
        return

    preview = getattr(self, "preview", None)
    card = preview.parentWidget() if preview is not None else None
    layout = card.layout() if card is not None else None
    if layout is None:
        return

    group = QGroupBox("Позиции для КП")
    group_layout = QVBoxLayout(group)
    group_layout.setContentsMargins(10, 8, 10, 8)
    group_layout.setSpacing(4)

    table = QTableWidget(0, 5)
    table.setHorizontalHeaderLabels(["Вкл", "Позиция из Calc", "Кол-во", "Цена", "Сумма"])
    table.verticalHeader().setVisible(False)
    table.setAlternatingRowColors(True)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setMinimumHeight(145)
    table.setMaximumHeight(260)
    table.setColumnWidth(0, 52)
    header = table.horizontalHeader()
    header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
    header.setSectionResizeMode(1, QHeaderView.Stretch)
    header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
    header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
    header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
    group_layout.addWidget(table)

    # Keep this block directly after the large calculation preview and before
    # the text-description option group when that group already exists.
    description_group = getattr(self, "_stulz_description_group", None)
    description_index = layout.indexOf(description_group) if description_group is not None else -1
    if description_index >= 0:
        layout.insertWidget(description_index, group)
    else:
        preview_index = layout.indexOf(preview)
        if preview_index >= 0:
            layout.insertWidget(preview_index + 1, group)
        else:
            layout.addWidget(group)

    table.itemChanged.connect(self._on_calc_position_item_changed)

    self._stulz_calc_position_group = group
    self._stulz_calc_position_table = table
    self._stulz_calc_position_enabled = getattr(self, "_stulz_calc_position_enabled", {})
    self._stulz_calc_position_controls_installed = True

    # The calculation block is a primary working area. Give it enough room and
    # keep long text on one line with a horizontal scrollbar if ever necessary.
    self.preview.setMinimumHeight(330)
    self.preview.setLineWrapMode(QTextEdit.NoWrap)


def _refresh_calc_position_controls(self) -> None:
    self._ensure_calc_position_controls()
    table = getattr(self, "_stulz_calc_position_table", None)
    if table is None or getattr(self, "_stulz_calc_position_table_busy", False):
        return

    try:
        context = self._stulz_original_make_context_for_positions()
        if not context.calc_path.exists():
            table.setRowCount(0)
            return
        calc = load_full_calc(context)
        items = list(getattr(calc, "items", []) or [])
    except Exception:
        table.setRowCount(0)
        return

    state = getattr(self, "_stulz_calc_position_enabled", {})
    current_keys: set[str] = set()

    self._stulz_calc_position_table_busy = True
    table.blockSignals(True)
    try:
        table.setRowCount(0)
        for source_index, item in enumerate(items):
            key = _calc_key(source_index, item)
            current_keys.add(key)
            enabled = bool(state.get(key, True))

            row = table.rowCount()
            table.insertRow(row)

            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            check_item.setCheckState(Qt.Checked if enabled else Qt.Unchecked)
            check_item.setData(Qt.UserRole, {"source_index": source_index, "calc_key": key})

            model = _plain(getattr(item, "name", ""))
            legend = _plain(getattr(item, "legend", ""))
            caption = f"{legend} — {model}" if legend else model
            position_item = QTableWidgetItem(caption)
            position_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            position_item.setToolTip(caption)

            qty_item = QTableWidgetItem(_format_qty(getattr(item, "qty", 0)))
            qty_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            price_item = QTableWidgetItem(_format_money(getattr(item, "unit_price", 0)))
            price_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            total_item = QTableWidgetItem(_format_money(getattr(item, "total_price", 0)))
            total_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            total_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            table.setItem(row, 0, check_item)
            table.setItem(row, 1, position_item)
            table.setItem(row, 2, qty_item)
            table.setItem(row, 3, price_item)
            table.setItem(row, 4, total_item)

        # Drop stale state after the user changes workbook/sheet while preserving
        # switches for rows that still have the same Calc identity.
        self._stulz_calc_position_enabled = {
            key: bool(state.get(key, True)) for key in current_keys
        }
        table.resizeRowsToContents()
        wanted_height = min(260, max(145, 40 + table.rowCount() * 30))
        table.setMinimumHeight(wanted_height)
    finally:
        table.blockSignals(False)
        self._stulz_calc_position_table_busy = False


def _on_calc_position_item_changed(self, item: QTableWidgetItem) -> None:
    if getattr(self, "_stulz_calc_position_table_busy", False) or item.column() != 0:
        return

    meta = item.data(Qt.UserRole)
    if not isinstance(meta, dict):
        return
    key = _plain(meta.get("calc_key"))
    if not key:
        return

    state = getattr(self, "_stulz_calc_position_enabled", {})
    state[key] = item.checkState() == Qt.Checked
    self._stulz_calc_position_enabled = state
    self.refresh_preview()


def make_context(self):
    context = self._stulz_original_make_context_for_positions()
    table = getattr(self, "_stulz_calc_position_table", None)
    disabled: list[int] = []
    if table is not None:
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item is None or item.checkState() == Qt.Checked:
                continue
            meta = item.data(Qt.UserRole)
            if not isinstance(meta, dict):
                continue
            try:
                disabled.append(int(meta.get("source_index")))
            except Exception:
                continue

    options = dict(getattr(context, "brand_options", None) or {})
    options["stulz_disabled_item_indexes"] = sorted(set(disabled))
    context.brand_options = options
    return context


def refresh_spec_models(self, context=None) -> None:
    # First let the mature Calc->spec mapper build and preserve its automatic or
    # manual mappings. Then hide rows excluded from the commercial offer.
    self._stulz_original_refresh_spec_models_for_positions(context)

    table = self.spec_models_table
    disabled_indexes: set[int] = set()
    position_table = getattr(self, "_stulz_calc_position_table", None)
    if position_table is not None:
        for row in range(position_table.rowCount()):
            check = position_table.item(row, 0)
            if check is None or check.checkState() == Qt.Checked:
                continue
            meta = check.data(Qt.UserRole)
            if isinstance(meta, dict):
                try:
                    disabled_indexes.add(int(meta.get("source_index")))
                except Exception:
                    pass

    if not disabled_indexes:
        return

    # Calc mapping keys begin with "calc:<source-index>:...".
    for row in range(table.rowCount() - 1, -1, -1):
        item = table.item(row, 0)
        raw = item.data(Qt.UserRole) if item is not None else None
        meta = dict(raw) if isinstance(raw, dict) else {}
        calc_key = _plain(meta.get("calc_key"))
        try:
            source_index = int(calc_key.split(":", 2)[1])
        except Exception:
            continue
        if source_index in disabled_indexes:
            table.removeRow(row)


def refresh_preview(self) -> None:
    # Existing UI logic already handles financing, specification summary and the
    # final total. Because make_context now carries disabled Calc indexes and the
    # final STULZ runtime filters by them, those values are recalculated instantly.
    self._stulz_original_refresh_preview_for_positions()

    # Remove the old textual position list; it is replaced by the interactive
    # full-width table directly below the calculation data.
    text = self.preview.toPlainText()
    marker = "\nПозиции для КП:"
    marker_index = text.find(marker)
    if marker_index >= 0:
        self.preview.setPlainText(text[:marker_index].rstrip())

    self._refresh_calc_position_controls()


def _install() -> None:
    page_module = sys.modules.get("gui.pages.stulz_page")
    page_class = getattr(page_module, "StulzPage", None) if page_module is not None else None
    if page_class is None:
        return

    # Capture the fully patched methods produced by the previous STULZ runtimes.
    if not hasattr(page_class, "_stulz_original_make_context_for_positions"):
        page_class._stulz_original_make_context_for_positions = page_class.make_context
    if not hasattr(page_class, "_stulz_original_refresh_spec_models_for_positions"):
        page_class._stulz_original_refresh_spec_models_for_positions = page_class.refresh_spec_models
    if not hasattr(page_class, "_stulz_original_refresh_preview_for_positions"):
        page_class._stulz_original_refresh_preview_for_positions = page_class.refresh_preview

    for name, function in {
        "_ensure_calc_position_controls": _ensure_calc_position_controls,
        "_refresh_calc_position_controls": _refresh_calc_position_controls,
        "_on_calc_position_item_changed": _on_calc_position_item_changed,
        "make_context": make_context,
        "refresh_spec_models": refresh_spec_models,
        "refresh_preview": refresh_preview,
    }.items():
        setattr(page_class, name, function)

    try:
        from PySide6.QtWidgets import QApplication
        app = QApplication.instance()
        if app is not None:
            for widget in app.allWidgets():
                if isinstance(widget, page_class):
                    widget._ensure_calc_position_controls()
                    widget.refresh_preview()
    except Exception:
        pass


_install()
