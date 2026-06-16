from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from brands.registry import get_brand_module
from core.manager_profile import find_manager_in_project
from core.models import OfferContext
from core.riello_price import (
    RielloPriceItem,
    default_price_path,
    format_price,
    item_backup_label,
    item_display_option,
    item_display_with_price,
    item_power_label,
    load_price_items,
    nearest_power_items,
    option_items,
)
from core.runtime_paths import resource_path
from gui.path_helpers import extract_client_from_project_dir, infer_output_dir


class RielloPage(QWidget):
    """Самостоятельная страница Riello.

    Правая часть приложения полностью принадлежит этой странице: папки, шаблон,
    подбор ИБП, быстрые опции, лист подбора оборудования и генерация Excel.
    """

    brand_name = "Riello"

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self._updating = False
        self._updating_path_display = False
        self.price_items: list[RielloPriceItem] = []
        self.filtered_items: list[RielloPriceItem] = []
        self.option_price_items: list[RielloPriceItem] = option_items()
        self.quote_rows: list[dict[str, Any]] = self._load_quote_rows()

        self.project_dir_path = self._saved("riello/project_dir", self._saved("project_dir", ""))
        self.output_dir_path = self._saved("riello/output_dir", self._saved("output_dir", ""))

        self.project_edit = QLineEdit(owner._display_dir(self.project_dir_path))
        self.project_edit.setToolTip(self.project_dir_path)
        self.client_edit = QLineEdit(self._saved("riello/client", self._saved("client", "ТОО Example")))
        self.calc_combo = QComboBox()
        self.calc_combo.setEditable(True)
        saved_calc = self._saved("riello/calc_path", "") or self._saved("calc_template_path", "")
        if saved_calc:
            owner._add_path_item(self.calc_combo, saved_calc, is_file=True)
            self.calc_combo.setCurrentIndex(0)
        self.output_edit = QLineEdit(owner._display_dir(self.output_dir_path))
        self.output_edit.setToolTip(self.output_dir_path)
        self.status_label = QLabel("Выберите папку проекта и добавьте оборудование в лист подбора")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        header = QHBoxLayout()
        h_text = QVBoxLayout()
        page_title = QLabel("Riello: расчет ИБП")
        page_title.setObjectName("PageTitle")
        page_subtitle = QLabel("Введите мощность, выберите модель из PDF-прайса, добавьте ИБП и опции в лист подбора оборудования.")
        page_subtitle.setObjectName("PageSubtitle")
        h_text.addWidget(page_title)
        h_text.addWidget(page_subtitle)
        self.generate_btn = QPushButton("Сформировать Excel")
        self.generate_btn.setObjectName("PrimaryButton")
        self.generate_btn.clicked.connect(self.generate)
        header.addLayout(h_text, stretch=1)
        header.addWidget(self.generate_btn, stretch=0, alignment=Qt.AlignTop)
        layout.addLayout(header)

        project_card = owner._card("Папка проекта")
        project_grid = QGridLayout()
        project_card.layout().addLayout(project_grid)
        project_grid.setColumnStretch(1, 1)
        project_grid.setVerticalSpacing(12)
        project_grid.setHorizontalSpacing(10)
        owner._add_row(project_grid, 0, "Папка проекта", self.project_edit, "Выбрать", self.browse_project_dir)
        layout.addWidget(project_card)

        files_card = owner._card("Riello: Excel-шаблон и результат")
        files_grid = QGridLayout()
        files_card.layout().addLayout(files_grid)
        files_grid.setColumnStretch(1, 1)
        files_grid.setVerticalSpacing(12)
        files_grid.setHorizontalSpacing(10)
        owner._add_row(files_grid, 0, "Клиент", self.client_edit, None, None)
        owner._add_row(files_grid, 1, "Excel-шаблон", self.calc_combo, "Выбрать", self.browse_calc_file)
        owner._add_row(files_grid, 2, "Папка результата", self.output_edit, "Выбрать", self.browse_output_dir)
        layout.addWidget(files_card)

        input_card = owner._card("1. Ввод данных")
        input_grid = QGridLayout()
        input_card.layout().addLayout(input_grid)
        input_grid.setColumnStretch(1, 1)
        input_grid.setColumnStretch(3, 0)
        input_grid.setColumnStretch(5, 0)
        input_grid.setVerticalSpacing(12)
        input_grid.setHorizontalSpacing(10)

        self.required_power_edit = QLineEdit(self._saved("riello/required_power_kw", "20"))
        self.required_power_edit.setPlaceholderText("например 20, 100 или 4000")
        self.ups_qty_edit = QLineEdit(self._saved("riello/ups_quantity", "1"))
        self.ups_qty_edit.setPlaceholderText("шт")
        self.ups_qty_edit.setMaximumWidth(95)
        self.ups_combo = QComboBox()
        self.ups_combo.setMinimumWidth(480)
        self.add_ups_btn = QPushButton("Включить в расчёт")
        self.add_ups_btn.clicked.connect(self.add_selected_ups_to_quote)

        self.option_combo = QComboBox()
        self.option_combo.setMinimumWidth(360)
        self.option_qty_edit = QLineEdit(self._saved("riello/option_quantity", "1"))
        self.option_qty_edit.setPlaceholderText("шт")
        self.option_qty_edit.setMaximumWidth(95)
        self.add_option_btn = QPushButton("Добавить опцию")
        self.add_option_btn.clicked.connect(self.add_selected_option_to_quote)

        input_grid.addWidget(QLabel("Мощность, кВА/кВт"), 0, 0)
        input_grid.addWidget(self.required_power_edit, 0, 1)
        input_grid.addWidget(QLabel("Кол-во ИБП"), 0, 2)
        input_grid.addWidget(self.ups_qty_edit, 0, 3)
        input_grid.addWidget(QLabel("Модель ИБП"), 1, 0)
        input_grid.addWidget(self.ups_combo, 1, 1, 1, 4)
        input_grid.addWidget(self.add_ups_btn, 1, 5)
        input_grid.addWidget(QLabel("Опция"), 2, 0)
        input_grid.addWidget(self.option_combo, 2, 1, 1, 3)
        input_grid.addWidget(QLabel("Кол-во"), 2, 4)
        input_grid.addWidget(self.option_qty_edit, 2, 5)
        input_grid.addWidget(self.add_option_btn, 2, 6)

        self.match_hint = QLabel("")
        self.match_hint.setObjectName("Hint")
        self.match_hint.setWordWrap(True)
        input_grid.addWidget(self.match_hint, 3, 1, 1, 6)
        layout.addWidget(input_card)

        details_card = owner._card("Карточка выбранной модели")
        self.details = QLabel("")
        self.details.setObjectName("Hint")
        self.details.setWordWrap(True)
        details_card.layout().addWidget(self.details)
        details_card.layout().addWidget(self.status_label)
        layout.addWidget(details_card)

        quote_card = owner._card("Лист подбора оборудования")
        quote_layout = QVBoxLayout()
        quote_card.layout().addLayout(quote_layout)
        self.quote_table = QTableWidget(0, 6)
        self.quote_table.setHorizontalHeaderLabels(["Тип", "Модель", "Код", "Кол-во", "Цена", "Сумма"])
        self.quote_table.verticalHeader().setVisible(False)
        self.quote_table.setAlternatingRowColors(True)
        self.quote_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.quote_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.quote_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.quote_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.quote_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.quote_table.setMinimumHeight(150)
        quote_layout.addWidget(self.quote_table)
        quote_buttons = QHBoxLayout()
        self.remove_quote_btn = QPushButton("Удалить выбранное")
        self.remove_quote_btn.clicked.connect(self.remove_selected_quote_rows)
        self.clear_quote_btn = QPushButton("Очистить лист")
        self.clear_quote_btn.clicked.connect(self.clear_quote_rows)
        quote_buttons.addStretch(1)
        quote_buttons.addWidget(self.remove_quote_btn)
        quote_buttons.addWidget(self.clear_quote_btn)
        quote_layout.addLayout(quote_buttons)
        layout.addWidget(quote_card)

        table_card = owner._card("Подходящие позиции из PDF-прайса")
        self.models_table = QTableWidget(0, 7)
        self.models_table.setHorizontalHeaderLabels(["Модель", "Код", "Мощность", "Стоимость", "Автономия", "Габариты", "Вес"])
        self.models_table.verticalHeader().setVisible(False)
        self.models_table.setAlternatingRowColors(True)
        self.models_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.models_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.models_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.models_table.setMinimumHeight(170)
        self.models_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.models_table.setEditTriggers(QTableWidget.NoEditTriggers)
        table_card.layout().addWidget(self.models_table)
        layout.addWidget(table_card)
        layout.addStretch(1)

        self._load_price_items()
        self._load_option_items()
        self.ensure_default_excel_template()
        self._connect_changes()
        self.refresh_summary()

    # ------------------------- helpers/settings -------------------------
    def _saved(self, key: str, default: str) -> str:
        settings = getattr(self.owner, "settings", None)
        if settings is None:
            return default
        value = settings.value(key, default)
        return str(value) if value is not None else default

    def _set_saved(self, key: str, value: str) -> None:
        settings = getattr(self.owner, "settings", None)
        if settings is None:
            return
        settings.setValue(key, value)

    def _path_from_combo(self, combo: QComboBox) -> str:
        return self.owner._path_from_combo(combo)

    def _set_line_path(self, line_edit: QLineEdit, path_text: str, is_file: bool = False) -> None:
        self._updating_path_display = True
        self.owner._set_line_path(line_edit, path_text, is_file=is_file)
        self._updating_path_display = False

    def project_path_text(self) -> str:
        return self.project_dir_path or self.project_edit.toolTip() or self.project_edit.text().strip()

    def output_path_text(self) -> str:
        return self.output_dir_path or self.output_edit.toolTip() or self.output_edit.text().strip()

    def ensure_default_excel_template(self) -> None:
        template_path = resource_path(Path("templates") / "riello" / "calc_08-04-26 UPS.xlsx")
        if not template_path.exists():
            return
        combo = self.calc_combo
        if self.owner._find_combo_path(combo, str(template_path)) < 0:
            self.owner._add_path_item(combo, str(template_path), is_file=True)
        if not self.owner._path_from_combo(combo).strip():
            index = self.owner._find_combo_path(combo, str(template_path))
            if index >= 0:
                combo.setCurrentIndex(index)

    def _connect_changes(self) -> None:
        self.project_edit.textChanged.connect(self._on_project_dir_changed)
        self.client_edit.textChanged.connect(self._on_changed)
        self.calc_combo.currentTextChanged.connect(self._on_changed)
        self.output_edit.textChanged.connect(self._on_output_dir_changed)
        self.required_power_edit.textChanged.connect(self._on_power_changed)
        self.ups_combo.currentIndexChanged.connect(self._on_ups_changed)
        self.ups_qty_edit.textChanged.connect(self._on_changed)
        self.option_combo.currentIndexChanged.connect(self._on_option_changed)
        self.option_qty_edit.textChanged.connect(self._on_changed)
        self.models_table.cellClicked.connect(self._on_table_row_clicked)

    def _load_price_items(self) -> None:
        self._updating = True
        try:
            self.price_items = load_price_items(default_price_path())
            self._reload_ups_models(keep_saved=True)
        finally:
            self._updating = False

    def _load_option_items(self) -> None:
        self.option_combo.blockSignals(True)
        try:
            self.option_combo.clear()
            for item in self.option_price_items:
                self.option_combo.addItem(item_display_option(item), item.code)
        finally:
            self.option_combo.blockSignals(False)

    def _load_quote_rows(self) -> list[dict[str, Any]]:
        raw = self._saved("riello/quote_rows_json", "")
        if not raw:
            return []
        try:
            rows = json.loads(raw)
        except Exception:
            return []
        if not isinstance(rows, list):
            return []
        clean_rows: list[dict[str, Any]] = []
        for row in rows:
            if isinstance(row, dict) and str(row.get("model") or "").strip():
                clean_rows.append(row)
        return clean_rows

    # ------------------------- browse/project -------------------------
    def browse_project_dir(self) -> None:
        old_project = self.project_path_text().strip()
        path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта Riello", old_project)
        if not path:
            return
        self.project_dir_path = path
        self._set_line_path(self.project_edit, path, is_file=False)
        if path != old_project:
            self.output_dir_path = ""
            self._set_line_path(self.output_edit, "", is_file=False)
        self.apply_project_dir(Path(path))
        self.save_options()

    def apply_project_dir(self, project_dir: Path) -> None:
        if not self.output_path_text().strip():
            try:
                guessed = infer_output_dir(str(project_dir))
            except Exception:
                guessed = str(project_dir)
            self.output_dir_path = guessed
            self._set_line_path(self.output_edit, guessed, is_file=False)
        client = extract_client_from_project_dir(str(project_dir))
        if client and (not self.client_edit.text().strip() or self.client_edit.text().strip() == "ТОО Example"):
            self.client_edit.setText(client)
        self.autofill_manager_from_project(project_dir, force=False)
        self.refresh_summary()

    def browse_calc_file(self) -> None:
        current_calc = self._path_from_combo(self.calc_combo)
        start_dir = str(Path(current_calc).parent) if current_calc else self.project_path_text()
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-шаблон Riello", start_dir, "Excel (*.xlsx *.xlsm)")
        if path:
            index = self.owner._find_combo_path(self.calc_combo, path)
            if index < 0:
                self.owner._add_path_item(self.calc_combo, path, is_file=True)
                index = self.calc_combo.count() - 1
            self.calc_combo.setCurrentIndex(index)
            self._set_saved("riello/calc_path", path)
            self.save_options()

    def browse_output_dir(self) -> None:
        start_dir = self.output_path_text() or self.project_path_text()
        path = QFileDialog.getExistingDirectory(self, "Выберите папку результата Riello", start_dir)
        if path:
            self.output_dir_path = path
            self._set_line_path(self.output_edit, path, is_file=False)
            self._set_saved("riello/output_dir", path)
            self.save_options()

    def autofill_manager_from_project(self, project_dir: Path, force: bool = False) -> None:
        if not force and self.owner._has_saved_manager_profile():
            return
        if not force and not self.owner._manager_profile().is_empty():
            return
        if not project_dir.exists():
            return
        profile = find_manager_in_project(project_dir)
        if profile.is_empty():
            return
        self.owner._set_manager_profile(profile)
        self.owner.settings.setValue("manager_name", profile.name)
        self.owner.settings.setValue("manager_position", profile.position)
        self.owner.settings.setValue("manager_email", profile.email)
        self.owner.settings.setValue("manager_phone", profile.phone)
        self.owner.settings.sync()
        self.status_label.setText("Данные исполнителя найдены в Word-файле проекта")

    def _on_project_dir_changed(self) -> None:
        if not self._updating_path_display:
            self.project_dir_path = self.project_edit.text().strip()
            self.project_edit.setToolTip(self.project_dir_path)
        self._on_changed()

    def _on_output_dir_changed(self) -> None:
        if not self._updating_path_display:
            self.output_dir_path = self.output_edit.text().strip()
            self.output_edit.setToolTip(self.output_dir_path)
        self._on_changed()

    # ------------------------- Riello picker -------------------------
    def _on_power_changed(self) -> None:
        if self._updating:
            return
        self._updating = True
        try:
            self._reload_ups_models(keep_saved=False)
        finally:
            self._updating = False
        self._on_changed()

    def _on_ups_changed(self) -> None:
        if self._updating:
            return
        self._on_changed()

    def _on_option_changed(self) -> None:
        if self._updating:
            return
        self._on_changed()

    def _on_table_row_clicked(self, row: int, _column: int) -> None:
        if row < 0 or row >= len(self.filtered_items):
            return
        item = self.filtered_items[row]
        index = self.ups_combo.findData(item.model)
        if index >= 0:
            self.ups_combo.setCurrentIndex(index)

    def _on_changed(self) -> None:
        if self._updating:
            return
        self.save_options()
        self.refresh_summary()

    def _reload_ups_models(self, keep_saved: bool = True) -> None:
        required_kw = self._as_float(self.required_power_edit.text(), 0.0)
        saved_ups = self._saved("riello/ups_model", "") if keep_saved else ""
        current_ups = str(self.ups_combo.currentData() or "")
        previous_ups = saved_ups or current_ups

        candidates = nearest_power_items(getattr(self, "price_items", []), required_kw)
        if not candidates:
            candidates = getattr(self, "price_items", [])[:]
        self.filtered_items = candidates

        self.ups_combo.blockSignals(True)
        try:
            self.ups_combo.clear()
            for item in candidates:
                self.ups_combo.addItem(item_display_with_price(item), item.model)
            index = self.ups_combo.findData(previous_ups)
            if index < 0 and self.ups_combo.count() > 0:
                index = 0
            if index >= 0:
                self.ups_combo.setCurrentIndex(index)
        finally:
            self.ups_combo.blockSignals(False)

    def _as_float(self, value: str, default: float = 0.0) -> float:
        try:
            return float(str(value or "").replace(" ", "").replace(",", "."))
        except Exception:
            return default

    def _fmt_qty(self, value: float | str) -> str:
        try:
            number = float(value)
            return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
        except Exception:
            return str(value)

    def _item_by_model(self, model: str) -> RielloPriceItem | None:
        model_upper = (model or "").upper()
        for item in getattr(self, "price_items", []):
            if item.model.upper() == model_upper:
                return item
        return None

    def _option_by_code(self, code: str) -> RielloPriceItem | None:
        code_upper = (code or "").upper()
        for item in self.option_price_items:
            if item.code.upper() == code_upper or item.model.upper() == code_upper:
                return item
        return None

    def _selected_item(self) -> RielloPriceItem | None:
        return self._item_by_model(str(self.ups_combo.currentData() or ""))

    def _selected_option(self) -> RielloPriceItem | None:
        return self._option_by_code(str(self.option_combo.currentData() or ""))

    def _row_from_item(self, item: RielloPriceItem, qty: float, kind: str) -> dict[str, Any]:
        return {
            "kind": kind,
            "model": item.model,
            "code": item.code,
            "qty": qty,
            "price": float(item.price or 0.0),
            "currency": item.currency or "EUR",
            "power": item.power or "",
            "backup_min": item.backup_min or "",
            "dimensions": item.dimensions or "",
            "weight_kg": float(item.weight_kg or 0.0),
            "section": item.section or "",
            "description": item.description or "",
        }

    def _add_or_update_quote_row(self, row: dict[str, Any]) -> None:
        code = str(row.get("code") or "").upper()
        kind = str(row.get("kind") or "")
        qty = self._as_float(str(row.get("qty") or "1"), 1.0) or 1.0
        for existing in self.quote_rows:
            if str(existing.get("code") or "").upper() == code and str(existing.get("kind") or "") == kind:
                existing["qty"] = self._as_float(str(existing.get("qty") or "0"), 0.0) + qty
                break
        else:
            self.quote_rows.append(row)
        self.save_options()
        self.refresh_summary()

    def add_selected_ups_to_quote(self) -> None:
        item = self._selected_item()
        if not item:
            QMessageBox.warning(self, "Riello", "Сначала выберите модель ИБП.")
            return
        qty = self._as_float(self.ups_qty_edit.text(), 1.0) or 1.0
        self._add_or_update_quote_row(self._row_from_item(item, qty, "ИБП"))
        self.status_label.setText(f"Добавлено в лист подбора: {item.model} — {self._fmt_qty(qty)} шт.")

    def add_selected_option_to_quote(self) -> None:
        item = self._selected_option()
        if not item:
            QMessageBox.warning(self, "Riello", "Сначала выберите опцию.")
            return
        qty = self._as_float(self.option_qty_edit.text(), 1.0) or 1.0
        self._add_or_update_quote_row(self._row_from_item(item, qty, "Опция"))
        self.status_label.setText(f"Добавлена опция: {item.model} — {self._fmt_qty(qty)} шт.")

    def remove_selected_quote_rows(self) -> None:
        rows = sorted({index.row() for index in self.quote_table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for row in rows:
            if 0 <= row < len(self.quote_rows):
                self.quote_rows.pop(row)
        self.save_options()
        self.refresh_summary()

    def clear_quote_rows(self) -> None:
        if not self.quote_rows:
            return
        self.quote_rows.clear()
        self.save_options()
        self.refresh_summary()
        self.status_label.setText("Лист подбора оборудования очищен.")

    def brand_options(self) -> dict[str, Any]:
        selected_model = str(self.ups_combo.currentData() or self.ups_combo.currentText()).strip()
        return {
            "price_path": str(default_price_path()),
            "required_power_kw": self.required_power_edit.text().strip() or "20",
            "ups_model": selected_model,
            "ups_quantity": self.ups_qty_edit.text().strip() or "1",
            "quote_lines": self.quote_rows,
            "city": self._saved("riello/city", "Алматы"),
            "rate": self._saved("riello/rate", "1"),
            "margin_percent": self._saved("riello/margin_percent", "15"),
            "vat_percent": self._saved("riello/vat_percent", "0"),
            "special_percent": self._saved("riello/special_percent", "0"),
            "transport_cost": self._saved("riello/transport_cost", "0"),
            "customs_clearance": self._saved("riello/customs_clearance", "0"),
            "certificate": self._saved("riello/certificate", "0"),
            "transport_to_customer": self._saved("riello/transport_to_customer", "0"),
            "site_inspection": self._saved("riello/site_inspection", "0"),
            "installation_startup": self._saved("riello/installation_startup", "0"),
            "extra_cost": self._saved("riello/extra_cost", "0"),
        }

    def refresh_summary(self) -> None:
        self._refresh_details()
        self._refresh_quote_table()
        self._refresh_models_table()

    def _refresh_models_table(self) -> None:
        table = self.models_table
        table.setRowCount(0)
        selected_model = str(self.ups_combo.currentData() or "")

        for item in self.filtered_items:
            row = table.rowCount()
            table.insertRow(row)
            values = [
                item.model,
                item.code,
                item_power_label(item),
                f"{format_price(item.price)} {item.currency}",
                item_backup_label(item),
                item.dimensions or "—",
                f"{self._fmt_qty(item.weight_kg)} кг" if item.weight_kg else "—",
            ]
            for col, value in enumerate(values):
                table.setItem(row, col, QTableWidgetItem(str(value)))
            if item.model == selected_model:
                table.selectRow(row)

        table.resizeRowsToContents()
        required = self._as_float(self.required_power_edit.text(), 0.0)
        if self.filtered_items:
            shown_power = item_power_label(self.filtered_items[0])
            self.match_hint.setText(
                f"Из PDF-прайса показаны позиции ближайшей подходящей мощности: "
                f"{shown_power}. Найдено: {len(self.filtered_items)}. "
                f"Список отсортирован по стоимости: сначала самые дешевые."
            )
        elif required:
            self.match_hint.setText(f"В прайсе не найдены позиции под {required:g} кВА/кВт.")
        else:
            self.match_hint.setText("Укажите требуемую мощность в кВА/кВт, например 20, 100 или 4000.")

    def _refresh_details(self) -> None:
        item = self._selected_item()
        if not item:
            self.details.setText("Модель пока не выбрана.")
            return
        power = item_power_label(item)
        self.details.setText(
            f"Модель: {item.model}\n"
            f"Код: {item.code}\n"
            f"Мощность: {power}\n"
            f"Стоимость: {format_price(item.price)} {item.currency}\n"
            f"{item_backup_label(item)}\n"
            f"Габариты: {item.dimensions or '—'}\n"
            f"Вес: {self._fmt_qty(item.weight_kg)} кг\n"
            f"Описание из прайса/раздела: {item.description or item.section or '—'}"
        )

    def _refresh_quote_table(self) -> None:
        table = self.quote_table
        table.setRowCount(0)
        total = 0.0
        currency = "EUR"
        for row_data in self.quote_rows:
            qty = self._as_float(str(row_data.get("qty") or "0"), 0.0)
            price = self._as_float(str(row_data.get("price") or "0"), 0.0)
            currency = str(row_data.get("currency") or currency)
            amount = qty * price
            total += amount
            row = table.rowCount()
            table.insertRow(row)
            values = [
                row_data.get("kind", ""),
                row_data.get("model", ""),
                row_data.get("code", ""),
                self._fmt_qty(qty),
                f"{format_price(price)} {currency}",
                f"{format_price(amount)} {currency}",
            ]
            for col, value in enumerate(values):
                table.setItem(row, col, QTableWidgetItem(str(value)))
        if self.quote_rows:
            self.status_label.setText(f"В листе подбора: {len(self.quote_rows)} позиций, сумма {format_price(total)} {currency}.")
        table.resizeRowsToContents()

    # ------------------------- context/generate -------------------------
    def make_context(self) -> OfferContext:
        project_text = self.project_path_text().strip()
        project_dir = Path(project_text) if project_text else Path(".")
        output_dir = Path(self.output_path_text().strip() or project_dir)
        signer = self.owner._selected_signer()
        return OfferContext(
            brand=self.brand_name,
            project_dir=project_dir,
            template_path=Path(""),
            calc_path=Path(self._path_from_combo(self.calc_combo)),
            output_dir=output_dir,
            client_name=self.client_edit.text().strip() or "Client",
            pdf_dir=project_dir if project_dir.exists() else None,
            manager_name=self.owner.manager_name_edit.text().strip(),
            manager_position=self.owner.manager_position_edit.text().strip(),
            manager_email=self.owner.manager_email_edit.text().strip(),
            manager_phone=self.owner.manager_phone_edit.text().strip(),
            signer_name=signer["name"],
            signer_position=signer["position"],
            brand_options=self.brand_options(),
        )

    def validate_context(self, context: OfferContext) -> None:
        if not context.project_dir.exists():
            raise FileNotFoundError("Выберите существующую папку проекта.")
        if not context.calc_path.exists():
            raise FileNotFoundError("Выберите существующий Excel-шаблон Riello.")
        if context.calc_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Excel-шаблон Riello должен быть файлом .xlsx или .xlsm")
        if not self.quote_rows:
            raise ValueError("Добавьте хотя бы одну позицию в лист подбора оборудования.")
        if not context.output_dir.exists():
            context.output_dir.mkdir(parents=True, exist_ok=True)

    def generate(self) -> None:
        try:
            self.generate_btn.setEnabled(False)
            self.status_label.setText("Формирую Excel-расчет...")
            QApplication.processEvents()
            context = self.make_context()
            self.validate_context(context)
            module = get_brand_module(context.brand)
            out = module.make_offer(context)
            self.remember_values()
            self.status_label.setText(f"Готово: {out.name}")

            msg = QMessageBox(self)
            msg.setWindowTitle("SAM Offer Generator")
            msg.setIcon(QMessageBox.Question)
            msg.setText("Excel-расчет успешно сформирован.")
            msg.setInformativeText(str(out))
            open_folder_btn = msg.addButton("Открыть папку", QMessageBox.ActionRole)
            open_file_btn = msg.addButton("Открыть расчет", QMessageBox.ActionRole)
            msg.addButton("Закрыть", QMessageBox.RejectRole)
            msg.exec()
            clicked = msg.clickedButton()
            if clicked == open_folder_btn:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(out.parent)))
            elif clicked == open_file_btn:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))
        except Exception as exc:
            self.status_label.setText("Ошибка формирования")
            QMessageBox.critical(self, "Ошибка", str(exc))
        finally:
            self.generate_btn.setEnabled(True)
            self.refresh_summary()

    # ------------------------- persistence/page hooks -------------------------
    def save_options(self) -> None:
        options = self.brand_options()
        self._set_saved("riello/project_dir", self.project_path_text())
        self._set_saved("riello/client", self.client_edit.text().strip())
        self._set_saved("riello/calc_path", self._path_from_combo(self.calc_combo))
        self._set_saved("riello/output_dir", self.output_path_text())
        self._set_saved("brand", self.brand_name)
        self._set_saved("riello/required_power_kw", str(options.get("required_power_kw", "")))
        self._set_saved("riello/ups_model", str(options.get("ups_model", "")))
        self._set_saved("riello/ups_quantity", str(options.get("ups_quantity", "")))
        self._set_saved("riello/option_quantity", self.option_qty_edit.text().strip() or "1")
        self._set_saved("riello/quote_rows_json", json.dumps(self.quote_rows, ensure_ascii=False))
        settings = getattr(self.owner, "settings", None)
        if settings is not None:
            settings.sync()

    def remember_values(self) -> None:
        self.save_options()

    def on_settings_changed(self) -> None:
        template_path = self._saved("calc_template_path", "")
        if template_path:
            index = self.owner._find_combo_path(self.calc_combo, template_path)
            if index < 0:
                self.owner._add_path_item(self.calc_combo, template_path, is_file=True)
                index = self.calc_combo.count() - 1
            self.calc_combo.setCurrentIndex(index)
        self.refresh_summary()

    def clear_cache(self) -> None:
        self.project_dir_path = ""
        self.output_dir_path = ""
        self.quote_rows.clear()
        self._set_line_path(self.project_edit, "", is_file=False)
        self.client_edit.clear()
        self.calc_combo.clear()
        self._set_line_path(self.output_edit, "", is_file=False)
        self.required_power_edit.setText("20")
        self.ups_qty_edit.setText("1")
        self.option_qty_edit.setText("1")
        self.ensure_default_excel_template()
        self.refresh_summary()
        self.status_label.setText("Кэш очищен. Выберите папку проекта заново.")

    def refresh_preview(self) -> None:
        self.refresh_summary()

    def apply_responsive_metrics(self, scale: float) -> None:
        self.generate_btn.setMinimumWidth(int(220 * scale))
        self.generate_btn.setMinimumHeight(int(42 * scale))
