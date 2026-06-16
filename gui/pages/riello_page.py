from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QFileDialog,
    QComboBox,
    QGridLayout,
    QLabel,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QVBoxLayout,
    QWidget,
)

from core.models import OfferContext
from core.riello_price import (
    default_price_path,
    format_price,
    item_backup_label,
    item_display_with_price,
    item_power_label,
    load_price_items,
    nearest_power_items,
    power_modules,
)
from core.runtime_paths import resource_path


class RielloPage(QWidget):
    """Страница Riello: сначала ввод мощности, затем выбор подходящей позиции из PDF-прайса."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self._updating = False
        self.price_items = []
        self.filtered_items = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.output_dir_path = self._saved("riello/output_dir", "")

        self.client_edit = QLineEdit(self._saved("riello/client", self._saved("client", "ТОО Example")))
        self.calc_combo = QComboBox()
        self.calc_combo.setEditable(True)
        saved_calc = self._saved("riello/calc_path", "") or self._saved("calc_template_path", "")
        if saved_calc:
            owner._add_path_item(self.calc_combo, saved_calc, is_file=True)
            self.calc_combo.setCurrentIndex(0)
        self.output_edit = QLineEdit(owner._display_dir(self.output_dir_path))
        self.output_edit.setToolTip(self.output_dir_path)

        files_card = owner._card("Riello: Excel-шаблон")
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
        self.required_power_edit.setPlaceholderText("например 20 или 100")

        self.autonomy_edit = QLineEdit(self._saved("riello/autonomy_min", "20"))
        self.autonomy_edit.setPlaceholderText("мин")
        self.autonomy_edit.setMaximumWidth(95)

        self.ups_qty_edit = QLineEdit(self._saved("riello/ups_quantity", "1"))
        self.ups_qty_edit.setPlaceholderText("шт")
        self.ups_qty_edit.setMaximumWidth(95)

        self.ups_combo = QComboBox()
        self.ups_combo.setMinimumWidth(480)

        input_grid.addWidget(QLabel("Мощность, кВА/кВт"), 0, 0)
        input_grid.addWidget(self.required_power_edit, 0, 1)
        input_grid.addWidget(QLabel("Автономия, мин"), 0, 2)
        input_grid.addWidget(self.autonomy_edit, 0, 3)
        input_grid.addWidget(QLabel("Кол-во ИБП"), 0, 4)
        input_grid.addWidget(self.ups_qty_edit, 0, 5)
        input_grid.addWidget(QLabel("Модель ИБП"), 1, 0)
        input_grid.addWidget(self.ups_combo, 1, 1, 1, 5)

        self.match_hint = QLabel("")
        self.match_hint.setObjectName("Hint")
        self.match_hint.setWordWrap(True)
        input_grid.addWidget(self.match_hint, 2, 1, 1, 5)
        layout.addWidget(input_card)

        table_card = owner._card("Подходящие позиции из PDF-прайса")
        self.models_table = QTableWidget(0, 7)
        self.models_table.setHorizontalHeaderLabels(["Модель", "Код", "Мощность", "Стоимость", "Автономия", "Габариты", "Вес"])
        self.models_table.verticalHeader().setVisible(False)
        self.models_table.setAlternatingRowColors(True)
        self.models_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.models_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.models_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.models_table.setMinimumHeight(180)
        self.models_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.models_table.setEditTriggers(QTableWidget.NoEditTriggers)
        table_card.layout().addWidget(self.models_table)
        layout.addWidget(table_card)

        details_card = owner._card("Карточка выбранной модели")
        self.details = QLabel("")
        self.details.setObjectName("Hint")
        self.details.setWordWrap(True)
        details_card.layout().addWidget(self.details)
        layout.addWidget(details_card)
        layout.addStretch(1)

        self._load_price_items()
        self.ensure_default_excel_template()
        self._connect_changes()
        self.refresh_summary()

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

    def _load_price_items(self) -> None:
        self._updating = True
        try:
            self.price_items = load_price_items(default_price_path())
            self._reload_ups_models(keep_saved=True)
        finally:
            self._updating = False

    def _connect_changes(self) -> None:
        self.client_edit.textChanged.connect(self._on_changed)
        self.calc_combo.currentTextChanged.connect(self._on_changed)
        self.output_edit.textChanged.connect(self._on_output_dir_changed)
        self.required_power_edit.textChanged.connect(self._on_power_changed)
        self.ups_combo.currentIndexChanged.connect(self._on_ups_changed)
        self.ups_qty_edit.textChanged.connect(self._on_changed)
        self.autonomy_edit.textChanged.connect(self._on_changed)
        self.models_table.cellClicked.connect(self._on_table_row_clicked)


    def _path_from_combo(self, combo: QComboBox) -> str:
        return self.owner._path_from_combo(combo)

    def _set_line_path(self, line_edit: QLineEdit, path_text: str, is_file: bool = False) -> None:
        self.owner._set_line_path(line_edit, path_text, is_file=is_file)

    def output_path_text(self) -> str:
        return self.output_dir_path or self.output_edit.text().strip()

    def browse_calc_file(self) -> None:
        current_calc = self._path_from_combo(self.calc_combo)
        start_dir = str(Path(current_calc).parent) if current_calc else self.owner._project_path_text()
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
        start_dir = self.output_path_text() or self.owner._project_path_text()
        path = QFileDialog.getExistingDirectory(self, "Выберите папку результата Riello", start_dir)
        if path:
            self.output_dir_path = path
            self._set_line_path(self.output_edit, path, is_file=False)
            self._set_saved("riello/output_dir", path)
            self.save_options()

    def _on_output_dir_changed(self) -> None:
        if not getattr(self.owner, "_updating_path_display", False):
            self.output_dir_path = self.output_edit.text().strip()
            self.output_edit.setToolTip(self.output_dir_path)
        self._on_changed()

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
        self.owner._refresh_preview()

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

    def _item_by_model(self, model: str):
        model_upper = (model or "").upper()
        for item in getattr(self, "price_items", []):
            if item.model.upper() == model_upper:
                return item
        return None

    def _selected_item(self):
        return self._item_by_model(str(self.ups_combo.currentData() or ""))

    def _fmt_qty(self, value: float | str) -> str:
        try:
            number = float(value)
            return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
        except Exception:
            return str(value)

    def _default_power_module(self, selected_model: str) -> str:
        """Технический fallback для текущего Excel-экспортера. На странице это поле пока не показываем."""
        selected = self._item_by_model(selected_model)
        prefix = selected_model.split(" ", 1)[0] if selected_model else "SRT"
        modules = power_modules(getattr(self, "price_items", []), prefix=prefix)
        if selected and " 20 PM" in selected.model.upper():
            return selected.model
        return modules[0].model if modules else selected_model

    def brand_options(self) -> dict[str, str]:
        selected_model = str(self.ups_combo.currentData() or self.ups_combo.currentText()).strip()
        return {
            "price_path": str(default_price_path()),
            "required_power_kw": self.required_power_edit.text().strip() or "20",
            "ups_model": selected_model,
            "ups_quantity": self.ups_qty_edit.text().strip() or "1",
            "autonomy_min": self.autonomy_edit.text().strip(),
            # Ниже — временные дефолты для старого генератора Excel. На странице эти поля пока не показываем.
            "power_module": self._default_power_module(selected_model),
            "modules_per_ups": "",
            "battery_cabinet_type": "",
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

    def save_options(self) -> None:
        options = self.brand_options()
        self._set_saved("riello/client", self.client_edit.text().strip())
        self._set_saved("riello/calc_path", self._path_from_combo(self.calc_combo))
        self._set_saved("riello/output_dir", self.output_path_text())
        for key in ("required_power_kw", "ups_model", "ups_quantity", "autonomy_min"):
            self._set_saved(f"riello/{key}", str(options.get(key, "")))
        settings = getattr(self.owner, "settings", None)
        if settings is not None:
            settings.sync()

    def refresh_summary(self) -> None:
        self._refresh_models_table()
        self._refresh_details()

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
                f"В списке модели сразу выводится ориентировочная стоимость."
            )
        elif required:
            self.match_hint.setText(f"В прайсе не найдены позиции под {required:g} кВА/кВт.")
        else:
            self.match_hint.setText("Укажите требуемую мощность в кВА/кВт, например 20 или 100.")

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
            f"Автономия: {item_backup_label(item)}\n"
            f"Габариты: {item.dimensions or '—'}\n"
            f"Вес: {self._fmt_qty(item.weight_kg)} кг\n"
            f"Описание из прайса/раздела: {item.description or item.section or '—'}"
        )


    def primary_button_text(self) -> str:
        return "Сформировать Excel"

    def clear_cache(self) -> None:
        self.client_edit.clear()
        self.calc_combo.clear()
        self.output_dir_path = ""
        self.output_edit.clear()
        self.output_edit.setToolTip("")
        self.required_power_edit.setText("20")
        self.autonomy_edit.setText("20")
        self.ups_qty_edit.setText("1")
        self.ensure_default_excel_template()
        self.refresh_summary()

    def apply_scan_results(self, project_dir: Path, found: dict, force: bool = False) -> None:
        if not self.output_path_text().strip():
            try:
                from gui.path_helpers import infer_output_dir
                guessed = infer_output_dir(str(project_dir))
            except Exception:
                guessed = str(project_dir)
            self.output_dir_path = guessed
            self._set_line_path(self.output_edit, guessed, is_file=False)
        client = self.owner._extract_client_from_project_dir(str(project_dir))
        if client and (not self.client_edit.text().strip() or self.client_edit.text().strip() == "ТОО Example"):
            self.client_edit.setText(client)
        self.save_options()

    def make_context(self) -> OfferContext:
        project_text = self.owner._project_path_text().strip()
        project_dir = Path(project_text) if project_text else Path(".")
        output_dir = Path(self.output_path_text().strip() or project_dir)
        signer = self.owner._selected_signer()
        return OfferContext(
            brand="Riello",
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

    def refresh_preview(self) -> None:
        # У Riello пока нет отдельного текстового блока проверки, поэтому обновляем только карточку.
        self.refresh_summary()

    def remember_values(self) -> None:
        self.save_options()
