from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
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

from core.riello_price import (
    default_price_path,
    item_display_with_power,
    item_power_kw,
    load_price_items,
    nearest_power_items,
    power_modules,
    rack_cabinets,
)
from core.runtime_paths import resource_path


class RielloPage(QWidget):
    """Страница Riello: выбор мощности, ИБП и формирование Excel-расчета по шаблону."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self._updating = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        files_card = owner._card("Riello: Excel-шаблон и параметры расчета")
        grid = QGridLayout()
        files_card.layout().addLayout(grid)
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)

        owner._add_row(grid, 0, "Клиент", owner.client_edit, None, None)
        owner._add_row(grid, 1, "Excel-шаблон", owner.calc_combo, "Выбрать", owner._browse_calc_file)
        owner._add_row(grid, 2, "Папка результата", owner.output_edit, "Выбрать", owner._browse_output_dir)
        layout.addWidget(files_card)

        config_card = owner._card("Конфигурация Riello")
        config_grid = QGridLayout()
        config_card.layout().addLayout(config_grid)
        config_grid.setColumnStretch(1, 1)
        config_grid.setVerticalSpacing(12)
        config_grid.setHorizontalSpacing(10)

        self.required_power_edit = QLineEdit(self._saved("riello/required_power_kw", "60"))
        self.ups_combo = QComboBox()
        self.power_module_combo = QComboBox()
        self.ups_qty_edit = QLineEdit(self._saved("riello/ups_quantity", "1"))
        self.modules_per_ups_edit = QLineEdit(self._saved("riello/modules_per_ups", "3"))
        self.autonomy_edit = QLineEdit(self._saved("riello/autonomy_min", ""))
        self.battery_cabinet_edit = QLineEdit(self._saved("riello/battery_cabinet_type", ""))
        self.city_edit = QLineEdit(self._saved("riello/city", "Алматы"))
        self.rate_edit = QLineEdit(self._saved("riello/rate", "1"))
        self.margin_edit = QLineEdit(self._saved("riello/margin_percent", "15"))
        self.vat_edit = QLineEdit(self._saved("riello/vat_percent", "0"))
        self.special_edit = QLineEdit(self._saved("riello/special_percent", "0"))
        self.transport_cost_edit = QLineEdit(self._saved("riello/transport_cost", "2000"))
        self.customs_edit = QLineEdit(self._saved("riello/customs_clearance", "200"))
        self.certificate_edit = QLineEdit(self._saved("riello/certificate", "200"))
        self.transport_to_customer_edit = QLineEdit(self._saved("riello/transport_to_customer", "1500"))
        self.site_inspection_edit = QLineEdit(self._saved("riello/site_inspection", "0"))
        self.installation_edit = QLineEdit(self._saved("riello/installation_startup", "0"))
        self.extra_cost_edit = QLineEdit(self._saved("riello/extra_cost", "0"))

        owner._add_row(config_grid, 0, "Мощность, кВт", self.required_power_edit, None, None)
        owner._add_row(config_grid, 1, "Модель ИБП", self.ups_combo, None, None)
        owner._add_row(config_grid, 2, "Кол-во ИБП", self.ups_qty_edit, None, None)
        owner._add_row(config_grid, 3, "Силовой модуль", self.power_module_combo, None, None)
        owner._add_row(config_grid, 4, "Модулей на ИБП", self.modules_per_ups_edit, None, None)
        owner._add_row(config_grid, 5, "Автономия", self.autonomy_edit, None, None)
        owner._add_row(config_grid, 6, "Бат. шкаф", self.battery_cabinet_edit, None, None)
        owner._add_row(config_grid, 7, "Город DDP", self.city_edit, None, None)
        owner._add_row(config_grid, 8, "Курс", self.rate_edit, None, None)
        owner._add_row(config_grid, 9, "Маржа, %", self.margin_edit, None, None)
        owner._add_row(config_grid, 10, "НДС, %", self.vat_edit, None, None)
        owner._add_row(config_grid, 11, "Спецусловие, %", self.special_edit, None, None)
        owner._add_row(config_grid, 12, "Доставка", self.transport_cost_edit, None, None)
        owner._add_row(config_grid, 13, "Таможня", self.customs_edit, None, None)
        owner._add_row(config_grid, 14, "Сертификат", self.certificate_edit, None, None)
        owner._add_row(config_grid, 15, "Дост. клиенту", self.transport_to_customer_edit, None, None)
        owner._add_row(config_grid, 16, "Обследование", self.site_inspection_edit, None, None)
        owner._add_row(config_grid, 17, "Монтаж/ПНР", self.installation_edit, None, None)
        owner._add_row(config_grid, 18, "Доп. расходы", self.extra_cost_edit, None, None)
        layout.addWidget(config_card)

        self.items_table = QTableWidget(0, 7)
        self.items_table.setHorizontalHeaderLabels(["Позиция", "Код", "Габариты", "Вес", "Цена", "Кол-во", "Сумма"])
        self.items_table.verticalHeader().setVisible(False)
        self.items_table.setAlternatingRowColors(True)
        self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.items_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.items_table.setMinimumHeight(145)
        table_card = owner._card("Состав расчета")
        table_card.layout().addWidget(self.items_table)
        layout.addWidget(table_card)

        hint_card = owner._card("Подсказка")
        self.hint = QLabel(
            "Сначала укажите требуемую мощность в кВт. Список «Модель ИБП» автоматически покажет "
            "модели ближайшей подходящей мощности из прайса Riello. Excel-шаблон копируется и заполняется; "
            "Word-шаблон для Riello не требуется."
        )
        self.hint.setObjectName("Hint")
        self.hint.setWordWrap(True)
        hint_card.layout().addWidget(self.hint)
        layout.addWidget(hint_card)
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
        combo = self.owner.calc_combo
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
            self.cabinet_items = rack_cabinets(self.price_items) or self.price_items[:]
            self._reload_ups_models(keep_saved=True)
            self._reload_power_modules(update_modules_count=True)
        finally:
            self._updating = False

    def _reload_ups_models(self, keep_saved: bool = True) -> None:
        required_kw = self._as_float(self.required_power_edit.text(), 0.0)
        saved_ups = self._saved("riello/ups_model", "SRT 60 PWC") if keep_saved else ""
        current_ups = str(self.ups_combo.currentData() or "")
        previous_ups = saved_ups or current_ups

        candidates = nearest_power_items(getattr(self, "cabinet_items", []), required_kw)
        if not candidates:
            candidates = getattr(self, "cabinet_items", [])[:]

        self.ups_combo.blockSignals(True)
        try:
            self.ups_combo.clear()
            for item in candidates:
                self.ups_combo.addItem(item_display_with_power(item), item.model)

            index = self.ups_combo.findData(previous_ups)
            if index < 0 and self.ups_combo.count() > 0:
                index = 0
            if index >= 0:
                self.ups_combo.setCurrentIndex(index)
        finally:
            self.ups_combo.blockSignals(False)

    def _reload_power_modules(self, update_modules_count: bool = False) -> None:
        selected_model = str(self.ups_combo.currentData() or self.ups_combo.currentText()).strip()
        prefix = selected_model.split(" ", 1)[0] if selected_model else "SRT"
        modules = power_modules(getattr(self, "price_items", []), prefix=prefix)
        if not modules:
            modules = power_modules(getattr(self, "price_items", []), prefix="SRT")
        saved_module = self._saved("riello/power_module", f"{prefix} 20 PM P")

        self.power_module_combo.blockSignals(True)
        try:
            self.power_module_combo.clear()
            for item in modules:
                self.power_module_combo.addItem(item_display_with_power(item), item.model)
            module_index = self.power_module_combo.findData(saved_module)
            if module_index < 0 and self.power_module_combo.count() > 0:
                module_index = 0
            if module_index >= 0:
                self.power_module_combo.setCurrentIndex(module_index)
        finally:
            self.power_module_combo.blockSignals(False)

        if update_modules_count:
            self._set_default_modules_per_ups()

    def _set_default_modules_per_ups(self) -> None:
        ups = self._item_by_model(str(self.ups_combo.currentData() or ""))
        module = self._item_by_model(str(self.power_module_combo.currentData() or ""))
        if not ups or not module:
            return
        ups_power = item_power_kw(ups)
        module_power = item_power_kw(module)
        if ups_power > 0 and module_power > 0:
            calculated = max(round(ups_power / module_power), 1)
            self.modules_per_ups_edit.blockSignals(True)
            try:
                self.modules_per_ups_edit.setText(str(calculated))
            finally:
                self.modules_per_ups_edit.blockSignals(False)

    def _connect_changes(self) -> None:
        self.required_power_edit.textChanged.connect(self._on_power_changed)
        self.ups_combo.currentIndexChanged.connect(self._on_ups_changed)
        self.power_module_combo.currentIndexChanged.connect(self._on_module_changed)
        widgets = [
            self.ups_qty_edit,
            self.modules_per_ups_edit,
            self.autonomy_edit,
            self.battery_cabinet_edit,
            self.city_edit,
            self.rate_edit,
            self.margin_edit,
            self.vat_edit,
            self.special_edit,
            self.transport_cost_edit,
            self.customs_edit,
            self.certificate_edit,
            self.transport_to_customer_edit,
            self.site_inspection_edit,
            self.installation_edit,
            self.extra_cost_edit,
        ]
        for widget in widgets:
            widget.textChanged.connect(self._on_changed)

    def _on_power_changed(self) -> None:
        if self._updating:
            return
        self._updating = True
        try:
            self._reload_ups_models(keep_saved=False)
            self._reload_power_modules(update_modules_count=True)
        finally:
            self._updating = False
        self._on_changed()

    def _on_ups_changed(self) -> None:
        if self._updating:
            return
        self._updating = True
        try:
            self._reload_power_modules(update_modules_count=True)
        finally:
            self._updating = False
        self._on_changed()

    def _on_module_changed(self) -> None:
        if self._updating:
            return
        self._set_default_modules_per_ups()
        self._on_changed()

    def _on_changed(self) -> None:
        if self._updating:
            return
        self.save_options()
        self.refresh_summary()
        self.owner._refresh_preview()

    def save_options(self) -> None:
        options = self.brand_options()
        for key, value in options.items():
            if key == "price_path":
                continue
            self._set_saved(f"riello/{key}", str(value))
        settings = getattr(self.owner, "settings", None)
        if settings is not None:
            settings.sync()

    def brand_options(self) -> dict[str, str]:
        return {
            "price_path": str(default_price_path()),
            "required_power_kw": self.required_power_edit.text().strip() or "60",
            "ups_model": str(self.ups_combo.currentData() or self.ups_combo.currentText()).strip(),
            "power_module": str(self.power_module_combo.currentData() or self.power_module_combo.currentText()).strip(),
            "ups_quantity": self.ups_qty_edit.text().strip() or "1",
            "modules_per_ups": self.modules_per_ups_edit.text().strip() or "3",
            "autonomy_min": self.autonomy_edit.text().strip(),
            "battery_cabinet_type": self.battery_cabinet_edit.text().strip(),
            "city": self.city_edit.text().strip() or "Алматы",
            "rate": self.rate_edit.text().strip() or "1",
            "margin_percent": self.margin_edit.text().strip() or "15",
            "vat_percent": self.vat_edit.text().strip() or "0",
            "special_percent": self.special_edit.text().strip() or "0",
            "transport_cost": self.transport_cost_edit.text().strip() or "0",
            "customs_clearance": self.customs_edit.text().strip() or "0",
            "certificate": self.certificate_edit.text().strip() or "0",
            "transport_to_customer": self.transport_to_customer_edit.text().strip() or "0",
            "site_inspection": self.site_inspection_edit.text().strip() or "0",
            "installation_startup": self.installation_edit.text().strip() or "0",
            "extra_cost": self.extra_cost_edit.text().strip() or "0",
        }

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

    def _format_money(self, value: float) -> str:
        try:
            return f"{float(value):,.2f}".replace(",", " ").replace(".", ",")
        except Exception:
            return str(value)

    def refresh_summary(self) -> None:
        table = self.items_table
        table.setRowCount(0)
        ups = self._item_by_model(str(self.ups_combo.currentData() or ""))
        module = self._item_by_model(str(self.power_module_combo.currentData() or ""))
        ups_qty = self._as_float(self.ups_qty_edit.text(), 1.0) or 1.0
        modules_per_ups = self._as_float(self.modules_per_ups_edit.text(), 3.0) or 3.0
        rows = []
        if ups:
            rows.append((ups, ups_qty))
        if module:
            rows.append((module, ups_qty * modules_per_ups))
        for item, qty in rows:
            row = table.rowCount()
            table.insertRow(row)
            values = [
                item.model,
                item.code,
                item.dimensions,
                str(item.weight_kg),
                f"{self._format_money(item.price)} {item.currency}",
                str(int(qty)) if float(qty).is_integer() else str(qty).replace(".", ","),
                f"{self._format_money(item.price * qty)} {item.currency}",
            ]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value)
                table.setItem(row, col, cell)
        table.resizeRowsToContents()
