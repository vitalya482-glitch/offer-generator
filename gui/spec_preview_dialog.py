from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QCheckBox,
    QDialogButtonBox,
    QGroupBox,
    QHeaderView,
    QLabel,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


def _to_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


DESCRIPTION_OPTION_LABELS: dict[str, str] = {
    "stulz_unit": "Прецизионный кондиционер/неры Stulz",
    "cooling_capacity": "Хладопроизводительность",
    "unit_dimensions": "Размеры внутреннего блока",
    "condenser": "Конденсор",
}


def default_description_options() -> dict[str, bool]:
    return {key: True for key in DESCRIPTION_OPTION_LABELS}


class SpecPreviewDialog(QDialog):
    """Preview parsed STULZ specification data before inserting it into Word."""

    def __init__(
        self,
        spec_blocks: list[dict[str, Any]],
        warnings: list[str] | None = None,
        parent=None,
        description_options: dict[str, bool] | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр спецификаций")
        self.resize(1200, 760)
        self._description_checkboxes: dict[str, QCheckBox] = {}
        self._initial_description_options = default_description_options()
        if description_options:
            self._initial_description_options.update({
                key: bool(value)
                for key, value in description_options.items()
                if key in self._initial_description_options
            })

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        intro = QLabel(
            "Здесь показано, что программа прочитала из Calc.pdf и WinPlan.pdf. "
            "Проверьте коды, переводы, количество и технические характеристики перед генерацией КП."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        if warnings:
            warning_label = QLabel("⚠ " + "\n⚠ ".join(warnings))
            warning_label.setWordWrap(True)
            warning_label.setStyleSheet("color: #b45309; font-weight: 600;")
            layout.addWidget(warning_label)

        layout.addWidget(self._make_description_options_group())

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, stretch=1)

        if spec_blocks:
            self.tabs.addTab(self._make_total_tab(spec_blocks), "Total")
            for block in spec_blocks:
                self.tabs.addTab(self._make_model_tab(block), _to_text(block.get("model") or "Модель"))
        else:
            empty = QLabel("Нет данных для предпросмотра. Проверьте Excel, список моделей и папку спецификаций.")
            empty.setAlignment(Qt.AlignCenter)
            self.tabs.addTab(empty, "Нет данных")

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _make_description_options_group(self) -> QGroupBox:
        group = QGroupBox("Текст для включения в описание в КП")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(4)

        for key, label in DESCRIPTION_OPTION_LABELS.items():
            checkbox = QCheckBox(label)
            checkbox.setChecked(bool(self._initial_description_options.get(key, True)))
            self._description_checkboxes[key] = checkbox
            layout.addWidget(checkbox)

        return group

    def description_options(self) -> dict[str, bool]:
        options = default_description_options()
        for key, checkbox in self._description_checkboxes.items():
            options[key] = checkbox.isChecked()
        return options


    def _make_total_tab(self, spec_blocks: list[dict[str, Any]]) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(10)

        title = QLabel("Итоговые цены из Calc.pdf")
        title.setStyleSheet("font-weight: 700;")
        layout.addWidget(title)

        hint = QLabel(
            "В этой вкладке показаны финальные суммы из строки Total per quantity. "
            "Именно эти значения уже учитывают выбранные опции, конденсаторы и скидки STULZ."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels([
            "Модель в КП",
            "Модель в Calc",
            "Кол-во",
            "List total",
            "Purchase total",
            "Purchase / unit",
            "Валюта",
            "Файл Calc",
        ])
        table.setAlternatingRowColors(True)
        table.setWordWrap(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setRowCount(len(spec_blocks))

        total_purchase = 0.0
        total_list = 0.0
        purchase_found = False
        list_found = False
        currencies: list[str] = []

        for row, block in enumerate(spec_blocks):
            currency = _to_text(block.get("currency"))
            if currency and currency not in currencies:
                currencies.append(currency)

            list_total = block.get("total_list_price")
            purchase_total = block.get("total_purchase_price")
            if isinstance(list_total, (int, float)):
                total_list += float(list_total)
                list_found = True
            if isinstance(purchase_total, (int, float)):
                total_purchase += float(purchase_total)
                purchase_found = True

            calc_pdf = block.get("calc_pdf")
            calc_name = Path(calc_pdf).name if calc_pdf else "не найден"
            values = [
                _to_text(block.get("model")),
                _to_text(block.get("calc_model")),
                self._format_qty(block.get("quantity")),
                self._format_money(list_total),
                self._format_money(purchase_total),
                self._format_money(block.get("unit_purchase_price")),
                currency,
                calc_name,
            ]

            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col in (2, 3, 4, 5, 6):
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                if col in (3, 4, 5) and not value:
                    item.setBackground(Qt.yellow)
                    item.setToolTip("Цена не найдена в Calc.pdf")
                table.setItem(row, col, item)

        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col in range(2, 7):
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.Stretch)

        table.resizeRowsToContents()
        layout.addWidget(table, stretch=1)

        currency_text = ", ".join(currencies) if currencies else ""
        summary = QLabel(
            "Итого List: "
            f"{self._format_money(total_list) if list_found else '-'} {currency_text}\n"
            "Итого Purchase: "
            f"{self._format_money(total_purchase) if purchase_found else '-'} {currency_text}"
        )
        summary.setStyleSheet("font-weight: 700;")
        layout.addWidget(summary)

        return widget

    def _format_money(self, value: Any) -> str:
        if value is None or value == "":
            return ""
        try:
            s = f"{float(value):,.2f}"
            s = s.replace(",", "TEMP").replace(".", ",").replace("TEMP", " ")
            return s
        except Exception:
            return _to_text(value)

    def _format_qty(self, value: Any) -> str:
        if value is None or value == "":
            return ""
        try:
            number = float(value)
            return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
        except Exception:
            return _to_text(value)

    def _make_model_tab(self, block: dict[str, Any]) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(10)

        files = QLabel(self._files_text(block))
        files.setWordWrap(True)
        layout.addWidget(files)

        options_title = QLabel(_to_text(block.get("options_title") or "Опции"))
        options_title.setStyleSheet("font-weight: 700;")
        layout.addWidget(options_title)

        options_table = QTableWidget(0, 5)
        options_table.setHorizontalHeaderLabels(["№", "Код", "Из PDF", "В КП", "Кол-во"])
        self._setup_options_table(options_table)

        options = block.get("options") or []
        options_table.setRowCount(len(options))

        for row, option in enumerate(options):
            translated = bool(option.get("translated", True))
            values = [
                str(row + 1),
                _to_text(option.get("code")),
                _to_text(option.get("source_name") or option.get("name")),
                _to_text(option.get("description")),
                _to_text(option.get("qty")),
            ]

            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col in (0, 1, 4):
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

                if not translated:
                    item.setBackground(Qt.yellow)
                    item.setToolTip("Нет перевода в базе stulz_options.json")

                options_table.setItem(row, col, item)

        for row in range(options_table.rowCount()):
            options_table.setRowHeight(row, 54)

        layout.addWidget(options_table, stretch=3)

        specs_title = QLabel(_to_text(block.get("technical_specs_title") or "Технические характеристики"))
        specs_title.setStyleSheet("font-weight: 700;")
        layout.addWidget(specs_title)

        specs_table = QTableWidget(0, 2)
        specs_table.setHorizontalHeaderLabels(["Параметр", "Значение"])
        self._setup_specs_table(specs_table)

        specs = block.get("technical_specs") or []
        specs_table.setRowCount(len(specs))

        for row, spec in enumerate(specs):
            name_item = QTableWidgetItem(_to_text(spec.get("name")))
            value_item = QTableWidgetItem(_to_text(spec.get("value")))

            if spec.get("is_section"):
                name_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                name_item.setBackground(Qt.lightGray)
                value_item.setBackground(Qt.lightGray)

            specs_table.setItem(row, 0, name_item)
            specs_table.setItem(row, 1, value_item)

        specs_table.resizeRowsToContents()
        layout.addWidget(specs_table, stretch=2)

        return widget

    def _setup_options_table(self, table: QTableWidget) -> None:
        table.setAlternatingRowColors(True)
        table.setWordWrap(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)

        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(54)

        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # №
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # код
        header.setSectionResizeMode(2, QHeaderView.Stretch)           # из PDF
        header.setSectionResizeMode(3, QHeaderView.Stretch)           # в КП
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # кол-во

    def _setup_specs_table(self, table: QTableWidget) -> None:
        table.setAlternatingRowColors(True)
        table.setWordWrap(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)

        table.verticalHeader().setVisible(False)

        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)

    def _files_text(self, block: dict[str, Any]) -> str:
        calc_pdf = block.get("calc_pdf")
        winplan_pdf = block.get("winplan_pdf")
        calc_text = Path(calc_pdf).name if calc_pdf else "не найден"
        winplan_text = Path(winplan_pdf).name if winplan_pdf else "не найден"
        return f"Calc PDF: {calc_text}\nWinPlan PDF: {winplan_text}"