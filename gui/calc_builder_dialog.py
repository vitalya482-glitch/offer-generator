from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


KAZAKHSTAN_CITIES = [
    "Алматы",
    "Астана",
    "Шымкент",
    "Караганда",
    "Атырау",
    "Актау",
    "Актобе",
    "Павлодар",
    "Усть-Каменогорск",
    "Семей",
    "Костанай",
    "Кокшетау",
    "Петропавловск",
    "Тараз",
    "Талдыкорган",
    "Уральск",
    "Кызылорда",
    "Туркестан",
    "Жезказган",
    "Экибастуз",
    "Темиртау",
    "Рудный",
    "Балхаш",
]

INCOTERMS = ["EXW", "FCA", "CPT", "CIP", "DAP", "DPU", "DDP", "FOB", "CFR", "CIF"]
CURRENCIES = ["EUR", "USD", "KZT"]

COST_ROWS = [
    ("vat", "НДС, %", "16"),
    ("margin", "Маржа, %", "25"),
    ("logistics", "Логистика", ""),
    ("certification", "Сертификация", ""),
    ("customs", "Таможенная очистка", ""),
    ("duty", "Пошлина", ""),
    ("installation", "Монтаж", ""),
    ("startup", "ПНР", ""),
    ("survey", "Обследование объекта", ""),
    ("storage", "Склад / хранение", ""),
    ("discount", "Скидка, %", ""),
]


def _to_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _format_money(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        s = f"{float(value):,.2f}"
        return s.replace(",", "TEMP").replace(".", ",").replace("TEMP", " ")
    except Exception:
        return _to_text(value)


def _format_qty(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
    except Exception:
        return _to_text(value)


class CalcBuilderDialog(QDialog):
    """First UI iteration for building an Excel calculation from STULZ Calc.pdf totals."""

    def __init__(self, spec_blocks: list[dict[str, Any]], parent=None, template_path: str = "") -> None:
        super().__init__(parent)
        self.spec_blocks = spec_blocks
        self.setWindowTitle("Расчёт из спецификаций")
        self.resize(1180, 760)
        self._cost_widgets: dict[str, tuple[QCheckBox, QLineEdit]] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title = QLabel("Расчёт из спецификаций")
        title.setStyleSheet("font-size: 16px; font-weight: 700;")
        layout.addWidget(title)

        hint = QLabel(
            "Первая итерация: здесь собраны данные из Calc.pdf и параметры будущего расчёта. "
            "Сохранение Excel и формулы будут подключены следующим шагом."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        layout.addWidget(self._make_template_group(template_path))
        layout.addWidget(self._make_models_table(), stretch=1)
        layout.addWidget(self._make_parameters_group())

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _make_template_group(self, template_path: str) -> QGroupBox:
        group = QGroupBox("Шаблон и сохранение")
        grid = QGridLayout(group)
        grid.setColumnStretch(1, 1)

        self.template_edit = QLineEdit(template_path)
        self.template_edit.setReadOnly(True)
        grid.addWidget(QLabel("Excel-шаблон Calc"), 0, 0)
        grid.addWidget(self.template_edit, 0, 1)

        browse_template = QPushButton("Выбрать шаблон")
        browse_template.setObjectName("GhostButton")
        browse_template.clicked.connect(self._browse_template)
        grid.addWidget(browse_template, 0, 2)

        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Путь будет выбран при сохранении Excel")
        grid.addWidget(QLabel("Куда сохранить"), 1, 0)
        grid.addWidget(self.output_edit, 1, 1)

        browse_output = QPushButton("Выбрать путь")
        browse_output.setObjectName("GhostButton")
        browse_output.clicked.connect(self._browse_output)
        grid.addWidget(browse_output, 1, 2)

        return group

    def _browse_template(self) -> None:
        current = self.template_edit.text().strip()
        start_dir = str(Path(current).parent) if current else ""
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-шаблон расчёта", start_dir, "Excel (*.xlsx)")
        if path:
            self.template_edit.setText(path)

    def _browse_output(self) -> None:
        current = self.output_edit.text().strip()
        start_dir = str(Path(current).parent) if current else ""
        path, _ = QFileDialog.getSaveFileName(self, "Куда сохранить расчёт", start_dir, "Excel (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_edit.setText(path)

    def _make_models_table(self) -> QWidget:
        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels([
            "Модель в КП",
            "Модель в Calc",
            "Кол-во",
            "List total",
            "Purchase total со скидкой",
            "Валюта",
            "Calc PDF",
        ])
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setRowCount(len(self.spec_blocks))

        for row, block in enumerate(self.spec_blocks):
            calc_pdf = block.get("calc_pdf")
            calc_name = Path(calc_pdf).name if calc_pdf else ""
            values = [
                _to_text(block.get("model")),
                _to_text(block.get("calc_model")),
                _format_qty(block.get("quantity")),
                _format_money(block.get("total_list_price")),
                _format_money(block.get("total_purchase_price")),
                _to_text(block.get("currency")),
                calc_name,
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter if col in (2, 3, 4, 5) else Qt.AlignLeft | Qt.AlignVCenter)
                table.setItem(row, col, item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col in range(2, 6):
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.Stretch)
        table.resizeRowsToContents()
        return table

    def _make_parameters_group(self) -> QGroupBox:
        group = QGroupBox("Параметры расчёта")
        grid = QGridLayout(group)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

        self.incoterms_combo = QComboBox()
        self.incoterms_combo.addItems(INCOTERMS)
        self.incoterms_combo.setCurrentText("DDP")
        grid.addWidget(QLabel("Условия поставки"), 0, 0)
        grid.addWidget(self.incoterms_combo, 0, 1)

        self.delivery_point_combo = QComboBox()
        self.delivery_point_combo.setEditable(True)
        self.delivery_point_combo.addItems(KAZAKHSTAN_CITIES)
        self.delivery_point_combo.setCurrentText("Алматы")
        self.delivery_point_combo.setInsertPolicy(QComboBox.NoInsert)
        self.delivery_point_combo.completer().setCaseSensitivity(Qt.CaseInsensitive)
        self.delivery_point_combo.completer().setFilterMode(Qt.MatchContains)
        grid.addWidget(QLabel("Пункт поставки"), 0, 2)
        grid.addWidget(self.delivery_point_combo, 0, 3)

        self.currency_combo = QComboBox()
        self.currency_combo.addItems(CURRENCIES)
        self.currency_combo.setCurrentText("EUR")
        grid.addWidget(QLabel("Валюта"), 1, 0)
        grid.addWidget(self.currency_combo, 1, 1)

        self.exchange_rate_edit = QLineEdit()
        self.exchange_rate_edit.setPlaceholderText("например 530")
        grid.addWidget(QLabel("Курс"), 1, 2)
        grid.addWidget(self.exchange_rate_edit, 1, 3)

        row = 2
        for key, label, default in COST_ROWS:
            checkbox = QCheckBox(label)
            checkbox.setChecked(True)
            edit = QLineEdit(default)
            edit.setPlaceholderText("пусто")
            grid.addWidget(checkbox, row, 0)
            grid.addWidget(edit, row, 1)
            self._cost_widgets[key] = (checkbox, edit)
            row += 1

        note = QLabel(
            "Если галочку убрать, строка останется в будущем Excel-шаблоне, "
            "но ячейка значения будет пустой."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #6b7280;")
        grid.addWidget(note, row, 0, 1, 4)

        return group
