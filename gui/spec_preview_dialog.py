from __future__ import annotations

from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
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


class SpecPreviewDialog(QDialog):
    """Preview parsed STULZ specification data before inserting it into Word."""

    def __init__(self, spec_blocks: list[dict[str, Any]], warnings: list[str] | None = None, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр спецификаций")
        self.resize(1200, 760)

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

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, stretch=1)

        if spec_blocks:
            for block in spec_blocks:
                self.tabs.addTab(self._make_model_tab(block), _to_text(block.get("model") or "Модель"))
        else:
            empty = QLabel("Нет данных для предпросмотра. Проверьте Excel, список моделей и папку спецификаций.")
            empty.setAlignment(Qt.AlignCenter)
            self.tabs.addTab(empty, "Нет данных")

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

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