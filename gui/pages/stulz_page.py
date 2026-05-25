from __future__ import annotations

from core.models import OfferContext

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGridLayout,
    QLabel,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QHBoxLayout,
    QVBoxLayout,
    QWidget,
)


class StulzPage(QWidget):
    """Страница STULZ: файлы КП, проверка данных и список моделей спецификации."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self._updating_spec_models = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        files_card = owner._card("Stulz: файлы и параметры КП")
        grid = QGridLayout()
        files_card.layout().addLayout(grid)
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)

        owner._add_row(grid, 0, "Клиент", owner.client_edit, None, None)
        owner._add_row(grid, 1, "Excel-расчет", owner.calc_combo, "Обновить", lambda: owner._scan_project(force=True))
        owner._add_row(grid, 2, "Лист Excel", owner.sheet_combo, "Листы", owner._load_sheets)
        owner._add_row(grid, 3, "Папка спецификаций", owner.spec_edit, "Выбрать", owner._browse_spec_dir)
        owner._add_row(grid, 4, "Папка результата", owner.output_edit, "Выбрать", owner._browse_output_dir)
        layout.addWidget(files_card)

        bottom = QHBoxLayout()
        preview_card = owner._card("Проверка данных")
        owner.preview = QTextEdit()
        owner.preview.setReadOnly(True)
        preview_card.layout().addWidget(owner.preview)
        bottom.addWidget(preview_card, stretch=2)

        spec_card = owner._card("Спецификации")
        spec_hint = QLabel(
            "Модели из расчета. Позже КП будет формироваться по этому списку: "
            "можно отключать позиции и менять количество."
        )
        spec_hint.setObjectName("Hint")
        spec_hint.setWordWrap(True)
        spec_card.layout().addWidget(spec_hint)

        owner.spec_models_table = QTableWidget(0, 3)
        owner.spec_models_table.setHorizontalHeaderLabels(["Вкл", "Модель", "Кол-во"])
        owner.spec_models_table.verticalHeader().setVisible(False)
        owner.spec_models_table.setAlternatingRowColors(True)
        owner.spec_models_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        owner.spec_models_table.setMinimumHeight(170)
        owner.spec_models_table.setColumnWidth(0, 52)
        owner.spec_models_table.setColumnWidth(2, 80)
        owner.spec_models_table.horizontalHeader().setStretchLastSection(False)
        owner.spec_models_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        spec_card.layout().addWidget(owner.spec_models_table)
        spec_card.layout().addWidget(owner.status_label)
        bottom.addWidget(spec_card, stretch=1)

        layout.addLayout(bottom)

    def clear_spec_models(self) -> None:
        self.owner.spec_models_table.setRowCount(0)

    def current_spec_model_state(self) -> dict[str, tuple[bool, str]]:
        state: dict[str, tuple[bool, str]] = {}
        table = self.owner.spec_models_table
        for row in range(table.rowCount()):
            enabled_item = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            if not model_item:
                continue
            model = model_item.text().strip()
            if not model:
                continue
            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
            qty = qty_item.text().strip() if qty_item else ""
            state[model] = (enabled, qty)
        return state


    def selected_spec_models(self) -> list[dict[str, object]]:
        models: list[dict[str, object]] = []
        table = self.owner.spec_models_table
        for row in range(table.rowCount()):
            enabled_item = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            if not model_item:
                continue
            model = model_item.text().strip()
            if not model:
                continue
            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
            qty_text = qty_item.text().strip() if qty_item else ""
            try:
                qty = float(qty_text.replace(",", ".")) if qty_text else 0.0
            except Exception:
                qty = 0.0
            models.append({"enabled": enabled, "model": model, "qty": qty_text, "qty_value": qty})
        return models

    def refresh_spec_models(self, context: OfferContext | None = None) -> None:
        owner = self.owner
        table = owner.spec_models_table
        if self._updating_spec_models:
            return

        previous = self.current_spec_model_state()
        self._updating_spec_models = True
        table.blockSignals(True)
        try:
            table.setRowCount(0)
            context = context or owner._make_context()
            if context.brand != "Stulz" or not context.calc_path.exists():
                return

            from core.excel_reader import parse_stulz_calc

            data = parse_stulz_calc(context.calc_path, context.sheet_name)
            for item in data.items:
                row = table.rowCount()
                table.insertRow(row)

                enabled_item = QTableWidgetItem("")
                enabled_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                enabled = previous.get(item.name, (True, ""))[0]
                enabled_item.setCheckState(Qt.Checked if enabled else Qt.Unchecked)

                model_item = QTableWidgetItem(item.name)
                model_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

                default_qty = str(int(item.qty)) if float(item.qty).is_integer() else str(item.qty)
                qty_text = previous.get(item.name, (True, ""))[1] or default_qty
                qty_item = QTableWidgetItem(qty_text)
                qty_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)

                table.setItem(row, 0, enabled_item)
                table.setItem(row, 1, model_item)
                table.setItem(row, 2, qty_item)

            table.resizeRowsToContents()
        except Exception:
            table.setRowCount(0)
        finally:
            table.blockSignals(False)
            self._updating_spec_models = False
