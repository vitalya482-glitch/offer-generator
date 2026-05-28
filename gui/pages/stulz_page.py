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
    QPushButton,
    QMessageBox,
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
            "Модели берутся из выбранной папки спецификаций, а не из Excel КП. "
            "Можно отключать позиции и менять количество."
        )
        spec_hint.setObjectName("Hint")
        spec_hint.setWordWrap(True)
        spec_card.layout().addWidget(spec_hint)

        owner.spec_error_label = QLabel("")
        owner.spec_error_label.setStyleSheet("color: #dc2626; font-weight: 700;")
        owner.spec_error_label.setWordWrap(True)
        owner.spec_error_label.setVisible(False)
        spec_card.layout().addWidget(owner.spec_error_label)

        owner.spec_models_table = QTableWidget(0, 3)
        owner.spec_models_table.setHorizontalHeaderLabels(["Вкл", "Модель", "Кол-во"])
        owner.spec_models_table.verticalHeader().setVisible(False)
        owner.spec_models_table.setAlternatingRowColors(True)
        owner.spec_models_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        owner.spec_models_table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        owner.spec_models_table.setMinimumHeight(170)
        owner.spec_models_table.setColumnWidth(0, 52)
        owner.spec_models_table.setColumnWidth(2, 80)
        owner.spec_models_table.horizontalHeader().setStretchLastSection(False)
        owner.spec_models_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        spec_card.layout().addWidget(owner.spec_models_table)

        owner.spec_preview_button = QPushButton("Предпросмотр спецификаций")

        owner.spec_preview_button.setMinimumHeight(42)

        owner.spec_preview_button.setStyleSheet("""
        QPushButton {
            background-color: #dc2626;
            color: white;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 600;
            padding: 8px 18px;
        }

        QPushButton:hover {
            background-color: #b91c1c;
        }

        QPushButton:pressed {
            background-color: #991b1b;
        }
        """)
        
        owner.spec_preview_button.clicked.connect(self.open_spec_preview)
        spec_card.layout().addWidget(owner.spec_preview_button)
        spec_card.layout().addWidget(owner.status_label)
        bottom.addWidget(spec_card, stretch=1)

        layout.addLayout(bottom)


    def open_spec_preview(self) -> None:
        try:
            from brands.stulz import build_specification_blocks, load_calc
            from gui.spec_preview_dialog import SpecPreviewDialog

            context = self.owner._make_context()
            self.owner._validate_context(context)
            calc = load_calc(context)
            spec_blocks, warnings = build_specification_blocks(context, calc)
            dialog = SpecPreviewDialog(spec_blocks, warnings, self)
            dialog.exec()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть предпросмотр спецификаций:\n{exc}")

    def clear_spec_models(self) -> None:
        self.owner.spec_models_table.setRowCount(0)
        if hasattr(self.owner, "spec_error_label"):
            self.owner.spec_error_label.setVisible(False)
            self.owner.spec_error_label.setText("")
        if hasattr(self.owner, "spec_preview_button"):
            self.owner.spec_preview_button.setEnabled(False)

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
        """Return enabled specification models from the table.

        The UI shows one row per model. If a user or an older saved state still
        leaves duplicate rows in the table, this method merges them before the
        context is passed to the offer generator. This prevents duplicate Excel
        positions from being silently reduced to the first quantity.
        """
        merged: dict[str, dict[str, object]] = {}
        order: list[str] = []
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

            if model not in merged:
                merged[model] = {"enabled": enabled, "model": model, "qty_value": 0.0}
                order.append(model)

            # If at least one duplicate row is enabled, keep the model enabled;
            # quantities from enabled duplicate rows are summed.
            merged[model]["enabled"] = bool(merged[model].get("enabled")) or enabled
            if enabled:
                merged[model]["qty_value"] = float(merged[model].get("qty_value", 0.0)) + qty

        result: list[dict[str, object]] = []
        for model in order:
            row = merged[model]
            qty = float(row.get("qty_value", 0.0))
            qty_text = str(int(qty)) if qty.is_integer() else str(qty)
            result.append({"enabled": row.get("enabled", True), "model": model, "qty": qty_text, "qty_value": qty})
        return result

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
            if context.brand != "Stulz":
                return

            from core.stulz_specification import list_stulz_specification_models

            spec_models = list_stulz_specification_models(context.pdf_dir)

            if hasattr(owner, "spec_error_label"):
                if spec_models:
                    owner.spec_error_label.setVisible(False)
                    owner.spec_error_label.setText("")
                else:
                    owner.spec_error_label.setText("Спецификации не найдены")
                    owner.spec_error_label.setVisible(True)

            if hasattr(owner, "spec_preview_button"):
                owner.spec_preview_button.setEnabled(bool(spec_models))

            for model, default_qty_value in spec_models:
                row = table.rowCount()
                table.insertRow(row)

                enabled_item = QTableWidgetItem("")
                enabled_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                enabled = previous.get(model, (True, ""))[0]
                enabled_item.setCheckState(Qt.Checked if enabled else Qt.Unchecked)

                model_item = QTableWidgetItem(model)
                model_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

                default_qty = str(int(default_qty_value)) if float(default_qty_value).is_integer() else str(default_qty_value)
                qty_text = previous.get(model, (True, ""))[1] or default_qty
                qty_item = QTableWidgetItem(qty_text)
                qty_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)

                table.setItem(row, 0, enabled_item)
                table.setItem(row, 1, model_item)
                table.setItem(row, 2, qty_item)

            table.resizeRowsToContents()
        except Exception:
            table.setRowCount(0)
            if hasattr(owner, "spec_error_label"):
                owner.spec_error_label.setText("Спецификации не найдены")
                owner.spec_error_label.setVisible(True)
            if hasattr(owner, "spec_preview_button"):
                owner.spec_preview_button.setEnabled(False)
        finally:
            table.blockSignals(False)
            self._updating_spec_models = False
