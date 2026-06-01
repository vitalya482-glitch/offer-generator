from __future__ import annotations

from core.models import OfferContext
from pathlib import Path

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
            "Количество берется из Calc.pdf. Можно отключать позиции и менять количество."
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

    def _format_model_for_table(self, model: str) -> str:
        # In the list we show compact model codes, but the source is still Calc.pdf.
        return (model or "").replace(" ", "").strip()

    def _format_qty_for_table(self, value: object) -> str:
        try:
            number = float(str(value).replace(",", "."))
            return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
        except Exception:
            return str(value or "")

    def _scan_calc_pdf_models(self, spec_dir: str | Path) -> list[dict[str, object]]:
        from core.pdf_parsers.stulz_calc_pdf import parse_stulz_calc_totals

        root = Path(spec_dir)
        if not root.exists():
            return []

        grouped: dict[str, dict[str, object]] = {}
        for pdf_path in sorted(root.rglob("*.pdf")):
            if "calc" not in pdf_path.stem.lower():
                continue
            try:
                totals = parse_stulz_calc_totals(pdf_path)
            except Exception:
                continue

            # Important: model is accepted only from the equipment row inside Calc.pdf.
            # We intentionally do not use project/file/folder names here, because they
            # can contain misleading codes like ASD221A.
            model = self._format_model_for_table(totals.model)
            if not model:
                continue

            qty = totals.quantity if totals.quantity not in (None, "") else 1
            try:
                qty_value = float(qty)
            except Exception:
                qty_value = 1.0

            if model not in grouped:
                grouped[model] = {"model": model, "qty": 0.0, "files": []}
            grouped[model]["qty"] = float(grouped[model]["qty"]) + qty_value
            grouped[model]["files"].append(str(pdf_path))

        return list(grouped.values())

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

            models = self._scan_calc_pdf_models(context.pdf_dir)
            if not models:
                return

            for entry in models:
                model = str(entry.get("model") or "").strip()
                if not model:
                    continue

                row = table.rowCount()
                table.insertRow(row)

                enabled_item = QTableWidgetItem("")
                enabled_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                enabled = previous.get(model, (True, ""))[0]
                enabled_item.setCheckState(Qt.Checked if enabled else Qt.Unchecked)

                model_item = QTableWidgetItem(model)
                model_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                files = entry.get("files") or []
                if files:
                    model_item.setToolTip("\n".join(files))

                default_qty = self._format_qty_for_table(entry.get("qty"))
                qty_text = previous.get(model, (True, ""))[1] or default_qty
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
