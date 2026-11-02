from __future__ import annotations

import re
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.models import OfferContext
from core.stulz_spec_catalog import discover_stulz_spec_entries, load_stulz_spec_entry
from gui.path_helpers import infer_specifications_dir
from gui.pages.stulz_page import StulzPage as _BaseStulzPage


DESCRIPTION_OPTION_LABELS: dict[str, str] = {
    "stulz_unit": "Прецизионный кондиционер Stulz",
    "cooling_capacity": "Хладопроизводительность",
    "unit_dimensions": "Размеры внутреннего блока (Ш×Г×В)",
    "condenser": "Наружный блок (конденсор)",
}


class StulzPage(_BaseStulzPage):
    """STULZ page with one GUI row per physical Calc.pdf specification set."""

    def _row_metadata(self, model_item: QTableWidgetItem | None) -> dict[str, object]:
        if not model_item:
            return {}
        raw = model_item.data(Qt.UserRole)
        return dict(raw) if isinstance(raw, dict) else {}

    def _row_key(self, model_item: QTableWidgetItem | None) -> str:
        meta = self._row_metadata(model_item)
        return str(meta.get("key") or (model_item.text().strip() if model_item else ""))

    def current_spec_model_state(self) -> dict[str, tuple[bool, str, bool]]:
        """Preserve row enable/quantity/legend switches across rescans."""

        state: dict[str, tuple[bool, str, bool]] = {}
        table = self.spec_models_table
        for row in range(table.rowCount()):
            enabled_item = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            legend_item = table.item(row, 3) if table.columnCount() > 3 else None
            key = self._row_key(model_item)
            if not key:
                continue
            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
            qty = qty_item.text().strip() if qty_item else ""
            legend_enabled = legend_item.checkState() == Qt.Checked if legend_item else True
            state[key] = (enabled, qty, legend_enabled)
        return state

    def selected_spec_models(self) -> list[dict[str, object]]:
        models: list[dict[str, object]] = []
        table = self.spec_models_table
        for row in range(table.rowCount()):
            enabled_item = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            legend_item = table.item(row, 3) if table.columnCount() > 3 else None
            if not model_item:
                continue

            meta = self._row_metadata(model_item)
            model = str(meta.get("model") or model_item.text()).strip()
            if not model:
                continue

            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
            legend_enabled = legend_item.checkState() == Qt.Checked if legend_item else True
            qty_text = qty_item.text().strip() if qty_item else ""
            try:
                qty = float(qty_text.replace(",", ".")) if qty_text else 0.0
            except Exception:
                qty = 0.0

            models.append(
                {
                    "enabled": enabled,
                    "model": model,
                    "qty": qty_text,
                    "qty_value": qty,
                    "legend_enabled": legend_enabled,
                    "key": str(meta.get("key") or ""),
                    "source_dir": str(meta.get("source_dir") or ""),
                    "source_label": str(meta.get("source_label") or ""),
                    "calc_pdf": str(meta.get("calc_pdf") or ""),
                    "unit_list_price": meta.get("unit_list_price"),
                    "total_list_price": meta.get("total_list_price"),
                    "unit_purchase_price": meta.get("unit_purchase_price"),
                    "total_purchase_price": meta.get("total_purchase_price"),
                    "currency": str(meta.get("currency") or "").upper(),
                }
            )
        return models

    def _scan_calc_pdf_models(self, spec_dir: str | Path) -> list[dict[str, object]]:
        """Return every physical Calc.pdf separately; never group by model name."""

        return [
            {
                "key": entry.key,
                "model": entry.model,
                "qty": entry.quantity,
                "files": [str(entry.calc_pdf)],
                "calc_pdf": str(entry.calc_pdf),
                "source_dir": str(entry.source_dir),
                "source_label": entry.source_label,
                "unit_list_price": entry.unit_list_price,
                "total_list_price": entry.total_list_price,
                "unit_purchase_price": entry.unit_purchase_price,
                "total_purchase_price": entry.total_purchase_price,
                "currency": entry.currency,
            }
            for entry in discover_stulz_spec_entries(spec_dir)
        ]

    def _manual_spec_files(self) -> list[str]:
        raw = getattr(self, "_stulz_manual_spec_files", [])
        return [str(path) for path in raw if str(path).strip()]

    def _manual_spec_models(self) -> list[dict[str, object]]:
        models: list[dict[str, object]] = []
        for path_text in self._manual_spec_files():
            entry = load_stulz_spec_entry(path_text)
            if entry is None:
                continue
            models.append(
                {
                    "key": entry.key,
                    "model": entry.model,
                    "qty": entry.quantity,
                    "files": [str(entry.calc_pdf)],
                    "calc_pdf": str(entry.calc_pdf),
                    "source_dir": str(entry.source_dir),
                    "source_label": entry.source_label,
                    "unit_list_price": entry.unit_list_price,
                    "total_list_price": entry.total_list_price,
                    "unit_purchase_price": entry.unit_purchase_price,
                    "total_purchase_price": entry.total_purchase_price,
                    "currency": entry.currency,
                }
            )
        return models

    def _ensure_manual_spec_controls(self) -> None:
        """Inject manual Calc.pdf selection controls into the existing STULZ card."""

        if getattr(self, "_stulz_manual_controls_installed", False):
            return

        parent = self.spec_models_table.parentWidget()
        layout = parent.layout() if parent else None
        if layout is None:
            return

        controls = QHBoxLayout()
        manual_btn = QPushButton("Выбрать спецификации вручную")
        manual_btn.setToolTip("Выбрать один или несколько STULZ Calc.pdf вручную")
        auto_btn = QPushButton("Автопоиск")
        auto_btn.setToolTip("Сбросить ручной выбор и снова искать все Calc.pdf в папке спецификаций")
        mode_label = QLabel("Режим: автопоиск")
        mode_label.setObjectName("Hint")

        manual_btn.clicked.connect(self.browse_manual_spec_files)
        auto_btn.clicked.connect(self.clear_manual_spec_files)

        controls.addWidget(manual_btn)
        controls.addWidget(auto_btn)
        controls.addWidget(mode_label, stretch=1)

        index = layout.indexOf(self.spec_preview_button)
        if index >= 0 and hasattr(layout, "insertLayout"):
            layout.insertLayout(index, controls)
        else:
            layout.addLayout(controls)

        self.manual_spec_button = manual_btn
        self.auto_spec_button = auto_btn
        self.manual_spec_mode_label = mode_label
        self._stulz_manual_controls_installed = True
        self._update_manual_mode_label()

    def _ensure_description_options_controls(self) -> None:
        """Move offer-description switches from the spec dialog into data preview."""

        if getattr(self, "_stulz_description_controls_installed", False):
            return

        preview = getattr(self, "preview", None)
        card = preview.parentWidget() if preview is not None else None
        layout = card.layout() if card is not None else None
        if layout is None:
            return

        group = QGroupBox("Текст для включения в описание КП")
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(10, 8, 10, 8)
        group_layout.setSpacing(4)

        saved_options = self.description_options()
        checkboxes: dict[str, QCheckBox] = {}
        for key, label in DESCRIPTION_OPTION_LABELS.items():
            checkbox = QCheckBox(label)
            checkbox.setChecked(bool(saved_options.get(key, True)))
            checkbox.toggled.connect(lambda _checked, option_key=key: self._on_description_option_changed(option_key))
            checkboxes[key] = checkbox
            group_layout.addWidget(checkbox)

        # preview is already in the card layout. Adding the group now places the
        # switches below all main preview data, exactly where they are needed.
        layout.addWidget(group)

        self._stulz_description_group = group
        self._stulz_description_checkboxes = checkboxes
        self._stulz_description_controls_installed = True

    def _on_description_option_changed(self, _changed_key: str = "") -> None:
        checkboxes = getattr(self, "_stulz_description_checkboxes", {})
        if not checkboxes:
            return
        options = {key: checkbox.isChecked() for key, checkbox in checkboxes.items()}
        self._save_description_options(options)
        self.refresh_preview()

    def _update_manual_mode_label(self) -> None:
        label = getattr(self, "manual_spec_mode_label", None)
        if label is None:
            return
        count = len(self._manual_spec_files())
        label.setText(f"Режим: вручную ({count})" if count else "Режим: автопоиск")

    def browse_manual_spec_files(self) -> None:
        """Select concrete Calc.pdf files. Manual mode uses only these specifications."""

        start_dir = self.spec_path_text() or self.project_path_text()
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите STULZ Calc.pdf",
            start_dir,
            "PDF (*.pdf)",
        )
        if not paths:
            return

        valid: list[str] = []
        invalid: list[str] = []
        seen: set[str] = set()

        for path_text in paths:
            entry = load_stulz_spec_entry(path_text)
            if entry is None:
                invalid.append(Path(path_text).name)
                continue
            if entry.key in seen:
                continue
            seen.add(entry.key)
            valid.append(str(entry.calc_pdf))

        if not valid:
            QMessageBox.warning(
                self,
                "STULZ",
                "В выбранных PDF не удалось определить ни одной STULZ-спецификации.",
            )
            return

        self._stulz_manual_spec_files = valid
        self._update_manual_mode_label()
        self.status_label.setText(f"Ручной выбор спецификаций: {len(valid)}")
        self.refresh_preview()

        if invalid:
            QMessageBox.warning(
                self,
                "STULZ",
                "Некоторые выбранные PDF пропущены, потому что не удалось прочитать модель из Calc.pdf:\n"
                + "\n".join(invalid),
            )

    def clear_manual_spec_files(self) -> None:
        self._stulz_manual_spec_files = []
        self._update_manual_mode_label()
        self.status_label.setText("Ручной выбор сброшен. Используется автопоиск спецификаций.")
        self.refresh_preview()

    def open_spec_preview(self) -> None:
        """Create a Calc workbook from checked PDF model rows."""

        try:
            from core.stulz_calc_writer import StulzCalcWriter

            self.validate_calc_parameters()
            context = self.make_context()
            if not context.project_dir.exists():
                raise FileNotFoundError("Выберите существующую папку проекта.")
            if context.pdf_dir and not context.pdf_dir.exists():
                raise FileNotFoundError("Выберите существующую папку спецификаций.")
            parameters = self.calc_parameter_values()
            self.refresh_spec_models(context)
            selected_models = [
                row for row in self.selected_spec_models()
                if row.get("enabled", True) and str(row.get("model") or "").strip()
            ]
            if not selected_models:
                if getattr(self, "spec_models_table", None) is not None and self.spec_models_table.rowCount() == 0:
                    raise ValueError(
                        "В выбранной папке спецификаций не найдены Calc.pdf. "
                        "Проверьте поле «Папка спецификаций» или обновите её содержимое."
                    )
                raise ValueError("Отметьте хотя бы одну модель в таблице на вкладке «1. Собрать расчёт».")

            output_path = Path(parameters["Сохранить Calc"])
            if output_path.exists():
                answer = QMessageBox.question(
                    self,
                    "Перезаписать Calc?",
                    f"Файл уже существует:\n{output_path}\n\nПерезаписать его?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return

            self.spec_preview_button.setEnabled(False)
            self.spec_preview_button.setText("Формирую Calc…")
            from PySide6.QtWidgets import QApplication
            QApplication.processEvents()
            saved_path = StulzCalcWriter().write(
                template_path=parameters["Шаблон Calc"],
                sheet_name=parameters["Лист шаблона"],
                output_path=output_path,
                models=selected_models,
                currency=parameters["Валюта"],
                delivery_terms=parameters["Условия поставки"],
                delivery_point=parameters["Пункт поставки"],
                exchange_rate=parameters["Курс"],
                vat_percent=parameters["НДС, %"],
                margin_percent=parameters["Маржа, %"],
            )
            self.remember_values()
            self.status_label.setText(f"Calc сформирован: {saved_path.name}")
            QMessageBox.information(self, "Calc сформирован", f"Файл сохранён:\n{saved_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать Calc:\n{exc}")
        finally:
            self.spec_preview_button.setEnabled(True)
            self.spec_preview_button.setText("Сформировать расчёт")

    def _short_source_label(self, entry: dict[str, object]) -> str:
        source_dir = str(entry.get("source_dir") or "")
        folder = Path(source_dir).name if source_dir else ""

        # Common supplier folder form:
        # SAM Trade LLP@KCELL_ASR552AS_Almaty@ASR 552 AS
        match = re.search(r"_([^_@]+)@[^@]+$", folder)
        if match:
            return match.group(1)

        label = str(entry.get("source_label") or "").strip()
        return Path(label).name if label else ""

    def refresh_spec_models(self, context: OfferContext | None = None) -> None:
        self._ensure_manual_spec_controls()
        self._ensure_description_options_controls()

        table = self.spec_models_table
        if self._updating_spec_models:
            return

        # The separate Calc input table lives on workflow step 1. This table
        # remains the legacy specification selection used while preparing КП.
        if table.columnCount() < 4:
            table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Вкл", "Модель", "Кол-во", "Легенда"])
        table.setColumnWidth(0, 52)
        table.setColumnWidth(2, 80)
        table.setColumnWidth(3, 82)

        previous = self.current_spec_model_state()
        self._updating_spec_models = True
        table.blockSignals(True)
        try:
            table.setRowCount(0)
            context = context or self.make_context()
            if context.brand != self.brand_name:
                return

            manual_models = self._manual_spec_models()
            models = manual_models if self._manual_spec_files() else self._scan_calc_pdf_models(context.pdf_dir)

            if not self._manual_spec_files() and not models and context.project_dir.exists():
                fallback_dirs = [
                    Path(infer_specifications_dir(str(context.project_dir))),
                    context.project_dir,
                ]
                seen: set[str] = set()
                for fallback_dir in fallback_dirs:
                    fallback_key = str(fallback_dir.resolve()) if fallback_dir.exists() else str(fallback_dir)
                    if not fallback_key or fallback_key in seen:
                        continue
                    seen.add(fallback_key)
                    if context.pdf_dir and fallback_dir == context.pdf_dir:
                        continue
                    fallback_models = self._scan_calc_pdf_models(fallback_dir)
                    if fallback_models:
                        models = fallback_models
                        context.pdf_dir = fallback_dir
                        self._set_spec_dir_path(str(fallback_dir))
                        break

            self._update_manual_mode_label()

            if not models:
                self.status_label.setText(
                    "Calc.pdf не найдены. Проверьте папку спецификаций или выберите PDF вручную."
                )
                return

            for entry in models:
                model = str(entry.get("model") or "").strip()
                key = str(entry.get("key") or "").strip()
                if not model:
                    continue

                row = table.rowCount()
                table.insertRow(row)

                old_state = previous.get(key, (True, "", True))

                enabled_item = QTableWidgetItem("")
                enabled_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                enabled_item.setCheckState(Qt.Checked if old_state[0] else Qt.Unchecked)

                source_short = self._short_source_label(entry)
                caption = f"{model} · {source_short}" if source_short else model
                model_item = QTableWidgetItem(caption)
                model_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                model_item.setData(
                    Qt.UserRole,
                    {
                        "key": key,
                        "model": model,
                        "source_dir": str(entry.get("source_dir") or ""),
                        "source_label": str(entry.get("source_label") or ""),
                        "calc_pdf": str(entry.get("calc_pdf") or ""),
                        "unit_list_price": entry.get("unit_list_price"),
                        "total_list_price": entry.get("total_list_price"),
                        "unit_purchase_price": entry.get("unit_purchase_price"),
                        "total_purchase_price": entry.get("total_purchase_price"),
                        "currency": str(entry.get("currency") or "").upper(),
                    },
                )
                files = entry.get("files") or []
                tooltip_parts = [str(entry.get("source_label") or "")]
                tooltip_parts.extend(str(path) for path in files)
                model_item.setToolTip("\n".join(part for part in tooltip_parts if part))

                default_qty = self._format_qty_for_table(entry.get("qty"))
                qty_text = old_state[1] or default_qty
                qty_item = QTableWidgetItem(qty_text)
                qty_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)

                legend_item = QTableWidgetItem("")
                legend_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                legend_item.setCheckState(Qt.Checked if old_state[2] else Qt.Unchecked)
                legend_item.setToolTip(
                    "Добавлять легенду из Excel (например, город/объект) в заголовок этой спецификации"
                )

                table.setItem(row, 0, enabled_item)
                table.setItem(row, 1, model_item)
                table.setItem(row, 2, qty_item)
                table.setItem(row, 3, legend_item)

            table.resizeRowsToContents()
            self.status_label.setText(f"Найдено позиций: {table.rowCount()}. Отметьте нужные для Calc.")
        except Exception as exc:
            table.setRowCount(0)
            self.status_label.setText(f"Не удалось прочитать позиции Calc.pdf: {exc}")
        finally:
            table.blockSignals(False)
            self._updating_spec_models = False

    @staticmethod
    def _money_table_item(value: object) -> QTableWidgetItem:
        item = QTableWidgetItem("")
        if value is not None:
            try:
                item.setText(f"{float(value):,.2f}".replace(",", " ").replace(".", ","))
            except (TypeError, ValueError):
                item.setText(str(value))
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        return item

    def _on_spec_model_table_changed(self, item: QTableWidgetItem) -> None:
        if self._updating_spec_models:
            return
        row = item.row()
        if item.column() == 2:
            meta = self._row_metadata(self.spec_models_table.item(row, 1))
            try:
                quantity = float(item.text().strip().replace(",", "."))
            except (TypeError, ValueError):
                quantity = 0.0
            for column, price_key in ((4, "unit_list_price"), (6, "unit_purchase_price")):
                unit_price = meta.get(price_key)
                total = float(unit_price) * quantity if unit_price is not None else None
                self.spec_models_table.setItem(row, column, self._money_table_item(total))
        checked = sum(
            1 for index in range(self.spec_models_table.rowCount())
            if self.spec_models_table.item(index, 0)
            and self.spec_models_table.item(index, 0).checkState() == Qt.Checked
        )
        self.status_label.setText(
            f"Найдено позиций: {self.spec_models_table.rowCount()}. В Calc отмечено: {checked}."
        )
