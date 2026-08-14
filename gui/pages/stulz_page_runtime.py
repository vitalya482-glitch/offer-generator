from __future__ import annotations

import re
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton, QTableWidgetItem

from core.models import OfferContext
from core.stulz_spec_catalog import discover_stulz_spec_entries, load_stulz_spec_entry
from gui.path_helpers import infer_specifications_dir
from gui.pages.stulz_page import StulzPage as _BaseStulzPage


class StulzPage(_BaseStulzPage):
    """STULZ page with one GUI row per physical Calc.pdf specification set."""

    def _row_metadata(self, model_item: QTableWidgetItem | None) -> dict[str, str]:
        if not model_item:
            return {}
        raw = model_item.data(Qt.UserRole)
        return dict(raw) if isinstance(raw, dict) else {}

    def _row_key(self, model_item: QTableWidgetItem | None) -> str:
        meta = self._row_metadata(model_item)
        return str(meta.get("key") or (model_item.text().strip() if model_item else ""))

    def current_spec_model_state(self) -> dict[str, tuple[bool, str]]:
        state: dict[str, tuple[bool, str]] = {}
        table = self.spec_models_table
        for row in range(table.rowCount()):
            enabled_item = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            key = self._row_key(model_item)
            if not key:
                continue
            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
            qty = qty_item.text().strip() if qty_item else ""
            state[key] = (enabled, qty)
        return state

    def selected_spec_models(self) -> list[dict[str, object]]:
        models: list[dict[str, object]] = []
        table = self.spec_models_table
        for row in range(table.rowCount()):
            enabled_item = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            if not model_item:
                continue

            meta = self._row_metadata(model_item)
            model = str(meta.get("model") or model_item.text()).strip()
            if not model:
                continue

            enabled = enabled_item.checkState() == Qt.Checked if enabled_item else True
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
                    "key": str(meta.get("key") or ""),
                    "source_dir": str(meta.get("source_dir") or ""),
                    "source_label": str(meta.get("source_label") or ""),
                    "calc_pdf": str(meta.get("calc_pdf") or ""),
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

        table = self.spec_models_table
        if self._updating_spec_models:
            return

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
                return

            for entry in models:
                model = str(entry.get("model") or "").strip()
                key = str(entry.get("key") or "").strip()
                if not model:
                    continue

                row = table.rowCount()
                table.insertRow(row)

                enabled_item = QTableWidgetItem("")
                enabled_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                enabled = previous.get(key, (True, ""))[0]
                enabled_item.setCheckState(Qt.Checked if enabled else Qt.Unchecked)

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
                    },
                )
                files = entry.get("files") or []
                tooltip_parts = [str(entry.get("source_label") or "")]
                tooltip_parts.extend(str(path) for path in files)
                model_item.setToolTip("\n".join(part for part in tooltip_parts if part))

                default_qty = self._format_qty_for_table(entry.get("qty"))
                qty_text = previous.get(key, (True, ""))[1] or default_qty
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
