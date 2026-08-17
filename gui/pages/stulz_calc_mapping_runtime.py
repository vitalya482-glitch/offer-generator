from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidgetItem,
)

from core.stulz_spec_catalog import discover_stulz_spec_entries
from gui.path_helpers import infer_specifications_dir


_TRANSLIT = str.maketrans({
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
})


def _plain(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\u00a0", " ")).strip()


def _norm(value: object) -> str:
    text = _plain(value).lower().translate(_TRANSLIT)
    return re.sub(r"[^a-z0-9]+", "", text)


def _model_key(value: object) -> str:
    return _norm(value)


def _format_qty(value: object) -> str:
    try:
        number = float(value)
    except Exception:
        return _plain(value)
    if number.is_integer():
        return str(int(number))
    return (f"{number:.3f}").rstrip("0").rstrip(".").replace(".", ",")


def _calc_key(index: int, item: object) -> str:
    model = _plain(getattr(item, "name", ""))
    legend = _plain(getattr(item, "legend", ""))
    return f"calc:{index}:{_norm(model)}:{_norm(legend)}"


def _entry_dict(entry: object) -> dict[str, object]:
    return {
        "key": str(getattr(entry, "key", "") or ""),
        "model": str(getattr(entry, "model", "") or ""),
        "qty": getattr(entry, "quantity", 1),
        "source_dir": str(getattr(entry, "source_dir", "") or ""),
        "source_label": str(getattr(entry, "source_label", "") or ""),
        "calc_pdf": str(getattr(entry, "calc_pdf", "") or ""),
    }


def _folder_name(entry: dict[str, object]) -> str:
    source = _plain(entry.get("source_dir"))
    return Path(source).name if source else ""


def _match_score(item: object, entry: dict[str, object]) -> int:
    item_model = _model_key(getattr(item, "name", ""))
    spec_model = _model_key(entry.get("model"))
    if not item_model or item_model != spec_model:
        return -100_000

    score = 1000
    legend = _norm(getattr(item, "legend", ""))
    source = _norm(f"{entry.get('source_dir', '')} {entry.get('source_label', '')}")
    if legend and source:
        if legend in source:
            score += 5000
        else:
            score -= 150

    try:
        calc_qty = float(getattr(item, "qty", 0) or 0)
        spec_qty = float(entry.get("qty", 0) or 0)
        if calc_qty > 0 and abs(calc_qty - spec_qty) < 1e-6:
            score += 250
    except Exception:
        pass
    return score


def _discover_for_context(widget: object, context: object) -> list[dict[str, object]]:
    entries = [_entry_dict(entry) for entry in discover_stulz_spec_entries(getattr(context, "pdf_dir", None))]
    if entries:
        return entries

    project_dir = Path(getattr(context, "project_dir", "") or "")
    if not project_dir.exists():
        return []

    candidates = [Path(infer_specifications_dir(str(project_dir))), project_dir]
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).lower()
        if not key or key in seen or not candidate.exists():
            continue
        seen.add(key)
        found = [_entry_dict(entry) for entry in discover_stulz_spec_entries(candidate)]
        if found:
            try:
                context.pdf_dir = candidate
                widget._set_spec_dir_path(str(candidate))
            except Exception:
                pass
            return found
    return []


def _row_metadata(widget: object, item: QTableWidgetItem | None) -> dict[str, Any]:
    if item is None:
        return {}
    raw = item.data(Qt.UserRole)
    return dict(raw) if isinstance(raw, dict) else {}


def current_spec_model_state(self) -> dict[str, dict[str, Any]]:
    state: dict[str, dict[str, Any]] = {}
    table = self.spec_models_table
    for row in range(table.rowCount()):
        item = table.item(row, 0)
        meta = _row_metadata(self, item)
        calc_key = _plain(meta.get("calc_key"))
        if calc_key:
            state[calc_key] = meta
    return state


def selected_spec_models(self) -> list[dict[str, object]]:
    """Return Calc-position rows with their concrete physical specification mapping."""
    rows: list[dict[str, object]] = []
    table = self.spec_models_table
    for row in range(table.rowCount()):
        meta = _row_metadata(self, table.item(row, 0))
        model = _plain(meta.get("model"))
        source_dir = _plain(meta.get("source_dir"))
        if not model:
            continue
        qty = meta.get("qty_value", 0)
        rows.append({
            "enabled": bool(source_dir),
            "model": model,
            "qty": _format_qty(qty),
            "qty_value": qty,
            "legend": _plain(meta.get("legend")),
            "legend_enabled": True,
            "key": _plain(meta.get("spec_key")),
            "calc_key": _plain(meta.get("calc_key")),
            "source_dir": source_dir,
            "source_label": _plain(meta.get("source_label")),
            "calc_pdf": _plain(meta.get("calc_pdf")),
            "mapping_mode": _plain(meta.get("mapping_mode")),
        })
    return rows


def _ensure_calc_mapping_layout(self) -> None:
    if getattr(self, "_stulz_calc_mapping_layout_installed", False):
        return

    preview_card = self.preview.parentWidget()
    spec_card = self.spec_models_table.parentWidget()
    root = self.layout()
    if preview_card is None or spec_card is None or root is None:
        return

    # Rename the calculation check card. The title may already live inside the
    # refresh-button row, so search descendants instead of assuming one layout index.
    for label in preview_card.findChildren(QLabel):
        if label.text().strip() == "Проверка данных":
            label.setText("Проверка данных расчёта Calc")
            break

    # Put both cards on separate full-width rows: Calc check first, mappings below.
    old_layout = None
    old_index = -1
    for index in range(root.count()):
        child = root.itemAt(index).layout()
        if child is not None and (child.indexOf(preview_card) >= 0 or child.indexOf(spec_card) >= 0):
            old_layout = child
            old_index = index
            break

    if old_layout is not None:
        old_layout.removeWidget(preview_card)
        old_layout.removeWidget(spec_card)
        insert_at = old_index if old_index >= 0 else root.count()
        root.insertWidget(insert_at, preview_card)
        root.insertWidget(insert_at + 1, spec_card)

    self.preview.setMinimumHeight(300)
    self.spec_models_table.setMinimumHeight(190)

    for label in spec_card.findChildren(QLabel):
        text = label.text().strip()
        if text.startswith("Модели берутся из выбранной папки"):
            label.setText(
                "Для каждой позиции из расчёта Calc автоматически подбирается физическая папка спецификации. "
                "Если соответствие не найдено или выбрано неверно — выделите строку и назначьте папку вручную."
            )
            label.setWordWrap(True)
            break

    setattr(self, "_stulz_calc_mapping_layout_installed", True)


def _ensure_manual_spec_controls(self) -> None:
    """Use per-Calc-row manual folder assignment instead of a global PDF picker."""
    if getattr(self, "_stulz_manual_controls_installed", False):
        manual_btn = getattr(self, "manual_spec_button", None)
        auto_btn = getattr(self, "auto_spec_button", None)
        mode_label = getattr(self, "manual_spec_mode_label", None)
        if manual_btn is not None:
            try:
                manual_btn.clicked.disconnect()
            except Exception:
                pass
            manual_btn.setText("Выбрать для позиции")
            manual_btn.setToolTip("Назначить папку спецификации для выделенной позиции Calc")
            manual_btn.clicked.connect(self.browse_manual_spec_files)
        if auto_btn is not None:
            try:
                auto_btn.clicked.disconnect()
            except Exception:
                pass
            auto_btn.setText("Автопоиск")
            auto_btn.setToolTip("Заново автоматически сопоставить все позиции Calc со спецификациями")
            auto_btn.clicked.connect(self.clear_manual_spec_files)
        if mode_label is not None:
            mode_label.setText("Позиции: Calc → папка спецификации")
        return

    parent = self.spec_models_table.parentWidget()
    layout = parent.layout() if parent is not None else None
    if layout is None:
        return

    controls = QHBoxLayout()
    manual_btn = QPushButton("Выбрать для позиции")
    auto_btn = QPushButton("Автопоиск")
    mode_label = QLabel("Позиции: Calc → папка спецификации")
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


def _update_manual_mode_label(self) -> None:
    label = getattr(self, "manual_spec_mode_label", None)
    if label is not None:
        label.setText("Позиции: Calc → папка спецификации")


def browse_manual_spec_files(self) -> None:
    """Assign one physical specification folder to the selected Calc position."""
    table = self.spec_models_table
    row = table.currentRow()
    if row < 0:
        QMessageBox.information(self, "STULZ", "Сначала выделите позицию Calc в таблице.")
        return

    position_item = table.item(row, 0)
    meta = _row_metadata(self, position_item)
    model = _plain(meta.get("model"))
    start_dir = _plain(meta.get("source_dir")) or self.spec_path_text() or self.project_path_text()
    folder = QFileDialog.getExistingDirectory(self, f"Выберите папку спецификации для {model}", start_dir)
    if not folder:
        return

    entries = [_entry_dict(entry) for entry in discover_stulz_spec_entries(folder)]
    matching = [entry for entry in entries if _model_key(entry.get("model")) == _model_key(model)]
    if not matching:
        QMessageBox.warning(
            self,
            "STULZ",
            f"В выбранной папке не найден читаемый Calc.pdf для модели {model}.\n\n{folder}",
        )
        return

    chosen = matching[0]
    meta.update({
        "spec_key": _plain(chosen.get("key")),
        "source_dir": _plain(chosen.get("source_dir")),
        "source_label": _plain(chosen.get("source_label")),
        "calc_pdf": _plain(chosen.get("calc_pdf")),
        "mapping_mode": "manual",
    })
    position_item.setData(Qt.UserRole, meta)
    table.item(row, 2).setText(_folder_name(chosen))
    table.item(row, 2).setToolTip(_plain(chosen.get("source_dir")))
    table.item(row, 3).setText("✓ Вручную")
    self.status_label.setText(f"Для позиции «{position_item.text()}» назначена папка спецификации вручную.")


def clear_manual_spec_files(self) -> None:
    setattr(self, "_stulz_force_auto_spec_rematch", True)
    self.status_label.setText("Выполняется повторный автопоиск спецификаций…")
    self.refresh_preview()


def refresh_spec_models(self, context=None) -> None:
    self._ensure_calc_mapping_layout()
    self._ensure_manual_spec_controls()
    self._ensure_description_options_controls()

    table = self.spec_models_table
    if self._updating_spec_models:
        return

    previous = current_spec_model_state(self)
    force_auto = bool(getattr(self, "_stulz_force_auto_spec_rematch", False))
    setattr(self, "_stulz_force_auto_spec_rematch", False)

    self._updating_spec_models = True
    table.blockSignals(True)
    try:
        context = context or self.make_context()
        if context.brand != self.brand_name:
            return

        # Use the legend-aware loader directly; calling the registry here while
        # the top-level STULZ runtime is importing could recurse into itself.
        from brands.stulz_legend_runtime import load_calc

        calc = load_calc(context)
        items = list(getattr(calc, "items", []) or [])
        specs = _discover_for_context(self, context)

        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Позиция из Calc", "Кол-во", "Папка спецификации", "Статус"])
        table.setRowCount(0)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        used_spec_keys: set[str] = set()
        matched_count = 0

        for index, item in enumerate(items):
            calc_key = _calc_key(index, item)
            model = _plain(getattr(item, "name", ""))
            legend = _plain(getattr(item, "legend", ""))
            qty = getattr(item, "qty", 0)

            chosen: dict[str, object] | None = None
            mapping_mode = "auto"

            old = previous.get(calc_key) if not force_auto else None
            if old and _plain(old.get("source_dir")):
                old_spec_key = _plain(old.get("spec_key"))
                # Preserve a manual assignment and any still-valid prior auto mapping.
                chosen = next((entry for entry in specs if _plain(entry.get("key")) == old_spec_key), None)
                if chosen is None and Path(_plain(old.get("source_dir"))).exists():
                    chosen = {
                        "key": old_spec_key,
                        "model": model,
                        "qty": qty,
                        "source_dir": old.get("source_dir", ""),
                        "source_label": old.get("source_label", ""),
                        "calc_pdf": old.get("calc_pdf", ""),
                    }
                mapping_mode = _plain(old.get("mapping_mode")) or "auto"

            if chosen is None:
                candidates = [
                    entry for entry in specs
                    if _plain(entry.get("key")) not in used_spec_keys
                    and _model_key(entry.get("model")) == _model_key(model)
                ]
                if candidates:
                    candidates.sort(key=lambda entry: _match_score(item, entry), reverse=True)
                    chosen = candidates[0]
                    mapping_mode = "auto"

            if chosen is not None:
                spec_key = _plain(chosen.get("key"))
                if spec_key:
                    used_spec_keys.add(spec_key)
                matched_count += 1

            row = table.rowCount()
            table.insertRow(row)
            caption = f"{legend} — {model}" if legend else model
            position_item = QTableWidgetItem(caption)
            position_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            meta = {
                "calc_key": calc_key,
                "model": model,
                "legend": legend,
                "qty_value": qty,
                "spec_key": _plain(chosen.get("key")) if chosen else "",
                "source_dir": _plain(chosen.get("source_dir")) if chosen else "",
                "source_label": _plain(chosen.get("source_label")) if chosen else "",
                "calc_pdf": _plain(chosen.get("calc_pdf")) if chosen else "",
                "mapping_mode": mapping_mode if chosen else "",
            }
            position_item.setData(Qt.UserRole, meta)

            qty_item = QTableWidgetItem(_format_qty(qty))
            qty_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            folder_item = QTableWidgetItem(_folder_name(chosen) if chosen else "Не выбрана")
            folder_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            if chosen:
                folder_item.setToolTip(_plain(chosen.get("source_dir")))

            status_text = "✓ Вручную" if chosen and mapping_mode == "manual" else ("✓ Авто" if chosen else "⚠ Не найдена")
            status_item = QTableWidgetItem(status_text)
            status_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            table.setItem(row, 0, position_item)
            table.setItem(row, 1, qty_item)
            table.setItem(row, 2, folder_item)
            table.setItem(row, 3, status_item)

        table.resizeRowsToContents()
        total = len(items)
        if total:
            if matched_count == total:
                self.status_label.setText(f"Автосопоставление: {matched_count} из {total} позиций Calc связаны со спецификациями.")
            else:
                self.status_label.setText(
                    f"Спецификации сопоставлены для {matched_count} из {total} позиций Calc. "
                    "Для строк со статусом «Не найдена» выберите папку вручную."
                )
    except Exception as exc:
        table.setRowCount(0)
        self.status_label.setText(f"Не удалось сопоставить позиции Calc со спецификациями: {exc}")
    finally:
        table.blockSignals(False)
        self._updating_spec_models = False


def _install() -> None:
    page_module = sys.modules.get("gui.pages.stulz_page")
    page_class = getattr(page_module, "StulzPage", None) if page_module is not None else None
    if page_class is None:
        return

    for name, function in {
        "current_spec_model_state": current_spec_model_state,
        "selected_spec_models": selected_spec_models,
        "_ensure_calc_mapping_layout": _ensure_calc_mapping_layout,
        "_ensure_manual_spec_controls": _ensure_manual_spec_controls,
        "_update_manual_mode_label": _update_manual_mode_label,
        "browse_manual_spec_files": browse_manual_spec_files,
        "clear_manual_spec_files": clear_manual_spec_files,
        "refresh_spec_models": refresh_spec_models,
    }.items():
        setattr(page_class, name, function)

    try:
        from PySide6.QtWidgets import QApplication
        for widget in QApplication.allWidgets():
            if isinstance(widget, page_class):
                widget._ensure_calc_mapping_layout()
                widget._ensure_manual_spec_controls()
                widget.refresh_spec_models()
    except Exception:
        pass


_install()
