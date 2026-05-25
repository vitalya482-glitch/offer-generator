from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class ReferenceTableDialog(QDialog):
    """Editable reference table for Stulz options and WinPlan labels."""

    def __init__(
        self,
        owner,
        *,
        title: str,
        columns: list[tuple[str, str]],
        load_rows: Callable[[], list[dict[str, str]]],
        save_rows: Callable[[list[dict[str, str]]], None],
        import_rows: Callable[[Path], list[dict[str, str]]] | None = None,
        untranslated_key: str | None = None,
    ) -> None:
        super().__init__(owner)
        self.owner = owner
        self.columns = columns
        self.load_rows = load_rows
        self.save_rows = save_rows
        self.import_rows = import_rows
        self.untranslated_key = untranslated_key
        self.rows: list[dict[str, str]] = [dict(row) for row in load_rows()]

        self.setWindowTitle(title)
        self.resize(1100, 680)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        top = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по таблице")
        self.search_edit.textChanged.connect(self._apply_filter)
        top.addWidget(QLabel("Поиск:"))
        top.addWidget(self.search_edit, stretch=1)

        self.only_untranslated_btn = QPushButton("Без перевода")
        self.only_untranslated_btn.setCheckable(True)
        self.only_untranslated_btn.clicked.connect(self._apply_filter)
        top.addWidget(self.only_untranslated_btn)
        layout.addLayout(top)

        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        layout.addWidget(self.info_label)

        self.table = QTableWidget(0, len(columns))
        self.table.setHorizontalHeaderLabels([header for header, _ in columns])
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        self.table.verticalHeader().setDefaultSectionSize(48)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, stretch=1)

        buttons = QHBoxLayout()
        add_btn = QPushButton("Добавить строку")
        add_btn.clicked.connect(self._add_row)
        buttons.addWidget(add_btn)

        delete_btn = QPushButton("Удалить строку")
        delete_btn.clicked.connect(self._delete_selected_rows)
        buttons.addWidget(delete_btn)

        if import_rows is not None:
            import_btn = QPushButton("Импорт из StulzMacros")
            import_btn.clicked.connect(self._import_from_file)
            buttons.addWidget(import_btn)

        buttons.addStretch(1)
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self._save)
        buttons.addWidget(save_btn)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.reject)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)

        self._fill_table(self.rows)
        self._update_info()

    def _is_untranslated(self, row: dict[str, str]) -> bool:
        if not self.untranslated_key:
            return False
        return not str(row.get(self.untranslated_key, "")).strip()

    def _visible_rows(self) -> list[dict[str, str]]:
        text = self.search_edit.text().strip().lower()
        only_untranslated = self.only_untranslated_btn.isChecked()
        result = []
        for row in self.rows:
            if only_untranslated and not self._is_untranslated(row):
                continue
            if text and text not in " ".join(str(value).lower() for value in row.values()):
                continue
            result.append(row)
        return result

    def _fill_table(self, rows: list[dict[str, str]]) -> None:
        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, (_, key) in enumerate(self.columns):
                item = QTableWidgetItem(str(row.get(key, "")))
                item.setData(Qt.UserRole, row)
                self.table.setItem(row_index, col_index, item)
        self.table.resizeColumnsToContents()
        self.table.blockSignals(False)

    def _sync_visible_table_to_rows(self) -> None:
        # Persist edits from currently visible rows back to source dictionaries.
        for row_index in range(self.table.rowCount()):
            first_item = self.table.item(row_index, 0)
            if first_item is None:
                continue
            source_row = first_item.data(Qt.UserRole)
            if not isinstance(source_row, dict):
                continue
            for col_index, (_, key) in enumerate(self.columns):
                item = self.table.item(row_index, col_index)
                source_row[key] = item.text().strip() if item else ""

    def _apply_filter(self) -> None:
        self._sync_visible_table_to_rows()
        self._fill_table(self._visible_rows())
        self._update_info()

    def _add_row(self) -> None:
        self._sync_visible_table_to_rows()
        row = {key: "" for _, key in self.columns}
        self.rows.append(row)
        self._fill_table(self._visible_rows())
        self._update_info()

    def _delete_selected_rows(self) -> None:
        selected = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not selected:
            return
        self._sync_visible_table_to_rows()
        visible = self._visible_rows()
        for row_index in selected:
            if 0 <= row_index < len(visible) and visible[row_index] in self.rows:
                self.rows.remove(visible[row_index])
        self._fill_table(self._visible_rows())
        self._update_info()

    def _import_from_file(self) -> None:
        if self.import_rows is None:
            return
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл StulzMacros", "", "Excel Macro (*.xlsm);;Excel (*.xlsx)")
        if not path:
            return
        try:
            imported = self.import_rows(Path(path))
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка импорта", str(exc))
            return
        self._sync_visible_table_to_rows()
        self.rows = imported
        self.search_edit.clear()
        self.only_untranslated_btn.setChecked(False)
        self._fill_table(self.rows)
        self._update_info()
        QMessageBox.information(self, "Импорт", f"Импортировано строк: {len(imported)}")

    def _save(self) -> None:
        self._sync_visible_table_to_rows()
        cleaned: list[dict[str, str]] = []
        for row in self.rows:
            normalized = {key: str(row.get(key, "")).strip() for _, key in self.columns}
            if any(normalized.values()):
                cleaned.append(normalized)
        try:
            self.save_rows(cleaned)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка сохранения", str(exc))
            return
        self.rows = cleaned
        self._fill_table(self._visible_rows())
        self._update_info()
        QMessageBox.information(self, "Сохранено", "Справочник сохранен.")

    def _update_info(self) -> None:
        untranslated = sum(1 for row in self.rows if self._is_untranslated(row))
        visible = len(self._visible_rows())
        text = f"Всего строк: {len(self.rows)}. Показано: {visible}."
        if self.untranslated_key:
            text += f" Без перевода: {untranslated}."
        self.info_label.setText(text)
