from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.manager_profile import find_manager_in_project
from core.project_scanner import clear_scan_cache, scan_project_files
from core.riello_offer_word import (
    generate_riello_word_offer,
    preview_riello_offer_calc,
    read_riello_offer_calc,
)
from gui.path_helpers import extract_client_from_project_dir, infer_output_dir


class RielloOfferPage(QWidget):
    """Riello stage 2: read a finished Calc and create a tagged Word offer."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self.settings = owner.settings
        self._updating_path_display = False

        self.project_dir_path = self._saved("riello_offer/project_dir", self._saved("riello/project_dir", ""))
        self.output_dir_path = self._saved("riello_offer/output_dir", self._saved("riello/output_dir", ""))

        self.project_edit = QLineEdit(owner._display_dir(self.project_dir_path))
        self.project_edit.setToolTip(self.project_dir_path)
        self.client_edit = QLineEdit(self._saved("riello_offer/client", self._saved("riello/client", "ТОО Example")))
        self.calc_combo = QComboBox()
        self.calc_combo.setEditable(True)
        self.template_combo = QComboBox()
        self.template_combo.setEditable(True)
        self.output_edit = QLineEdit(owner._display_dir(self.output_dir_path))
        self.output_edit.setToolTip(self.output_dir_path)
        self.status_label = QLabel("Выберите готовый Riello Calc и Word-шаблон.")
        self.status_label.setObjectName("Hint")
        self.status_label.setWordWrap(True)

        self._load_saved_paths()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        header = QHBoxLayout()
        titles = QVBoxLayout()
        title = QLabel("Riello: формирование коммерческого предложения")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            "Готовый Calc является источником данных. Проверьте позиции и суммы, выберите Word-шаблон и сформируйте КП."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        titles.addWidget(title)
        titles.addWidget(subtitle)
        self.generate_btn = QPushButton("Сформировать КП")
        self.generate_btn.setObjectName("PrimaryButton")
        self.generate_btn.clicked.connect(self.generate)
        header.addLayout(titles, stretch=1)
        header.addWidget(self.generate_btn, alignment=Qt.AlignTop)
        layout.addLayout(header)

        project_card = owner._card("Папка проекта")
        project_grid = QGridLayout()
        project_card.layout().addLayout(project_grid)
        project_grid.setColumnStretch(1, 1)
        project_grid.setVerticalSpacing(12)
        project_grid.setHorizontalSpacing(10)
        owner._add_row(project_grid, 0, "Папка проекта", self.project_edit, "Выбрать", self.browse_project_dir)
        layout.addWidget(project_card)

        files_card = owner._card("Riello: файлы и параметры КП")
        files_grid = QGridLayout()
        files_card.layout().addLayout(files_grid)
        files_grid.setColumnStretch(1, 1)
        files_grid.setVerticalSpacing(12)
        files_grid.setHorizontalSpacing(10)
        owner._add_row(files_grid, 0, "Клиент", self.client_edit, None, None)
        owner._add_row(files_grid, 1, "Calc", self.calc_combo, "Выбрать", self.browse_calc_file)
        owner._add_row(files_grid, 2, "Word-шаблон", self.template_combo, "Выбрать", self.browse_template_file)
        owner._add_row(files_grid, 3, "Папка результата", self.output_edit, "Выбрать", self.browse_output_dir)
        layout.addWidget(files_card)

        preview_card = owner._card("Проверка данных расчёта Calc")
        preview_top = QHBoxLayout()
        preview_top.addStretch(1)
        self.updated_label = QLabel("")
        self.updated_label.setObjectName("Hint")
        self.refresh_btn = QPushButton("↻ Обновить")
        self.refresh_btn.setObjectName("GhostButton")
        self.refresh_btn.clicked.connect(self.refresh_preview)
        preview_top.addWidget(self.updated_label)
        preview_top.addWidget(self.refresh_btn)
        preview_card.layout().addLayout(preview_top)

        self.preview_label = QLabel("Calc пока не выбран.")
        self.preview_label.setWordWrap(True)
        self.preview_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.preview_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.preview_label.setSizePolicy(self.preview_label.sizePolicy().horizontalPolicy(), self.preview_label.sizePolicy().verticalPolicy())
        preview_card.layout().addWidget(self.preview_label)
        preview_card.layout().addWidget(self.status_label)
        layout.addWidget(preview_card)
        layout.addStretch(1)

        self.project_edit.textChanged.connect(self._on_project_text_changed)
        self.client_edit.textChanged.connect(self._remember_and_refresh)
        self.calc_combo.currentTextChanged.connect(self._remember_and_refresh)
        self.template_combo.currentTextChanged.connect(self._remember_and_refresh)
        self.output_edit.textChanged.connect(self._on_output_text_changed)

        self.refresh_shortcut = QShortcut(QKeySequence("F5"), self)
        self.refresh_shortcut.activated.connect(self.refresh_preview)

        self.scan_project(force=False)
        self.refresh_preview()

    def _saved(self, key: str, default: str) -> str:
        value = self.settings.value(key, default)
        return str(value) if value is not None else default

    def _set_saved(self, key: str, value: str) -> None:
        self.settings.setValue(key, value)

    def _path_from_combo(self, combo: QComboBox) -> str:
        return self.owner._path_from_combo(combo)

    def _set_line_path(self, line_edit: QLineEdit, path_text: str) -> None:
        self._updating_path_display = True
        self.owner._set_line_path(line_edit, path_text, is_file=False)
        self._updating_path_display = False

    def project_path_text(self) -> str:
        return self.project_dir_path or self.project_edit.toolTip() or self.project_edit.text().strip()

    def output_path_text(self) -> str:
        return self.output_dir_path or self.output_edit.toolTip() or self.output_edit.text().strip()

    def _load_saved_paths(self) -> None:
        calc = self._saved("riello_offer/calc_path", "")
        if calc:
            self.owner._add_path_item(self.calc_combo, calc, is_file=True)
            self.calc_combo.setCurrentIndex(0)
        template = self._saved("riello_offer/template_path", "")
        if template:
            self.owner._add_path_item(self.template_combo, template, is_file=True)
            self.template_combo.setCurrentIndex(0)

    def remember_values(self) -> None:
        self._set_saved("riello_offer/project_dir", self.project_path_text())
        self._set_saved("riello_offer/client", self.client_edit.text().strip())
        self._set_saved("riello_offer/calc_path", self._path_from_combo(self.calc_combo))
        self._set_saved("riello_offer/template_path", self._path_from_combo(self.template_combo))
        self._set_saved("riello_offer/output_dir", self.output_path_text())
        self.settings.sync()

    def _remember_and_refresh(self, *_args) -> None:
        self.remember_values()
        self.refresh_preview()

    def _on_project_text_changed(self) -> None:
        if not self._updating_path_display:
            self.project_dir_path = self.project_edit.text().strip()
            self.project_edit.setToolTip(self.project_dir_path)
        self.remember_values()

    def _on_output_text_changed(self) -> None:
        if not self._updating_path_display:
            self.output_dir_path = self.output_edit.text().strip()
            self.output_edit.setToolTip(self.output_dir_path)
        self.remember_values()

    def adopt_project(self, path_text: str) -> None:
        path_text = str(path_text or "").strip()
        if not path_text or path_text == self.project_path_text():
            return
        self.project_dir_path = path_text
        self._set_line_path(self.project_edit, path_text)
        self.autofill_client(force=False)
        self.scan_project(force=True)

    def browse_project_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта Riello", self.project_path_text())
        if not path:
            return
        self.project_dir_path = path
        self._set_line_path(self.project_edit, path)
        self.autofill_client(force=True)
        if not self.output_path_text():
            self.output_dir_path = infer_output_dir(path)
            self._set_line_path(self.output_edit, self.output_dir_path)
        self.autofill_manager(force=False)
        self.scan_project(force=True)

    def browse_calc_file(self) -> None:
        current = self._path_from_combo(self.calc_combo)
        start = str(Path(current).parent) if current else self.project_path_text()
        path, _ = QFileDialog.getOpenFileName(self, "Выберите готовый Riello Calc", start, "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        self._select_combo_path(self.calc_combo, path)
        self.remember_values()
        self.refresh_preview()

    def browse_template_file(self) -> None:
        current = self._path_from_combo(self.template_combo)
        start = str(Path(current).parent) if current else self.project_path_text()
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Word-шаблон Riello", start, "Word (*.docx)")
        if not path:
            return
        self._select_combo_path(self.template_combo, path)
        self.remember_values()
        self.refresh_preview()

    def browse_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите папку результата", self.output_path_text() or self.project_path_text())
        if not path:
            return
        self.output_dir_path = path
        self._set_line_path(self.output_edit, path)
        self.remember_values()

    def _select_combo_path(self, combo: QComboBox, path: str) -> None:
        index = self.owner._find_combo_path(combo, path)
        if index < 0:
            self.owner._add_path_item(combo, path, is_file=True)
            index = combo.count() - 1
        combo.setCurrentIndex(index)

    def scan_project(self, force: bool = False) -> None:
        text = self.project_path_text().strip()
        root = Path(text) if text else None
        if root is None or not root.exists():
            return
        if force:
            clear_scan_cache()
        found = scan_project_files(root, use_cache=not force)

        old_calc = self._path_from_combo(self.calc_combo)
        self.calc_combo.blockSignals(True)
        try:
            self.calc_combo.clear()
            excel_files = list(found.get("excel", []))
            excel_files.sort(key=lambda p: ("riello" not in p.name.lower(), -p.stat().st_mtime))
            for path in excel_files:
                self.owner._add_path_item(self.calc_combo, str(path), is_file=True)
            old_index = self.owner._find_combo_path(self.calc_combo, old_calc) if old_calc else -1
            if old_index >= 0:
                self.calc_combo.setCurrentIndex(old_index)
            elif self.calc_combo.count():
                self.calc_combo.setCurrentIndex(0)
        finally:
            self.calc_combo.blockSignals(False)

        old_template = self._path_from_combo(self.template_combo)
        self.template_combo.blockSignals(True)
        try:
            self.template_combo.clear()
            if old_template and Path(old_template).exists():
                self.owner._add_path_item(self.template_combo, old_template, is_file=True)
            for path in found.get("word", []):
                if self.owner._find_combo_path(self.template_combo, str(path)) < 0:
                    self.owner._add_path_item(self.template_combo, str(path), is_file=True)
            if old_template:
                idx = self.owner._find_combo_path(self.template_combo, old_template)
                if idx >= 0:
                    self.template_combo.setCurrentIndex(idx)
        finally:
            self.template_combo.blockSignals(False)

        if not self.output_path_text():
            self.output_dir_path = infer_output_dir(str(root))
            self._set_line_path(self.output_edit, self.output_dir_path)
        self.remember_values()
        self.refresh_preview()

    def autofill_client(self, force: bool = False) -> None:
        client = extract_client_from_project_dir(self.project_path_text())
        if not client:
            return
        current = self.client_edit.text().strip()
        if force or not current or current == "ТОО Example":
            self.client_edit.setText(client)

    def autofill_manager(self, force: bool = False) -> None:
        if not force and self.owner._has_saved_manager_profile():
            return
        root = Path(self.project_path_text())
        if not root.exists():
            return
        profile = find_manager_in_project(root)
        if profile.is_empty():
            return
        self.owner._set_manager_profile(profile)

    def refresh_preview(self) -> None:
        from datetime import datetime

        calc_text = self._path_from_combo(self.calc_combo).strip()
        if not calc_text:
            self.preview_label.setText("Calc пока не выбран.")
            self.status_label.setText("Выберите готовый Riello Calc.")
            return
        try:
            text = preview_riello_offer_calc(calc_text)
            client = self.client_edit.text().strip() or "-"
            template = Path(self._path_from_combo(self.template_combo)).name if self._path_from_combo(self.template_combo) else "не выбран"
            self.preview_label.setText(f"Заказчик: {client}\nWord-шаблон: {template}\n\n{text}")
            calc = read_riello_offer_calc(calc_text)
            self.status_label.setText(
                f"Calc прочитан: {len(calc.items)} поз.; ИБП: {len(calc.equipment)}; опций: {len(calc.options)}. F5 — перечитать файл."
            )
            self.updated_label.setText(f"Обновлено: {datetime.now():%d.%m.%Y %H:%M:%S}")
        except Exception as exc:
            self.preview_label.setText(f"Не удалось прочитать Calc:\n{exc}")
            self.status_label.setText("Исправьте выбор Calc или его структуру.")

    def generate(self) -> None:
        calc_path = Path(self._path_from_combo(self.calc_combo))
        template_path = Path(self._path_from_combo(self.template_combo))
        output_dir = Path(self.output_path_text())
        if not calc_path.exists():
            QMessageBox.warning(self, "Riello", "Выберите существующий готовый Calc.")
            return
        if not template_path.exists():
            QMessageBox.warning(self, "Riello", "Выберите Word-шаблон Riello.")
            return
        if not output_dir:
            QMessageBox.warning(self, "Riello", "Выберите папку результата.")
            return

        try:
            signer = self.owner._selected_signer()
            manager = self.owner._manager_profile()
            output = generate_riello_word_offer(
                calc_path=calc_path,
                template_path=template_path,
                output_dir=output_dir,
                client_name=self.client_edit.text().strip() or "Client",
                signer_name=signer.get("name", ""),
                signer_position=signer.get("position", ""),
                manager_name=manager.name,
                manager_position=manager.position,
                manager_email=manager.email,
                manager_phone=manager.phone,
            )
            self.status_label.setText(f"КП сформировано: {output.name}")
            answer = QMessageBox.question(
                self,
                "Riello",
                f"КП сформировано:\n{output}\n\nОткрыть файл?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(output)))
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка формирования КП Riello", str(exc))

    def clear_cache(self) -> None:
        for key in (
            "riello_offer/project_dir",
            "riello_offer/client",
            "riello_offer/calc_path",
            "riello_offer/template_path",
            "riello_offer/output_dir",
        ):
            self.settings.remove(key)
        self.settings.sync()

    def on_settings_changed(self) -> None:
        self.refresh_preview()
