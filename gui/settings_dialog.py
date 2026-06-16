from __future__ import annotations

from pathlib import Path

from core.stulz_reference import (
    import_options_from_xlsm,
    import_winplan_from_xlsm,
    load_stulz_options,
    load_stulz_winplan,
    save_stulz_options,
    save_stulz_winplan,
)
from gui.reference_table_dialog import ReferenceTableDialog

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)


class SettingsDialog(QDialog):
    """Окно редко изменяемых настроек программы."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self.setWindowTitle("Настройки")
        self.setMinimumWidth(640)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        self.template_combo = QComboBox()
        self.template_combo.setEditable(True)
        current_template = str(owner.settings.value("template_path", "") or "")
        if current_template:
            owner._add_path_item(self.template_combo, current_template, is_file=True)
            self.template_combo.setCurrentIndex(0)

        self.calc_template_combo = QComboBox()
        self.calc_template_combo.setEditable(True)
        current_calc_template = str(owner.settings.value("calc_template_path", "") or "")
        if current_calc_template:
            owner._add_path_item(self.calc_template_combo, current_calc_template, is_file=True)
            self.calc_template_combo.setCurrentIndex(0)

        template_card = owner._card("Шаблоны")
        template_grid = QGridLayout()
        template_card.layout().addLayout(template_grid)
        template_grid.setColumnStretch(1, 1)

        lab = QLabel("Word-шаблон КП")
        lab.setObjectName("FormLabel")
        template_grid.addWidget(lab, 0, 0)
        template_grid.addWidget(self.template_combo, 0, 1)

        template_btn = QPushButton("Обзор")
        template_btn.setObjectName("GhostButton")
        template_btn.clicked.connect(self._browse_template)
        template_grid.addWidget(template_btn, 0, 2)

        calc_lab = QLabel("Excel-шаблон расчёта")
        calc_lab.setObjectName("FormLabel")
        template_grid.addWidget(calc_lab, 1, 0)
        template_grid.addWidget(self.calc_template_combo, 1, 1)

        calc_template_btn = QPushButton("Обзор")
        calc_template_btn.setObjectName("GhostButton")
        calc_template_btn.clicked.connect(self._browse_calc_template)
        template_grid.addWidget(calc_template_btn, 1, 2)

        layout.addWidget(template_card)

        self.manager_edits = {
            "manager_name": QLineEdit(owner.manager_name_edit.text()),
            "manager_position": QLineEdit(owner.manager_position_edit.text()),
            "manager_email": QLineEdit(owner.manager_email_edit.text()),
            "manager_phone": QLineEdit(owner.manager_phone_edit.text()),
        }

        manager_card = owner._card("Исполнитель")
        manager_grid = QGridLayout()
        manager_card.layout().addLayout(manager_grid)
        manager_grid.setColumnStretch(1, 1)
        rows = [
            ("ФИО", "manager_name"),
            ("Должность", "manager_position"),
            ("Email", "manager_email"),
            ("Телефон", "manager_phone"),
        ]
        for row, (label_text, key) in enumerate(rows):
            label = QLabel(label_text)
            label.setObjectName("FormLabel")
            manager_grid.addWidget(label, row, 0)
            manager_grid.addWidget(self.manager_edits[key], row, 1, 1, 2)
        layout.addWidget(manager_card)


        reference_card = owner._card("Справочники STULZ")
        reference_hint = QLabel(
            "Здесь редактируются базы переводов для спецификаций: опции STULZ и параметры WinPlan. "
            "Изменения сохраняются в config/*.json и будут использоваться дальше при сборке спецификации."
        )
        reference_hint.setWordWrap(True)
        reference_card.layout().addWidget(reference_hint)
        reference_buttons = QGridLayout()
        reference_card.layout().addLayout(reference_buttons)
        options_btn = QPushButton("Опции Stulz")
        options_btn.setObjectName("GhostButton")
        options_btn.clicked.connect(self._open_stulz_options)
        reference_buttons.addWidget(options_btn, 0, 0)
        winplan_btn = QPushButton("WinPlan")
        winplan_btn.setObjectName("GhostButton")
        winplan_btn.clicked.connect(self._open_winplan)
        reference_buttons.addWidget(winplan_btn, 0, 1)
        layout.addWidget(reference_card)

        cache_card = owner._card("Кэш")
        cache_hint = QLabel("Очистка сбрасывает сохраненные пути проекта, расчета, шаблона и папки результата.")
        cache_hint.setWordWrap(True)
        cache_card.layout().addWidget(cache_hint)
        clear_btn = QPushButton("Очистить кэш")
        clear_btn.setObjectName("GhostButton")
        clear_btn.clicked.connect(owner._clear_cache)
        clear_btn.clicked.connect(self.accept)
        cache_card.layout().addWidget(clear_btn)
        layout.addWidget(cache_card)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText("Сохранить")
        buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _open_stulz_options(self) -> None:
        dialog = ReferenceTableDialog(
            self,
            title="Опции Stulz",
            columns=[
                ("Код", "code"),
                ("Название в Calculation", "source_name"),
                ("Русское описание для КП", "ru_description"),
            ],
            load_rows=load_stulz_options,
            save_rows=save_stulz_options,
            import_rows=import_options_from_xlsm,
            untranslated_key="ru_description",
        )
        dialog.exec()

    def _open_winplan(self) -> None:
        dialog = ReferenceTableDialog(
            self,
            title="WinPlan",
            columns=[
                ("Название в WinPlan", "source_name"),
                ("Русское название", "ru_name"),
                ("Раздел", "section"),
                ("Ед. изм.", "unit"),
            ],
            load_rows=load_stulz_winplan,
            save_rows=save_stulz_winplan,
            import_rows=import_winplan_from_xlsm,
            untranslated_key="ru_name",
        )
        dialog.exec()

    def _browse_template(self) -> None:
        owner = self.owner
        selected = str(self.template_combo.currentData() or self.template_combo.currentText()).strip()
        start_dir = str(Path(selected).parent) if selected else owner._saved("template_dir", owner.current_project_dir())
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Word-шаблон", start_dir, "Word (*.docx)")
        if path:
            index = owner._find_combo_path(self.template_combo, path)
            if index < 0:
                owner._add_path_item(self.template_combo, path, is_file=True)
                index = self.template_combo.count() - 1
            self.template_combo.setCurrentIndex(index)
            owner.settings.setValue("template_dir", str(Path(path).parent))


    def _browse_calc_template(self) -> None:
        owner = self.owner
        selected = str(self.calc_template_combo.currentData() or self.calc_template_combo.currentText()).strip()
        start_dir = str(Path(selected).parent) if selected else owner._saved("calc_template_dir", owner.current_project_dir())
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-шаблон расчёта", start_dir, "Excel (*.xlsx)")
        if path:
            index = owner._find_combo_path(self.calc_template_combo, path)
            if index < 0:
                owner._add_path_item(self.calc_template_combo, path, is_file=True)
                index = self.calc_template_combo.count() - 1
            self.calc_template_combo.setCurrentIndex(index)
            owner.settings.setValue("calc_template_dir", str(Path(path).parent))

    def apply_to_owner(self) -> None:
        owner = self.owner
        template_path = str(self.template_combo.currentData() or self.template_combo.currentText()).strip()
        if template_path:
            owner.settings.setValue("template_path", template_path)
            owner.settings.setValue("template_dir", str(Path(template_path).parent))

        calc_template_path = str(self.calc_template_combo.currentData() or self.calc_template_combo.currentText()).strip()
        owner.settings.setValue("calc_template_path", calc_template_path)
        if calc_template_path:
            owner.settings.setValue("calc_template_dir", str(Path(calc_template_path).parent))

        owner.manager_name_edit.setText(self.manager_edits["manager_name"].text().strip())
        owner.manager_position_edit.setText(self.manager_edits["manager_position"].text().strip())
        owner.manager_email_edit.setText(self.manager_edits["manager_email"].text().strip())
        owner.manager_phone_edit.setText(self.manager_edits["manager_phone"].text().strip())
        owner._save_manager_profile()
        owner._notify_pages_settings_changed()
