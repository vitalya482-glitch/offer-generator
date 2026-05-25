from __future__ import annotations

from pathlib import Path

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
        current_template = owner._path_from_combo(owner.template_combo)
        if current_template:
            owner._add_path_item(self.template_combo, current_template, is_file=True)
            self.template_combo.setCurrentIndex(0)

        template_card = owner._card("Шаблон КП")
        template_grid = QGridLayout()
        template_card.layout().addLayout(template_grid)
        template_grid.setColumnStretch(1, 1)

        lab = QLabel("Word-шаблон")
        lab.setObjectName("FormLabel")
        template_grid.addWidget(lab, 0, 0)
        template_grid.addWidget(self.template_combo, 0, 1)

        template_btn = QPushButton("Обзор")
        template_btn.setObjectName("GhostButton")
        template_btn.clicked.connect(self._browse_template)
        template_grid.addWidget(template_btn, 0, 2)
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

    def _browse_template(self) -> None:
        owner = self.owner
        selected = str(self.template_combo.currentData() or self.template_combo.currentText()).strip()
        start_dir = str(Path(selected).parent) if selected else owner._saved("template_dir", owner._project_path_text())
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Word-шаблон", start_dir, "Word (*.docx)")
        if path:
            index = owner._find_combo_path(self.template_combo, path)
            if index < 0:
                owner._add_path_item(self.template_combo, path, is_file=True)
                index = self.template_combo.count() - 1
            self.template_combo.setCurrentIndex(index)
            owner.settings.setValue("template_dir", str(Path(path).parent))

    def apply_to_owner(self) -> None:
        owner = self.owner
        template_path = str(self.template_combo.currentData() or self.template_combo.currentText()).strip()
        if template_path:
            owner.template_combo.blockSignals(True)
            owner.template_combo.clear()
            owner._add_path_item(owner.template_combo, template_path, is_file=True)
            owner.template_combo.setCurrentIndex(0)
            owner.template_combo.blockSignals(False)

        owner.manager_name_edit.setText(self.manager_edits["manager_name"].text().strip())
        owner.manager_position_edit.setText(self.manager_edits["manager_position"].text().strip())
        owner.manager_email_edit.setText(self.manager_edits["manager_email"].text().strip())
        owner.manager_phone_edit.setText(self.manager_edits["manager_phone"].text().strip())
        owner._save_manager_profile()
        owner._remember_values()
        owner._refresh_preview()
