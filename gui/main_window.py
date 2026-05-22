from __future__ import annotations

APP_FOOTER = """
Направления:
Stulz · Riello · DC Eltek · Generator

Разработчик:
Литвинов Виталий Константинович
"""

SIGNERS = {
    "saniya": {
        "name": "Сания Санаткызы",
        "position": "Коммерческий директор",
    },
    "alisher": {
        "name": "Анаркулов Алишер",
        "position": "Исполнительный директор",
    },
}

import sys
from pathlib import Path

from brands.registry import BRANDS, get_brand_module
from core.excel_reader import list_sheets
from core.models import OfferContext
from core.manager_profile import ManagerProfile, find_manager_in_project
from core.project_scanner import clear_scan_cache, scan_project_files
from gui.path_helpers import extract_brand_from_project_dir, extract_client_from_project_dir
from gui.ui_style import stylesheet, ui_scale


def run_gui() -> None:
    try:
        from PySide6.QtCore import Qt, QSettings, QUrl
        from PySide6.QtGui import QDesktopServices, QFont, QIcon
        from PySide6.QtWidgets import (
            QApplication,
            QComboBox,
            QFileDialog,
            QFrame,
            QGridLayout,
            QGroupBox,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QButtonGroup,
            QPushButton,
            QRadioButton,
            QScrollArea,
            QSizePolicy,
            QSpacerItem,
            QTextEdit,
            QVBoxLayout,
            QWidget,
        )
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("Для запуска GUI установите PySide6: pip install PySide6") from exc

    class OfferGeneratorWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.settings = QSettings("SAM Group", "SAM Offer Generator")
            self.setWindowTitle("SAM Offer Generator")
            self.setMinimumSize(900, 620)
            self.setWindowIcon(QIcon())

            self._updating_path_display = False
            self.project_dir_path = self._saved("project_dir", "")
            self.output_dir_path = self._saved("output_dir", "")

            self.project_edit = QLineEdit(self._display_dir(self.project_dir_path))
            self.project_edit.setToolTip(self.project_dir_path)
            self.client_edit = QLineEdit(self._saved("client", "ТОО Example"))
            self.brand_combo = QComboBox()
            self.brand_combo.addItems(BRANDS.keys())
            self.brand_combo.setCurrentText(self._saved("brand", "Stulz"))
            self.calc_combo = QComboBox()
            self.calc_combo.setEditable(True)
            saved_calc = self._saved("calc_path", "")
            if saved_calc:
                self._add_path_item(self.calc_combo, saved_calc, is_file=True)
                self.calc_combo.setCurrentIndex(0)
            self.template_combo = QComboBox()
            self.template_combo.setEditable(True)
            saved_template = self._saved("template_path", "")
            if saved_template:
                self._add_path_item(self.template_combo, saved_template, is_file=True)
                self.template_combo.setCurrentIndex(0)
            self.sheet_combo = QComboBox()
            self.sheet_combo.setEditable(True)
            saved_sheet = self._saved("sheet_name", "")
            if saved_sheet:
                self.sheet_combo.addItem(saved_sheet)
                self.sheet_combo.setCurrentText(saved_sheet)
            self.output_edit = QLineEdit(self._display_dir(self.output_dir_path))
            self.output_edit.setToolTip(self.output_dir_path)
            self.preview = QTextEdit()
            self.preview.setReadOnly(True)
            self.status_label = QLabel("Выберите папку проекта")
            self._auto_client_value = ""

            self._base_font_size = 10
            self._last_scale = 0.0
            self._responsive_widgets: list[QWidget] = []

            self.manager_name_edit = QLineEdit(self._saved("manager_name", ""))
            self.manager_position_edit = QLineEdit(self._saved("manager_position", ""))
            self.manager_email_edit = QLineEdit(self._saved("manager_email", ""))
            self.manager_phone_edit = QLineEdit(self._saved("manager_phone", ""))
            for edit in (
                self.manager_name_edit,
                self.manager_position_edit,
                self.manager_email_edit,
                self.manager_phone_edit,
            ):
                edit.setObjectName("SidebarInput")
                self._responsive_widgets.append(edit)
            self.use_manager_btn = QPushButton("Использовать")
            self.use_manager_btn.setObjectName("SidebarButton")
            self.use_manager_btn.clicked.connect(self._save_manager_profile)
            self._responsive_widgets.append(self.use_manager_btn)

            self.signer_group = QButtonGroup(self)

            self.signer_saniya_radio = QRadioButton(
                "Сания Санаткызы\nКоммерческий директор"
            )
            self.signer_saniya_radio.setObjectName("SidebarRadio")

            self.signer_alisher_radio = QRadioButton(
                "Анаркулов Алишер\nИсполнительный директор"
            )
            self.signer_alisher_radio.setObjectName("SidebarRadio")

            self.signer_group.addButton(self.signer_saniya_radio)
            self.signer_group.addButton(self.signer_alisher_radio)

            saved_signer = self._saved("signer_key", "saniya")
            if saved_signer == "alisher":
                self.signer_alisher_radio.setChecked(True)
            else:
                self.signer_saniya_radio.setChecked(True)

            self.signer_saniya_radio.toggled.connect(self._on_signer_changed)
            self.signer_alisher_radio.toggled.connect(self._on_signer_changed)

            self._responsive_widgets.append(self.signer_saniya_radio)
            self._responsive_widgets.append(self.signer_alisher_radio)

            self._build_ui()
            self._apply_style()
            self._autofill_client_from_project_dir()
            self._autofill_brand_from_project_dir()
            self._scan_project(force=False)
            self._autofill_manager_from_project(force=False)
            self._refresh_preview()

        def _clear_cache(self) -> None:
            keys_to_clear = [
                "project_dir",
                "output_dir",
                "calc_path",
                "template_path",
                "sheet_name",
                "template_dir",
            ]

            for key in keys_to_clear:
                self.settings.remove(key)

            self.settings.sync()
            clear_scan_cache()

            self.project_dir_path = ""
            self.output_dir_path = ""

            self._updating_path_display = True
            self.project_edit.clear()
            self.output_edit.clear()
            self._updating_path_display = False

            self.project_edit.setToolTip("")
            self.output_edit.setToolTip("")

            self.calc_combo.blockSignals(True)
            self.template_combo.blockSignals(True)
            self.sheet_combo.blockSignals(True)

            self.calc_combo.clear()
            self.template_combo.clear()
            self.sheet_combo.clear()

            self.calc_combo.blockSignals(False)
            self.template_combo.blockSignals(False)
            self.sheet_combo.blockSignals(False)

            self.preview.setPlainText("Кэш очищен. Выберите папку проекта заново.")
            self.status_label.setText("Кэш очищен. Выберите папку проекта заново.")

        def _saved(self, key: str, default: str) -> str:
            value = self.settings.value(key, default)
            return str(value) if value is not None else default

        def _display_file(self, path_text: str) -> str:
            return Path(path_text).name if path_text else ""

        def _display_dir(self, path_text: str) -> str:
            if not path_text:
                return ""
            p = Path(path_text)
            return p.name or path_text

        def _path_from_combo(self, combo: QComboBox) -> str:
            data = combo.currentData()
            return str(data) if data else combo.currentText().strip()

        def _add_path_item(self, combo: QComboBox, path_text: str, is_file: bool = True) -> None:
            full = str(path_text)
            display = self._display_file(full) if is_file else self._display_dir(full)
            combo.addItem(display, full)
            index = combo.count() - 1
            combo.setItemData(index, full, Qt.ToolTipRole)

        def _find_combo_path(self, combo: QComboBox, path_text: str) -> int:
            full = str(path_text)
            for i in range(combo.count()):
                if str(combo.itemData(i) or combo.itemText(i)) == full:
                    return i
            return -1

        def _set_line_path(self, line_edit: QLineEdit, path_text: str, is_file: bool = False) -> None:
            full = str(path_text)
            display = self._display_file(full) if is_file else self._display_dir(full)
            self._updating_path_display = True
            line_edit.setText(display)
            self._updating_path_display = False
            line_edit.setToolTip(full)

        def _project_path_text(self) -> str:
            return self.project_dir_path or self.project_edit.text().strip()

        def _output_path_text(self) -> str:
            return self.output_dir_path or self.output_edit.text().strip()

        def _build_ui(self) -> None:
            central = QWidget()
            root = QHBoxLayout(central)
            root.setContentsMargins(0, 0, 0, 0)
            root.setSpacing(0)

            self.sidebar = QFrame()
            sidebar = self.sidebar
            sidebar.setObjectName("Sidebar")
            sidebar.setMinimumWidth(220)
            side = QVBoxLayout(sidebar)
            side.setContentsMargins(18, 16, 18, 14)
            side.setSpacing(7)

            brand = QLabel("SAM\nGROUP")
            brand.setObjectName("Brand")
            title = QLabel("Offer Generator")
            title.setObjectName("SideTitle")
            subtitle = QLabel("Папка проекта → расчет Excel → шаблон Word → готовое КП")
            subtitle.setObjectName("SideSubtitle")
            subtitle.setWordWrap(True)
            self.clear_cache_btn = QPushButton("Очистить кэш")
            self.clear_cache_btn.setObjectName("Badge")
            self.clear_cache_btn.clicked.connect(self._clear_cache)
            self._responsive_widgets.append(self.clear_cache_btn)

            side.addWidget(brand)
            side.addWidget(title)
            side.addWidget(subtitle)
            side.addSpacing(6)
            side.addWidget(self.clear_cache_btn)
            side.addSpacing(4)
            manager_title = QLabel("Исполнитель")
            manager_title.setObjectName("SidebarSectionTitle")
            side.addWidget(manager_title)
            self._add_sidebar_field(side, "ФИО", self.manager_name_edit)
            self._add_sidebar_field(side, "Должность", self.manager_position_edit)
            self._add_sidebar_field(side, "Email", self.manager_email_edit)
            self._add_sidebar_field(side, "Телефон", self.manager_phone_edit)
            side.addWidget(self.use_manager_btn)

            side.addSpacing(6)
            signer_title = QLabel("Подписант")
            signer_title.setObjectName("SidebarSectionTitle")
            side.addWidget(signer_title)
            side.addWidget(self.signer_saniya_radio)
            side.addWidget(self.signer_alisher_radio)

            manager_hint = QLabel("Если поля пустые, программа попробует взять данные из Word-файла КП в папке проекта.")
            manager_hint.setObjectName("SidebarHint")
            manager_hint.setWordWrap(True)
            side.addWidget(manager_hint)
            side.addSpacerItem(QSpacerItem(20, 12, QSizePolicy.Minimum, QSizePolicy.Expanding))

            footer = QLabel(APP_FOOTER)
            footer.setObjectName("SidebarFooter")
            footer.setWordWrap(True)
            side.addWidget(footer)

            self.content = QFrame()
            content = self.content
            content.setObjectName("Content")
            content.setMinimumWidth(560)
            content_layout = QVBoxLayout(content)
            content_layout.setContentsMargins(34, 28, 34, 28)
            content_layout.setSpacing(18)

            header = QHBoxLayout()
            h_text = QVBoxLayout()
            page_title = QLabel("Новое коммерческое предложение")
            page_title.setObjectName("PageTitle")
            page_subtitle = QLabel("Сначала выберите папку проекта на сервере. Программа найдет Excel и Word внутри папки.")
            page_subtitle.setObjectName("PageSubtitle")
            h_text.addWidget(page_title)
            h_text.addWidget(page_subtitle)
            self.generate_btn = QPushButton("Сформировать КП")
            self.generate_btn.setObjectName("PrimaryButton")
            self.generate_btn.clicked.connect(self._generate)
            header.addLayout(h_text, stretch=1)
            header.addWidget(self.generate_btn, stretch=0, alignment=Qt.AlignTop)
            content_layout.addLayout(header)

            form_card = self._card("Папка проекта и файлы")
            grid = QGridLayout()
            form_card.layout().addLayout(grid)
            grid.setColumnStretch(1, 1)
            grid.setVerticalSpacing(12)
            grid.setHorizontalSpacing(10)

            self._add_row(grid, 0, "Папка проекта", self.project_edit, "Выбрать", self._browse_project_dir)
            self._add_row(grid, 1, "Направление", self.brand_combo, None, None)
            self._add_row(grid, 2, "Клиент", self.client_edit, None, None)
            self._add_row(grid, 3, "Excel-расчет", self.calc_combo, "Обновить", lambda: self._scan_project(force=True))
            self._add_row(grid, 4, "Word-шаблон", self.template_combo, "Обзор", self._browse_template_file)
            self._add_row(grid, 5, "Лист Excel", self.sheet_combo, "Листы", self._load_sheets)
            self._add_row(grid, 6, "Папка результата", self.output_edit, "Выбрать", self._browse_output_dir)
            content_layout.addWidget(form_card)

            bottom = QHBoxLayout()
            preview_card = self._card("Проверка данных")
            preview_card.layout().addWidget(self.preview)
            bottom.addWidget(preview_card, stretch=2)

            status_card = self._card("Статус")
            status_text = QLabel(
                "1. Выберите папку проекта\n"
                "2. Выберите направление\n"
                "3. Выберите Excel и Word\n"
                "4. Нажмите \"Сформировать КП\"\n\n"
                "Результат сохраняется в выбранную папку результата."
            )
            status_text.setWordWrap(True)
            status_card.layout().addWidget(status_text)
            status_card.layout().addWidget(self.status_label)
            bottom.addWidget(status_card, stretch=1)
            content_layout.addLayout(bottom)

            scroll = QScrollArea()
            scroll.setObjectName("ContentScroll")
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setWidget(content)

            self.sidebar_scroll = QScrollArea()
            self.sidebar_scroll.setObjectName("SidebarScroll")
            self.sidebar_scroll.setWidgetResizable(True)
            self.sidebar_scroll.setFrameShape(QFrame.NoFrame)
            self.sidebar_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.sidebar_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.sidebar_scroll.setWidget(sidebar)

            root.addWidget(self.sidebar_scroll)
            root.addWidget(scroll, stretch=1)
            self.setCentralWidget(central)
            self._apply_responsive_metrics(force=True)

            self.project_edit.textChanged.connect(self._on_project_dir_changed)
            self.calc_combo.currentTextChanged.connect(self._load_sheets)
            self.calc_combo.currentTextChanged.connect(self._refresh_preview)
            self.template_combo.currentTextChanged.connect(self._refresh_preview)
            self.sheet_combo.currentTextChanged.connect(self._refresh_preview)
            self.brand_combo.currentTextChanged.connect(self._refresh_preview)
            self.client_edit.textChanged.connect(self._refresh_preview)
            self.output_edit.textChanged.connect(self._on_output_dir_changed)
            self.manager_name_edit.textChanged.connect(self._refresh_preview)
            self.manager_position_edit.textChanged.connect(self._refresh_preview)
            self.manager_email_edit.textChanged.connect(self._refresh_preview)
            self.manager_phone_edit.textChanged.connect(self._refresh_preview)

            # Save user input immediately, not only after successful generation.
            self.project_edit.textChanged.connect(self._remember_values)
            self.client_edit.textChanged.connect(self._remember_values)
            self.output_edit.textChanged.connect(self._remember_values)
            self.brand_combo.currentTextChanged.connect(self._remember_values)
            self.calc_combo.currentTextChanged.connect(self._remember_values)
            self.template_combo.currentTextChanged.connect(self._remember_values)
            self.sheet_combo.currentTextChanged.connect(self._remember_values)

        def _add_sidebar_field(self, layout, label: str, widget) -> None:
            lab = QLabel(label)
            lab.setObjectName("SidebarFormLabel")
            layout.addWidget(lab)
            layout.addWidget(widget)

        def _add_row(self, grid, row: int, label: str, widget, button_text: str | None, command) -> None:
            lab = QLabel(label)
            lab.setObjectName("FormLabel")
            lab.setMinimumWidth(110)
            lab.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            widget.setMinimumWidth(0)
            if isinstance(widget, QComboBox):
                widget.setMinimumContentsLength(1)
                widget.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            self._responsive_widgets.append(widget)

            grid.addWidget(lab, row, 0)
            grid.addWidget(widget, row, 1)
            if button_text:
                btn = QPushButton(button_text)
                btn.setObjectName("GhostButton")
                btn.setMinimumWidth(92)
                btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                if command:
                    btn.clicked.connect(command)
                self._responsive_widgets.append(btn)
                grid.addWidget(btn, row, 2)
            else:
                grid.setColumnMinimumWidth(2, 92)

        def _card(self, title: str) -> QFrame:
            frame = QFrame()
            frame.setObjectName("Card")
            layout = QVBoxLayout(frame)
            layout.setContentsMargins(20, 18, 20, 20)
            layout.setSpacing(12)
            label = QLabel(title)
            label.setObjectName("CardTitle")
            layout.addWidget(label)
            return frame

        def _ui_scale(self) -> float:
            return ui_scale(self.width(), self.height())

        def _apply_responsive_metrics(self, force: bool = False) -> None:
            scale = self._ui_scale()
            if not force and abs(scale - self._last_scale) < 0.03:
                return
            self._last_scale = scale

            sidebar_width = int(max(220, min(240, self.width() * 0.20)))
            self.sidebar.setFixedWidth(sidebar_width)
            if hasattr(self, "sidebar_scroll"):
                self.sidebar_scroll.setFixedWidth(sidebar_width)

            content_margin_x = int(22 * scale)
            content_margin_y = int(20 * scale)
            self.content.layout().setContentsMargins(content_margin_x, content_margin_y, content_margin_x, content_margin_y)
            self.content.layout().setSpacing(int(16 * scale))

            sidebar_input_h = int(28 * scale)
            for widget in self._responsive_widgets:
                widget.setMinimumHeight(sidebar_input_h)

            self.generate_btn.setMinimumWidth(int(220 * scale))
            self.generate_btn.setMinimumHeight(int(42 * scale))
            self._apply_style(scale)

        def resizeEvent(self, event) -> None:  # noqa: N802 - Qt API name
            super().resizeEvent(event)
            self._apply_responsive_metrics()

        def _apply_style(self, scale: float | None = None) -> None:
            scale = scale if scale is not None else self._ui_scale()
            app = QApplication.instance()
            if app:
                app.setFont(QFont("Segoe UI", max(9, int(10 * scale))))
            self.setStyleSheet(stylesheet(scale))

        def _extract_client_from_project_dir(self, path_text: str) -> str:
            return extract_client_from_project_dir(path_text)

        def _extract_brand_from_project_dir(self, path_text: str) -> str:
            return extract_brand_from_project_dir(path_text, tuple(BRANDS.keys()))

        def _autofill_brand_from_project_dir(self) -> None:
            brand = self._extract_brand_from_project_dir(self._project_path_text())
            if not brand or self.brand_combo.currentText() == brand:
                return

            self.brand_combo.blockSignals(True)
            self.brand_combo.setCurrentText(brand)
            self.brand_combo.blockSignals(False)
            self._refresh_preview()

        def _on_project_dir_changed(self) -> None:
            # Do not scan the project tree on every typed character.
            # Scanning is done on startup, after folder selection, or via "Обновить".
            if not self._updating_path_display:
                self.project_dir_path = self.project_edit.text().strip()
                self.project_edit.setToolTip(self.project_dir_path)
            self._autofill_client_from_project_dir()
            self._autofill_brand_from_project_dir()
            self.status_label.setText("Папка изменена. Нажмите «Обновить» или выберите папку через кнопку.")

        def _autofill_client_from_project_dir(self, force: bool = False) -> None:
            client = self._extract_client_from_project_dir(self._project_path_text())
            if not client:
                return

            current = self.client_edit.text().strip()
            should_update = (
                force
                or not current
                or current == "ТОО Example"
                or current == self._auto_client_value
            )
            if not should_update:
                return

            self.client_edit.blockSignals(True)
            self.client_edit.setText(client)
            self.client_edit.blockSignals(False)
            self._auto_client_value = client
            self._refresh_preview()

        def _manager_profile(self) -> ManagerProfile:
            return ManagerProfile(
                name=self.manager_name_edit.text().strip(),
                position=self.manager_position_edit.text().strip(),
                email=self.manager_email_edit.text().strip(),
                phone=self.manager_phone_edit.text().strip(),
            )

        def _selected_signer_key(self) -> str:
            if self.signer_alisher_radio.isChecked():
                return "alisher"
            return "saniya"

        def _selected_signer(self) -> dict[str, str]:
            return SIGNERS[self._selected_signer_key()]

        def _on_signer_changed(self) -> None:
            self.settings.setValue("signer_key", self._selected_signer_key())
            self.settings.sync()
            self._refresh_preview()

        def _set_manager_profile(self, profile: ManagerProfile) -> None:
            self.manager_name_edit.setText(profile.name)
            self.manager_position_edit.setText(profile.position)
            self.manager_email_edit.setText(profile.email)
            self.manager_phone_edit.setText(profile.phone)

        def _save_manager_profile(self) -> None:
            profile = self._manager_profile()
            self.settings.setValue("manager_name", profile.name)
            self.settings.setValue("manager_position", profile.position)
            self.settings.setValue("manager_email", profile.email)
            self.settings.setValue("manager_phone", profile.phone)
            self.settings.setValue("manager_profile_locked", "1")
            self.settings.sync()
            self.status_label.setText("Данные исполнителя сохранены")
            self._refresh_preview()

        def _has_saved_manager_profile(self) -> bool:
            locked = self._saved("manager_profile_locked", "") == "1"
            saved = any(
                self._saved(key, "").strip()
                for key in ("manager_name", "manager_position", "manager_email", "manager_phone")
            )
            return locked or saved

        def _autofill_manager_from_project(self, force: bool = False) -> None:
            if not force and self._has_saved_manager_profile():
                return
            if not force and not self._manager_profile().is_empty():
                return

            project_text = self._project_path_text().strip()
            project_dir = Path(project_text) if project_text else None
            if not project_dir or not project_dir.exists():
                return

            profile = find_manager_in_project(project_dir)
            if profile.is_empty():
                return

            self._set_manager_profile(profile)
            self.settings.setValue("manager_name", profile.name)
            self.settings.setValue("manager_position", profile.position)
            self.settings.setValue("manager_email", profile.email)
            self.settings.setValue("manager_phone", profile.phone)
            self.settings.sync()
            self.status_label.setText("Данные исполнителя найдены в Word-файле проекта")

        def _on_output_dir_changed(self) -> None:
            if not self._updating_path_display:
                self.output_dir_path = self.output_edit.text().strip()
                self.output_edit.setToolTip(self.output_dir_path)
            self._refresh_preview()

        def _browse_project_dir(self) -> None:
            old_project = self._project_path_text().strip()
            path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта", old_project)
            if path:
                # Changing project must not reset the selected Word template.
                # Templates are stored separately and are remembered in QSettings.
                self.project_dir_path = path
                self._set_line_path(self.project_edit, path, is_file=False)

                # Keep a custom result folder. Update it only when it was empty
                # or previously pointed to the old project folder.
                current_output = self._output_path_text().strip()
                if not current_output or current_output == old_project:
                    self.output_dir_path = path
                    self._set_line_path(self.output_edit, path, is_file=False)

                self._autofill_client_from_project_dir(force=True)
                self._autofill_brand_from_project_dir()
                self._scan_project(force=True)
                self._autofill_manager_from_project(force=False)

        def _browse_output_dir(self) -> None:
            path = QFileDialog.getExistingDirectory(self, "Выберите папку результата", self._output_path_text() or self._project_path_text())
            if path:
                self.output_dir_path = path
                self._set_line_path(self.output_edit, path, is_file=False)

        def _browse_template_file(self) -> None:
            current_template = self._path_from_combo(self.template_combo)
            start_dir = str(Path(current_template).parent) if current_template else self._saved("template_dir", self._project_path_text())
            path, _ = QFileDialog.getOpenFileName(self, "Выберите Word-шаблон", start_dir, "Word (*.docx)")
            if path:
                index = self._find_combo_path(self.template_combo, path)
                if index < 0:
                    self._add_path_item(self.template_combo, path, is_file=True)
                    index = self.template_combo.count() - 1
                self.template_combo.setCurrentIndex(index)
                self.settings.setValue("template_dir", str(Path(path).parent))
                self._remember_values()

        def _scan_project(self, force: bool = False) -> None:
            project_text = self._project_path_text().strip()
            project_dir = Path(project_text) if project_text else None
            if not project_dir or not project_dir.exists():
                self.status_label.setText("Папка проекта не выбрана")
                return

            if force:
                clear_scan_cache()

            found = scan_project_files(project_dir, use_cache=not force)

            old_calc = self._path_from_combo(self.calc_combo)
            old_template = self._path_from_combo(self.template_combo)

            self.calc_combo.blockSignals(True)
            self.calc_combo.clear()
            for p in found["excel"]:
                self._add_path_item(self.calc_combo, str(p), is_file=True)
            old_calc_index = self._find_combo_path(self.calc_combo, old_calc) if old_calc else -1
            if old_calc_index >= 0:
                self.calc_combo.setCurrentIndex(old_calc_index)
            self.calc_combo.blockSignals(False)

            self.template_combo.blockSignals(True)
            self.template_combo.clear()

            for p in found["word"]:
                self._add_path_item(self.template_combo, str(p), is_file=True)

            selected_template_index = -1

            # Если старый шаблон находится в текущей папке проекта — оставляем его
            if old_template:
                try:
                    old_path = Path(old_template).resolve()
                    project_path = project_dir.resolve()
                    if old_path.exists() and old_path.is_relative_to(project_path):
                        selected_template_index = self._find_combo_path(self.template_combo, str(old_path))
                except Exception:
                    selected_template_index = -1

            # Иначе выбираем самый свежий Word-файл из проекта
            if selected_template_index < 0 and found["word"]:
                newest_template = max(found["word"], key=lambda p: p.stat().st_mtime)
                selected_template_index = self._find_combo_path(self.template_combo, str(newest_template))

            if selected_template_index >= 0:
                self.template_combo.setCurrentIndex(selected_template_index)

            self.template_combo.blockSignals(False)

            if not self._output_path_text().strip():
                self.output_dir_path = str(project_dir)
                self._set_line_path(self.output_edit, str(project_dir), is_file=False)

            self.status_label.setText(f"Найдено Excel: {len(found['excel'])}, Word: {len(found['word'])}")
            self._load_sheets()
            self._refresh_preview()

        def _load_sheets(self) -> None:
            current = self.sheet_combo.currentText().strip() or self._saved("sheet_name", "")
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            try:
                calc_path = Path(self._path_from_combo(self.calc_combo))
                if calc_path.exists():
                    sheets = list_sheets(calc_path)
                    self.sheet_combo.addItems(sheets)
                    if current and current in sheets:
                        self.sheet_combo.setCurrentText(current)
            except Exception:
                self.sheet_combo.addItem("")
            finally:
                self.sheet_combo.blockSignals(False)
            self._refresh_preview()

        def _make_context(self) -> OfferContext:
            project_dir = Path(self._project_path_text().strip())
            output_dir = Path(self._output_path_text().strip() or project_dir)
            pdf_dir = project_dir if project_dir.exists() else None
            return OfferContext(
                brand=self.brand_combo.currentText(),
                project_dir=project_dir,
                template_path=Path(self._path_from_combo(self.template_combo)),
                calc_path=Path(self._path_from_combo(self.calc_combo)),
                output_dir=output_dir,
                client_name=self.client_edit.text().strip() or "Client",
                sheet_name=self.sheet_combo.currentText().strip() or None,
                pdf_dir=pdf_dir,
                manager_name=self.manager_name_edit.text().strip(),
                manager_position=self.manager_position_edit.text().strip(),
                manager_email=self.manager_email_edit.text().strip(),
                manager_phone=self.manager_phone_edit.text().strip(),
                signer_name=self._selected_signer()["name"],
                signer_position=self._selected_signer()["position"],
            )

        def _validate_context(self, context: OfferContext) -> None:
            if not context.project_dir.exists():
                raise FileNotFoundError("Выберите существующую папку проекта.")
            if not context.calc_path.exists():
                raise FileNotFoundError("Выберите существующий Excel-файл калькуляции.")
            if not context.template_path.exists():
                raise FileNotFoundError("Выберите существующий Word-шаблон.")
            if context.template_path.suffix.lower() != ".docx":
                raise ValueError("Word-шаблон должен быть файлом .docx")

        def _refresh_preview(self) -> None:
            try:
                context = self._make_context()
                if not context.calc_path.exists():
                    self.preview.setPlainText("Excel-файл пока не выбран или не найден.")
                    return
                module = get_brand_module(context.brand)
                self.preview.setPlainText(module.preview(context))
            except Exception as exc:
                self.preview.setPlainText(f"Не удалось прочитать данные: {exc}")

        def _remember_values(self) -> None:
            self.settings.setValue("project_dir", self._project_path_text())
            self.settings.setValue("client", self.client_edit.text())
            self.settings.setValue("brand", self.brand_combo.currentText())
            self.settings.setValue("calc_path", self._path_from_combo(self.calc_combo))
            self.settings.setValue("template_path", self._path_from_combo(self.template_combo))
            self.settings.setValue("sheet_name", self.sheet_combo.currentText())
            self.settings.setValue("output_dir", self._output_path_text())
            self.settings.setValue("signer_key", self._selected_signer_key())
            self.settings.sync()

        def closeEvent(self, event) -> None:  # noqa: N802 - Qt API name
            self._remember_values()
            super().closeEvent(event)

        def _generate(self) -> None:
            try:
                self.generate_btn.setEnabled(False)
                self.status_label.setText("Формирую документ...")
                QApplication.processEvents()

                context = self._make_context()
                self._validate_context(context)
                module = get_brand_module(context.brand)
                out = module.make_offer(context)

                self._remember_values()
                self.status_label.setText(f"Готово: {out.name}")

                msg = QMessageBox(self)
                msg.setWindowTitle("SAM Offer Generator")
                msg.setIcon(QMessageBox.Question)
                msg.setText("Коммерческое предложение успешно сформировано.")
                msg.setInformativeText(str(out))
                open_folder_btn = msg.addButton("Открыть папку", QMessageBox.ActionRole)
                open_file_btn = msg.addButton("Открыть КП", QMessageBox.ActionRole)
                msg.addButton("Закрыть", QMessageBox.RejectRole)
                msg.exec()
                clicked = msg.clickedButton()
                if clicked == open_folder_btn:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(out.parent)))
                elif clicked == open_file_btn:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))

            except Exception as exc:
                self.status_label.setText("Ошибка формирования")
                QMessageBox.critical(self, "Ошибка", str(exc))
            finally:
                self.generate_btn.setEnabled(True)
                self._refresh_preview()

    app = QApplication.instance() or QApplication(sys.argv)
    window = OfferGeneratorWindow()
    window.show()
    app.exec()
