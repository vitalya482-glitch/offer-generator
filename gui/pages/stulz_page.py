from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QHeaderView,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from brands.registry import get_brand_module
from core.excel_reader import list_sheets
from core.manager_profile import find_manager_in_project
from core.models import OfferContext
from core.runtime_paths import resource_path
from gui.project_folder_selector import ProjectFolderSelector
from gui.path_helpers import (
    infer_output_dir,
    infer_specifications_dir,
)

STULZ_DESCRIPTION_OPTION_DEFAULTS: dict[str, bool] = {
    "stulz_unit": True,
    "cooling_capacity": True,
    "unit_dimensions": True,
    "condenser": True,
}

STULZ_DESCRIPTION_OPTION_SETTINGS_PREFIX = "stulz_description_options/"


def _settings_bool(value: object, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off"}:
        return False
    return default


class StulzPage(QWidget):
    """Полностью самостоятельная страница STULZ.

    Страница сама хранит поля проекта, Excel, Word-шаблона, спецификаций,
    предпросмотр и кнопку формирования КП. MainWindow нужен только как каркас
    и источник общих данных: подписант, исполнитель, настройки.
    """

    brand_name = "Stulz"

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self.settings = owner.settings
        self._updating_path_display = False
        self._updating_spec_models = False
        self._auto_client_value = ""
        self.stulz_description_options = self._load_description_options()

        self.project_dir_path = self._saved("project_dir", "")
        self.output_dir_path = self._saved("output_dir", "")
        self.spec_dir_path = self._saved("spec_dir", "")

        self.project_selector = ProjectFolderSelector(
            owner,
            self.settings,
            settings_key="project_dir",
            available_brands=("Stulz", "Riello", "DC Eltek", "Generator", "HVAC", "Battery"),
            initial_path=self.project_dir_path,
            parent=self,
        )
        self.project_edit = self.project_selector.path_edit
        self.project_selector.projectSelected.connect(self._on_project_folder_selected)
        offer_project_dir = self._saved("stulz/offer_project_dir", self.project_dir_path)
        self.offer_project_selector = ProjectFolderSelector(
            owner,
            self.settings,
            settings_key="stulz/offer_project_dir",
            available_brands=("Stulz", "Riello", "DC Eltek", "Generator", "HVAC", "Battery"),
            initial_path=offer_project_dir,
            parent=self,
        )
        self.offer_project_selector.projectSelected.connect(self._on_offer_project_folder_selected)
        self.offer_calc_combo = QComboBox()
        self.offer_calc_combo.setEditable(True)
        self.offer_sheet_combo = QComboBox()
        self.offer_sheet_combo.setEditable(True)
        self.client_edit = QLineEdit(self._saved("client", "ТОО Example"))
        self.calc_combo = QComboBox()
        self.calc_combo.setEditable(True)
        self.template_combo = QComboBox()
        self.template_combo.setEditable(True)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setEditable(True)
        self.spec_edit = QLineEdit(self._display_dir(self.spec_dir_path))
        self.spec_edit.setToolTip(self.spec_dir_path)
        self.output_edit = QLineEdit(self._display_dir(self.output_dir_path))
        self.output_edit.setToolTip(self.output_dir_path)

        # Параметры нового Calc вынесены на первый шаг, чтобы пользователь
        # видел все обязательные поля до запуска формирования расчёта.
        saved_calc_template = self._saved("calc_template_path", "").strip()
        default_calc_template = resource_path(Path("templates") / "calc exel" / "Calc_14-08-26 Stulz.xlsx")
        if not saved_calc_template or not Path(saved_calc_template).exists():
            saved_calc_template = str(default_calc_template) if default_calc_template.exists() else ""
        self.calc_template_edit = QLineEdit(saved_calc_template)
        self.calc_template_edit.setReadOnly(True)
        self.calc_template_edit.setToolTip(self.calc_template_edit.text().strip())
        self.calc_template_sheet_combo = QComboBox()
        self.calc_template_sheet_combo.addItem("Сначала выберите шаблон Calc", "")
        self.calc_output_edit = QLineEdit(self._saved("calc_output_path", ""))
        self.calc_output_edit.setToolTip(self.calc_output_edit.text().strip())

        self.calc_delivery_combo = QComboBox()
        self.calc_delivery_combo.setEditable(True)
        self.calc_delivery_combo.addItem("Не указано", "")
        for term in ("EXW", "FCA", "CPT", "CIP", "DAP", "DPU", "DDP", "FOB", "CFR", "CIF"):
            self.calc_delivery_combo.addItem(term, term)
        self.calc_delivery_combo.setCurrentIndex(0)

        self.calc_delivery_point_combo = QComboBox()
        self.calc_delivery_point_combo.setEditable(True)
        self.calc_delivery_point_combo.addItem("Не указано", "")
        for city in (
            "Алматы", "Астана", "Шымкент", "Караганда", "Атырау", "Актау", "Актобе",
            "Павлодар", "Усть-Каменогорск", "Семей", "Костанай", "Кокшетау",
            "Петропавловск", "Тараз", "Талдыкорган", "Уральск", "Кызылорда",
            "Туркестан", "Жезказган", "Экибастуз", "Темиртау", "Рудный", "Балхаш",
        ):
            self.calc_delivery_point_combo.addItem(city, city)
        self.calc_delivery_point_combo.setCurrentIndex(0)

        self.calc_currency_combo = QComboBox()
        self.calc_currency_combo.addItem("Не указана", "")
        self.calc_currency_combo.addItem("KZT", "KZT")
        self.calc_currency_combo.addItem("EUR", "EUR")
        self.calc_currency_combo.addItem("USD", "USD")
        self.calc_currency_combo.setCurrentIndex(0)

        self.calc_exchange_rate_edit = QLineEdit(self._saved("calc_exchange_rate", ""))
        self.calc_exchange_rate_edit.setPlaceholderText("например 530")
        self.calc_vat_edit = QLineEdit(self._saved("calc_vat_percent", ""))
        self.calc_vat_edit.setPlaceholderText("например 16")
        self.calc_margin_edit = QLineEdit(self._saved("calc_margin_percent", ""))
        self.calc_margin_edit.setPlaceholderText("например 20")
        self.status_label = QLabel("Выберите папку проекта")

        self._load_saved_combo_paths()
        self.load_calc_template_sheets(refresh=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = QHBoxLayout()
        h_text = QVBoxLayout()
        page_title = QLabel("Stulz: новое коммерческое предложение")
        page_title.setObjectName("PageTitle")
        page_subtitle = QLabel("Выберите папку проекта. Страница Stulz сама найдет Excel, Word и PDF-спецификации внутри папки.")
        page_subtitle.setObjectName("PageSubtitle")
        h_text.addWidget(page_title)
        h_text.addWidget(page_subtitle)
        header.addLayout(h_text, stretch=1)
        layout.addLayout(header)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setObjectName("StulzWorkflowTabs")

        self.calc_stage = QWidget()
        self.offer_stage = QWidget()
        self.positions_stage = QWidget()
        self.tabs.addTab(self.calc_stage, "1. Собрать расчёт")
        self.tabs.addTab(self.offer_stage, "2. Собрать КП")
        self.tabs.addTab(self.positions_stage, "3. Позиции и цены")
        self.tabs.currentChanged.connect(self._on_stage_changed)
        saved_stage = str(self.settings.value("stulz/workspace_tab", "0") or "0")
        try:
            self.tabs.setCurrentIndex(max(0, min(2, int(saved_stage))))
        except Exception:
            self.tabs.setCurrentIndex(0)
        layout.addWidget(self.tabs)

        calc_layout = QVBoxLayout(self.calc_stage)
        calc_layout.setContentsMargins(18, 18, 18, 18)
        calc_layout.setSpacing(14)
        calc_layout.addWidget(self.project_selector.widget)

        files_card = owner._card("Stulz: источники расчёта")
        grid = QGridLayout()
        files_card.layout().addLayout(grid)
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)

        owner._add_row(grid, 0, "Excel-расчет", self.calc_combo, "Обновить", lambda: self.scan_project(force=True))
        owner._add_row(grid, 1, "Лист Excel", self.sheet_combo, "Листы", self.load_sheets)
        owner._add_row(grid, 2, "Папка спецификаций", self.spec_edit, "Выбрать", self.browse_spec_dir)
        calc_layout.addWidget(files_card)

        calc_parameters_card = owner._card("Параметры нового расчёта")
        parameters_grid = QGridLayout()
        calc_parameters_card.layout().addLayout(parameters_grid)
        parameters_grid.setColumnStretch(1, 1)
        parameters_grid.setColumnStretch(3, 1)
        parameters_grid.setVerticalSpacing(12)
        parameters_grid.setHorizontalSpacing(10)
        owner._add_row(
            parameters_grid,
            0,
            "Шаблон Calc",
            self.calc_template_edit,
            "Выбрать",
            self.browse_calc_template,
        )
        owner._add_row(
            parameters_grid,
            1,
            "Лист шаблона",
            self.calc_template_sheet_combo,
            None,
            None,
        )
        owner._add_row(
            parameters_grid,
            2,
            "Сохранить Calc",
            self.calc_output_edit,
            "Выбрать",
            self.browse_calc_output,
        )
        owner._add_row(parameters_grid, 3, "Условия поставки", self.calc_delivery_combo, None, None)
        owner._add_row(parameters_grid, 4, "Пункт поставки", self.calc_delivery_point_combo, None, None)
        owner._add_row(parameters_grid, 5, "Валюта", self.calc_currency_combo, None, None)
        owner._add_row(parameters_grid, 6, "Курс", self.calc_exchange_rate_edit, None, None)
        owner._add_row(parameters_grid, 7, "НДС, %", self.calc_vat_edit, None, None)
        owner._add_row(parameters_grid, 8, "Маржа, %", self.calc_margin_edit, None, None)
        calc_parameters_hint = QLabel(
            "Все поля обязательны для формирования Calc. Если значение не найдено автоматически, "
            "оставьте его пустым и заполните вручную. Дополнительные расходы пока не используются."
        )
        calc_parameters_hint.setObjectName("Hint")
        calc_parameters_hint.setWordWrap(True)
        calc_parameters_card.layout().addWidget(calc_parameters_hint)
        calc_layout.addWidget(calc_parameters_card)

        calc_action_card = owner._card("Расчёт Calc")
        calc_action_layout = QVBoxLayout()
        calc_action_card.layout().addLayout(calc_action_layout)
        calc_hint = QLabel(
            "Проверьте найденные позиции и спецификации, затем откройте форму параметров расчёта."
        )
        calc_hint.setObjectName("Hint")
        calc_hint.setWordWrap(True)
        calc_action_layout.addWidget(calc_hint)
        self.spec_preview_button = QPushButton("Сформировать расчёт")
        self.spec_preview_button.setMinimumHeight(42)
        self.spec_preview_button.setObjectName("PrimaryButton")
        self.spec_preview_button.clicked.connect(self.generate_calc_file)
        calc_action_layout.addWidget(self.spec_preview_button)
        calc_layout.addWidget(calc_action_card)
        calc_layout.addStretch(1)

        offer_layout = QVBoxLayout(self.offer_stage)
        offer_layout.setContentsMargins(18, 18, 18, 18)
        offer_layout.setSpacing(14)
        offer_layout.addWidget(self.offer_project_selector.widget)

        offer_files_card = owner._card("Stulz: параметры КП")
        offer_grid = QGridLayout()
        offer_files_card.layout().addLayout(offer_grid)
        offer_grid.setColumnStretch(1, 1)
        offer_grid.setVerticalSpacing(12)
        offer_grid.setHorizontalSpacing(10)
        owner._add_row(offer_grid, 0, "Excel Calc", self.offer_calc_combo, "Обновить", self.refresh_offer_project_files)
        owner._add_row(offer_grid, 1, "Лист Excel", self.offer_sheet_combo, "Листы", self.load_offer_sheets)
        owner._add_row(offer_grid, 2, "Клиент", self.client_edit, None, None)
        owner._add_row(offer_grid, 3, "Word-шаблон", self.template_combo, "Выбрать", self.browse_template_file)
        owner._add_row(offer_grid, 4, "Папка результата", self.output_edit, "Выбрать", self.browse_output_dir)
        offer_layout.addWidget(offer_files_card)

        offer_header = QHBoxLayout()
        offer_title = QLabel("Проверка и формирование КП")
        offer_title.setObjectName("SectionTitle")
        offer_header.addWidget(offer_title)
        offer_header.addStretch(1)
        self.generate_btn = QPushButton("Сформировать КП")
        self.generate_btn.setObjectName("PrimaryButton")
        self.generate_btn.clicked.connect(self.generate)
        offer_header.addWidget(self.generate_btn)
        offer_layout.addLayout(offer_header)

        preview_card = owner._card("Проверка данных КП")
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        preview_card.layout().addWidget(self.preview)
        offer_layout.addWidget(preview_card, stretch=1)
        offer_layout.addStretch(1)

        calc_positions_card = owner._card("Позиции для нового Calc")
        spec_hint = QLabel(
            "Позиции загружаются из Calc.pdf. Отметьте модели, которые должны попасть в новый Calc; "
            "количество можно скорректировать перед формированием."
        )
        spec_hint.setObjectName("Hint")
        spec_hint.setWordWrap(True)
        calc_positions_card.layout().addWidget(spec_hint)

        self.calc_models_table = QTableWidget(0, 8)
        self.calc_models_table.setHorizontalHeaderLabels([
            "В Calc", "Модель", "Кол-во", "List total",
            "Покупка / шт.", "Purchase total", "Валюта", "Файл Calc.pdf",
        ])
        self.calc_models_table.verticalHeader().setVisible(False)
        self.calc_models_table.setAlternatingRowColors(True)
        self.calc_models_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.calc_models_table.setMinimumHeight(190)
        self.calc_models_table.setColumnWidth(0, 52)
        self.calc_models_table.setColumnWidth(2, 80)
        self.calc_models_table.setColumnWidth(3, 115)
        self.calc_models_table.setColumnWidth(4, 120)
        self.calc_models_table.setColumnWidth(5, 125)
        self.calc_models_table.setColumnWidth(6, 72)
        self.calc_models_table.horizontalHeader().setStretchLastSection(True)
        self.calc_models_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        calc_positions_card.layout().addWidget(self.calc_models_table)
        self.calc_models_status_label = QLabel("Выберите папку проекта для поиска Calc.pdf.")
        self.calc_models_status_label.setObjectName("Hint")
        calc_models_status_row = QHBoxLayout()
        calc_models_status_row.addWidget(self.calc_models_status_label, stretch=1)
        self.open_calc_button = QPushButton("Открыть расчёт")
        self.open_calc_button.setObjectName("GhostButton")
        self.open_calc_button.clicked.connect(self.open_generated_calc)
        self.open_calc_folder_button = QPushButton("Открыть папку расчёта")
        self.open_calc_folder_button.setObjectName("GhostButton")
        self.open_calc_folder_button.clicked.connect(self.open_calc_output_folder)
        calc_models_status_row.addWidget(self.open_calc_button)
        calc_models_status_row.addWidget(self.open_calc_folder_button)
        calc_positions_card.layout().addLayout(calc_models_status_row)

        calc_positions_controls = QHBoxLayout()
        self.calc_models_refresh_button = QPushButton("Обновить позиции")
        self.calc_models_refresh_button.setObjectName("GhostButton")
        self.calc_models_refresh_button.clicked.connect(self.refresh_calc_models_table)
        self.calc_models_manual_button = QPushButton("Выбрать Calc.pdf вручную")
        self.calc_models_manual_button.setObjectName("GhostButton")
        self.calc_models_manual_button.clicked.connect(self.browse_calc_model_pdfs)
        self.calc_models_auto_button = QPushButton("Автопоиск")
        self.calc_models_auto_button.setObjectName("GhostButton")
        self.calc_models_auto_button.clicked.connect(self.clear_calc_model_manual_selection)
        calc_positions_controls.addWidget(self.calc_models_refresh_button)
        calc_positions_controls.addWidget(self.calc_models_manual_button)
        calc_positions_controls.addWidget(self.calc_models_auto_button)
        calc_positions_controls.addStretch(1)
        calc_positions_card.layout().addLayout(calc_positions_controls)

        calc_layout.insertWidget(calc_layout.indexOf(calc_parameters_card), calc_positions_card)

        positions_layout = QVBoxLayout(self.positions_stage)
        positions_layout.setContentsMargins(18, 18, 18, 18)
        positions_layout.setSpacing(14)
        positions_title = QLabel("Stulz: позиции и цены")
        positions_title.setObjectName("PageTitle")
        positions_layout.addWidget(positions_title)
        positions_hint = QLabel(
            "Выбор моделей для нового Calc находится на подвкладке «1. Собрать расчёт». "
            "Там можно проверить цены и отметить позиции, которые попадут в файл."
        )
        positions_hint.setObjectName("PageSubtitle")
        positions_hint.setWordWrap(True)
        positions_layout.addWidget(positions_hint)
        spec_card = owner._card("Сопоставление позиций исходного Calc со спецификациями")
        spec_map_hint = QLabel(
            "Эта таблица используется при подготовке КП. Выбор позиций для нового файла Calc "
            "находится на подвкладке «1. Собрать расчёт»."
        )
        spec_map_hint.setObjectName("Hint")
        spec_map_hint.setWordWrap(True)
        spec_card.layout().addWidget(spec_map_hint)
        self.spec_models_table = QTableWidget(0, 3)
        self.spec_models_table.setHorizontalHeaderLabels(["Позиция", "Кол-во", "Папка спецификации"])
        self.spec_models_table.verticalHeader().setVisible(False)
        self.spec_models_table.setAlternatingRowColors(True)
        self.spec_models_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.spec_models_table.setMinimumHeight(190)
        self.spec_models_table.horizontalHeader().setStretchLastSection(True)
        self.spec_models_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        spec_card.layout().addWidget(self.spec_models_table)
        spec_card.layout().addWidget(self.status_label)
        positions_layout.addWidget(spec_card, stretch=1)
        positions_layout.addStretch(1)

        self._ensure_runtime_layers_loaded()
        self._connect_changes()
        self._update_calc_open_buttons()
        self.load_sheets(refresh=False)
        self.scan_project(force=False)
        if self.offer_project_path_text() and Path(self.offer_project_path_text()).is_dir():
            self.refresh_offer_project_files()
        self.refresh_calc_models_table()
        self.refresh_preview()
        self.settings.setValue("stulz/offer_project_dir", self.offer_project_path_text())

    def _on_stage_changed(self, index: int) -> None:
        """Сохраняет выбранную подвкладку и обновляет связанные данные."""
        self.settings.setValue("stulz/workspace_tab", index)
        self.settings.sync()
        if index == 1 and hasattr(self, "preview"):
            self.refresh_preview()
        if index == 0 and hasattr(self, "calc_models_table"):
            self.refresh_calc_models_table()

    # ------------------------- settings/path helpers -------------------------
    def _saved(self, key: str, default: str) -> str:
        value = self.settings.value(key, default)
        return str(value) if value is not None else default

    def _set_saved(self, key: str, value: str) -> None:
        self.settings.setValue(key, value)

    def _display_file(self, path_text: str) -> str:
        return self.owner._display_file(path_text)

    def _display_dir(self, path_text: str) -> str:
        return self.owner._display_dir(path_text)

    def _path_from_combo(self, combo: QComboBox) -> str:
        return self.owner._path_from_combo(combo)

    def _set_line_path(self, line_edit: QLineEdit, path_text: str, is_file: bool = False) -> None:
        if line_edit is self.project_edit:
            self._updating_path_display = True
            try:
                self.project_selector.set_project_path(path_text)
            finally:
                self._updating_path_display = False
            return
        self._updating_path_display = True
        self.owner._set_line_path(line_edit, path_text, is_file=is_file)
        self._updating_path_display = False

    def _set_spec_dir_path(self, path_text: str) -> None:
        self.spec_dir_path = path_text
        self.spec_edit.blockSignals(True)
        try:
            self._set_line_path(self.spec_edit, path_text, is_file=False)
        finally:
            self.spec_edit.blockSignals(False)

    def project_path_text(self) -> str:
        return self.project_selector.project_path_text()

    def offer_project_path_text(self) -> str:
        return self.offer_project_selector.project_path_text()

    def _on_offer_project_folder_selected(self, _old_project: str, path: str) -> None:
        self.settings.setValue("stulz/offer_project_dir", path)
        self.settings.sync()
        if path:
            self.refresh_offer_project_files()

    def refresh_offer_project_files(self) -> None:
        path = self.offer_project_path_text().strip()
        if not path or not Path(path).is_dir():
            self.status_label.setText("Выберите папку проекта для КП.")
            return
        self.status_label.setText("Сканирую файлы проекта для КП…")
        self.offer_project_selector.scan_async(self._apply_offer_project_scan)

    def load_offer_sheets(self) -> None:
        current = self.offer_sheet_combo.currentText().strip() or self._saved("stulz/offer_sheet_name", "")
        self.offer_sheet_combo.blockSignals(True)
        self.offer_sheet_combo.clear()
        calc_path = self._path_from_combo(self.offer_calc_combo).strip()
        try:
            if calc_path and Path(calc_path).exists():
                sheets = list_sheets(Path(calc_path))
                self.offer_sheet_combo.addItems(sheets)
                if current and current in sheets:
                    self.offer_sheet_combo.setCurrentText(current)
        except Exception as exc:
            self.offer_sheet_combo.addItem(f"Не удалось прочитать листы: {exc}")
        finally:
            self.offer_sheet_combo.blockSignals(False)

    def _apply_offer_project_scan(self, found: dict[str, list[Path]]) -> None:
        project_info = self.offer_project_selector.project_info(found.get("pdf_dirs", []))
        current_calc = self._path_from_combo(self.offer_calc_combo)
        self.offer_calc_combo.blockSignals(True)
        self.offer_calc_combo.clear()
        for path in found.get("excel", []):
            self.owner._add_path_item(self.offer_calc_combo, str(path), is_file=True)
        calc_index = self.owner._find_combo_path(self.offer_calc_combo, current_calc) if current_calc else -1
        if calc_index < 0 and self.offer_calc_combo.count():
            calc_index = 0
        if calc_index >= 0:
            self.offer_calc_combo.setCurrentIndex(calc_index)
        self.offer_calc_combo.blockSignals(False)

        current_template = self._path_from_combo(self.template_combo)
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        for path in found.get("word", []):
            self.owner._add_path_item(self.template_combo, str(path), is_file=True)
        template_index = self.owner._find_combo_path(self.template_combo, current_template) if current_template else -1
        if template_index < 0 and current_template and Path(current_template).exists():
            self.owner._add_path_item(self.template_combo, current_template, is_file=True)
            template_index = self.template_combo.count() - 1
        if template_index < 0 and self.template_combo.count():
            template_index = 0
        if template_index >= 0:
            self.template_combo.setCurrentIndex(template_index)
        self.template_combo.blockSignals(False)

        client = self.offer_project_selector.client_name()
        if client:
            self.client_edit.setText(client)
        if project_info.output_dir:
            self.output_dir_path = project_info.output_dir
            self._set_line_path(self.output_edit, project_info.output_dir, is_file=False)
        self.load_offer_sheets()
        self.status_label.setText(
            f"Файлы проекта для КП: Excel — {len(found.get('excel', []))}, "
            f"Word — {len(found.get('word', []))}, PDF — {len(found.get('pdf_dirs', []))} папок."
        )
        self.remember_values()
        self.refresh_preview()

    def output_path_text(self) -> str:
        return self.output_dir_path or self.output_edit.toolTip() or self.output_edit.text().strip()

    def spec_path_text(self) -> str:
        return self.spec_dir_path or self.spec_edit.toolTip() or self.spec_edit.text().strip()

    def _load_saved_combo_paths(self) -> None:
        saved_calc = self._saved("calc_path", "")
        if saved_calc:
            self.owner._add_path_item(self.calc_combo, saved_calc, is_file=True)
            self.calc_combo.setCurrentIndex(0)

        saved_template = self._saved("template_path", "")
        if saved_template:
            self.owner._add_path_item(self.template_combo, saved_template, is_file=True)
            self.template_combo.setCurrentIndex(0)

        saved_sheet = self._saved("sheet_name", "")
        if saved_sheet:
            self.sheet_combo.addItem(saved_sheet)
            self.sheet_combo.setCurrentText(saved_sheet)

        saved_offer_calc = self._saved("stulz/offer_calc_path", "")
        if saved_offer_calc:
            self.owner._add_path_item(self.offer_calc_combo, saved_offer_calc, is_file=True)
            self.offer_calc_combo.setCurrentIndex(0)
        saved_offer_sheet = self._saved("stulz/offer_sheet_name", "")
        if saved_offer_sheet:
            self.offer_sheet_combo.addItem(saved_offer_sheet)
            self.offer_sheet_combo.setCurrentText(saved_offer_sheet)

        saved_delivery = self._saved("calc_delivery_terms", "")
        if saved_delivery:
            self.calc_delivery_combo.setCurrentText(saved_delivery)
        saved_delivery_point = self._saved("calc_delivery_point", "")
        if saved_delivery_point:
            self.calc_delivery_point_combo.setCurrentText(saved_delivery_point)
        saved_currency = self._saved("calc_currency", "").upper()
        if saved_currency:
            currency_index = self.calc_currency_combo.findData(saved_currency)
            if currency_index >= 0:
                self.calc_currency_combo.setCurrentIndex(currency_index)

    def _connect_changes(self) -> None:
        self.project_selector.pathEdited.connect(self.on_project_dir_changed)
        self.project_selector.pathEdited.connect(lambda _path: self.remember_values())
        self.offer_project_selector.pathEdited.connect(
            lambda path: self.settings.setValue("stulz/offer_project_dir", path)
        )
        self.offer_project_selector.path_edit.editingFinished.connect(self.refresh_offer_project_files)
        self.offer_calc_combo.currentTextChanged.connect(self.load_offer_sheets)
        self.offer_calc_combo.currentTextChanged.connect(self.refresh_preview)
        self.offer_sheet_combo.currentTextChanged.connect(self.refresh_preview)
        self.client_edit.textChanged.connect(self.refresh_preview)
        self.calc_combo.currentTextChanged.connect(lambda _text: self.load_sheets())
        self.calc_combo.currentTextChanged.connect(self.refresh_preview)
        self.template_combo.currentTextChanged.connect(self.refresh_preview)
        self.sheet_combo.currentTextChanged.connect(self.refresh_preview)
        self.output_edit.textChanged.connect(self.on_output_dir_changed)
        self.spec_edit.textChanged.connect(self.on_spec_dir_changed)
        self.calc_models_table.itemChanged.connect(self.on_calc_model_table_changed)
        self.calc_output_edit.textChanged.connect(self.on_calc_output_changed)
        self.calc_template_sheet_combo.currentIndexChanged.connect(self._on_calc_template_sheet_changed)
        self.calc_exchange_rate_edit.textChanged.connect(self.refresh_preview)
        self.calc_vat_edit.textChanged.connect(self.refresh_preview)
        self.calc_margin_edit.textChanged.connect(self.refresh_preview)
        self.calc_delivery_combo.currentTextChanged.connect(self.refresh_preview)
        self.calc_delivery_point_combo.currentTextChanged.connect(self.refresh_preview)
        self.calc_currency_combo.currentIndexChanged.connect(self.refresh_preview)

        for widget_signal in (
            self.project_edit.textChanged,
            self.client_edit.textChanged,
            self.output_edit.textChanged,
            self.spec_edit.textChanged,
            self.calc_combo.currentTextChanged,
            self.template_combo.currentTextChanged,
            self.sheet_combo.currentTextChanged,
            self.calc_template_edit.textChanged,
            self.calc_template_sheet_combo.currentTextChanged,
            self.calc_output_edit.textChanged,
            self.calc_exchange_rate_edit.textChanged,
            self.calc_vat_edit.textChanged,
            self.calc_margin_edit.textChanged,
            self.calc_delivery_combo.currentTextChanged,
            self.calc_delivery_point_combo.currentTextChanged,
            self.calc_currency_combo.currentTextChanged,
        ):
            widget_signal.connect(lambda *_: self.remember_values())

    def _ensure_runtime_layers_loaded(self):
        if getattr(self, "_stulz_runtime_layers_loaded", False):
            return getattr(self, "_stulz_runtime_module", None)
        if getattr(self, "_stulz_runtime_layers_loading", False):
            return getattr(self, "_stulz_runtime_module", None)

        self._stulz_runtime_layers_loading = True
        try:
            module = get_brand_module(self.brand_name)
        except Exception:
            return None
        finally:
            self._stulz_runtime_layers_loading = False

        self._stulz_runtime_module = module
        self._stulz_runtime_layers_loaded = True
        return module

    # ------------------------- description/spec options -------------------------
    def _load_description_options(self) -> dict[str, bool]:
        options = dict(STULZ_DESCRIPTION_OPTION_DEFAULTS)
        for key, default in STULZ_DESCRIPTION_OPTION_DEFAULTS.items():
            value = self.settings.value(STULZ_DESCRIPTION_OPTION_SETTINGS_PREFIX + key, default)
            options[key] = _settings_bool(value, default)
        return options

    def _save_description_options(self, options: dict[str, bool]) -> None:
        normalized = dict(STULZ_DESCRIPTION_OPTION_DEFAULTS)
        normalized.update({key: bool(value) for key, value in options.items() if key in normalized})
        self.stulz_description_options = normalized
        for key, value in normalized.items():
            self.settings.setValue(STULZ_DESCRIPTION_OPTION_SETTINGS_PREFIX + key, bool(value))
        self.settings.sync()

    def description_options(self) -> dict[str, bool]:
        defaults = dict(STULZ_DESCRIPTION_OPTION_DEFAULTS)
        defaults.update({key: bool(value) for key, value in self.stulz_description_options.items() if key in defaults})
        return defaults

    # ------------------------- browse/scan -------------------------
    def on_project_dir_changed(self) -> None:
        if not self._updating_path_display:
            self.project_dir_path = self.project_edit.text().strip()
            self.project_edit.setToolTip(self.project_dir_path)
        self.autofill_client_from_project_dir()
        self.status_label.setText("Папка изменена. Нажмите «Обновить» или выберите папку через кнопку.")
        self.refresh_calc_models_table()
        self.refresh_preview()

    def on_output_dir_changed(self) -> None:
        if not self._updating_path_display:
            self.output_dir_path = self.output_edit.text().strip()
            self.output_edit.setToolTip(self.output_dir_path)
        self.refresh_preview()

    def on_spec_dir_changed(self) -> None:
        if not self._updating_path_display:
            self.spec_dir_path = self.spec_edit.text().strip()
            self.spec_edit.setToolTip(self.spec_dir_path)
        self.refresh_preview()
        self.refresh_calc_models_table()

    def on_calc_output_changed(self) -> None:
        self.calc_output_edit.setToolTip(self.calc_output_edit.text().strip())
        self._update_calc_open_buttons()
        self.refresh_preview()

    def _update_calc_open_buttons(self) -> None:
        if not hasattr(self, "open_calc_button"):
            return
        output_text = self.calc_output_edit.text().strip()
        output_path = Path(output_text) if output_text else None
        self.open_calc_button.setEnabled(bool(output_path and output_path.is_file()))
        self.open_calc_folder_button.setEnabled(bool(output_path and output_path.parent.is_dir()))

    def open_generated_calc(self) -> None:
        output_text = self.calc_output_edit.text().strip()
        output_path = Path(output_text) if output_text else None
        if not output_path or not output_path.is_file():
            QMessageBox.warning(self, "Calc", "Сначала сформируйте расчёт или укажите существующий файл.")
            self._update_calc_open_buttons()
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path.resolve()))):
            QMessageBox.warning(self, "Calc", f"Не удалось открыть файл:\n{output_path}")

    def open_calc_output_folder(self) -> None:
        output_text = self.calc_output_edit.text().strip()
        output_path = Path(output_text) if output_text else None
        folder = output_path.parent if output_path else None
        if not folder or not folder.is_dir():
            QMessageBox.warning(self, "Calc", "Папка для расчёта не найдена.")
            self._update_calc_open_buttons()
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve()))):
            QMessageBox.warning(self, "Calc", f"Не удалось открыть папку:\n{folder}")

    def browse_project_dir(self) -> None:
        self.project_selector.browse()

    def _on_project_folder_selected(self, old_project: str, path: str) -> None:
        self.project_dir_path = path
        self.autofill_client_from_project_dir(force=True)

        if path != old_project:
            self.spec_dir_path = ""
            self.output_dir_path = ""
            self._set_line_path(self.spec_edit, "", is_file=False)
            self._set_line_path(self.output_edit, "", is_file=False)

        brand = self.project_selector.brand_name()
        if brand and brand != self.brand_name:
            self.owner._select_tab_for_brand(brand)
        self.scan_project(force=True)
        self.autofill_manager_from_project(force=False)
        self.remember_values()

    def browse_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите папку результата", self.output_path_text() or self.project_path_text())
        if path:
            self.output_dir_path = path
            self._set_line_path(self.output_edit, path, is_file=False)

    def browse_spec_dir(self) -> None:
        start_dir = self.spec_path_text() or self.project_path_text()
        path = QFileDialog.getExistingDirectory(self, "Выберите папку спецификаций", start_dir)
        if path:
            self.spec_dir_path = path
            self._set_line_path(self.spec_edit, path, is_file=False)
            self.refresh_preview()

    def browse_calc_template(self) -> None:
        current = self.calc_template_edit.text().strip()
        start_dir = str(Path(current).parent) if current else self.project_path_text()
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-шаблон расчёта",
            start_dir,
            "Excel (*.xlsx *.xlsm)",
        )
        if path:
            self.calc_template_edit.setText(path)
            self.calc_template_edit.setToolTip(path)
            self.load_calc_template_sheets()
            self.refresh_preview()

    def load_calc_template_sheets(self, refresh: bool = True) -> None:
        """Загружает доступные листы выбранного Excel-шаблона."""
        template_path = Path(self.calc_template_edit.text().strip())
        saved_sheet = self._saved("calc_template_sheet", "").strip()
        self.calc_template_sheet_combo.blockSignals(True)
        try:
            self.calc_template_sheet_combo.clear()
            if not template_path.exists():
                self.calc_template_sheet_combo.addItem("Сначала выберите шаблон Calc", "")
                return

            sheets = list_sheets(template_path)
            self.calc_template_sheet_combo.addItem("Выберите лист шаблона", "")
            for sheet in sheets:
                self.calc_template_sheet_combo.addItem(sheet, sheet)
            saved_index = self.calc_template_sheet_combo.findData(saved_sheet)
            if saved_index >= 0:
                self.calc_template_sheet_combo.setCurrentIndex(saved_index)
            elif sheets:
                self.calc_template_sheet_combo.setCurrentIndex(1)
        except Exception as exc:
            self.calc_template_sheet_combo.clear()
            self.calc_template_sheet_combo.addItem(f"Не удалось прочитать листы: {exc}", "")
        finally:
            self.calc_template_sheet_combo.blockSignals(False)

        self._on_calc_template_sheet_changed()
        if refresh:
            self.refresh_preview()

    def _on_calc_template_sheet_changed(self, *_args) -> None:
        self.settings.setValue(
            "calc_template_sheet",
            str(self.calc_template_sheet_combo.currentData() or ""),
        )
        self.settings.sync()

    def browse_calc_output(self) -> None:
        current = self.calc_output_edit.text().strip()
        if current:
            initial_path = current
        else:
            output_dir = self.output_path_text().strip() or self.project_path_text().strip()
            project_name = Path(self.project_path_text().strip()).name or "Project"
            safe_project_name = re.sub(r'[<>:"/\\|?*]+', "_", project_name).strip(" ._") or "Project"
            filename = f"Calc_Stulz_{safe_project_name}_{datetime.now():%d-%m-%y}.xlsx"
            initial_path = str(Path(output_dir) / filename) if output_dir else filename
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Куда сохранить расчёт Calc",
            initial_path,
            "Excel (*.xlsx)",
        )
        if path:
            if not path.lower().endswith((".xlsx", ".xlsm")):
                path += ".xlsx"
            self.calc_output_edit.setText(path)

    def calc_parameter_values(self) -> dict[str, str]:
        """Возвращает значения обязательных полей нового расчёта."""

        def combo_value(combo: QComboBox) -> str:
            text = combo.currentText().strip()
            return "" if not text or text == "Не указано" else text

        return {
            "Шаблон Calc": self.calc_template_edit.text().strip(),
            "Лист шаблона": str(self.calc_template_sheet_combo.currentData() or "").strip(),
            "Сохранить Calc": self.calc_output_edit.text().strip(),
            "Условия поставки": combo_value(self.calc_delivery_combo),
            "Пункт поставки": combo_value(self.calc_delivery_point_combo),
            "Валюта": str(self.calc_currency_combo.currentData() or "").strip().upper(),
            "Курс": self.calc_exchange_rate_edit.text().strip(),
            "НДС, %": self.calc_vat_edit.text().strip(),
            "Маржа, %": self.calc_margin_edit.text().strip(),
        }

    def validate_calc_parameters(self) -> None:
        """Проверяет обязательные параметры до открытия предпросмотра Calc."""
        values = self.calc_parameter_values()
        for label, value in values.items():
            if not value:
                raise ValueError(f"Заполните поле «{label}» перед формированием расчёта.")

        template_path = Path(values["Шаблон Calc"])
        if not template_path.exists():
            raise FileNotFoundError(f"Файл из поля «Шаблон Calc» не найден: {template_path}")
        if template_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Поле «Шаблон Calc» должно содержать Excel-файл .xlsx или .xlsm.")
        if values["Лист шаблона"] not in list_sheets(template_path):
            raise ValueError("Выберите существующий лист в поле «Лист шаблона».")

        output_path = Path(values["Сохранить Calc"])
        if output_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Поле «Сохранить Calc» должно заканчиваться на .xlsx или .xlsm.")

        for label in ("Курс", "НДС, %", "Маржа, %"):
            try:
                number = float(values[label].replace(",", "."))
            except ValueError as exc:
                raise ValueError(f"Поле «{label}» должно быть числом.") from exc
            if number < 0:
                raise ValueError(f"Поле «{label}» не может быть отрицательным.")
        if float(values["Курс"].replace(",", ".")) <= 0:
            raise ValueError("Поле «Курс» должно быть больше нуля.")
        for label in ("НДС, %", "Маржа, %"):
            if float(values[label].replace(",", ".")) > 100:
                raise ValueError(f"Поле «{label}» не может быть больше 100%.")

    def browse_calc_file(self) -> None:
        current_calc = self._path_from_combo(self.calc_combo)
        start_dir = str(Path(current_calc).parent) if current_calc else self.project_path_text()
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel-файл", start_dir, "Excel (*.xlsx *.xlsm)")
        if path:
            index = self.owner._find_combo_path(self.calc_combo, path)
            if index < 0:
                self.owner._add_path_item(self.calc_combo, path, is_file=True)
                index = self.calc_combo.count() - 1
            self.calc_combo.setCurrentIndex(index)
            self.load_sheets()
            self.remember_values()

    def browse_template_file(self) -> None:
        current_template = self._path_from_combo(self.template_combo)
        start_dir = str(Path(current_template).parent) if current_template else self._saved("template_dir", self.project_path_text())
        path, _ = QFileDialog.getOpenFileName(self, "Выберите Word-шаблон", start_dir, "Word (*.docx)")
        if path:
            index = self.owner._find_combo_path(self.template_combo, path)
            if index < 0:
                self.owner._add_path_item(self.template_combo, path, is_file=True)
                index = self.template_combo.count() - 1
            self.template_combo.setCurrentIndex(index)
            self.settings.setValue("template_dir", str(Path(path).parent))
            self.remember_values()

    def scan_project(self, force: bool = False) -> None:
        project_text = self.project_path_text().strip()
        project_dir = Path(project_text) if project_text else None
        if not project_dir or not project_dir.exists():
            self.status_label.setText("Папка проекта не выбрана")
            return

        self.project_selector.scan_async(self._apply_project_scan, force=force)

    def _apply_project_scan(self, found: dict[str, list[Path]]) -> None:
        old_calc = self._path_from_combo(self.calc_combo)
        old_template = self._path_from_combo(self.template_combo)
        project_info = self.project_selector.project_info(found.get("pdf_dirs", []))

        self.calc_combo.blockSignals(True)
        self.calc_combo.clear()
        for p in found.get("excel", []):
            self.owner._add_path_item(self.calc_combo, str(p), is_file=True)
        old_calc_index = self.owner._find_combo_path(self.calc_combo, old_calc) if old_calc else -1
        if old_calc_index >= 0:
            self.calc_combo.setCurrentIndex(old_calc_index)
        self.calc_combo.blockSignals(False)

        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        if old_template and Path(old_template).exists():
            self.owner._add_path_item(self.template_combo, old_template, is_file=True)
        for p in found.get("word", []):
            if self.owner._find_combo_path(self.template_combo, str(p)) < 0:
                self.owner._add_path_item(self.template_combo, str(p), is_file=True)
        selected_template_index = self.owner._find_combo_path(self.template_combo, old_template) if old_template else -1
        if selected_template_index < 0 and not old_template and found.get("word"):
            newest_template = max(found["word"], key=lambda p: p.stat().st_mtime)
            selected_template_index = self.owner._find_combo_path(self.template_combo, str(newest_template))
        if selected_template_index >= 0:
            self.template_combo.setCurrentIndex(selected_template_index)
        self.template_combo.blockSignals(False)

        if not self.output_path_text().strip():
            guessed_output_dir = project_info.output_dir
            self.output_dir_path = guessed_output_dir
            self._set_line_path(self.output_edit, guessed_output_dir, is_file=False)

        current_spec_text = self.spec_path_text().strip()
        current_spec_dir = Path(current_spec_text) if current_spec_text else None
        should_guess_spec_dir = not current_spec_text or bool(current_spec_dir and not current_spec_dir.exists())
        if should_guess_spec_dir:
            guessed_spec_dir = project_info.specifications_dir
            self._set_spec_dir_path(guessed_spec_dir)

        spec_hint = self._display_dir(self.spec_dir_path) if self.spec_dir_path else "не выбрана"
        self.status_label.setText(
            f"Найдено Excel: {len(found.get('excel', []))}, Word: {len(found.get('word', []))}, "
            f"папок PDF: {len(found.get('pdf_dirs', []))}. Спецификации: {spec_hint}"
        )
        self.load_sheets()
        self.refresh_preview()
        self.refresh_calc_models_table()

    def load_sheets(self, refresh: bool = True) -> None:
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
        if refresh:
            self.refresh_preview()

    # ------------------------- autofill manager/client -------------------------
    def autofill_client_from_project_dir(self, force: bool = False) -> None:
        client = self.project_selector.client_name()
        if not client:
            return
        current = self.client_edit.text().strip()
        should_update = force or not current or current == "ТОО Example" or current == self._auto_client_value
        if not should_update:
            return
        self.client_edit.blockSignals(True)
        self.client_edit.setText(client)
        self.client_edit.blockSignals(False)
        self._auto_client_value = client
        self.refresh_preview()

    def autofill_manager_from_project(self, force: bool = False) -> None:
        if not force and self.owner._has_saved_manager_profile():
            return
        if not force and not self.owner._manager_profile().is_empty():
            return
        project_text = self.project_path_text().strip()
        project_dir = Path(project_text) if project_text else None
        if not project_dir or not project_dir.exists():
            return
        profile = find_manager_in_project(project_dir)
        if profile.is_empty():
            return
        self.owner._set_manager_profile(profile)
        self.settings.setValue("manager_name", profile.name)
        self.settings.setValue("manager_position", profile.position)
        self.settings.setValue("manager_email", profile.email)
        self.settings.setValue("manager_phone", profile.phone)
        self.settings.sync()
        self.status_label.setText("Данные исполнителя найдены в Word-файле проекта")

    # ------------------------- specifications -------------------------
    def open_spec_preview(self) -> None:
        """Compatibility handler for older UI integrations."""
        self.generate_calc_file()

    def _calc_model_table_state(self) -> dict[str, tuple[bool, str]]:
        state: dict[str, tuple[bool, str]] = {}
        table = self.calc_models_table
        for row in range(table.rowCount()):
            check = table.item(row, 0)
            model_item = table.item(row, 1)
            quantity = table.item(row, 2)
            meta = model_item.data(Qt.UserRole) if model_item else None
            key = str(meta.get("key") or "") if isinstance(meta, dict) else ""
            if key:
                state[key] = (
                    check.checkState() == Qt.Checked if check else True,
                    quantity.text().strip() if quantity else "",
                )
        return state

    def refresh_calc_models_table(self) -> None:
        """Load PDF Calc models into the first workflow step, independently of source Excel."""
        if not hasattr(self, "calc_models_table"):
            return
        if not getattr(self, "_calc_models_entries_ready", False):
            previous = self._calc_model_table_state()
            manual_paths = list(getattr(self, "_calc_manual_spec_files", []) or [])
            project_text = self.project_path_text().strip()
            spec_text = self.spec_path_text().strip()
            self.calc_models_status_label.setText("Идёт загрузка позиций Calc.pdf…")
            self.owner.run_background(
                lambda: self._discover_calc_entries(manual_paths, project_text, spec_text),
                lambda entries: self._continue_calc_models_refresh(entries, previous),
                self._show_calc_models_error,
            )
            return

        self._calc_models_entries_ready = False
        table = self.calc_models_table
        previous = getattr(self, "_calc_models_previous_state", {})
        entries = getattr(self, "_calc_models_preloaded_entries", [])
        table.blockSignals(True)
        try:
            table.setRowCount(0)
            if not entries:
                self.calc_models_status_label.setText(
                    "В этой папке не найдены Calc.pdf. Проверьте папку спецификаций "
                    "или выберите PDF вручную."
                )
                return

            for entry in entries:
                row = table.rowCount()
                table.insertRow(row)
                old_enabled, old_qty = previous.get(entry.key, (True, ""))

                check = QTableWidgetItem("")
                check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
                check.setCheckState(Qt.Checked if old_enabled else Qt.Unchecked)

                model_item = QTableWidgetItem(entry.model)
                model_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                model_item.setData(Qt.UserRole, {
                    "key": entry.key,
                    "model": entry.model,
                    "calc_pdf": str(entry.calc_pdf),
                    "unit_list_price": entry.unit_list_price,
                    "unit_purchase_price": entry.unit_purchase_price,
                    "currency": entry.currency,
                })
                model_item.setToolTip(str(entry.calc_pdf))

                qty_text = old_qty or self._format_qty_for_table(entry.quantity)
                qty_item = QTableWidgetItem(qty_text)
                qty_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                try:
                    quantity = float(qty_text.replace(",", "."))
                except ValueError:
                    quantity = 0.0

                table.setItem(row, 0, check)
                table.setItem(row, 1, model_item)
                table.setItem(row, 2, qty_item)
                table.setItem(row, 3, self._calc_money_item(
                    entry.unit_list_price * quantity if entry.unit_list_price is not None else None
                ))
                table.setItem(row, 4, self._calc_money_item(entry.unit_purchase_price))
                table.setItem(row, 5, self._calc_money_item(
                    entry.unit_purchase_price * quantity if entry.unit_purchase_price is not None else None
                ))
                table.setItem(row, 6, QTableWidgetItem(entry.currency))
                table.setItem(row, 7, QTableWidgetItem(entry.calc_pdf.name))

            table.resizeRowsToContents()
            self.calc_models_status_label.setText(
                f"Найдено Calc.pdf: {table.rowCount()}. Отметьте позиции для включения в новый файл."
            )
        except Exception as exc:
            table.setRowCount(0)
            self.calc_models_status_label.setText(f"Не удалось прочитать Calc.pdf: {exc}")
        finally:
            table.blockSignals(False)

    @staticmethod
    def _discover_calc_entries(manual_paths, project_text: str, spec_text: str):
        from core.stulz_spec_catalog import discover_stulz_spec_entries, load_stulz_spec_entry

        if manual_paths:
            return [entry for path in manual_paths if (entry := load_stulz_spec_entry(path)) is not None]
        project_dir = Path(project_text) if project_text else None
        roots = [Path(spec_text)] if spec_text else []
        if project_dir and project_dir.exists():
            roots.extend((Path(infer_specifications_dir(str(project_dir))), project_dir))
        entries = []
        seen_roots: set[str] = set()
        seen_entries: set[str] = set()
        for root in roots:
            root_key = str(root.resolve()) if root.exists() else str(root)
            if not root.exists() or root_key in seen_roots:
                continue
            seen_roots.add(root_key)
            for entry in discover_stulz_spec_entries(root):
                if entry.key not in seen_entries:
                    seen_entries.add(entry.key)
                    entries.append(entry)
        return entries

    def _continue_calc_models_refresh(self, entries, previous: dict[str, tuple[bool, str]]) -> None:
        self._calc_models_preloaded_entries = entries
        self._calc_models_previous_state = previous
        self._calc_models_entries_ready = True
        self.refresh_calc_models_table()

    def _show_calc_models_error(self, error: Exception) -> None:
        self.calc_models_table.setRowCount(0)
        self.calc_models_status_label.setText(f"Не удалось прочитать Calc.pdf: {error}")

    @staticmethod
    def _calc_money_item(value: object) -> QTableWidgetItem:
        item = QTableWidgetItem("")
        if value is not None:
            try:
                item.setText(f"{float(value):,.2f}".replace(",", " ").replace(".", ","))
            except (TypeError, ValueError):
                item.setText(str(value))
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        return item

    def on_calc_model_table_changed(self, item: QTableWidgetItem) -> None:
        if item.column() == 2:
            model_item = self.calc_models_table.item(item.row(), 1)
            meta = model_item.data(Qt.UserRole) if model_item else None
            if isinstance(meta, dict):
                try:
                    quantity = float(item.text().strip().replace(",", "."))
                except ValueError:
                    quantity = 0.0
                for column, key in ((3, "unit_list_price"), (5, "unit_purchase_price")):
                    unit_price = meta.get(key)
                    total = float(unit_price) * quantity if unit_price is not None else None
                    self.calc_models_table.setItem(item.row(), column, self._calc_money_item(total))

        selected_count = sum(
            1 for row in range(self.calc_models_table.rowCount())
            if self.calc_models_table.item(row, 0)
            and self.calc_models_table.item(row, 0).checkState() == Qt.Checked
        )
        self.calc_models_status_label.setText(
            f"Найдено Calc.pdf: {self.calc_models_table.rowCount()}. Отмечено для нового Calc: {selected_count}."
        )

    def selected_calc_models(self) -> list[dict[str, object]]:
        models: list[dict[str, object]] = []
        table = self.calc_models_table
        for row in range(table.rowCount()):
            check = table.item(row, 0)
            model_item = table.item(row, 1)
            qty_item = table.item(row, 2)
            meta = model_item.data(Qt.UserRole) if model_item else None
            if not isinstance(meta, dict):
                continue
            try:
                quantity = float(qty_item.text().strip().replace(",", ".")) if qty_item else 0.0
            except ValueError:
                quantity = 0.0
            models.append({
                **meta,
                "enabled": check.checkState() == Qt.Checked if check else True,
                "qty_value": quantity,
            })
        return models

    def browse_calc_model_pdfs(self) -> None:
        start_dir = self.spec_path_text() or self.project_path_text()
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите Calc.pdf", start_dir, "PDF (*.pdf)"
        )
        if paths:
            self._calc_manual_spec_files = paths
            self.refresh_calc_models_table()

    def clear_calc_model_manual_selection(self) -> None:
        self._calc_manual_spec_files = []
        self.refresh_calc_models_table()

    def generate_calc_file(self) -> None:
        try:
            from core.stulz_calc_writer import StulzCalcWriter

            self.validate_calc_parameters()
            self.refresh_calc_models_table()
            selected = [row for row in self.selected_calc_models() if row.get("enabled")]
            if not selected:
                raise ValueError("Отметьте хотя бы одну позицию в таблице на вкладке «1. Собрать расчёт».")

            values = self.calc_parameter_values()
            output_path = Path(values["Сохранить Calc"])
            if output_path.resolve() == Path(values["Шаблон Calc"]).resolve():
                raise ValueError("Нельзя сохранять результат поверх файла шаблона.")
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
            QApplication.processEvents()
            result = StulzCalcWriter().write(
                template_path=values["Шаблон Calc"],
                sheet_name=values["Лист шаблона"],
                output_path=output_path,
                models=selected,
                currency=values["Валюта"],
                delivery_terms=values["Условия поставки"],
                delivery_point=values["Пункт поставки"],
                exchange_rate=values["Курс"],
                vat_percent=values["НДС, %"],
                margin_percent=values["Маржа, %"],
            )
            self.remember_values()
            self.calc_models_status_label.setText(f"Calc сформирован: {result.name}")
            self._update_calc_open_buttons()
            message = QMessageBox(self)
            message.setWindowTitle("Calc сформирован")
            message.setIcon(QMessageBox.Information)
            message.setText("Расчёт успешно сформирован.")
            message.setInformativeText(f"Файл сохранён:\n{result}")
            open_file_action = message.addButton("Открыть расчёт", QMessageBox.ActionRole)
            open_folder_action = message.addButton("Открыть папку расчёта", QMessageBox.ActionRole)
            message.addButton("Закрыть", QMessageBox.AcceptRole)
            message.exec()
            clicked = message.clickedButton()
            if clicked == open_file_action:
                self.open_generated_calc()
            elif clicked == open_folder_action:
                self.open_calc_output_folder()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать Calc:\n{exc}")
        finally:
            self.spec_preview_button.setEnabled(True)
            self.spec_preview_button.setText("Сформировать расчёт")

    def clear_spec_models(self) -> None:
        self.spec_models_table.setRowCount(0)

    def current_spec_model_state(self) -> dict[str, tuple[bool, str]]:
        state: dict[str, tuple[bool, str]] = {}
        table = self.spec_models_table
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
        table = self.spec_models_table
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
            models = self._scan_calc_pdf_models(context.pdf_dir)
            if not models and context.project_dir.exists():
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

    # ------------------------- context/generate/preview -------------------------
    def make_context(self) -> OfferContext:
        project_dir = Path(self.offer_project_path_text().strip())
        output_dir = Path(self.output_path_text().strip() or project_dir)
        spec_text = infer_specifications_dir(str(project_dir)) if project_dir.exists() else ""
        pdf_dir = Path(spec_text) if spec_text and Path(spec_text).exists() else (project_dir if project_dir.exists() else None)
        calc_text = self._path_from_combo(self.offer_calc_combo).strip()
        template_text = self._path_from_combo(self.template_combo).strip()
        signer = self.owner._selected_signer()
        calc_parameters = self.calc_parameter_values()
        return OfferContext(
            brand=self.brand_name,
            project_dir=project_dir,
            template_path=Path(template_text) if template_text else Path("__template_not_selected__.docx"),
            calc_path=Path(calc_text) if calc_text else Path("__calc_not_selected__.xlsx"),
            output_dir=output_dir,
            client_name=self.client_edit.text().strip() or "Client",
            sheet_name=self.offer_sheet_combo.currentText().strip() or None,
            pdf_dir=pdf_dir,
            manager_name=self.owner.manager_name_edit.text().strip(),
            manager_position=self.owner.manager_position_edit.text().strip(),
            manager_email=self.owner.manager_email_edit.text().strip(),
            manager_phone=self.owner.manager_phone_edit.text().strip(),
            signer_name=signer["name"],
            signer_position=signer["position"],
            spec_models=self.selected_spec_models(),
            description_options=self.description_options(),
            brand_options={"stulz_calc_parameters": calc_parameters},
        )

    def validate_context(self, context: OfferContext) -> None:
        if not context.project_dir.exists():
            raise FileNotFoundError("Выберите существующую папку проекта.")
        if not context.calc_path.exists():
            raise FileNotFoundError("Выберите существующий Excel-файл калькуляции.")
        if not context.template_path.exists():
            raise FileNotFoundError("Выберите существующий Word-шаблон.")
        if context.pdf_dir and not context.pdf_dir.exists():
            raise FileNotFoundError("Выберите существующую папку спецификаций.")
        if context.template_path.suffix.lower() != ".docx":
            raise ValueError("Word-шаблон должен быть файлом .docx")

    def refresh_preview(self) -> None:
        try:
            context = self.make_context()
            module = self._ensure_runtime_layers_loaded()
            self.refresh_spec_models(context)
            if not context.calc_path.exists():
                self.preview.setPlainText("Excel-файл пока не выбран или не найден.")
                return
            module = module or get_brand_module(context.brand)
            self.preview.setPlainText(module.preview(context))
        except Exception as exc:
            self.preview.setPlainText(f"Не удалось прочитать данные: {exc}")
            self.refresh_spec_models(None)

    def generate(self) -> None:
        try:
            self.generate_btn.setEnabled(False)
            self.status_label.setText("Формирую КП...")
            QApplication.processEvents()

            context = self.make_context()
            self.validate_context(context)
            module = get_brand_module(context.brand)
            out = module.make_offer(context)

            self.remember_values()
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
            self.refresh_preview()

    # ------------------------- persistence/page hooks -------------------------
    def remember_values(self) -> None:
        self._set_saved("project_dir", self.project_path_text())
        self._set_saved("stulz/offer_project_dir", self.offer_project_path_text())
        self._set_saved("stulz/offer_calc_path", self._path_from_combo(self.offer_calc_combo))
        self._set_saved("stulz/offer_sheet_name", self.offer_sheet_combo.currentText().strip())
        self._set_saved("client", self.client_edit.text().strip())
        self._set_saved("calc_path", self._path_from_combo(self.calc_combo))
        self._set_saved("template_path", self._path_from_combo(self.template_combo))
        self._set_saved("sheet_name", self.sheet_combo.currentText().strip())
        self._set_saved("output_dir", self.output_path_text())
        self._set_saved("spec_dir", self.spec_path_text())
        self._set_saved("calc_template_path", self.calc_template_edit.text().strip())
        self._set_saved("calc_template_sheet", str(self.calc_template_sheet_combo.currentData() or "").strip())
        self._set_saved("calc_output_path", self.calc_output_edit.text().strip())
        self._set_saved("calc_delivery_terms", self.calc_delivery_combo.currentText().strip())
        self._set_saved("calc_delivery_point", self.calc_delivery_point_combo.currentText().strip())
        self._set_saved("calc_currency", str(self.calc_currency_combo.currentData() or "").strip().upper())
        self._set_saved("calc_exchange_rate", self.calc_exchange_rate_edit.text().strip())
        self._set_saved("calc_vat_percent", self.calc_vat_edit.text().strip())
        self._set_saved("calc_margin_percent", self.calc_margin_edit.text().strip())
        self.settings.sync()

    def on_settings_changed(self) -> None:
        template_path = self._saved("template_path", "")
        if template_path:
            index = self.owner._find_combo_path(self.template_combo, template_path)
            if index < 0:
                self.owner._add_path_item(self.template_combo, template_path, is_file=True)
                index = self.template_combo.count() - 1
            self.template_combo.setCurrentIndex(index)
        self.refresh_preview()

    def clear_cache(self) -> None:
        self.project_dir_path = ""
        self.output_dir_path = ""
        self.spec_dir_path = ""
        self._set_line_path(self.project_edit, "", is_file=False)
        self.offer_project_selector.set_project_path("")
        self._set_line_path(self.output_edit, "", is_file=False)
        self._set_line_path(self.spec_edit, "", is_file=False)
        default_calc_template = resource_path(Path("templates") / "calc exel" / "Calc_14-08-26 Stulz.xlsx")
        self.calc_template_edit.setText(str(default_calc_template) if default_calc_template.exists() else "")
        self.calc_template_edit.setToolTip(self.calc_template_edit.text().strip())
        self.load_calc_template_sheets(refresh=False)
        self.calc_output_edit.clear()
        self.calc_delivery_combo.setCurrentIndex(0)
        self.calc_delivery_point_combo.setCurrentIndex(0)
        self.calc_currency_combo.setCurrentIndex(0)
        self.calc_exchange_rate_edit.clear()
        self.calc_vat_edit.clear()
        self.calc_margin_edit.clear()
        self.calc_combo.clear()
        self.offer_calc_combo.clear()
        self.offer_sheet_combo.clear()
        self.template_combo.clear()
        self.sheet_combo.clear()
        self.clear_spec_models()
        self.calc_models_table.setRowCount(0)
        self._calc_manual_spec_files = []
        self.calc_models_status_label.setText("Кэш очищен. Выберите папку проекта заново.")
        self.preview.setPlainText("Кэш очищен. Выберите папку проекта заново.")
        self.status_label.setText("Кэш очищен. Выберите папку проекта заново.")
        for key in (
            "calc_template_path",
            "calc_template_sheet",
            "calc_output_path",
            "calc_delivery_terms",
            "calc_delivery_point",
            "calc_currency",
            "calc_exchange_rate",
            "calc_vat_percent",
            "calc_margin_percent",
            "stulz/offer_project_dir",
            "stulz/offer_calc_path",
            "stulz/offer_sheet_name",
        ):
            self.settings.remove(key)
        self.settings.sync()

    def apply_responsive_metrics(self, scale: float) -> None:
        self.generate_btn.setMinimumWidth(int(220 * scale))
        self.generate_btn.setMinimumHeight(int(42 * scale))
