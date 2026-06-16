from __future__ import annotations

from pathlib import Path

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
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from brands.registry import get_brand_module
from core.excel_reader import list_sheets
from core.manager_profile import find_manager_in_project
from core.models import OfferContext
from core.project_scanner import clear_scan_cache, scan_project_files
from gui.path_helpers import (
    extract_brand_from_project_dir,
    extract_client_from_project_dir,
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

        self.project_edit = QLineEdit(self._display_dir(self.project_dir_path))
        self.project_edit.setToolTip(self.project_dir_path)
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
        self.status_label = QLabel("Выберите папку проекта")

        self._load_saved_combo_paths()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        header = QHBoxLayout()
        h_text = QVBoxLayout()
        page_title = QLabel("Stulz: новое коммерческое предложение")
        page_title.setObjectName("PageTitle")
        page_subtitle = QLabel("Выберите папку проекта. Страница Stulz сама найдет Excel, Word и PDF-спецификации внутри папки.")
        page_subtitle.setObjectName("PageSubtitle")
        h_text.addWidget(page_title)
        h_text.addWidget(page_subtitle)
        self.generate_btn = QPushButton("Сформировать КП")
        self.generate_btn.setObjectName("PrimaryButton")
        self.generate_btn.clicked.connect(self.generate)
        header.addLayout(h_text, stretch=1)
        header.addWidget(self.generate_btn, stretch=0, alignment=Qt.AlignTop)
        layout.addLayout(header)

        project_card = owner._card("Папка проекта")
        project_grid = QGridLayout()
        project_card.layout().addLayout(project_grid)
        project_grid.setColumnStretch(1, 1)
        project_grid.setVerticalSpacing(12)
        project_grid.setHorizontalSpacing(10)
        owner._add_row(project_grid, 0, "Папка проекта", self.project_edit, "Выбрать", self.browse_project_dir)
        layout.addWidget(project_card)

        files_card = owner._card("Stulz: файлы и параметры КП")
        grid = QGridLayout()
        files_card.layout().addLayout(grid)
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)

        owner._add_row(grid, 0, "Клиент", self.client_edit, None, None)
        owner._add_row(grid, 1, "Excel-расчет", self.calc_combo, "Обновить", lambda: self.scan_project(force=True))
        owner._add_row(grid, 2, "Лист Excel", self.sheet_combo, "Листы", self.load_sheets)
        owner._add_row(grid, 3, "Word-шаблон", self.template_combo, "Выбрать", self.browse_template_file)
        owner._add_row(grid, 4, "Папка спецификаций", self.spec_edit, "Выбрать", self.browse_spec_dir)
        owner._add_row(grid, 5, "Папка результата", self.output_edit, "Выбрать", self.browse_output_dir)
        layout.addWidget(files_card)

        bottom = QHBoxLayout()
        preview_card = owner._card("Проверка данных")
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        preview_card.layout().addWidget(self.preview)
        bottom.addWidget(preview_card, stretch=2)

        spec_card = owner._card("Спецификации")
        spec_hint = QLabel(
            "Модели берутся из выбранной папки спецификаций, а не из Excel КП. "
            "Количество берется из Calc.pdf. Можно отключать позиции и менять количество."
        )
        spec_hint.setObjectName("Hint")
        spec_hint.setWordWrap(True)
        spec_card.layout().addWidget(spec_hint)

        self.spec_models_table = QTableWidget(0, 3)
        self.spec_models_table.setHorizontalHeaderLabels(["Вкл", "Модель", "Кол-во"])
        self.spec_models_table.verticalHeader().setVisible(False)
        self.spec_models_table.setAlternatingRowColors(True)
        self.spec_models_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.spec_models_table.setMinimumHeight(170)
        self.spec_models_table.setColumnWidth(0, 52)
        self.spec_models_table.setColumnWidth(2, 80)
        self.spec_models_table.horizontalHeader().setStretchLastSection(False)
        self.spec_models_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        spec_card.layout().addWidget(self.spec_models_table)

        self.spec_preview_button = QPushButton("Предпросмотр спецификаций")
        self.spec_preview_button.setMinimumHeight(42)
        self.spec_preview_button.setObjectName("PrimaryButton")
        self.spec_preview_button.clicked.connect(self.open_spec_preview)
        spec_card.layout().addWidget(self.spec_preview_button)
        spec_card.layout().addWidget(self.status_label)
        bottom.addWidget(spec_card, stretch=1)
        layout.addLayout(bottom)

        self._connect_changes()
        self.load_sheets(refresh=False)
        self.scan_project(force=False)
        self.refresh_preview()

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
        return self.project_dir_path or self.project_edit.toolTip() or self.project_edit.text().strip()

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

    def _connect_changes(self) -> None:
        self.project_edit.textChanged.connect(self.on_project_dir_changed)
        self.client_edit.textChanged.connect(self.refresh_preview)
        self.calc_combo.currentTextChanged.connect(lambda _text: self.load_sheets())
        self.calc_combo.currentTextChanged.connect(self.refresh_preview)
        self.template_combo.currentTextChanged.connect(self.refresh_preview)
        self.sheet_combo.currentTextChanged.connect(self.refresh_preview)
        self.output_edit.textChanged.connect(self.on_output_dir_changed)
        self.spec_edit.textChanged.connect(self.on_spec_dir_changed)

        for widget_signal in (
            self.project_edit.textChanged,
            self.client_edit.textChanged,
            self.output_edit.textChanged,
            self.spec_edit.textChanged,
            self.calc_combo.currentTextChanged,
            self.template_combo.currentTextChanged,
            self.sheet_combo.currentTextChanged,
        ):
            widget_signal.connect(lambda *_: self.remember_values())

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

    def browse_project_dir(self) -> None:
        old_project = self.project_path_text().strip()
        path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта", old_project)
        if not path:
            return

        self.project_dir_path = path
        self._set_line_path(self.project_edit, path, is_file=False)

        if path != old_project:
            self.spec_dir_path = ""
            self.output_dir_path = ""
            self._set_line_path(self.spec_edit, "", is_file=False)
            self._set_line_path(self.output_edit, "", is_file=False)

        self.autofill_client_from_project_dir(force=True)
        brand = extract_brand_from_project_dir(path, tuple(["Stulz", "Riello", "DC Eltek", "Generator"]))
        if brand and brand != self.brand_name:
            self.owner._select_tab_for_brand(brand)
        self.scan_project(force=True)
        self.autofill_manager_from_project(force=False)

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

        if force:
            clear_scan_cache()

        found = scan_project_files(project_dir, use_cache=not force)
        old_calc = self._path_from_combo(self.calc_combo)
        old_template = self._path_from_combo(self.template_combo)

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
            guessed_output_dir = infer_output_dir(str(project_dir))
            self.output_dir_path = guessed_output_dir
            self._set_line_path(self.output_edit, guessed_output_dir, is_file=False)

        current_spec_text = self.spec_path_text().strip()
        current_spec_dir = Path(current_spec_text) if current_spec_text else None
        should_guess_spec_dir = not current_spec_text or bool(current_spec_dir and not current_spec_dir.exists())
        if should_guess_spec_dir:
            guessed_spec_dir = infer_specifications_dir(str(project_dir), found.get("pdf_dirs", []))
            self._set_spec_dir_path(guessed_spec_dir)

        spec_hint = self._display_dir(self.spec_dir_path) if self.spec_dir_path else "не выбрана"
        self.status_label.setText(
            f"Найдено Excel: {len(found.get('excel', []))}, Word: {len(found.get('word', []))}, "
            f"папок PDF: {len(found.get('pdf_dirs', []))}. Спецификации: {spec_hint}"
        )
        self.load_sheets()
        self.refresh_preview()

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
        client = extract_client_from_project_dir(self.project_path_text())
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
        try:
            from brands.stulz import build_specification_blocks, load_calc
            from gui.spec_preview_dialog import SpecPreviewDialog

            context = self.make_context()
            self.validate_context(context)
            calc = load_calc(context)
            spec_blocks, warnings = build_specification_blocks(context, calc)
            dialog = SpecPreviewDialog(
                spec_blocks,
                warnings,
                self,
                description_options=self.description_options(),
            )
            dialog.exec()
            self._save_description_options(dialog.description_options())
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть предпросмотр спецификаций:\n{exc}")

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
        project_dir = Path(self.project_path_text().strip())
        output_dir = Path(self.output_path_text().strip() or project_dir)
        spec_text = self.spec_path_text().strip()
        pdf_dir = Path(spec_text) if spec_text else (project_dir if project_dir.exists() else None)
        calc_text = self._path_from_combo(self.calc_combo).strip()
        template_text = self._path_from_combo(self.template_combo).strip()
        signer = self.owner._selected_signer()
        return OfferContext(
            brand=self.brand_name,
            project_dir=project_dir,
            template_path=Path(template_text) if template_text else Path("__template_not_selected__.docx"),
            calc_path=Path(calc_text) if calc_text else Path("__calc_not_selected__.xlsx"),
            output_dir=output_dir,
            client_name=self.client_edit.text().strip() or "Client",
            sheet_name=self.sheet_combo.currentText().strip() or None,
            pdf_dir=pdf_dir,
            manager_name=self.owner.manager_name_edit.text().strip(),
            manager_position=self.owner.manager_position_edit.text().strip(),
            manager_email=self.owner.manager_email_edit.text().strip(),
            manager_phone=self.owner.manager_phone_edit.text().strip(),
            signer_name=signer["name"],
            signer_position=signer["position"],
            spec_models=self.selected_spec_models(),
            description_options=self.description_options(),
            brand_options={},
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
            self.refresh_spec_models(context)
            if not context.calc_path.exists():
                self.preview.setPlainText("Excel-файл пока не выбран или не найден.")
                return
            module = get_brand_module(context.brand)
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
        self._set_saved("client", self.client_edit.text().strip())
        self._set_saved("calc_path", self._path_from_combo(self.calc_combo))
        self._set_saved("template_path", self._path_from_combo(self.template_combo))
        self._set_saved("sheet_name", self.sheet_combo.currentText().strip())
        self._set_saved("output_dir", self.output_path_text())
        self._set_saved("spec_dir", self.spec_path_text())
        self._set_saved("brand", self.brand_name)
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
        self._set_line_path(self.output_edit, "", is_file=False)
        self._set_line_path(self.spec_edit, "", is_file=False)
        self.calc_combo.clear()
        self.template_combo.clear()
        self.sheet_combo.clear()
        self.clear_spec_models()
        self.preview.setPlainText("Кэш очищен. Выберите папку проекта заново.")
        self.status_label.setText("Кэш очищен. Выберите папку проекта заново.")

    def apply_responsive_metrics(self, scale: float) -> None:
        self.generate_btn.setMinimumWidth(int(220 * scale))
        self.generate_btn.setMinimumHeight(int(42 * scale))
