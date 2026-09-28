from __future__ import annotations

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from brands.dc_eltek import (
    detect_dc_eltek_currency,
    find_default_dc_eltek_template,
    make_offer as make_dc_eltek_offer_ru,
    preview as build_dc_eltek_preview,
)
from brands.dc_eltek_en import make_offer as make_dc_eltek_offer_en


PROJECTS_MARKER = "02_Projects"
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
SALES_DIR_MARKERS = (
    "sales",
    "sale",
    "commercial",
    "коммер",
    "кп",
)
CALC_FILE_MARKERS = (
    "calc",
    "calculation",
    "расчет",
    "расчёт",
)


def extract_client_from_project_path(path_text: str) -> str:
    if not path_text:
        return ""
    parts = [part for part in path_text.replace("/", "\\").split("\\") if part]
    for index, part in enumerate(parts):
        if part.lower() == PROJECTS_MARKER.lower() and index + 1 < len(parts):
            return parts[index + 1].strip()
    return ""


def read_excel_sheet_names(path_text: str) -> list[str]:
    path = Path(path_text)
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        raise ValueError("Пока поддерживаются только Excel-файлы .xlsx и .xlsm")
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    with zipfile.ZipFile(path) as archive:
        workbook_xml = archive.read("xl/workbook.xml")
        namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        root = ET.fromstring(workbook_xml)
        sheets = root.find("main:sheets", namespace)
        if sheets is None:
            return []
        return [
            name
            for sheet in sheets.findall("main:sheet", namespace)
            if (name := sheet.attrib.get("name", "").strip())
        ]


def _safe_exists(path: Path) -> bool:
    try:
        return path.exists()
    except Exception:
        return False


def _safe_is_dir(path: Path) -> bool:
    try:
        return path.is_dir()
    except Exception:
        return False


def _safe_is_file(path: Path) -> bool:
    try:
        return path.is_file()
    except Exception:
        return False


def _safe_iterdir(path: Path):
    try:
        yield from path.iterdir()
    except Exception:
        return


def _safe_mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except Exception:
        return 0.0


def _is_sales_dir_name(name: str) -> bool:
    normalized = name.lower().replace("_", " ").replace("-", " ")
    return any(marker in normalized for marker in SALES_DIR_MARKERS)


def _is_excel_calc_candidate(path: Path) -> bool:
    if path.name.startswith("~$"):
        return False
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        return False
    if not _safe_is_file(path):
        return False
    return True


def _iter_dirs_limited(root: Path, max_depth: int = 4):
    stack: list[tuple[Path, int]] = [(root, 0)]
    seen: set[str] = set()
    while stack:
        current, depth = stack.pop()
        current_key = str(current).lower()
        if current_key in seen:
            continue
        seen.add(current_key)
        yield current, depth
        if depth >= max_depth:
            continue
        for child in _safe_iterdir(current):
            if _safe_is_dir(child):
                stack.append((child, depth + 1))


def _find_sales_dirs(project_dir: Path) -> list[Path]:
    found: list[Path] = []
    seen: set[str] = set()

    def add(path: Path) -> None:
        key = str(path).lower()
        if key not in seen and _safe_is_dir(path):
            seen.add(key)
            found.append(path)

    for name in (
        "sales",
        "Sales",
        "03_sales docs",
        "02_sales docs",
        "sales docs",
        "Sales docs",
    ):
        add(project_dir / name)

    for folder, _depth in _iter_dirs_limited(project_dir, max_depth=4):
        if folder == project_dir:
            continue
        if _is_sales_dir_name(folder.name):
            add(folder)

    return found


def _score_calc_candidate(path: Path, sales_dir: Path) -> tuple[int, float, str]:
    name = path.name.lower()
    stem = path.stem.lower()
    score = 0
    if any(marker in stem for marker in CALC_FILE_MARKERS):
        score += 1000
    if "dc" in stem or "eltek" in stem:
        score += 200
    if "hvac" in stem:
        score -= 250
    if "offer" in stem or "кп" in stem:
        score -= 200
    if path.parent == sales_dir:
        score += 80
    if name.startswith("calc"):
        score += 80
    return (score, _safe_mtime(path), str(path))


def find_calc_file_in_sales(project_dir_text: str) -> str:
    project_dir = Path(project_dir_text.strip()) if project_dir_text else Path()
    if not project_dir_text or not _safe_is_dir(project_dir):
        return ""

    sales_dirs = _find_sales_dirs(project_dir)
    search_dirs = sales_dirs or [project_dir]
    candidates: list[tuple[tuple[int, float, str], Path]] = []

    for sales_dir in search_dirs:
        for folder, depth in _iter_dirs_limited(sales_dir, max_depth=3):
            if depth > 3:
                continue
            for child in _safe_iterdir(folder):
                if _is_excel_calc_candidate(child):
                    candidates.append((_score_calc_candidate(child, sales_dir), child))

    if not candidates:
        return ""

    candidates.sort(key=lambda item: item[0], reverse=True)
    return str(candidates[0][1])


class DcEltekPage(QWidget):
    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self.settings = owner.settings
        self._setting_currency_programmatically = False
        self.last_output_path = self._saved("dc_eltek_last_output_path", "")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        card = owner._card("DC Eltek")
        form = QGridLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(10)
        form.setColumnStretch(1, 1)
        card.layout().addLayout(form)

        self.project_dir_edit = QLineEdit(self._saved("dc_eltek_project_dir", ""))
        self.project_dir_edit.setPlaceholderText("Папка проекта")
        self.project_dir_edit.editingFinished.connect(self._on_project_dir_changed)
        owner._add_row(form, 0, "Папка проекта", self.project_dir_edit, "Выбрать", self.select_project_dir)

        self.client_edit = QLineEdit(self._saved("dc_eltek_client", ""))
        self.client_edit.setPlaceholderText("Клиент")
        self.client_edit.editingFinished.connect(self._on_field_changed)
        owner._add_row(form, 1, "Клиент", self.client_edit, None, None)

        self.calc_path_edit = QLineEdit(self._saved("dc_eltek_calc_path", ""))
        self.calc_path_edit.setPlaceholderText("Excel calc из папки sales")
        self.calc_path_edit.setReadOnly(True)
        owner._add_row(form, 2, "Расчёт Excel", self.calc_path_edit, None, None)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setEditable(False)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        owner._add_row(form, 3, "Лист для КП", self.sheet_combo, None, None)

        self.currency_combo = QComboBox()
        self.currency_combo.addItem("Не указана", "")
        self.currency_combo.addItem("KZT", "KZT")
        self.currency_combo.addItem("EUR", "EUR")
        self.currency_combo.addItem("USD", "USD")
        self.currency_combo.currentIndexChanged.connect(self._on_currency_changed)
        owner._add_row(form, 4, "Валюта", self.currency_combo, None, None)

        saved_template = self._saved("dc_eltek_template_path", "") or find_default_dc_eltek_template()
        self.template_path_edit = QLineEdit(saved_template)
        self.template_path_edit.setPlaceholderText("Шаблон КП .docx")
        self.template_path_edit.editingFinished.connect(self._on_field_changed)
        owner._add_row(form, 5, "Шаблон КП", self.template_path_edit, "Выбрать", self.select_template_file)

        buttons = QHBoxLayout()
        buttons.setContentsMargins(0, 0, 0, 0)
        buttons.setSpacing(8)

        self.generate_ru_btn = QPushButton("Сформировать КП RU")
        self.generate_ru_btn.setObjectName("PrimaryButton")
        self.generate_ru_btn.clicked.connect(lambda: self.generate_offer("ru"))
        buttons.addWidget(self.generate_ru_btn, 2)

        self.generate_en_btn = QPushButton("Generate offer EN")
        self.generate_en_btn.clicked.connect(lambda: self.generate_offer("en"))
        buttons.addWidget(self.generate_en_btn, 2)

        self.refresh_btn = QPushButton("Обновить")
        self.refresh_btn.clicked.connect(self.refresh_data)
        buttons.addWidget(self.refresh_btn, 1)

        self.open_offer_btn = QPushButton("Открыть КП")
        self.open_offer_btn.clicked.connect(self.open_generated_offer)
        buttons.addWidget(self.open_offer_btn, 1)

        self.open_folder_btn = QPushButton("Открыть папку")
        self.open_folder_btn.clicked.connect(self.open_generated_folder)
        buttons.addWidget(self.open_folder_btn, 1)

        card.layout().addLayout(buttons)

        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setMinimumHeight(260)
        card.layout().addWidget(self.preview_box)

        layout.addWidget(card)
        layout.addStretch(1)

        if self.project_dir_edit.text().strip() and not self.calc_path_edit.text().strip():
            self._auto_find_calc_from_project(show_warning=False)
        self._load_sheet_names(initial=True)
        self._restore_or_detect_currency()
        self._update_open_buttons()
        self._update_preview()

    def _saved(self, key: str, default: str) -> str:
        value = self.settings.value(key, default)
        return str(value) if value is not None else default

    def project_path_text(self) -> str:
        return self.project_dir_edit.text().strip()

    def _calc_output_dir(self) -> str:
        calc_path = self.calc_path_edit.text().strip()
        if calc_path:
            return str(Path(calc_path).parent)
        return self.project_dir_edit.text().strip()

    def _currency_value(self) -> str:
        return str(self.currency_combo.currentData() or "").upper().strip()

    def _set_currency_value(self, value: str) -> None:
        value = (value or "").upper().strip()
        self._setting_currency_programmatically = True
        try:
            for index in range(self.currency_combo.count()):
                if str(self.currency_combo.itemData(index) or "").upper() == value:
                    self.currency_combo.setCurrentIndex(index)
                    return
            self.currency_combo.setCurrentIndex(0)
        finally:
            self._setting_currency_programmatically = False

    def remember_values(self) -> None:
        self.settings.setValue("brand", "DC Eltek")
        self.settings.setValue("dc_eltek_project_dir", self.project_dir_edit.text().strip())
        self.settings.setValue("dc_eltek_client", self.client_edit.text().strip())
        self.settings.setValue("dc_eltek_calc_path", self.calc_path_edit.text().strip())
        self.settings.setValue("dc_eltek_sheet_name", self.sheet_combo.currentText().strip())
        self.settings.setValue("dc_eltek_currency", self._currency_value())
        self.settings.setValue("dc_eltek_template_path", self.template_path_edit.text().strip())
        self.settings.setValue("dc_eltek_last_output_path", self.last_output_path)
        self.settings.sync()

    def clear_cache(self) -> None:
        for key in (
            "dc_eltek_project_dir",
            "dc_eltek_client",
            "dc_eltek_calc_path",
            "dc_eltek_sheet_name",
            "dc_eltek_currency",
            "dc_eltek_template_path",
            "dc_eltek_output_dir",
            "dc_eltek_last_output_path",
        ):
            self.settings.remove(key)
        self.project_dir_edit.clear()
        self.client_edit.clear()
        self.calc_path_edit.clear()
        self.sheet_combo.clear()
        self._set_currency_value("")
        self.template_path_edit.setText(find_default_dc_eltek_template())
        self.last_output_path = ""
        self.preview_box.clear()
        self._update_open_buttons()
        self.settings.sync()

    def apply_responsive_metrics(self, scale: float) -> None:
        self.preview_box.setMinimumHeight(int(260 * scale))

    def on_settings_changed(self) -> None:
        self._update_preview()

    def _on_field_changed(self) -> None:
        self.remember_values()
        self._update_preview()

    def _on_currency_changed(self) -> None:
        if self._setting_currency_programmatically:
            return
        self.remember_values()
        self._update_preview()

    def _on_sheet_changed(self, _text: str = "") -> None:
        self._auto_detect_currency(force=True)
        self.remember_values()
        self._update_preview()

    def _reload_calc_dependent_data(self, *, initial: bool = False) -> None:
        self._load_sheet_names(initial=initial)
        self._auto_detect_currency(force=True)
        self.remember_values()
        self._update_preview()
        self._update_open_buttons()

    def select_project_dir(self) -> None:
        current = self.project_dir_edit.text().strip() or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта", current)
        if not path:
            return
        self.project_dir_edit.setText(path)
        self._on_project_dir_changed(force_client=True, show_calc_warning=True)

    def _on_project_dir_changed(self, force_client: bool = False, show_calc_warning: bool = False) -> None:
        path_text = self.project_dir_edit.text().strip()
        extracted_client = extract_client_from_project_path(path_text)
        if extracted_client and (force_client or not self.client_edit.text().strip()):
            self.client_edit.setText(extracted_client)
        self._auto_find_calc_from_project(show_warning=show_calc_warning)
        self.remember_values()
        self._update_preview()
        self._update_open_buttons()

    def _auto_find_calc_from_project(self, show_warning: bool = False) -> bool:
        project_dir = self.project_dir_edit.text().strip()
        if not project_dir:
            return False

        calc_path = find_calc_file_in_sales(project_dir)
        if not calc_path:
            self.calc_path_edit.clear()
            self.sheet_combo.clear()
            self._set_currency_value("")
            self.last_output_path = ""
            if show_warning:
                QMessageBox.warning(
                    self,
                    "DC Eltek",
                    "В папке проекта не найден Excel calc в папке sales / sales docs.",
                )
            return False

        if self.calc_path_edit.text().strip() != calc_path:
            self.calc_path_edit.setText(calc_path)
            self.last_output_path = ""
        self._reload_calc_dependent_data(initial=False)
        return True

    def select_template_file(self) -> None:
        start_dir = self.project_dir_edit.text().strip() or str(Path.home())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите шаблон КП",
            start_dir,
            "Word templates (*.docx);;All files (*.*)",
        )
        if not path:
            return
        self.template_path_edit.setText(path)
        self.remember_values()
        self._update_preview()

    def refresh_data(self) -> None:
        if self.project_dir_edit.text().strip():
            self._auto_find_calc_from_project(show_warning=False)
        else:
            self._reload_calc_dependent_data(initial=False)
        self._update_open_buttons()

    def _load_sheet_names(self, initial: bool) -> None:
        current_sheet = self._saved("dc_eltek_sheet_name", "") if initial else self.sheet_combo.currentText().strip()
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        calc_path = self.calc_path_edit.text().strip()
        if not calc_path:
            self.sheet_combo.blockSignals(False)
            return
        try:
            sheet_names = read_excel_sheet_names(calc_path)
        except Exception as exc:
            self.sheet_combo.blockSignals(False)
            if not initial:
                QMessageBox.warning(self, "DC Eltek", f"Не удалось прочитать листы Excel:\n{exc}")
            return
        self.sheet_combo.addItems(sheet_names)
        if current_sheet:
            index = self.sheet_combo.findText(current_sheet, Qt.MatchFixedString)
            if index >= 0:
                self.sheet_combo.setCurrentIndex(index)
        self.sheet_combo.blockSignals(False)

    def _restore_or_detect_currency(self) -> None:
        saved_currency = self._saved("dc_eltek_currency", "").upper().strip()
        if saved_currency:
            self._set_currency_value(saved_currency)
        else:
            self._auto_detect_currency(force=True)

    def _auto_detect_currency(self, force: bool = False) -> None:
        calc_path = self.calc_path_edit.text().strip()
        sheet_name = self.sheet_combo.currentText().strip()
        if not calc_path or not sheet_name:
            if force:
                self._set_currency_value("")
            return
        try:
            detected = detect_dc_eltek_currency(calc_path, sheet_name)
        except Exception:
            detected = ""
        if force or detected:
            self._set_currency_value(detected)

    def _selected_signer(self) -> dict[str, str]:
        if hasattr(self.owner, "_selected_signer"):
            try:
                return dict(self.owner._selected_signer())
            except Exception:
                pass
        return {"name": "Сания Санаткызы", "position": "Коммерческий директор"}

    def _manager_profile(self):
        if hasattr(self.owner, "_manager_profile"):
            try:
                return self.owner._manager_profile()
            except Exception:
                pass
        return None

    def _context_dict(self, language: str = "ru") -> dict[str, str]:
        signer = self._selected_signer()
        manager = self._manager_profile()
        return {
            "project_dir": self.project_dir_edit.text().strip(),
            "output_dir": self._calc_output_dir(),
            "client": self.client_edit.text().strip(),
            "calc_path": self.calc_path_edit.text().strip(),
            "sheet_name": self.sheet_combo.currentText().strip(),
            "currency": self._currency_value(),
            "template_path": self.template_path_edit.text().strip(),
            "offer_language": language,
            "signer_name": str(signer.get("name", "")),
            "signer_position": str(signer.get("position", "")),
            "manager_name": str(getattr(manager, "name", "") if manager else ""),
            "manager_position": str(getattr(manager, "position", "") if manager else ""),
            "manager_email": str(getattr(manager, "email", "") if manager else ""),
            "manager_phone": str(getattr(manager, "phone", "") if manager else ""),
        }

    def _update_preview(self) -> None:
        self.preview_box.setPlainText(build_dc_eltek_preview(self._context_dict("ru")))

    def _update_open_buttons(self) -> None:
        path = Path(self.last_output_path) if self.last_output_path else None
        exists = bool(path and path.exists())
        self.open_offer_btn.setEnabled(exists)
        calc_folder = Path(self._calc_output_dir()) if self._calc_output_dir() else None
        self.open_folder_btn.setEnabled(bool((path and path.parent.exists()) or (calc_folder and calc_folder.exists())))

    def generate_offer(self, language: str = "ru") -> None:
        self.remember_values()
        language = "en" if language == "en" else "ru"
        data = self._context_dict(language)
        missing: list[str] = []
        if not data["project_dir"]:
            missing.append("папка проекта")
        if not data["client"]:
            missing.append("клиент")
        if not data["calc_path"]:
            missing.append("Excel calc")
        if not data["sheet_name"]:
            missing.append("лист для КП")
        if not data["template_path"]:
            missing.append("шаблон КП")
        if missing:
            QMessageBox.warning(self, "DC Eltek", "Заполните поля:\n- " + "\n- ".join(missing))
            return
        if not data["currency"]:
            QMessageBox.warning(
                self,
                "DC Eltek",
                "В расчёте не удалось определить валюту. Выберите валюту вручную перед формированием КП.",
            )
            return
        try:
            maker = make_dc_eltek_offer_en if language == "en" else make_dc_eltek_offer_ru
            result = maker(data)
        except Exception as exc:
            QMessageBox.critical(self, "DC Eltek", f"Не удалось сформировать КП:\n{exc}")
            return
        if isinstance(result, dict):
            output_value = result.get("output_path", "")
        else:
            output_value = result
        output_path = Path(str(output_value)).resolve()
        self.last_output_path = str(output_path)
        self.remember_values()
        self._update_preview()
        self._update_open_buttons()
        title = "КП сформировано" if language == "ru" else "English offer generated"
        QMessageBox.information(self, "DC Eltek", f"{title}:\n{output_path}")

    def open_generated_offer(self) -> None:
        if not self.last_output_path:
            return
        path = Path(self.last_output_path)
        if not path.exists():
            QMessageBox.warning(self, "DC Eltek", "Файл КП не найден. Сформируйте КП заново.")
            self._update_open_buttons()
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def open_generated_folder(self) -> None:
        folder: Path | None = None
        if self.last_output_path:
            path = Path(self.last_output_path)
            if path.parent.exists():
                folder = path.parent
        if folder is None and self._calc_output_dir():
            candidate = Path(self._calc_output_dir())
            if candidate.exists():
                folder = candidate
        if folder is None:
            QMessageBox.warning(self, "DC Eltek", "Папка не найдена.")
            self._update_open_buttons()
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))
