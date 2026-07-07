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
    make_offer as make_dc_eltek_offer,
    preview as build_dc_eltek_preview,
)


PROJECTS_MARKER = "02_Projects"


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
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
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
        self.calc_path_edit.setPlaceholderText("Excel calc")
        self.calc_path_edit.editingFinished.connect(self._on_calc_path_edited)
        owner._add_row(form, 2, "Расчёт Excel", self.calc_path_edit, "Выбрать", self.select_calc_file)

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

        saved_output_dir = self._saved("dc_eltek_output_dir", "")
        if not saved_output_dir:
            saved_output_dir = self.project_dir_edit.text().strip()
        self.output_path_edit = QLineEdit(saved_output_dir)
        self.output_path_edit.setPlaceholderText("Папка для сохранения КП")
        self.output_path_edit.editingFinished.connect(self._on_field_changed)
        owner._add_row(form, 6, "Путь сохранения КП", self.output_path_edit, "Выбрать", self.select_output_dir)

        buttons = QHBoxLayout()
        buttons.setContentsMargins(0, 0, 0, 0)
        buttons.setSpacing(8)

        self.generate_btn = QPushButton("Сформировать КП")
        self.generate_btn.setObjectName("PrimaryButton")
        self.generate_btn.clicked.connect(self.generate_offer)
        buttons.addWidget(self.generate_btn, 2)

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

        self._load_sheet_names(initial=True)
        self._restore_or_detect_currency()
        self._update_open_buttons()
        self._update_preview()

    def _saved(self, key: str, default: str) -> str:
        value = self.settings.value(key, default)
        return str(value) if value is not None else default

    def project_path_text(self) -> str:
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
        self.settings.setValue("dc_eltek_output_dir", self.output_path_edit.text().strip())
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
        self.output_path_edit.clear()
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

    def _on_calc_path_edited(self) -> None:
        self._load_sheet_names(initial=False)
        self._auto_detect_currency(force=True)
        self.remember_values()
        self._update_preview()

    def select_project_dir(self) -> None:
        current = self.project_dir_edit.text().strip() or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта", current)
        if not path:
            return
        self.project_dir_edit.setText(path)
        self._on_project_dir_changed(force_client=True)

    def _on_project_dir_changed(self, force_client: bool = False) -> None:
        path_text = self.project_dir_edit.text().strip()
        extracted_client = extract_client_from_project_path(path_text)
        if extracted_client and (force_client or not self.client_edit.text().strip()):
            self.client_edit.setText(extracted_client)
        if path_text and not self.output_path_edit.text().strip():
            self.output_path_edit.setText(path_text)
        self.remember_values()
        self._update_preview()

    def select_calc_file(self) -> None:
        start_dir = self.project_dir_edit.text().strip() or str(Path.home())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel calc",
            start_dir,
            "Excel files (*.xlsx *.xlsm);;All files (*.*)",
        )
        if not path:
            return
        self.calc_path_edit.setText(path)
        self._load_sheet_names(initial=False)
        self._auto_detect_currency(force=True)
        self.remember_values()
        self._update_preview()

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

    def select_output_dir(self) -> None:
        current = (
            self.output_path_edit.text().strip()
            or self.project_dir_edit.text().strip()
            or str(Path.home())
        )
        path = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения КП", current)
        if not path:
            return
        self.output_path_edit.setText(path)
        self.remember_values()
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

    def _context_dict(self) -> dict[str, str]:
        signer = self._selected_signer()
        manager = self._manager_profile()
        return {
            "project_dir": self.project_dir_edit.text().strip(),
            "output_dir": self.output_path_edit.text().strip() or self.project_dir_edit.text().strip(),
            "client": self.client_edit.text().strip(),
            "calc_path": self.calc_path_edit.text().strip(),
            "sheet_name": self.sheet_combo.currentText().strip(),
            "currency": self._currency_value(),
            "template_path": self.template_path_edit.text().strip(),
            "signer_name": str(signer.get("name", "")),
            "signer_position": str(signer.get("position", "")),
            "manager_name": str(getattr(manager, "name", "") if manager else ""),
            "manager_position": str(getattr(manager, "position", "") if manager else ""),
            "manager_email": str(getattr(manager, "email", "") if manager else ""),
            "manager_phone": str(getattr(manager, "phone", "") if manager else ""),
        }

    def _update_preview(self) -> None:
        self.preview_box.setPlainText(build_dc_eltek_preview(self._context_dict()))

    def _update_open_buttons(self) -> None:
        path = Path(self.last_output_path) if self.last_output_path else None
        exists = bool(path and path.exists())
        self.open_offer_btn.setEnabled(exists)
        selected_folder = Path(self.output_path_edit.text().strip()) if self.output_path_edit.text().strip() else None
        self.open_folder_btn.setEnabled(bool((path and path.parent.exists()) or (selected_folder and selected_folder.exists())))

    def generate_offer(self) -> None:
        self.remember_values()
        data = self._context_dict()
        missing: list[str] = []
        if not data["project_dir"]:
            missing.append("папка проекта")
        if not data["client"]:
            missing.append("клиент")
        if not data["calc_path"]:
            missing.append("Excel calc")
        if not data["sheet_name"]:
            missing.append("лист для КП")
        if not data["currency"]:
            QMessageBox.warning(
                self,
                "DC Eltek",
                "Валюта не указана в расчёте. Выберите валюту вручную перед формированием КП.",
            )
            return
        if not data["template_path"]:
            missing.append("шаблон КП")
        if not data["output_dir"]:
            missing.append("путь сохранения КП")
        if missing:
            QMessageBox.warning(self, "DC Eltek", "Заполните поля: " + ", ".join(missing))
            return

        output_dir = Path(data["output_dir"])
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            QMessageBox.critical(self, "DC Eltek", f"Не удалось создать папку сохранения:\n{exc}")
            return

        try:
            result_path = make_dc_eltek_offer(data)
        except Exception as exc:
            QMessageBox.critical(self, "DC Eltek", f"Не удалось сформировать КП:\n{exc}")
            return

        self.last_output_path = str(result_path)
        self.remember_values()
        self._update_open_buttons()
        self._update_preview()
        QMessageBox.information(self, "DC Eltek", f"КП сформировано:\n{result_path}")

    def open_generated_offer(self) -> None:
        path = Path(self.last_output_path) if self.last_output_path else Path()
        if not self.last_output_path or not path.exists():
            QMessageBox.warning(self, "DC Eltek", "Файл КП не найден.")
            self._update_open_buttons()
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def open_generated_folder(self) -> None:
        if self.last_output_path and Path(self.last_output_path).parent.exists():
            folder = Path(self.last_output_path).parent
        else:
            folder = Path(self.output_path_edit.text().strip() or self.project_dir_edit.text().strip() or ".")
        if not folder.exists():
            QMessageBox.warning(self, "DC Eltek", "Папка КП не найдена.")
            self._update_open_buttons()
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))


DCEltekPage = DcEltekPage
