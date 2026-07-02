from __future__ import annotations

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


PROJECTS_MARKER = "02_Projects"


def extract_client_from_project_path(path_text: str) -> str:
    """Берет клиента из пути вида ...\\02_Projects\\КЛИЕНТ\\..."""
    if not path_text:
        return ""

    parts = [part for part in path_text.replace("/", "\\").split("\\") if part]
    for index, part in enumerate(parts):
        if part.lower() == PROJECTS_MARKER.lower() and index + 1 < len(parts):
            return parts[index + 1].strip()

    return ""


def read_excel_sheet_names(path_text: str) -> list[str]:
    """Читает список листов из .xlsx/.xlsm без зависимости от openpyxl.

    Старый .xls намеренно не читаем: для него нужна отдельная библиотека.
    """
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

    result: list[str] = []
    for sheet in sheets.findall("main:sheet", namespace):
        name = sheet.attrib.get("name", "").strip()
        if name:
            result.append(name)
    return result


class DCEltekPage(QWidget):
    """Страница DC Eltek: выбор проекта, расчета, листа и шаблона КП."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        self.settings = owner.settings

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
        owner._add_row(form, 1, "Клиент", self.client_edit, None, None)

        self.calc_path_edit = QLineEdit(self._saved("dc_eltek_calc_path", ""))
        self.calc_path_edit.setPlaceholderText("Excel calc")
        owner._add_row(form, 2, "Расчёт Excel", self.calc_path_edit, "Выбрать", self.select_calc_file)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setEditable(False)
        owner._add_row(form, 3, "Лист для КП", self.sheet_combo, None, None)

        self.template_path_edit = QLineEdit(self._saved("dc_eltek_template_path", ""))
        self.template_path_edit.setPlaceholderText("Шаблон КП .docx")
        owner._add_row(form, 4, "Шаблон КП", self.template_path_edit, "Выбрать", self.select_template_file)

        self.generate_btn = QPushButton("Сформировать КП")
        self.generate_btn.setObjectName("PrimaryButton")
        self.generate_btn.clicked.connect(self.generate_offer)
        card.layout().addWidget(self.generate_btn)

        self.preview_box = QTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setMinimumHeight(120)
        card.layout().addWidget(self.preview_box)

        layout.addWidget(card)
        layout.addStretch(1)

        self._load_sheet_names(initial=True)
        self._update_preview()

    def _saved(self, key: str, default: str) -> str:
        value = self.settings.value(key, default)
        return str(value) if value is not None else default

    def project_path_text(self) -> str:
        return self.project_dir_edit.text().strip()

    def remember_values(self) -> None:
        self.settings.setValue("brand", "DC Eltek")
        self.settings.setValue("dc_eltek_project_dir", self.project_dir_edit.text().strip())
        self.settings.setValue("dc_eltek_client", self.client_edit.text().strip())
        self.settings.setValue("dc_eltek_calc_path", self.calc_path_edit.text().strip())
        self.settings.setValue("dc_eltek_sheet_name", self.sheet_combo.currentText().strip())
        self.settings.setValue("dc_eltek_template_path", self.template_path_edit.text().strip())
        self.settings.sync()

    def clear_cache(self) -> None:
        for key in (
            "dc_eltek_project_dir",
            "dc_eltek_client",
            "dc_eltek_calc_path",
            "dc_eltek_sheet_name",
            "dc_eltek_template_path",
        ):
            self.settings.remove(key)
        self.project_dir_edit.clear()
        self.client_edit.clear()
        self.calc_path_edit.clear()
        self.sheet_combo.clear()
        self.template_path_edit.clear()
        self.preview_box.clear()
        self.settings.sync()

    def apply_responsive_metrics(self, scale: float) -> None:
        self.preview_box.setMinimumHeight(int(120 * scale))

    def on_settings_changed(self) -> None:
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

    def _load_sheet_names(self, initial: bool) -> None:
        current_sheet = self._saved("dc_eltek_sheet_name", "") if initial else self.sheet_combo.currentText().strip()
        self.sheet_combo.clear()

        calc_path = self.calc_path_edit.text().strip()
        if not calc_path:
            return

        try:
            sheet_names = read_excel_sheet_names(calc_path)
        except Exception as exc:
            if not initial:
                QMessageBox.warning(self, "DC Eltek", f"Не удалось прочитать листы Excel:\n{exc}")
            return

        self.sheet_combo.addItems(sheet_names)
        if current_sheet:
            index = self.sheet_combo.findText(current_sheet, Qt.MatchFixedString)
            if index >= 0:
                self.sheet_combo.setCurrentIndex(index)

    def _context_dict(self) -> dict[str, str]:
        return {
            "project_dir": self.project_dir_edit.text().strip(),
            "client": self.client_edit.text().strip(),
            "calc_path": self.calc_path_edit.text().strip(),
            "sheet_name": self.sheet_combo.currentText().strip(),
            "template_path": self.template_path_edit.text().strip(),
        }

    def _update_preview(self) -> None:
        data = self._context_dict()
        lines = [
            "DC Eltek — подготовка КП",
            "",
            f"Папка проекта: {data['project_dir'] or 'не выбрана'}",
            f"Клиент: {data['client'] or 'не указан'}",
            f"Расчёт Excel: {data['calc_path'] or 'не выбран'}",
            f"Лист для КП: {data['sheet_name'] or 'не выбран'}",
            f"Шаблон КП: {data['template_path'] or 'не выбран'}",
        ]
        self.preview_box.setPlainText("\n".join(lines))

    def generate_offer(self) -> None:
        self.remember_values()
        data = self._context_dict()

        missing = []
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
            QMessageBox.warning(self, "DC Eltek", "Заполните поля: " + ", ".join(missing))
            return

        QMessageBox.information(
            self,
            "DC Eltek",
            "Каркас вкладки готов.\nГенерацию КП подключим следующим этапом после разбора Excel и тегов шаблона.",
        )
