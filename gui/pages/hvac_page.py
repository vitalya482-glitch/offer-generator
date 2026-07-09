from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any
import re

from docx import Document
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.excel_calc_parser import (
    CalcItem,
    CalcParseResult,
    format_money,
    format_qty,
    parse_calculation,
    read_sheet_names,
)

try:
    from brands.hvac.template_finder import find_default_hvac_template
except ImportError:  # pragma: no cover
    def find_default_hvac_template() -> str:
        return ""


_TAG_RE = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")
_ITEM_TAGS = {"item_no", "item_name", "item_qty", "item_unit_price", "item_total"}


class HVACPage(QWidget):
    """HVAC: universal calculation parser -> tagged Word offer."""

    def __init__(self, parent: QWidget | None = None, project_root: str | Path | None = None):
        super().__init__(parent)
        self.project_root = Path(project_root) if project_root else Path.cwd()
        self.parse_result: CalcParseResult | None = None
        self.items: list[CalcItem] = []
        self.item_checks: list[QCheckBox] = []
        self._build_ui()
        self._set_defaults()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)

        title = QLabel("HVAC — формирование КП по calculation Excel")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        root.addWidget(title)

        files_box = QGroupBox("Файлы")
        files = QGridLayout(files_box)
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("Выберите calculation Excel")
        excel_btn = QPushButton("Обзор")
        excel_btn.clicked.connect(self._choose_excel)

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._parse_excel)

        self.template_path = QLineEdit()
        self.template_path.setPlaceholderText("Шаблон HVAC DOCX")
        template_btn = QPushButton("Обзор")
        template_btn.clicked.connect(self._choose_template)
        self.template_status = QLabel()

        self.output_dir = QLineEdit()
        output_btn = QPushButton("Обзор")
        output_btn.clicked.connect(self._choose_output_dir)

        files.addWidget(QLabel("Excel calculation:"), 0, 0)
        files.addWidget(self.excel_path, 0, 1)
        files.addWidget(excel_btn, 0, 2)
        files.addWidget(QLabel("Лист Excel:"), 1, 0)
        files.addWidget(self.sheet_combo, 1, 1, 1, 2)
        files.addWidget(QLabel("Шаблон КП:"), 2, 0)
        files.addWidget(self.template_path, 2, 1)
        files.addWidget(template_btn, 2, 2)
        files.addWidget(self.template_status, 3, 1, 1, 2)
        files.addWidget(QLabel("Папка сохранения:"), 4, 0)
        files.addWidget(self.output_dir, 4, 1)
        files.addWidget(output_btn, 4, 2)
        root.addWidget(files_box)

        data_box = QGroupBox("Данные КП")
        data = QGridLayout(data_box)
        self.offer_date = QLineEdit()
        self.offer_version = QLineEdit()
        self.client_company = QLineEdit()
        self.project_name = QLineEdit()
        self.intro_text = QTextEdit()
        self.intro_text.setFixedHeight(48)
        self.data_files = QTextEdit()
        self.data_files.setFixedHeight(64)
        self.data_files.setPlaceholderText("Каждый документ с новой строки")
        self.delivery_terms = QLineEdit()
        self.delivery_time = QLineEdit()
        self.payment_terms = QLineEdit()
        self.signer_name = QLineEdit()
        self.signer_position = QLineEdit()
        self.manager_name = QLineEdit()
        self.manager_position = QLineEdit()
        self.manager_email = QLineEdit()
        self.manager_phone = QLineEdit()

        fields = [
            ("Дата:", self.offer_date),
            ("Версия:", self.offer_version),
            ("Заказчик:", self.client_company),
            ("Проект:", self.project_name),
            ("Условия поставки:", self.delivery_terms),
            ("Срок поставки:", self.delivery_time),
            ("Условия оплаты:", self.payment_terms),
            ("Подписант:", self.signer_name),
            ("Должность подписанта:", self.signer_position),
            ("Исполнитель:", self.manager_name),
            ("Должность исполнителя:", self.manager_position),
            ("Email:", self.manager_email),
            ("Телефон:", self.manager_phone),
        ]
        for index, (label, widget) in enumerate(fields):
            row = index // 2
            col = (index % 2) * 2
            data.addWidget(QLabel(label), row, col)
            data.addWidget(widget, row, col + 1)
        row = (len(fields) + 1) // 2
        data.addWidget(QLabel("Вводный текст:"), row, 0)
        data.addWidget(self.intro_text, row, 1, 1, 3)
        data.addWidget(QLabel("На основании:"), row + 1, 0)
        data.addWidget(self.data_files, row + 1, 1, 1, 3)
        root.addWidget(data_box)

        services_box = QGroupBox("Работы и услуги")
        services = QHBoxLayout(services_box)
        self.engineering_check = QCheckBox("Инжиниринг включён")
        self.installation_check = QCheckBox("Монтаж включён")
        self.startup_check = QCheckBox("Пусконаладка включена")
        services.addWidget(self.engineering_check)
        services.addWidget(self.installation_check)
        services.addWidget(self.startup_check)
        services.addStretch(1)
        root.addWidget(services_box)

        items_box = QGroupBox("Позиции из Excel")
        items_layout = QVBoxLayout(items_box)
        self.items_table = QTableWidget(0, 6)
        self.items_table.setHorizontalHeaderLabels(
            ["В КП", "Наименование", "Кол-во", "Цена за ед.", "Сумма", "Колонка"]
        )
        self.items_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        header = self.items_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for column in range(2, 6):
            header.setSectionResizeMode(column, QHeaderView.ResizeToContents)
        items_layout.addWidget(self.items_table)

        row_buttons = QHBoxLayout()
        parse_btn = QPushButton("Прочитать Excel")
        parse_btn.clicked.connect(self._parse_excel)
        all_btn = QPushButton("Выбрать все")
        all_btn.clicked.connect(lambda: self._set_all_items(True))
        none_btn = QPushButton("Снять все")
        none_btn.clicked.connect(lambda: self._set_all_items(False))
        row_buttons.addWidget(parse_btn)
        row_buttons.addWidget(all_btn)
        row_buttons.addWidget(none_btn)
        row_buttons.addStretch(1)
        items_layout.addLayout(row_buttons)
        root.addWidget(items_box, 1)

        self.calc_status = QLabel("Calculation не прочитан")
        self.calc_status.setWordWrap(True)
        root.addWidget(self.calc_status)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        generate_btn = QPushButton("Сформировать КП")
        generate_btn.setMinimumHeight(36)
        generate_btn.clicked.connect(self._generate_offer)
        bottom.addWidget(generate_btn)
        root.addLayout(bottom)

    def _set_defaults(self) -> None:
        self.offer_date.setText(datetime.now().strftime("%d.%m.%Y г."))
        self.offer_version.setText("1")
        self.delivery_terms.setText("DDP Алматы")
        self.delivery_time.setText("16–20 недель")
        self.payment_terms.setText("70% предоплата, 30% после поставки")
        self.signer_name.setText("Сания Санаткызы")
        self.signer_position.setText("Коммерческий директор")
        self.manager_name.setText("Виталий Литвинов")
        self.manager_position.setText("менеджер по продажам")
        self.manager_email.setText("Vitaliy@sam.kz")
        self.output_dir.setText(str(self.project_root / "output"))

        template = find_default_hvac_template()
        self.template_path.setText(template)
        self.template_status.setText(
            "Шаблон найден автоматически" if template else "Шаблон не найден — выберите вручную"
        )

    def _choose_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите calculation Excel",
            str(self.project_root),
            "Excel (*.xlsx *.xlsm)",
        )
        if not path:
            return
        self.excel_path.setText(path)
        self._load_sheets(path)
        preferred = "DDP_Almaty_20-10(v2)"
        index = self.sheet_combo.findText(preferred)
        if index >= 0:
            self.sheet_combo.setCurrentIndex(index)
        self._parse_excel()

    def _load_sheets(self, path: str) -> None:
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        try:
            self.sheet_combo.addItems(read_sheet_names(path))
        except Exception as exc:
            QMessageBox.warning(self, "HVAC", f"Не удалось прочитать листы Excel:\n{exc}")
        finally:
            self.sheet_combo.blockSignals(False)

    def _choose_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите шаблон КП HVAC",
            str(self.project_root),
            "Word (*.docx)",
        )
        if path:
            self.template_path.setText(path)
            self.template_status.setText("Шаблон выбран вручную")

    def _choose_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку сохранения",
            self.output_dir.text() or str(self.project_root),
        )
        if path:
            self.output_dir.setText(path)

    def _parse_excel(self) -> None:
        path = self.excel_path.text().strip()
        if not path:
            return
        try:
            result = parse_calculation(path, sheet_name=self.sheet_combo.currentText() or None)
        except Exception as exc:
            QMessageBox.warning(self, "HVAC", f"Не удалось разобрать calculation:\n{exc}")
            return

        self.parse_result = result
        self.items = result.items
        self.delivery_terms.setText(result.delivery_basis or self.delivery_terms.text())
        self.engineering_check.setChecked(result.engineering.included is True)
        self.installation_check.setChecked(result.installation.included is True)
        self.startup_check.setChecked(result.startup.included is True)
        self._fill_items_table()
        self._update_status()

    def _fill_items_table(self) -> None:
        self.items_table.setRowCount(0)
        self.item_checks.clear()
        for row, item in enumerate(self.items):
            self.items_table.insertRow(row)
            check = QCheckBox()
            check.setChecked(True)
            check.stateChanged.connect(self._update_status)
            self.item_checks.append(check)
            self.items_table.setCellWidget(row, 0, _centered_widget(check))
            self.items_table.setItem(row, 1, QTableWidgetItem(item.name))
            self.items_table.setItem(row, 2, QTableWidgetItem(format_qty(item.qty)))
            self.items_table.setItem(row, 3, _right_item(format_money(item.unit_price)))
            self.items_table.setItem(row, 4, _right_item(format_money(item.total_price)))
            self.items_table.setItem(row, 5, QTableWidgetItem(str(item.source_col)))
        self.items_table.resizeRowsToContents()

    def _set_all_items(self, checked: bool) -> None:
        for check in self.item_checks:
            check.setChecked(checked)
        self._update_status()

    def _selected_items(self) -> list[CalcItem]:
        return [item for item, check in zip(self.items, self.item_checks) if check.isChecked()]

    def _update_status(self) -> None:
        if not self.parse_result:
            self.calc_status.setText("Calculation не прочитан")
            return
        selected = self._selected_items()
        total = sum(float(item.total_price or 0) for item in selected)
        result = self.parse_result
        details = [
            f"Лист: {result.sheet_name}",
            f"Позиций: {len(selected)} из {len(result.items)}",
            f"Итого выбранных: {format_money(total)} {result.currency or ''}".strip(),
            f"Курс: {format_money(result.exchange_rate) if result.exchange_rate else 'не найден'}",
            f"НДС: {format_qty(result.vat_percent) + '%' if result.vat_percent is not None else 'не найден'}",
        ]
        if result.warnings:
            details.append("Предупреждения: " + "; ".join(result.warnings))
        self.calc_status.setText(" | ".join(details))

    def _generate_offer(self) -> None:
        template = Path(self.template_path.text().strip())
        if not template.exists():
            QMessageBox.warning(self, "HVAC", "Не найден шаблон КП.")
            return
        items = self._selected_items()
        if not items:
            QMessageBox.warning(self, "HVAC", "Не выбраны позиции для КП.")
            return

        output_dir = Path(self.output_dir.text().strip() or self.project_root / "output")
        output_dir.mkdir(parents=True, exist_ok=True)
        client = _safe_filename(self.client_company.text().strip() or "Company")
        version = _safe_filename(self.offer_version.text().strip() or "1")
        output = output_dir / f"Offer_{client}_{datetime.now():%d-%m-%y}(v{version}) HVAC.docx"

        tags = self._collect_tags(items)
        try:
            _render_hvac_template(template, output, tags, items)
        except Exception as exc:
            QMessageBox.critical(self, "HVAC", f"Не удалось сформировать КП:\n{exc}")
            return
        QMessageBox.information(self, "HVAC", f"КП сформировано:\n{output}")

    def _collect_tags(self, items: list[CalcItem]) -> dict[str, str]:
        result = self.parse_result
        currency = (result.currency if result else None) or "EUR"
        docs = [line.strip() for line in self.data_files.toPlainText().splitlines() if line.strip()]
        data_file_name = "\n".join(f"• {name}" for name in docs)
        grand_total = sum(float(item.total_price or 0) for item in items)
        total_words = f"{format_money(grand_total)} {currency}"

        return {
            "offer_date": self.offer_date.text().strip(),
            "offer_version": self.offer_version.text().strip(),
            "client_company_full": self.client_company.text().strip(),
            "intro_text": self.intro_text.toPlainText().strip(),
            "project_name": self.project_name.text().strip(),
            "data_file_name": data_file_name,
            "delivery_terms": self.delivery_terms.text().strip(),
            "installation_terms": (
                "Монтажные работы включены." if self.installation_check.isChecked()
                else "Монтажные работы не включены."
            ),
            "startup_terms": (
                "Пуско-наладочные работы включены." if self.startup_check.isChecked()
                else "Пуско-наладочные работы не включены."
            ),
            "engineering_terms": (
                "Инжиниринг включен, срок выполнения от 7 недель."
                if self.engineering_check.isChecked()
                else "Инжиниринг не включен."
            ),
            "delivery_time": self.delivery_time.text().strip(),
            "unit_price_header": f"Цена за ед., {currency}",
            "total_price_header": f"Сумма, {currency}",
            "total_label": "Итого",
            "grand_total": format_money(grand_total),
            "total_price_block": total_words,
            "payment_terms": self.payment_terms.text().strip(),
            "signer_name": self.signer_name.text().strip(),
            "signer_position": self.signer_position.text().strip(),
            "manager_name": self.manager_name.text().strip(),
            "manager_position": self.manager_position.text().strip(),
            "manager_email": self.manager_email.text().strip(),
            "manager_phone": self.manager_phone.text().strip(),
        }


def _render_hvac_template(
    template_path: Path,
    output_path: Path,
    tags: dict[str, str],
    items: list[CalcItem],
) -> None:
    document = Document(str(template_path))
    item_table = None
    template_row_index = None

    for table in document.tables:
        for row_index, row in enumerate(table.rows):
            row_text = " | ".join(cell.text for cell in row.cells)
            found = {match.group(1).casefold() for match in _TAG_RE.finditer(row_text)}
            if found & _ITEM_TAGS:
                item_table = table
                template_row_index = row_index
                break
        if item_table is not None:
            break

    if item_table is None or template_row_index is None:
        raise ValueError("В шаблоне не найдена строка {{item_name}} ценовой таблицы.")

    template_row = item_table.rows[template_row_index]
    row_xml = deepcopy(template_row._tr)
    item_table._tbl.remove(template_row._tr)

    insert_index = template_row_index
    for number, item in enumerate(items, start=1):
        new_xml = deepcopy(row_xml)
        item_table._tbl.insert(insert_index, new_xml)
        new_row = item_table.rows[insert_index]
        item_tags = {
            "item_no": str(number),
            "item_name": item.name,
            "item_qty": format_qty(item.qty),
            "item_unit_price": format_money(item.unit_price),
            "item_total": format_money(item.total_price),
        }
        for cell in new_row.cells:
            for paragraph in cell.paragraphs:
                _replace_tags_in_paragraph(paragraph, item_tags)
        insert_index += 1

    _replace_tags_everywhere(document, tags)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))


def _replace_tags_everywhere(document: Document, tags: dict[str, str]) -> None:
    for paragraph in document.paragraphs:
        _replace_tags_in_paragraph(paragraph, tags)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_tags_in_paragraph(paragraph, tags)
    for section in document.sections:
        for part in (section.header, section.footer):
            for paragraph in part.paragraphs:
                _replace_tags_in_paragraph(paragraph, tags)
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            _replace_tags_in_paragraph(paragraph, tags)


def _replace_tags_in_paragraph(paragraph, tags: dict[str, str]) -> None:
    text = paragraph.text
    if "{{" not in text:
        return

    normalized_tags = {key.casefold(): str(value or "") for key, value in tags.items()}
    replaced = _TAG_RE.sub(lambda match: normalized_tags.get(match.group(1).casefold(), ""), text)
    if replaced == text:
        return
    if paragraph.runs:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(replaced)


def _right_item(text: str) -> QTableWidgetItem:
    item = QTableWidgetItem(text)
    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    return item


def _centered_widget(widget: QWidget) -> QWidget:
    box = QWidget()
    layout = QHBoxLayout(box)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addStretch(1)
    layout.addWidget(widget)
    layout.addStretch(1)
    return box


def _safe_filename(value: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', "_", value).strip()
    return "_".join(value.split()) or "HVAC"
