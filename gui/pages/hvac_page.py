from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
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

from brands.hvac import (
    HVACPosition,
    build_hvac_offer,
    find_default_hvac_template,
    read_hvac_positions,
    read_sheet_names,
)


class HVACPage(QWidget):
    """Simple HVAC tab: Excel calculation -> DOCX commercial offer."""

    def __init__(self, parent: QWidget | None = None, project_root: str | Path | None = None):
        super().__init__(parent)
        self.project_root = Path(project_root) if project_root else Path.cwd()
        self.positions: list[HVACPosition] = []
        self.variant1_checks: list[QCheckBox] = []
        self.variant2_checks: list[QCheckBox] = []
        self._build_ui()
        self._set_defaults()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)

        title = QLabel("HVAC — формирование КП по calculation Excel")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        root.addWidget(title)

        files_box = QGroupBox("Файлы")
        files_layout = QGridLayout(files_box)

        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("Выберите HVAC calculation Excel")
        excel_btn = QPushButton("Обзор")
        excel_btn.clicked.connect(self._choose_excel)

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._reload_positions)

        self.template_path = QLineEdit()
        self.template_path.setPlaceholderText("Если встроенный шаблон не найден — выберите вручную")
        template_btn = QPushButton("Обзор")
        template_btn.clicked.connect(self._choose_template)

        self.output_dir = QLineEdit()
        output_btn = QPushButton("Обзор")
        output_btn.clicked.connect(self._choose_output_dir)

        self.template_status = QLabel("")
        self.template_status.setStyleSheet("color: #666;")

        files_layout.addWidget(QLabel("Excel calculation:"), 0, 0)
        files_layout.addWidget(self.excel_path, 0, 1)
        files_layout.addWidget(excel_btn, 0, 2)
        files_layout.addWidget(QLabel("Лист Excel:"), 1, 0)
        files_layout.addWidget(self.sheet_combo, 1, 1, 1, 2)
        files_layout.addWidget(QLabel("Шаблон КП:"), 2, 0)
        files_layout.addWidget(self.template_path, 2, 1)
        files_layout.addWidget(template_btn, 2, 2)
        files_layout.addWidget(self.template_status, 3, 1, 1, 2)
        files_layout.addWidget(QLabel("Папка сохранения:"), 4, 0)
        files_layout.addWidget(self.output_dir, 4, 1)
        files_layout.addWidget(output_btn, 4, 2)
        root.addWidget(files_box)

        data_box = QGroupBox("Данные КП")
        data_layout = QGridLayout(data_box)

        self.client_name = QLineEdit()
        self.project_name = QLineEdit()
        self.city = QLineEdit()
        self.offer_date = QLineEdit()
        self.offer_version = QLineEdit()
        self.delivery_terms = QLineEdit()
        self.engineering_term = QLineEdit()
        self.supply_term = QLineEdit()
        self.currency = QLineEdit()
        self.currency_rate_text = QLineEdit()
        self.payment_terms = QLineEdit()
        self.validity_text = QLineEdit()
        self.variant1_title = QLineEdit()
        self.variant2_title = QLineEdit()
        self.signatory_name = QLineEdit()
        self.signatory_position = QLineEdit()
        self.executor_name = QLineEdit()
        self.executor_position = QLineEdit()
        self.executor_email = QLineEdit()

        self.basis_docs = QTextEdit()
        self.basis_docs.setFixedHeight(72)
        self.basis_docs.setPlaceholderText("Каждый документ с новой строки")

        labels_fields = [
            ("Заказчик:", self.client_name),
            ("Проект:", self.project_name),
            ("Город:", self.city),
            ("Дата:", self.offer_date),
            ("Версия:", self.offer_version),
            ("Условия поставки:", self.delivery_terms),
            ("Инжиниринг:", self.engineering_term),
            ("Поставка:", self.supply_term),
            ("Валюта:", self.currency),
            ("Курс/текст:", self.currency_rate_text),
            ("Оплата:", self.payment_terms),
            ("Срок действия:", self.validity_text),
            ("Вариант 1:", self.variant1_title),
            ("Вариант 2:", self.variant2_title),
            ("Подписант:", self.signatory_name),
            ("Должность подписанта:", self.signatory_position),
            ("Исполнитель:", self.executor_name),
            ("Должность исполнителя:", self.executor_position),
            ("Email исполнителя:", self.executor_email),
        ]
        for idx, (label, field) in enumerate(labels_fields):
            row = idx // 2
            col = (idx % 2) * 2
            data_layout.addWidget(QLabel(label), row, col)
            data_layout.addWidget(field, row, col + 1)

        basis_row = (len(labels_fields) + 1) // 2
        data_layout.addWidget(QLabel("Основание:"), basis_row, 0)
        data_layout.addWidget(self.basis_docs, basis_row, 1, 1, 3)
        root.addWidget(data_box)

        positions_box = QGroupBox("Позиции из Excel")
        positions_layout = QVBoxLayout(positions_box)
        self.positions_table = QTableWidget(0, 6)
        self.positions_table.setHorizontalHeaderLabels(["Наименование", "Кол-во", "Сумма", "Строка цены", "Вариант 1", "Вариант 2"])
        self.positions_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.positions_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        header = self.positions_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        positions_layout.addWidget(self.positions_table)

        buttons_row = QHBoxLayout()
        reload_btn = QPushButton("Прочитать Excel")
        reload_btn.clicked.connect(self._reload_positions)
        v1_all_btn = QPushButton("Все в вариант 1")
        v1_all_btn.clicked.connect(lambda: self._set_checks(self.variant1_checks, True))
        v2_all_btn = QPushButton("Все в вариант 2")
        v2_all_btn.clicked.connect(lambda: self._set_checks(self.variant2_checks, True))
        clear_btn = QPushButton("Снять галочки")
        clear_btn.clicked.connect(self._clear_all_checks)
        buttons_row.addWidget(reload_btn)
        buttons_row.addWidget(v1_all_btn)
        buttons_row.addWidget(v2_all_btn)
        buttons_row.addWidget(clear_btn)
        buttons_row.addStretch(1)
        positions_layout.addLayout(buttons_row)
        root.addWidget(positions_box, 1)

        self.summary_label = QLabel("Итого: вариант 1 — 0.00 EUR, вариант 2 — 0.00 EUR")
        root.addWidget(self.summary_label)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        generate_btn = QPushButton("Сформировать КП")
        generate_btn.setMinimumHeight(36)
        generate_btn.clicked.connect(self._generate_offer)
        bottom.addWidget(generate_btn)
        root.addLayout(bottom)

    def _set_defaults(self) -> None:
        today = datetime.now().strftime("%d.%m.%Y")
        self.city.setText("Алматы")
        self.offer_date.setText(today)
        self.offer_version.setText("1")
        self.delivery_terms.setText("DDP Алматы")
        self.engineering_term.setText("от 7 недель")
        self.supply_term.setText("16-20 недель")
        self.currency.setText("ЕВРО")
        self.currency_rate_text.setText("Взаиморасчет осуществляется в тенге по курсу АО Банк ЦентрКредит на день оплаты.")
        self.payment_terms.setText("70% предоплата, 30% после поставки")
        self.validity_text.setText("Предложение действительно в течение 30 дней.")
        self.variant1_title.setText("Вариант 1 – AHU")
        self.variant2_title.setText("Вариант 2 – PACU")
        self.signatory_name.setText("Алишер Анаркулов")
        self.signatory_position.setText("Исполнительный директор")
        self.executor_name.setText("Виталий Литвинов")
        self.executor_position.setText("менеджер по продажам")
        self.executor_email.setText("Vitaliy@sam.kz")
        self.output_dir.setText(str(self.project_root / "output"))

        template = find_default_hvac_template()
        self.template_path.setText(template)
        if template:
            self.template_status.setText("Шаблон найден автоматически")
        else:
            self.template_status.setText("Шаблон не найден, выберите вручную")

    def _choose_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Выберите HVAC calculation", str(self.project_root), "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        self.excel_path.setText(path)
        self._load_sheets(path)
        self._reload_positions()

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
        path, _ = QFileDialog.getOpenFileName(self, "Выберите шаблон КП HVAC", str(self.project_root), "Word (*.docx)")
        if path:
            self.template_path.setText(path)
            self.template_status.setText("Шаблон выбран вручную")

    def _choose_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите папку сохранения", self.output_dir.text() or str(self.project_root))
        if path:
            self.output_dir.setText(path)

    def _reload_positions(self) -> None:
        path = self.excel_path.text().strip()
        if not path:
            return
        try:
            sheet = self.sheet_combo.currentText() or None
            self.positions = read_hvac_positions(path, sheet_name=sheet)
        except Exception as exc:
            QMessageBox.warning(self, "HVAC", f"Не удалось прочитать позиции из Excel:\n{exc}")
            return
        self._fill_positions_table()

    def _fill_positions_table(self) -> None:
        self.positions_table.setRowCount(0)
        self.variant1_checks.clear()
        self.variant2_checks.clear()

        for row, pos in enumerate(self.positions):
            self.positions_table.insertRow(row)
            self.positions_table.setItem(row, 0, QTableWidgetItem(pos.name))
            self.positions_table.setItem(row, 1, QTableWidgetItem(pos.qty_text))
            amount_item = QTableWidgetItem(pos.amount_text)
            amount_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.positions_table.setItem(row, 2, amount_item)
            self.positions_table.setItem(row, 3, QTableWidgetItem(pos.price_row_label))

            v1 = QCheckBox()
            v2 = QCheckBox()
            v1.stateChanged.connect(self._update_summary)
            v2.stateChanged.connect(self._update_summary)
            self.variant1_checks.append(v1)
            self.variant2_checks.append(v2)
            self.positions_table.setCellWidget(row, 4, _centered_widget(v1))
            self.positions_table.setCellWidget(row, 5, _centered_widget(v2))

        self.positions_table.resizeRowsToContents()
        self._update_summary()

    def _set_checks(self, checks: list[QCheckBox], checked: bool) -> None:
        for check in checks:
            check.setChecked(checked)
        self._update_summary()

    def _clear_all_checks(self) -> None:
        self._set_checks(self.variant1_checks, False)
        self._set_checks(self.variant2_checks, False)

    def _selected_positions(self, checks: list[QCheckBox]) -> list[HVACPosition]:
        return [pos for pos, check in zip(self.positions, checks) if check.isChecked()]

    def _update_summary(self) -> None:
        v1_total = sum(float(pos.amount or 0) for pos in self._selected_positions(self.variant1_checks))
        v2_total = sum(float(pos.amount or 0) for pos in self._selected_positions(self.variant2_checks))
        self.summary_label.setText(
            f"Итого: вариант 1 — {v1_total:,.2f} EUR, вариант 2 — {v2_total:,.2f} EUR".replace(",", " ")
        )

    def _collect_fields(self) -> dict[str, Any]:
        return {
            "CLIENT_NAME": self.client_name.text().strip(),
            "PROJECT_NAME": self.project_name.text().strip(),
            "CITY": self.city.text().strip(),
            "OFFER_DATE": self.offer_date.text().strip(),
            "OFFER_VERSION": self.offer_version.text().strip(),
            "DELIVERY_TERMS": self.delivery_terms.text().strip(),
            "ENGINEERING_TERM": self.engineering_term.text().strip(),
            "SUPPLY_TERM": self.supply_term.text().strip(),
            "CURRENCY": self.currency.text().strip(),
            "CURRENCY_RATE_TEXT": self.currency_rate_text.text().strip(),
            "PAYMENT_TERMS": self.payment_terms.text().strip(),
            "VALIDITY_TEXT": self.validity_text.text().strip(),
            "VARIANT_1_TITLE": self.variant1_title.text().strip(),
            "VARIANT_2_TITLE": self.variant2_title.text().strip(),
            "SIGNATORY_NAME": self.signatory_name.text().strip(),
            "SIGNATORY_POSITION": self.signatory_position.text().strip(),
            "EXECUTOR_NAME": self.executor_name.text().strip(),
            "EXECUTOR_POSITION": self.executor_position.text().strip(),
            "EXECUTOR_EMAIL": self.executor_email.text().strip(),
            "BASIS_DOCS": [x.strip() for x in self.basis_docs.toPlainText().splitlines() if x.strip()],
        }

    def _generate_offer(self) -> None:
        template = self.template_path.text().strip()
        if not template:
            QMessageBox.warning(self, "HVAC", "Не выбран шаблон КП.")
            return
        if not self.positions:
            QMessageBox.warning(self, "HVAC", "Нет позиций из Excel. Сначала выберите calculation и нажмите 'Прочитать Excel'.")
            return

        v1_items = self._selected_positions(self.variant1_checks)
        v2_items = self._selected_positions(self.variant2_checks)
        if not v1_items and not v2_items:
            QMessageBox.warning(self, "HVAC", "Выберите хотя бы одну позицию для варианта 1 или варианта 2.")
            return

        fields = self._collect_fields()
        client = _safe_filename(fields.get("CLIENT_NAME") or "HVAC")
        date = _safe_filename(fields.get("OFFER_DATE") or datetime.now().strftime("%d.%m.%Y"))
        version = _safe_filename(fields.get("OFFER_VERSION") or "1")
        output = Path(self.output_dir.text().strip() or self.project_root / "output") / f"Offer_{client}_{date}_v{version}_HVAC.docx"

        try:
            result = build_hvac_offer(template, output, fields, v1_items, v2_items)
        except Exception as exc:
            QMessageBox.critical(self, "HVAC", f"Не удалось сформировать КП:\n{exc}")
            return

        QMessageBox.information(self, "HVAC", f"КП сформировано:\n{result}")


def _centered_widget(widget: QWidget) -> QWidget:
    box = QWidget()
    layout = QHBoxLayout(box)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addStretch(1)
    layout.addWidget(widget)
    layout.addStretch(1)
    return box


def _safe_filename(value: str) -> str:
    value = value.replace("/", "-").replace("\\", "-").replace(":", "-")
    forbidden = '<>"|?*'
    for ch in forbidden:
        value = value.replace(ch, "")
    return "_".join(value.split()) or "HVAC"
