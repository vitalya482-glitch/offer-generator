from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QCheckBox, QComboBox, QGridLayout, QLabel

from core.excel_calc_parser import format_money
from core.final_offer_word_maker import format_amount_in_words


_CURRENCY_OPTIONS = ("Авто", "KZT", "EUR", "USD", "RUB")
_PAYMENT_DEFAULT = "70% предоплата, 30% после поставки оборудования"
_PAYMENT_OLD_DEFAULT = "70% предоплата, 30% после поставки"


def _load_legacy_module():
    legacy_path = Path(__file__).resolve().parents[1] / "hvac_page.py"
    spec = importlib.util.spec_from_file_location("_sam_legacy_hvac_page", legacy_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Не удалось загрузить базовую HVAC страницу: {legacy_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_legacy = _load_legacy_module()


class HVACPage(_legacy.HVACPage):
    """HVAC page hotfix: manual currency, VAT flag and payment default."""

    def _build_ui(self) -> None:
        _legacy.HVACPage._build_ui(self)
        self.currency_combo = QComboBox()
        self.currency_combo.addItems(_CURRENCY_OPTIONS)
        self.vat_check = QCheckBox("НДС включён")

        card = self.owner._card("Валюта и НДС")
        grid = QGridLayout()
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)
        card.layout().addLayout(grid)
        grid.addWidget(QLabel("Валюта:"), 0, 0)
        grid.addWidget(self.currency_combo, 0, 1)
        grid.addWidget(self.vat_check, 0, 2, alignment=Qt.AlignLeft)

        layout = self.layout()
        insert_index = max(0, layout.count() - 2)
        layout.insertWidget(insert_index, card)

        self.currency_combo.currentTextChanged.connect(lambda *_: (self.remember_values(), self._update_status()))
        self.vat_check.stateChanged.connect(lambda *_: (self.remember_values(), self._update_status()))
        self.installation_check.stateChanged.connect(self._sync_startup_with_installation)

    def _load_saved_values(self) -> None:
        _legacy.HVACPage._load_saved_values(self)
        saved_payment = self.payment_terms.text().strip()
        if not saved_payment or saved_payment == _PAYMENT_OLD_DEFAULT:
            self.payment_terms.setText(_PAYMENT_DEFAULT)

        saved_currency = self._saved("hvac/currency", "Авто").strip()
        self.currency_combo.setCurrentText(saved_currency if saved_currency in _CURRENCY_OPTIONS else "Авто")

    def remember_values(self) -> None:
        _legacy.HVACPage.remember_values(self)
        if hasattr(self, "currency_combo"):
            self.settings.setValue("hvac/currency", self.currency_combo.currentText().strip())
            self.settings.sync()

    def clear_cache(self) -> None:
        _legacy.HVACPage.clear_cache(self)
        self.currency_combo.setCurrentText("Авто")
        self.vat_check.setChecked(False)
        self.vat_check.setText("НДС включён")
        self.payment_terms.setText(_PAYMENT_DEFAULT)
        self.remember_values()

    def _parse_excel(self) -> None:
        _legacy.HVACPage._parse_excel(self)
        result = getattr(self, "parse_result", None)
        if result is None or not hasattr(self, "vat_check"):
            return
        if result.vat_included is not None:
            self.vat_check.setChecked(bool(result.vat_included))
        percent = _format_percent(result.vat_percent)
        self.vat_check.setText(f"НДС включён ({percent}%)" if percent else "НДС включён")
        self._update_status()
        self.remember_values()

    def _sync_startup_with_installation(self, *_args: Any) -> None:
        if self.installation_check.isChecked():
            self.startup_check.setChecked(True)

    def _offer_currency(self) -> str:
        selected = ""
        if hasattr(self, "currency_combo"):
            selected = self.currency_combo.currentText().strip().upper()
        if selected and selected != "АВТО":
            return selected
        return _legacy.HVACPage._offer_currency(self)

    def _update_status(self) -> None:
        _legacy.HVACPage._update_status(self)
        if not hasattr(self, "status_label") or not hasattr(self, "currency_combo"):
            return
        result = getattr(self, "parse_result", None)
        if result is None:
            return
        extra = [
            f"валюта Excel: {result.currency or 'не определена'}",
            f"валюта КП: {self._offer_currency()}",
            _vat_status_text(bool(self.vat_check.isChecked()), getattr(result, "vat_percent", None)),
        ]
        current = self.status_label.text().strip()
        self.status_label.setText((current + " | " if current else "") + " | ".join(extra))

    def _collect_tags(self, items):
        tags = _legacy.HVACPage._collect_tags(self, items)
        currency = self._offer_currency()
        grand_total = sum(float(item.total_price or 0) for item in items)
        vat_percent = getattr(getattr(self, "parse_result", None), "vat_percent", None)

        tags["unit_price_header"] = f"Цена за ед., {currency}"
        tags["total_price_header"] = f"Сумма, {currency}"
        tags["grand_total"] = format_money(grand_total)
        tags["total_price_block"] = _format_total_price_block(
            grand_total,
            currency,
            bool(self.vat_check.isChecked()),
            vat_percent,
        )
        tags["currency_name"] = _legacy._currency_name_ru(currency)
        tags["payment_terms"] = self.payment_terms.text().strip() or _PAYMENT_DEFAULT
        return tags


def _format_percent(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return ""
    if number.is_integer():
        return str(int(number))
    return str(number).rstrip("0").rstrip(".").replace(".", ",")


def _vat_status_text(vat_included: bool, percent: Any) -> str:
    pct = _format_percent(percent)
    if vat_included:
        return f"НДС включён ({pct}%)" if pct else "НДС включён"
    return "НДС не включён"


def _format_total_price_block(total: Any, currency: str, vat_included: bool, vat_percent: Any) -> str:
    amount = format_amount_in_words(total, currency)
    pct = _format_percent(vat_percent)
    if vat_included:
        suffix = f"с учетом стоимости НДС {pct}%" if pct else "с учетом стоимости НДС"
    else:
        suffix = "без учета НДС"
    return f"{amount} {suffix}"


for _name, _value in vars(_legacy).items():
    if _name not in globals() and not (_name.startswith("__") and _name.endswith("__")):
        globals()[_name] = _value
