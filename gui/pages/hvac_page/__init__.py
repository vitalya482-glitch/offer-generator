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
        self._sync_output_dir_to_calc_or_sales()

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
        self._sync_output_dir_to_calc_or_sales()
        self.remember_values()

    def scan_project(self) -> None:
        _legacy.HVACPage.scan_project(self)
        self._sync_output_dir_to_calc_or_sales()

    def _on_calc_changed(self) -> None:
        _legacy.HVACPage._on_calc_changed(self)
        self._sync_output_dir_to_calc_or_sales()

    def _parse_excel(self) -> None:
        _legacy.HVACPage._parse_excel(self)
        result = getattr(self, "parse_result", None)
        if result is None or not hasattr(self, "vat_check"):
            self._sync_output_dir_to_calc_or_sales()
            return
        if result.vat_included is not None:
            self.vat_check.setChecked(bool(result.vat_included))
        percent = _format_percent(result.vat_percent)
        self.vat_check.setText(f"НДС включён ({percent}%)" if percent else "НДС включён")
        self._sync_output_dir_to_calc_or_sales()
        self._update_status()
        self.remember_values()

    def generate(self) -> None:
        self._sync_output_dir_to_calc_or_sales()
        _legacy.HVACPage.generate(self)

    def _preferred_output_dir(self) -> Path | None:
        """Prefer the folder where the selected calc file is located.

        In real project folders the HVAC calculation is normally placed directly
        in Sales docs. Saving the offer next to that calc is the safest rule and
        also avoids stale saved paths from older projects.
        """

        try:
            calc_value = self._path_from_combo(self.calc_combo)
        except Exception:
            calc_value = ""
        calc_path = Path(str(calc_value or "").strip())
        if calc_path.is_file():
            return calc_path.parent

        project_text = self.project_path_text().strip() if hasattr(self, "project_path_text") else ""
        project_dir = Path(project_text)
        if project_dir.is_dir():
            try:
                inferred = _legacy.infer_output_dir(str(project_dir))
            except Exception:
                inferred = ""
            inferred_path = Path(str(inferred or "").strip())
            if inferred_path:
                return inferred_path
        return None

    def _sync_output_dir_to_calc_or_sales(self) -> None:
        preferred = self._preferred_output_dir()
        if preferred is None:
            return
        preferred_text = str(preferred)
        current = self.output_path_text().strip() if hasattr(self, "output_path_text") else ""
        if current and _paths_equal(current, preferred_text):
            return
        self.output_dir_path = preferred_text
        try:
            self._set_line_path(self.output_edit, preferred_text, is_file=False)
        except Exception:
            try:
                self.output_edit.setText(preferred_text)
            except Exception:
                pass

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
        preferred = self._preferred_output_dir()
        if preferred is not None:
            extra.append(f"папка КП: {preferred.name}")
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


def _paths_equal(left: str, right: str) -> bool:
    try:
        legacy_equal = getattr(_legacy, "_same_windows_path", None)
        if legacy_equal is not None:
            return bool(legacy_equal(left, right))
    except Exception:
        pass
    return str(left or "").strip().casefold() == str(right or "").strip().casefold()


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
