from __future__ import annotations

import copy
import sys

from PySide6.QtWidgets import (
    QApplication,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QWidget,
)


DEFAULT_PAYMENT_TERMS = "70% предоплата, 30% после поставки оборудования"


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


def _find_files_grid(self) -> QGridLayout | None:
    client_edit = getattr(self, "client_edit", None)
    card = client_edit.parentWidget() if client_edit is not None else None
    layout = card.layout() if card is not None else None
    if layout is None:
        return None

    for index in range(layout.count()):
        child = layout.itemAt(index).layout()
        if isinstance(child, QGridLayout):
            return child
    return None


def payment_terms_value(self) -> str:
    edit = getattr(self, "payment_terms_edit", None)
    value = edit.text().strip() if edit is not None else ""
    return value or DEFAULT_PAYMENT_TERMS


def specifications_enabled(self) -> bool:
    button = getattr(self, "specifications_toggle_button", None)
    if button is None:
        return True
    return bool(button.isChecked())


def _ensure_payment_control(self) -> None:
    if getattr(self, "_stulz_payment_control_installed", False):
        return

    currency_combo = getattr(self, "currency_combo", None)
    grid = _find_files_grid(self)
    if currency_combo is None or grid is None:
        return

    # Currency was inserted by stulz_currency_runtime at row 3, column 1.
    # Replace that single widget with a horizontal container so payment terms sit
    # immediately next to the verified currency without adding another tall row.
    item = grid.itemAtPosition(3, 1)
    if item is None or item.widget() is not currency_combo:
        return

    parent = currency_combo.parentWidget()
    grid.removeWidget(currency_combo)

    container = QWidget(parent)
    row = QHBoxLayout(container)
    row.setContentsMargins(0, 0, 0, 0)
    row.setSpacing(10)

    currency_combo.setMinimumWidth(110)
    currency_combo.setMaximumWidth(170)
    row.addWidget(currency_combo, 0)

    payment_label = QLabel("Условия оплаты", container)
    payment_label.setObjectName("FormLabel")
    row.addWidget(payment_label, 0)

    saved = str(self.settings.value("stulz_payment_terms", DEFAULT_PAYMENT_TERMS) or "").strip()
    payment_edit = QLineEdit(saved or DEFAULT_PAYMENT_TERMS, container)
    payment_edit.setPlaceholderText(DEFAULT_PAYMENT_TERMS)
    payment_edit.setToolTip("Текст подставляется в поле условий оплаты готового коммерческого предложения.")
    payment_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    row.addWidget(payment_edit, 1)

    grid.addWidget(container, 3, 1)
    payment_edit.editingFinished.connect(self._on_stulz_payment_terms_changed)

    self.payment_terms_edit = payment_edit
    self._stulz_payment_control_container = container
    self._stulz_payment_control_installed = True


def _apply_specifications_toggle_state(self) -> None:
    enabled = specifications_enabled(self)
    button = getattr(self, "specifications_toggle_button", None)
    if button is not None:
        button.setText("Спецификации: ВКЛ" if enabled else "Спецификации: ВЫКЛ")
        button.setToolTip(
            "ВКЛ — в Word добавляются опции, технические характеристики и чертежи. "
            "ВЫКЛ — формируется только основная часть коммерческого предложения."
        )

    for name in (
        "spec_models_table",
        "spec_preview_button",
        "manual_spec_button",
        "auto_spec_button",
    ):
        widget = getattr(self, name, None)
        if widget is not None:
            widget.setEnabled(enabled)


def _ensure_specifications_toggle(self) -> None:
    if getattr(self, "_stulz_specifications_toggle_installed", False):
        return

    table = getattr(self, "spec_models_table", None)
    card = table.parentWidget() if table is not None else None
    layout = card.layout() if card is not None else None
    if layout is None:
        return

    button = QPushButton(card)
    button.setCheckable(True)
    button.setObjectName("GhostButton")
    button.setMinimumHeight(36)

    saved = _settings_bool(self.settings.value("stulz_specifications_enabled", True), True)
    button.setChecked(saved)
    button.clicked.connect(self._on_stulz_specifications_toggled)

    table_index = layout.indexOf(table)
    if table_index >= 0:
        layout.insertWidget(table_index, button)
    else:
        layout.addWidget(button)

    self.specifications_toggle_button = button
    self._stulz_specifications_toggle_installed = True
    self._apply_specifications_toggle_state()


def _ensure_offer_controls(self) -> None:
    self._ensure_payment_control()
    self._ensure_specifications_toggle()


def _on_payment_terms_changed(self) -> None:
    value = payment_terms_value(self)
    edit = getattr(self, "payment_terms_edit", None)
    if edit is not None and not edit.text().strip():
        edit.setText(value)
    try:
        self.settings.setValue("stulz_payment_terms", value)
        self.settings.sync()
    except Exception:
        pass
    try:
        self.refresh_preview()
    except Exception:
        pass


def _on_specifications_toggled(self, checked: bool) -> None:
    try:
        self.settings.setValue("stulz_specifications_enabled", bool(checked))
        self.settings.sync()
    except Exception:
        pass

    if not checked:
        table = getattr(self, "spec_models_table", None)
        if table is not None:
            table.setRowCount(0)

    self._apply_specifications_toggle_state()
    try:
        self.refresh_preview()
    except Exception:
        pass


def make_context(self):
    context = self._stulz_offer_options_original_make_context()
    options = dict(getattr(context, "brand_options", None) or {})
    options["stulz_payment_terms"] = payment_terms_value(self)
    options["stulz_include_specifications"] = specifications_enabled(self)
    context.brand_options = options
    return context


def validate_context(self, context) -> None:
    # When specification generation is disabled, a supplier/specification folder
    # must not block creation of the ordinary commercial offer.
    validate_context = context
    if not bool((getattr(context, "brand_options", None) or {}).get("stulz_include_specifications", True)):
        validate_context = copy.copy(context)
        validate_context.pdf_dir = None
    self._stulz_offer_options_original_validate_context(validate_context)


def remember_values(self) -> None:
    self._stulz_offer_options_original_remember_values()
    try:
        self.settings.setValue("stulz_payment_terms", payment_terms_value(self))
        self.settings.setValue("stulz_specifications_enabled", specifications_enabled(self))
        self.settings.sync()
    except Exception:
        pass


def clear_cache(self) -> None:
    self._stulz_offer_options_original_clear_cache()
    try:
        self.settings.remove("stulz_payment_terms")
        self.settings.remove("stulz_specifications_enabled")
        self.settings.sync()
    except Exception:
        pass

    edit = getattr(self, "payment_terms_edit", None)
    if edit is not None:
        edit.setText(DEFAULT_PAYMENT_TERMS)
    button = getattr(self, "specifications_toggle_button", None)
    if button is not None:
        button.setChecked(True)
    self._apply_specifications_toggle_state()


def refresh_spec_models(self, context=None) -> None:
    self._ensure_offer_controls()
    if not specifications_enabled(self):
        table = getattr(self, "spec_models_table", None)
        if table is not None:
            table.setRowCount(0)
        self._apply_specifications_toggle_state()
        return

    self._stulz_offer_options_original_refresh_spec_models(context)
    self._apply_specifications_toggle_state()


def _install() -> None:
    page_module = sys.modules.get("gui.pages.stulz_page")
    page_class = getattr(page_module, "StulzPage", None) if page_module is not None else None
    if page_class is None:
        return

    originals = {
        "make_context": "_stulz_offer_options_original_make_context",
        "validate_context": "_stulz_offer_options_original_validate_context",
        "remember_values": "_stulz_offer_options_original_remember_values",
        "clear_cache": "_stulz_offer_options_original_clear_cache",
        "refresh_spec_models": "_stulz_offer_options_original_refresh_spec_models",
    }
    for method_name, saved_name in originals.items():
        if not hasattr(page_class, saved_name):
            setattr(page_class, saved_name, getattr(page_class, method_name))

    for name, function in {
        "payment_terms_value": payment_terms_value,
        "specifications_enabled": specifications_enabled,
        "_ensure_payment_control": _ensure_payment_control,
        "_ensure_specifications_toggle": _ensure_specifications_toggle,
        "_ensure_offer_controls": _ensure_offer_controls,
        "_apply_specifications_toggle_state": _apply_specifications_toggle_state,
        "_on_stulz_payment_terms_changed": _on_payment_terms_changed,
        "_on_stulz_specifications_toggled": _on_specifications_toggled,
        "make_context": make_context,
        "validate_context": validate_context,
        "remember_values": remember_values,
        "clear_cache": clear_cache,
        "refresh_spec_models": refresh_spec_models,
    }.items():
        setattr(page_class, name, function)

    app = QApplication.instance()
    if app is None:
        return

    for widget in app.allWidgets():
        if isinstance(widget, page_class):
            widget._ensure_offer_controls()
            widget._apply_specifications_toggle_state()
            widget.refresh_preview()


_install()
