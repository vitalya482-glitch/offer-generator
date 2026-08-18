from __future__ import annotations

import copy

import brands.stulz_position_selection_runtime as _previous


# Re-export the complete mature STULZ runtime: physical specification mapping,
# legends, cable quantities, compressor sections and selectable Calc positions.
#
# IMPORTANT: build a stable snapshot before writing anything into globals().
# The previous runtime itself contains a private attribute named ``_previous``.
# The old loop overwrote this module's own ``_previous`` variable midway through
# the export and the very next ``getattr(_previous, ...)`` was executed against
# brands.stulz_compressor_runtime instead of stulz_position_selection_runtime.
# That produced:
#   module 'brands.stulz_compressor_runtime' has no attribute '_selection_ui'
# during commercial-offer generation/import.
_runtime_exports = tuple(
    (name, getattr(_previous, name))
    for name in dir(_previous)
    if not name.startswith("__")
)
for _name, _value in _runtime_exports:
    globals()[_name] = _value


# ---------------------------------------------------------------------------
# Final commercial currency chosen/verified in the STULZ page
# ---------------------------------------------------------------------------
# The legacy Excel reader assumes EUR when it cannot recognise a currency.
# The GUI now performs a conservative check and stores the confirmed value in
# OfferContext.brand_options. Apply that value at the top runtime so preview,
# totals, amount-in-words and Word generation all use exactly the same currency.
_BASE_LOAD_CALC = globals()["load_calc"]
_SUPPORTED_CURRENCIES = {"KZT", "EUR", "USD"}


def _selected_currency(context) -> str:
    options = getattr(context, "brand_options", None) or {}
    if not isinstance(options, dict):
        return ""
    value = str(options.get("stulz_currency", "") or "").upper().strip()
    return value if value in _SUPPORTED_CURRENCIES else ""


def load_calc(context):
    calc = _BASE_LOAD_CALC(context)
    currency = _selected_currency(context)
    if not currency or str(getattr(calc, "currency", "") or "").upper().strip() == currency:
        return calc

    # Do not mutate a CalcData object that may originate from the parser cache.
    overridden = copy.copy(calc)
    overridden.currency = currency
    return overridden


# preview() and make_offer() were defined in stulz_position_selection_runtime and
# resolve load_calc from that module's globals. Patch that namespace as well as
# this top-level re-export.
_previous.load_calc = load_calc
globals()["load_calc"] = load_calc


# Install the final GUI-only layout layers last. The full-content patch removes
# nested scrolling; the currency patch inserts the verification field and makes
# it mandatory before offer generation.
import gui.pages.stulz_full_content_runtime as _full_content_ui  # noqa: E402,F401
import gui.pages.stulz_currency_runtime as _currency_ui  # noqa: E402,F401
