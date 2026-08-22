from __future__ import annotations

import copy
import tempfile
from pathlib import Path

from docx import Document

import brands.stulz_position_selection_runtime as _previous
from core.docx_renderer import _find_table_with_tags, _find_tag_paragraph, _remove_element


# Re-export the complete mature STULZ runtime: physical specification mapping,
# legends, cable quantities, compressor sections and selectable Calc positions.
#
# IMPORTANT: build a stable snapshot before writing anything into globals().
# The previous runtime itself contains a private attribute named ``_previous``.
# Keep a separate module reference because the export loop also writes private
# names into globals() and can otherwise replace this module's ``_previous``.
_POSITION_SELECTION_RUNTIME = _previous
_runtime_exports = tuple(
    (name, getattr(_POSITION_SELECTION_RUNTIME, name))
    for name in dir(_POSITION_SELECTION_RUNTIME)
    if not name.startswith("__")
)
for _name, _value in _runtime_exports:
    globals()[_name] = _value
_previous = _POSITION_SELECTION_RUNTIME


# ---------------------------------------------------------------------------
# Final commercial settings chosen/verified on the STULZ page
# ---------------------------------------------------------------------------
_BASE_LOAD_CALC = globals()["load_calc"]
_BASE_PREVIEW = _previous.preview
_BASE_MAKE_OFFER = _previous.make_offer
_BASE_BUILD_SPECIFICATION_BLOCKS = _previous._BUILD_SPECIFICATION_BLOCKS
_BASE_BUILD_REPLACEMENTS = _previous._BUILD_REPLACEMENTS
_SUPPORTED_CURRENCIES = {"KZT", "EUR", "USD"}
_DEFAULT_PAYMENT_TERMS = "70% предоплата, 30% после поставки оборудования"
_SPEC_INTRO_SENTENCE = (
    "Опции, включенные в комплектацию и технические характеристики "
    "указаны в спецификации коммерческого предложения."
)


def _brand_options(context) -> dict:
    options = getattr(context, "brand_options", None) or {}
    return options if isinstance(options, dict) else {}


def _selected_currency(context) -> str:
    value = str(_brand_options(context).get("stulz_currency", "") or "").upper().strip()
    return value if value in _SUPPORTED_CURRENCIES else ""


def _payment_terms(context) -> str:
    value = str(_brand_options(context).get("stulz_payment_terms", "") or "").strip()
    return value or _DEFAULT_PAYMENT_TERMS


def _include_specifications(context) -> bool:
    value = _brand_options(context).get("stulz_include_specifications", True)
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"0", "false", "no", "off", "выкл"}:
        return False
    return True


def load_calc(context):
    calc = _BASE_LOAD_CALC(context)
    currency = _selected_currency(context)
    if not currency or str(getattr(calc, "currency", "") or "").upper().strip() == currency:
        return calc

    # Do not mutate a CalcData object that may originate from the parser cache.
    overridden = copy.copy(calc)
    overridden.currency = currency
    return overridden


def _build_specification_blocks_for_offer(context, calc):
    """Skip every supplier-PDF lookup when specifications are disabled."""
    if not _include_specifications(context):
        return [], []
    return _BASE_BUILD_SPECIFICATION_BLOCKS(context, calc)


def build_replacements(context, calc, offer_version=None):
    replacements = dict(_BASE_BUILD_REPLACEMENTS(context, calc, offer_version=offer_version))
    replacements["{{payment_terms}}"] = _payment_terms(context)
    replacements["{{offer_validity}}"] = "30 дней"

    if not _include_specifications(context):
        intro = str(replacements.get("{{intro_text}}", "") or "")
        intro = intro.replace(_SPEC_INTRO_SENTENCE, "")
        replacements["{{intro_text}}"] = " ".join(intro.split())

    return replacements


def preview(context) -> str:
    """Keep the mature preview, but expose final offer controls explicitly."""
    text = _BASE_PREVIEW(context)
    calc = load_calc(context)
    include_specs = _include_specifications(context)
    payment = _payment_terms(context)
    currency = str(getattr(calc, "currency", "") or "").upper().strip()

    result: list[str] = []
    payment_written = False
    total_written = False

    for line in text.splitlines():
        stripped = line.strip()

        if not include_specs:
            if stripped.startswith("Спецификации:"):
                result.append("Спецификации: ВЫКЛ")
                continue
            if stripped.startswith((
                "Чертежи для вставки:",
                "Опций для спецификации:",
                "Строк тех. характеристик:",
                "Файлы чертежей:",
            )):
                continue

        if stripped.startswith("Условия оплаты:"):
            result.append(f"Условия оплаты: {payment}")
            payment_written = True
            continue

        if stripped.startswith(("Сумма:", "Итоговая сумма:", "Итоговая стоимость:")):
            result.append(f"Итоговая стоимость: {format_money(calc.total_price)} {currency}")
            total_written = True
            continue

        result.append(line)
        if stripped.startswith("Условия поставки:") and not payment_written:
            result.append(f"Условия оплаты: {payment}")
            payment_written = True

    if not payment_written:
        result.append(f"Условия оплаты: {payment}")
    if not total_written:
        result.append(f"Итоговая стоимость: {format_money(calc.total_price)} {currency}")

    return "\n".join(result)


def _template_without_specifications(template_path: str | Path) -> Path:
    """Create a temporary Word template with the complete STULZ spec block removed."""
    doc = Document(str(template_path))

    # Remove template-row tables first. Afterwards remove their title/tag
    # paragraphs, including the older generic table placeholders if present.
    for tags in (
        ["{{opt_no}}", "{{opt_name}}", "{{opt_qty}}"],
        ["{{technical_specs_parameter}}", "{{technical_specs_value}}"],
    ):
        table = _find_table_with_tags(doc, tags)
        if table is not None:
            _remove_element(table._tbl)

    for tag in (
        "{{options_title}}",
        "{{technical_specs_title}}",
        "{{options_table}}",
        "{{technical_specs_table}}",
    ):
        paragraph = _find_tag_paragraph(doc, tag)
        if paragraph is not None:
            _remove_element(paragraph._p)

    handle = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    temp_path = Path(handle.name)
    handle.close()
    doc.save(str(temp_path))
    return temp_path


def make_offer(context) -> Path:
    if _include_specifications(context):
        return _BASE_MAKE_OFFER(context)

    temporary_template = _template_without_specifications(context.template_path)
    stripped_context = copy.copy(context)
    stripped_context.template_path = temporary_template
    try:
        return _BASE_MAKE_OFFER(stripped_context)
    finally:
        try:
            temporary_template.unlink(missing_ok=True)
        except Exception:
            pass


# preview() and make_offer() were defined in stulz_position_selection_runtime and
# resolve these private helpers from that module's globals. Patch those helpers
# while keeping the full public specification builder available to the manual
# preview dialog when specifications are switched back on.
_previous.load_calc = load_calc
_previous._BUILD_SPECIFICATION_BLOCKS = _build_specification_blocks_for_offer
_previous._BUILD_REPLACEMENTS = build_replacements
globals()["load_calc"] = load_calc
globals()["build_replacements"] = build_replacements
globals()["preview"] = preview
globals()["make_offer"] = make_offer


# Install GUI-only layout layers last: full content -> currency validation ->
# payment/specification controls. Static imports also ensure PyInstaller includes
# every runtime module in the small App package.
import gui.pages.stulz_full_content_runtime as _full_content_ui  # noqa: E402,F401
import gui.pages.stulz_currency_runtime as _currency_ui  # noqa: E402,F401
import gui.pages.stulz_offer_options_runtime as _offer_options_ui  # noqa: E402,F401
