from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import brands.stulz as _base
from core.models import CalcData, OfferContext, OfferItem
from core.stulz_spec_catalog import discover_stulz_spec_entries
from core.stulz_specification import build_stulz_specification


# Re-export the public surface of the existing STULZ module. The original
# module remains the single place for pricing and Word rendering; this wrapper
# replaces physical specification matching and a few STULZ description rules.
for _name in dir(_base):
    if not _name.startswith("__"):
        globals()[_name] = getattr(_base, _name)


def _model_key(value: str) -> str:
    return _base._spec_model_key(value)


def _positive_qty(value: object, default: float = 1.0) -> float:
    try:
        qty = float(str(value).replace(",", "."))
    except Exception:
        qty = default
    return qty if qty > 0 else default


def _explicit_physical_rows(context: OfferContext) -> list[dict[str, Any]]:
    """Return enabled GUI rows that already point to a concrete Calc folder."""

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in getattr(context, "spec_models", []) or []:
        if not row.get("enabled", True):
            continue

        model = str(row.get("model") or "").strip()
        source_dir = str(row.get("source_dir") or "").strip()
        calc_pdf = str(row.get("calc_pdf") or "").strip()
        key = str(row.get("key") or calc_pdf or source_dir or "").strip().lower()
        if not model or not (source_dir or calc_pdf):
            continue
        if key and key in seen:
            continue

        if calc_pdf and not source_dir:
            source_dir = str(Path(calc_pdf).parent)

        rows.append(
            {
                "key": key,
                "model": model,
                "qty": _positive_qty(row.get("qty_value", row.get("qty", 1))),
                "source_dir": source_dir,
                "calc_pdf": calc_pdf,
                "source_label": str(row.get("source_label") or ""),
            }
        )
        if key:
            seen.add(key)

    return rows


def _selected_model_keys(context: OfferContext, calc: CalcData) -> set[str]:
    """Models allowed by the GUI; fall back to the models from the calculation."""

    keys: set[str] = set()
    has_rows = False
    for row in getattr(context, "spec_models", []) or []:
        has_rows = True
        if not row.get("enabled", True):
            continue
        key = _model_key(str(row.get("model") or ""))
        if key:
            keys.add(key)

    if has_rows:
        return keys

    for item in calc.items:
        key = _model_key(item.name)
        if key:
            keys.add(key)
    return keys


def _discover_selected_physical_specs(context: OfferContext, calc: CalcData) -> list[dict[str, Any]]:
    """Discover every physical specification set instead of grouping by model.

    This is also the compatibility path for saved/older GUI state where one row
    such as ASR552AS / qty 5 represented several Calc folders.
    """

    allowed_keys = _selected_model_keys(context, calc)
    result: list[dict[str, Any]] = []
    for entry in discover_stulz_spec_entries(context.pdf_dir):
        if allowed_keys and _model_key(entry.model) not in allowed_keys:
            continue
        result.append(
            {
                "key": entry.key,
                "model": entry.model,
                "qty": entry.quantity,
                "source_dir": str(entry.source_dir),
                "calc_pdf": str(entry.calc_pdf),
                "source_label": entry.source_label,
            }
        )
    return result


def _physical_specs(context: OfferContext, calc: CalcData) -> list[dict[str, Any]]:
    explicit = _explicit_physical_rows(context)
    if explicit:
        return explicit

    discovered = _discover_selected_physical_specs(context, calc)
    if discovered:
        return discovered

    # Last-resort compatibility fallback: keep the old selection semantics.
    result: list[dict[str, Any]] = []
    for index, row in enumerate(getattr(context, "spec_models", []) or []):
        if not row.get("enabled", True):
            continue
        model = str(row.get("model") or "").strip()
        if not model:
            continue
        result.append(
            {
                "key": f"legacy:{index}:{_model_key(model)}",
                "model": model,
                "qty": _positive_qty(row.get("qty_value", row.get("qty", 1))),
                "source_dir": str(context.pdf_dir or ""),
                "calc_pdf": "",
                "source_label": "",
            }
        )
    return result


def _calc_for_physical_spec(calc: CalcData, model: str, qty: float) -> CalcData:
    """Build an isolated CalcData only for specification parsing.

    Price values are irrelevant for the PDF parsers. Using a synthetic item also
    avoids accidentally binding all duplicate ASR552AS rows to the first Excel row.
    """

    item = OfferItem(no=1, name=model, qty=qty, unit_price=0.0, total_price=0.0)
    return CalcData(
        sheet_name=calc.sheet_name,
        version=calc.version,
        currency=calc.currency,
        vat_percent=calc.vat_percent,
        exchange_rate=calc.exchange_rate,
        delivery_basis=calc.delivery_basis,
        items=[item],
        options=list(calc.options),
        installation_included=calc.installation_included,
    )


def build_specification_blocks(context: OfferContext, calc: CalcData) -> tuple[list[dict[str, Any]], list[str]]:
    blocks: list[dict[str, Any]] = []
    warnings: list[str] = []

    for selected in _physical_specs(context, calc):
        model = str(selected["model"])
        qty = _positive_qty(selected.get("qty", 1))
        source_dir = Path(str(selected.get("source_dir") or context.pdf_dir or ""))
        model_calc = _calc_for_physical_spec(calc, model, qty)

        # Critical rule: once Calc.pdf has identified a physical specification
        # folder, Calc/WinPlan/drawing are searched only inside that folder.
        specification = build_stulz_specification(source_dir, model_calc)
        label = str(selected.get("source_label") or source_dir.name or model)
        warnings.extend(f"{model} [{label}]: {warning}" for warning in specification.warnings)
        totals = specification.totals

        blocks.append(
            {
                "model": model,
                "calc_model": getattr(totals, "model", "") if totals else "",
                "quantity": getattr(totals, "quantity", None) if totals else qty,
                "total_list_price": getattr(totals, "total_list_price", None) if totals else None,
                "total_purchase_price": getattr(totals, "total_purchase_price", None) if totals else None,
                "unit_list_price": getattr(totals, "unit_list_price", None) if totals else None,
                "unit_purchase_price": getattr(totals, "unit_purchase_price", None) if totals else None,
                "currency": getattr(totals, "currency", "") if totals else "",
                "options_title": f"Опции, включенные в комплектацию кондиционеров {model}:",
                "options": [
                    {
                        "description": option.description,
                        "qty": option.qty,
                        "code": option.code,
                        "source_name": option.source_name,
                        "translated": option.translated,
                    }
                    for option in specification.options
                ],
                "technical_specs_title": f"Технические характеристики кондиционеров {model}:",
                "technical_specs": [
                    {"name": row.name, "value": row.value, "is_section": row.is_section}
                    for row in specification.technical_specs
                ],
                "calc_pdf": specification.calc_pdf,
                "winplan_pdf": specification.winplan_pdf,
                "drawing_pdf": specification.drawing_pdf,
                "source_dir": source_dir,
                "source_label": label,
                "spec_key": str(selected.get("key") or ""),
            }
        )

    return blocks, warnings


def _dimension_component(block: dict[str, Any], ru_name: str, en_name: str) -> str:
    """Read one unit dimension from translated or original WinPlan labels."""

    return (
        _base._tech_value(block, (ru_name,), ("unit",))
        or _base._tech_value(block, (en_name,), ("unit",))
    )


def _clean_dimension_component(value: str) -> str:
    text = str(value or "").replace("\u00a0", " ").strip()
    text = re.sub(r"\s*(?:мм|mm)\b\.?\s*$", "", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()


def _unit_dimensions_text(block: dict[str, Any]) -> str:
    """Return complete indoor-unit dimensions as W x D x H.

    The old implementation first matched any row containing the word
    "Габариты", which meant it often returned only the first component. WinPlan
    provides height, width and depth separately, so require all three values and
    combine them explicitly. A pre-combined dimensions row is used only as a
    fallback when it visibly contains three dimensions.
    """

    width = _dimension_component(block, "ширина", "width")
    depth = _dimension_component(block, "глубина", "depth")
    height = _dimension_component(block, "высота", "height")

    if width and depth and height:
        values = [_clean_dimension_component(value) for value in (width, depth, height)]
        if all(values):
            return " × ".join(values) + " мм"

    direct = _base._tech_value(block, ("габарит",), ("unit",))
    if direct:
        normalized = str(direct).replace("х", "×").replace("x", "×").replace("X", "×")
        if normalized.count("×") >= 2:
            return normalized.strip()

    return ""


def _build_offer_item_description(
    item: OfferItem,
    block: dict[str, Any] | None,
    options: dict[str, bool],
) -> str:
    if not block:
        return item.name

    model = str(block.get("calc_model") or block.get("model") or item.name).strip() or item.name
    parts: list[str] = []

    if options.get("stulz_unit", True):
        parts.append(f"Прецизионный кондиционер Stulz {model}")
    else:
        parts.append(model)

    if options.get("cooling_capacity", True):
        cooling = _base._cooling_capacity_text(block)
        if cooling:
            parts.append(f"хладопроизводительность {cooling}")

    if options.get("unit_dimensions", True):
        dimensions = _unit_dimensions_text(block)
        if dimensions:
            parts.append(f"размеры внутреннего блока (Ш×Г×В) {dimensions}")

    if options.get("condenser", True):
        condenser = _base._condenser_text(block)
        if condenser:
            parts.append(f"наружный блок (конденсор) - {condenser}")

    return ", ".join(parts)


# Existing functions such as make_offer() execute in the namespace of
# brands.stulz. Patch the globals they call so the rest of the mature STULZ
# implementation stays untouched.
_base.build_specification_blocks = build_specification_blocks
_base._unit_dimensions_text = _unit_dimensions_text
_base._build_offer_item_description = _build_offer_item_description


def preview(context: OfferContext) -> str:
    """Preview the calculation and the exact item descriptions that will enter КП."""

    calc = _base.load_calc(context)
    spec_blocks, spec_warnings = build_specification_blocks(context, calc)
    offer_items = _base.build_offer_items(context, calc, spec_blocks)
    models: list[str] = []
    for item in calc.items:
        if item.name and item.name not in models:
            models.append(item.name)

    lines = [
        f"Заказчик: {context.client_name}",
        f"Лист Excel: {calc.sheet_name}",
        f"Версия расчета: {calc.version}",
        f"Версия КП: {_base.find_next_offer_version(context.output_dir, context.client_name, calc.sheet_name)}",
        f"Валюта: {calc.currency}",
        f"Курс: {_base.format_money(calc.exchange_rate)}",
        f"НДС: {_base.format_qty(calc.vat_percent)}%",
        f"Условия поставки: {calc.delivery_basis}",
        f"Монтаж/ПНР: {'включены' if getattr(calc, 'installation_included', False) else 'не включены'}",
        f"Моделей для спецификации: {len(spec_blocks)}",
        f"Опций для спецификации: {sum(len(block.get('options', [])) for block in spec_blocks)}",
        f"Строк тех. характеристик: {sum(len(block.get('technical_specs', [])) for block in spec_blocks)}",
        f"Модели: {', '.join(models) if models else '-'}",
        f"Количество: {_base.format_qty(calc.quantity)}",
        f"Сумма: {_base.format_money(calc.total_price)} {_base.currency_name(calc.currency)}",
        "",
        "Предупреждения спецификации:",
    ]
    if spec_warnings:
        lines.extend(f"- {warning}" for warning in spec_warnings)
    else:
        lines.append("- нет")

    lines.extend(["", "Позиции для КП:"])
    for source_item, offer_item in zip(calc.items, offer_items):
        lines.append(
            f"{source_item.no}. {offer_item['item_name']} | кол-во {_base.format_qty(source_item.qty)} | "
            f"цена {_base.format_money(source_item.unit_price)} | "
            f"сумма {_base.format_money(source_item.total_price)}"
        )

    return "\n".join(lines)


def make_offer(context: OfferContext) -> Path:
    return _base.make_offer(context)
