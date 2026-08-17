from __future__ import annotations

import copy
from pathlib import Path
from typing import Any

import brands.stulz as _base
import brands.stulz_compressor_runtime as _previous
import brands.stulz_runtime as _physical_runtime
from core.docx_renderer import render_docx


# Preserve every mature STULZ feature accumulated in the previous runtime:
# physical spec mapping, legends, cable units and compressor block splitting.
for _name in dir(_previous):
    if not _name.startswith("__"):
        globals()[_name] = getattr(_previous, _name)


_FULL_LOAD_CALC = globals()["load_calc"]
_BUILD_SPECIFICATION_BLOCKS = globals()["build_specification_blocks"]
_BUILD_OFFER_ITEMS = globals()["build_offer_items"]
_BUILD_REPLACEMENTS = globals()["build_replacements"]
_FIND_NEXT_OFFER_VERSION = globals()["find_next_offer_version"]
_BUILD_OFFER_FILENAME = globals()["build_offer_filename"]
_FORMAT_MONEY = globals()["format_money"]
_FORMAT_QTY = globals()["format_qty"]
_CURRENCY_NAME = globals()["currency_name"]


def _disabled_item_indexes(context: Any) -> set[int]:
    options = getattr(context, "brand_options", None) or {}
    raw = options.get("stulz_disabled_item_indexes", []) if isinstance(options, dict) else []
    result: set[int] = set()
    for value in raw or []:
        try:
            result.add(int(value))
        except Exception:
            continue
    return result


def _filter_calc(context: Any, calc: Any) -> Any:
    """Return a shallow CalcData copy containing only user-enabled Calc rows."""

    disabled = _disabled_item_indexes(context)
    if not disabled:
        return calc

    filtered = copy.copy(calc)
    items = []
    for source_index, source_item in enumerate(getattr(calc, "items", []) or []):
        if source_index in disabled:
            continue
        item = copy.copy(source_item)
        # Commercial-offer numbering must remain contiguous after exclusions.
        item.no = len(items) + 1
        items.append(item)
    filtered.items = items
    return filtered


def load_calc(context):
    """Public STULZ loader respects Calc-position switches from the GUI."""

    return _filter_calc(context, _FULL_LOAD_CALC(context))


def preview(context) -> str:
    """Preview only rows currently enabled for the commercial offer."""

    calc = load_calc(context)
    spec_blocks, spec_warnings = _BUILD_SPECIFICATION_BLOCKS(context, calc)
    offer_items = _BUILD_OFFER_ITEMS(context, calc, spec_blocks)

    models: list[str] = []
    for item in calc.items:
        if item.name and item.name not in models:
            models.append(item.name)

    drawing_files = [
        Path(str(block.get("drawing_pdf"))).name
        for block in spec_blocks
        if block.get("drawing_pdf")
    ]

    lines = [
        f"Заказчик: {context.client_name}",
        f"Лист Excel: {calc.sheet_name}",
        f"Версия расчета: {calc.version}",
        f"Версия КП: {_FIND_NEXT_OFFER_VERSION(context.output_dir, context.client_name, calc.sheet_name)}",
        f"Валюта: {calc.currency}",
        f"Курс: {_FORMAT_MONEY(calc.exchange_rate)}",
        f"НДС: {_FORMAT_QTY(calc.vat_percent)}%",
        f"Условия поставки: {calc.delivery_basis}",
        f"Монтаж/ПНР: {'включены' if getattr(calc, 'installation_included', False) else 'не включены'}",
        f"Спецификации: найдено {len(spec_blocks)}",
        f"Чертежи для вставки: {len(drawing_files)}",
        f"Опций для спецификации: {sum(len(block.get('options', [])) for block in spec_blocks)}",
        f"Строк тех. характеристик: {sum(len(block.get('technical_specs', [])) for block in spec_blocks)}",
        f"Модели: {', '.join(models) if models else '-'}",
        f"Количество: {_FORMAT_QTY(calc.quantity)}",
        f"Сумма: {_FORMAT_MONEY(calc.total_price)} {_CURRENCY_NAME(calc.currency)}",
    ]

    if drawing_files:
        lines.append(f"Файлы чертежей: {'; '.join(drawing_files)}")

    if spec_warnings:
        lines.extend(["", "Предупреждения:"])
        lines.extend(f"- {warning}" for warning in spec_warnings)
    else:
        lines.extend(["", "Предупреждения: нет"])

    # The interactive GUI table renders offer positions. Keep a text fallback for
    # CLI/legacy consumers, but the GUI removes this tail from its QTextEdit.
    lines.extend(["", "Позиции для КП:"])
    for source_item, offer_item in zip(calc.items, offer_items):
        lines.append(
            f"{source_item.no}. {offer_item['item_name']} | кол-во {_FORMAT_QTY(source_item.qty)} | "
            f"цена {_FORMAT_MONEY(source_item.unit_price)} | "
            f"сумма {_FORMAT_MONEY(source_item.total_price)}"
        )

    return "\n".join(lines)


def make_offer(context) -> Path:
    """Generate Word using only Calc rows enabled by the user."""

    calc = load_calc(context)
    if not calc.items:
        raise ValueError("Выберите хотя бы одну позицию Calc для включения в коммерческое предложение.")

    spec_blocks, _warnings = _BUILD_SPECIFICATION_BLOCKS(context, calc)
    offer_version = _FIND_NEXT_OFFER_VERSION(context.output_dir, context.client_name, calc.sheet_name)
    replacements = _BUILD_REPLACEMENTS(context, calc, offer_version=offer_version)
    items = _BUILD_OFFER_ITEMS(context, calc, spec_blocks)

    filename = _BUILD_OFFER_FILENAME(context.client_name, offer_version)
    output_path = context.output_dir / filename

    return render_docx(
        template_path=context.template_path,
        output_path=output_path,
        replacements=replacements,
        items=items,
        stulz_spec_blocks=spec_blocks,
    )


# stulz_ui_runtime keeps a module-object reference to brands.stulz_runtime.
# Patch it so the existing calculation preview automatically uses the filtered
# CalcData without duplicating the mature financing/spec-summary UI code.
_physical_runtime.load_calc = load_calc
_physical_runtime.preview = preview

# Some mature helpers execute in brands.stulz globals. Keep direct calls made by
# older code consistent with the final runtime too.
_base.preview = preview

# Install the final GUI layer last, after every previous STULZ patch is active.
import gui.pages.stulz_position_selection_runtime as _selection_ui  # noqa: E402,F401
