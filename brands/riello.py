from __future__ import annotations

from pathlib import Path
from typing import Any

from core.models import OfferContext
from core.riello_excel_exporter import (
    RielloQuoteConfig,
    RielloQuoteLine,
    build_output_filename,
    export_riello_excel,
    find_next_excel_revision,
)
from core.riello_price import (
    RielloPriceItem,
    default_price_path,
    find_item,
    item_power_kw,
    load_price_items,
    nearest_power_items,
    power_modules,
    rack_cabinets,
)

BRAND_NAME = "Riello"


def _as_float(value: Any, default: float = 0.0) -> float:
    text = str(value or "").strip().replace("\xa0", " ").replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except Exception:
        return default


def _fmt_money(value: float) -> str:
    try:
        text = f"{float(value):,.2f}".replace(",", "TEMP").replace(".", ",").replace("TEMP", " ")
        return text
    except Exception:
        return str(value)


def _fmt_qty(value: float) -> str:
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number).replace(".", ",")
    except Exception:
        return str(value)


def _options(context: OfferContext) -> dict[str, Any]:
    raw = getattr(context, "brand_options", None)
    return raw if isinstance(raw, dict) else {}


def _selected_item(items: list[RielloPriceItem], value: str, fallback: RielloPriceItem | None = None) -> RielloPriceItem:
    found = find_item(items, value, contains=False) or find_item(items, value, contains=True)
    if found:
        return found
    if fallback:
        return fallback
    if items:
        return items[0]
    raise ValueError("В прайсе Riello не найдены позиции для выбора.")


def build_quote_config(context: OfferContext) -> RielloQuoteConfig:
    options = _options(context)
    price_path = options.get("price_path") or str(default_price_path())
    items = load_price_items(price_path)

    required_power_kw = _as_float(options.get("required_power_kw"), 20.0)
    nearest_items = nearest_power_items(items, required_power_kw)

    requested_ups = str(options.get("ups_model") or "").strip()
    selected_ups = find_item(nearest_items, requested_ups, contains=False) if requested_ups else None
    if selected_ups is None and requested_ups:
        selected_ups = find_item(items, requested_ups, contains=False)
    if selected_ups is None:
        selected_ups = nearest_items[0] if nearest_items else _selected_item(items, "SRT 20 PM P", items[0] if items else None)

    prefix = selected_ups.model.split(" ", 1)[0]
    modules = power_modules(items, prefix=prefix)
    requested_module = str(options.get("power_module") or "").strip()
    if " 20 PM" in selected_ups.model.upper():
        selected_module = selected_ups
    else:
        selected_module = _selected_item(items, requested_module or f"{prefix} 20 PM P", modules[0] if modules else None)

    ups_qty = max(_as_float(options.get("ups_quantity"), 1.0), 0.0) or 1.0
    raw_modules_per_ups = str(options.get("modules_per_ups") or "").strip()
    if raw_modules_per_ups:
        modules_per_ups = max(_as_float(raw_modules_per_ups, 3.0), 0.0) or 3.0
    else:
        ups_power = item_power_kw(selected_ups)
        module_power = item_power_kw(selected_module)
        modules_per_ups = max(round(ups_power / module_power), 1) if ups_power > 0 and module_power > 0 else 3.0
    module_qty = ups_qty * modules_per_ups

    city = str(options.get("city") or context.city or "Алматы").replace("г.", "").strip() or "Алматы"
    currency = selected_ups.currency or selected_module.currency or str(options.get("currency") or "EUR")

    lines = [
        RielloQuoteLine(selected_ups, ups_qty, note=f"Позиция из прайса; запрос {required_power_kw:g} кВт" if required_power_kw else "Позиция из прайса"),
    ]
    if selected_module.model != selected_ups.model:
        lines.append(RielloQuoteLine(selected_module, module_qty, note=f"{_fmt_qty(modules_per_ups)} мод. на 1 шкаф"))

    return RielloQuoteConfig(
        client_name=context.client_name,
        city=city,
        currency=currency,
        ups_quantity=ups_qty,
        required_power_kw=required_power_kw,
        autonomy_min=str(options.get("autonomy_min") or "").strip(),
        battery_cabinet_type=str(options.get("battery_cabinet_type") or "").strip(),
        rate=_as_float(options.get("rate"), 1.0) or 1.0,
        margin_percent=_as_float(options.get("margin_percent"), 15.0),
        vat_percent=_as_float(options.get("vat_percent"), 0.0),
        special_percent=_as_float(options.get("special_percent"), 0.0),
        transport_cost=_as_float(options.get("transport_cost"), 2000.0),
        customs_clearance=_as_float(options.get("customs_clearance"), 200.0),
        certificate=_as_float(options.get("certificate"), 200.0),
        transport_to_customer=_as_float(options.get("transport_to_customer"), 1500.0),
        site_inspection=_as_float(options.get("site_inspection"), 0.0),
        installation_startup=_as_float(options.get("installation_startup"), 0.0),
        extra_cost=_as_float(options.get("extra_cost"), 0.0),
        lines=lines,
    )


def make_offer(context: OfferContext) -> Path:
    config = build_quote_config(context)
    template_path = context.calc_path
    if not template_path.exists():
        raise FileNotFoundError("Выберите Excel-шаблон расчета Riello.")

    base_prefix = "Riello_"
    revision = find_next_excel_revision(context.output_dir, base_prefix)
    output_path = context.output_dir / build_output_filename(config, revision=revision)
    return export_riello_excel(template_path, output_path, config)


def preview(context: OfferContext) -> str:
    lines: list[str] = [
        f"Заказчик: {context.client_name}",
        f"Excel-шаблон: {context.calc_path.name if context.calc_path and context.calc_path.exists() else 'не выбран'}",
        f"Прайс Riello: {default_price_path().name if default_price_path().exists() else 'не найден, будут использованы резервные позиции'}",
    ]

    try:
        config = build_quote_config(context)
    except Exception as exc:
        lines.append(f"Не удалось собрать конфигурацию Riello: {exc}")
        return "\n".join(lines)

    lines.extend([
        "",
        f"Требуемая мощность: {_fmt_qty(config.required_power_kw)} кВт",
        "Состав расчета:",
    ])
    for idx, line in enumerate(config.lines, start=1):
        item = line.item
        lines.append(
            f"{idx}. {item.model} — {item.code}; кол-во {_fmt_qty(line.qty)}; "
            f"габариты {item.dimensions}; вес {_fmt_qty(item.weight_kg)} кг; "
            f"цена {_fmt_money(item.price)} {item.currency}; сумма {_fmt_money(line.total)} {item.currency}"
        )
    lines.extend([
        "",
        f"Суммарный вес: {_fmt_qty(config.total_weight)} кг",
        f"Оборудование: {_fmt_money(config.total_equipment)} {config.currency}",
        f"Город/DDP: {config.city}",
        f"Курс: {_fmt_money(config.rate)}",
        f"Маржа: {_fmt_qty(config.margin_percent)}%",
        f"НДС: {_fmt_qty(config.vat_percent)}%",
        f"Спецусловие: {_fmt_qty(config.special_percent)}%",
    ])
    if config.autonomy_min:
        lines.append(f"Время автономии: {config.autonomy_min}")
    if config.battery_cabinet_type:
        lines.append(f"Тип батарейного шкафа: {config.battery_cabinet_type}")
    return "\n".join(lines)
