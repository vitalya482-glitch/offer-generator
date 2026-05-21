from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from core.docx_renderer import render_docx
from core.excel_reader import parse_stulz_calc
from core.models import CalcData, OfferContext, OfferItem

BRAND_NAME = "Stulz"


def sanitize_filename(value: str) -> str:
    bad = '<>:"/\\|?*'
    for ch in bad:
        value = value.replace(ch, "")
    return value.strip() or "Client"


def format_money(value: float) -> str:
    try:
        return f"{float(value):,.2f}".replace(",", " ")
    except Exception:
        return str(value)


def format_qty(value: float) -> str:
    try:
        value = float(value)
        return str(int(value)) if value.is_integer() else str(value)
    except Exception:
        return str(value)


def currency_name(currency: str) -> str:
    code = (currency or "").upper()
    if code == "KZT":
        return "тенге"
    if code == "EUR":
        return "евро"
    if code == "USD":
        return "долларов США"
    return code or ""


def item_to_template_dict(item: OfferItem, calc: CalcData) -> dict[str, Any]:
    return {
        "item_no": item.no,
        "item_name": item.name,
        "item_qty": format_qty(item.qty),
        "item_unit_price": format_money(item.unit_price),
        "item_total": format_money(item.total_price),
    }


def build_intro_text(calc: CalcData) -> str:
    models: list[str] = []
    for item in calc.items:
        model = (item.name or "").strip()
        if model and model not in models:
            models.append(model)

    model_text = ", ".join(models) if models else "оборудование Stulz"
    qty = format_qty(calc.quantity)

    return (
        "В ответ на Ваш запрос направляем коммерческое предложение "
        "на поставку прецизионных кондиционеров Stulz "
        f"в количестве {qty} шт.: {model_text}. "
        "Опции, включенные в комплектацию и технические характеристики "
        "указаны в спецификации коммерческого предложения."
    )


def build_total_price_block(calc: CalcData) -> str:
    return (
        "Итого, стоимость оборудования составляет: "
        f"{format_money(calc.total_price)} {currency_name(calc.currency)}, "
        f"с учетом НДС {format_qty(calc.vat_percent)}%."
    )


def build_currency_terms(calc: CalcData) -> str:
    if calc.currency.upper() == "EUR":
        return "Взаиморасчет осуществляется в тенге по курсу БанкаЦентрКредит РК на день оплаты."

    if calc.exchange_rate and calc.exchange_rate > 1.01:
        return (
            "Расчет был сделан в тенге согласно курса "
            f"1 EUR = {format_money(calc.exchange_rate)} тенге, "
            "при изменении данного курса более чем на 3 тенге будет сделан перерасчет."
        )

    return "Стоимость указана в валюте коммерческого предложения."


def build_replacements(context: OfferContext, calc: CalcData) -> dict[str, Any]:
    cur_name = currency_name(calc.currency)
    return {
        "{{offer_date}}": datetime.now().strftime("%d.%m.%Y"),
        "{{offer_version}}": context.version or calc.version or "1",
        "{{client_company_full}}": context.client_name,
        "{{intro_text}}": build_intro_text(calc),
        "{{unit_price_header}}": f"Цена за единицу, {cur_name}",
        "{{total_price_header}}": f"Сумма, {cur_name}",
        "{{total_label}}": "ИТОГО",
        "{{grand_total}}": format_money(calc.total_price),
        "{{total_price_block}}": build_total_price_block(calc),
        "{{payment_terms}}": "100% предоплата, если иное не согласовано сторонами.",
        "{{delivery_time}}": "Срок поставки уточняется после размещения заказа.",
        "{{delivery_terms}}": calc.delivery_basis,
        "{{installation_terms}}": "Монтажные работы не включены, если иное не указано в предложении.",
        "{{startup_terms}}": "Пусконаладочные работы не включены, если иное не указано в предложении.",
        "{{offer_validity}}": "Коммерческое предложение действительно в течение 14 календарных дней.",
        "{{currency_terms}}": build_currency_terms(calc),
        "{{signer_name}}": "",
        "{{signer_position}}": "",
        "{{manager_name}}": "",
        "{{manager_position}}": "",
        "{{manager_email}}": "",
        "{{manager_phone}}": "",
        "{{options_title}}": "Опции, включенные в комплектацию оборудования:",
        "{{options_table}}": "",
        "{{technical_specs_table}}": "",
    }


def load_calc(context: OfferContext) -> CalcData:
    calc = parse_stulz_calc(context.calc_path, context.sheet_name)
    if not calc.items:
        raise ValueError(
            "В выбранном листе Excel не найдены позиции оборудования. "
            "Проверьте лист Excel или структуру расчета."
        )
    return calc


def make_offer(context: OfferContext) -> Path:
    calc = load_calc(context)
    replacements = build_replacements(context, calc)
    items = [item_to_template_dict(item, calc) for item in calc.items]

    filename = f"КП_{sanitize_filename(context.client_name)}_Stulz_{datetime.now():%Y-%m-%d}.docx"
    output_path = context.output_dir / filename

    render_docx(
        template_path=context.template_path,
        output_path=output_path,
        replacements=replacements,
        items=items,
    )

    return output_path


def preview(context: OfferContext) -> str:
    calc = load_calc(context)
    models = []
    for item in calc.items:
        if item.name and item.name not in models:
            models.append(item.name)

    lines = [
        f"Заказчик: {context.client_name}",
        f"Лист Excel: {calc.sheet_name}",
        f"Версия расчета: {calc.version}",
        f"Валюта: {calc.currency}",
        f"Курс: {format_money(calc.exchange_rate)}",
        f"НДС: {format_qty(calc.vat_percent)}%",
        f"Условия поставки: {calc.delivery_basis}",
        f"Модели: {', '.join(models) if models else '-'}",
        f"Количество: {format_qty(calc.quantity)}",
        f"Сумма: {format_money(calc.total_price)} {currency_name(calc.currency)}",
        "",
        "Позиции:",
    ]
    for item in calc.items:
        lines.append(
            f"{item.no}. {item.name} | кол-во {format_qty(item.qty)} | "
            f"цена {format_money(item.unit_price)} | сумма {format_money(item.total_price)}"
        )
    return "\n".join(lines)


# Backward-compatible aliases for older code/tests.
generate_offer = make_offer
build_preview = preview
