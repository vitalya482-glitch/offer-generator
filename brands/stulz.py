from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from core.docx_renderer import render_docx
from core.excel_reader import read_calc_excel


@dataclass
class OfferContext:
    client_name: str
    template_path: Path
    excel_path: Path
    output_dir: Path


def sanitize_filename(value: str) -> str:
    bad = '<>:"/\\|?*'
    for ch in bad:
        value = value.replace(ch, "")
    return value.strip()


def build_intro_text(calc) -> str:
    models = []

    for item in calc.items:
        model = item.get("item_name", "").strip()
        if model and model not in models:
            models.append(model)

    model_text = ", ".join(models)

    qty = sum(int(item.get("item_qty", 0)) for item in calc.items)

    return (
        "В ответ на Ваш запрос направляем коммерческое предложение "
        f"на поставку прецизионных кондиционеров Stulz "
        f"в количестве {qty} шт.: {model_text}. "
        "Опции, включенные в комплектацию и технические характеристики "
        "указаны в спецификации коммерческого предложения."
    )


def build_total_price_block(calc) -> str:
    return (
        f"Итого, стоимость оборудования составляет: "
        f"{calc.total_words} "
        f"({calc.grand_total}) "
        f"{calc.currency_name}, "
        f"с учетом НДС {calc.vat_percent}%."
    )


def build_currency_terms(calc) -> str:
    if calc.currency == "EUR":
        return (
            "Взаиморасчет осуществляется в тенге "
            "по курсу БанкаЦентрКредит РК на день оплаты."
        )

    return (
        f"Расчет был сделан в тенге согласно курса "
        f"1 EUR = {calc.exchange_rate} тенге, "
        f"при изменении данного курса более чем на 3 тенге "
        f"будет сделан перерасчет."
    )


def build_replacements(calc) -> dict[str, Any]:
    return {
        "{{offer_date}}": datetime.now().strftime("%d %B %Y г."),
        "{{offer_version}}": "1",
        "{{client_company_full}}": calc.client_name,
        "{{intro_text}}": build_intro_text(calc),
        "{{unit_price_header}}": f"Цена за единицу, {calc.currency_name}",
        "{{total_price_header}}": f"Сумма, {calc.currency_name}",
        "{{total_label}}": "ИТОГО",
        "{{grand_total}}": calc.grand_total,
        "{{total_price_block}}": build_total_price_block(calc),
        "{{payment_terms}}": calc.payment_terms,
        "{{delivery_time}}": calc.delivery_time,
        "{{delivery_terms}}": calc.delivery_terms,
        "{{installation_terms}}": calc.installation_terms,
        "{{startup_terms}}": calc.startup_terms,
        "{{offer_validity}}": calc.offer_validity,
        "{{currency_terms}}": build_currency_terms(calc),
        "{{signer_name}}": calc.signer_name,
        "{{signer_position}}": calc.signer_position,
        "{{manager_name}}": calc.manager_name,
        "{{manager_position}}": calc.manager_position,
        "{{manager_email}}": calc.manager_email,
        "{{manager_phone}}": calc.manager_phone,
        "{{options_title}}": "Опции, включенные в комплектацию оборудования:",
        "{{options_table}}": "",
        "{{technical_specs_table}}": "",
    }


def generate_offer(context: OfferContext) -> Path:
    calc = read_calc_excel(context.excel_path)

    replacements = build_replacements(calc)

    now = datetime.now()

    filename = (
        f"КП_"
        f"{sanitize_filename(calc.client_name)}_"
        f"Stulz_"
        f"{now:%Y-%m-%d}.docx"
    )

    output_path = context.output_dir / filename

    render_docx(
        template_path=context.template_path,
        output_path=output_path,
        replacements=replacements,
        items=calc.items,
    )

    return output_path


def build_preview(calc) -> list[str]:
    models = []

    for item in calc.items:
        model = item.get("item_name", "").strip()
        if model and model not in models:
            models.append(model)

    model_text = ", ".join(models)

    return [
        f"Заказчик: {calc.client_name}",
        f"Модели: {model_text}",
        f"Количество позиций: {len(calc.items)}",
        f"Сумма: {calc.grand_total} {calc.currency_name}",
    ]