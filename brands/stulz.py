from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from docx import Document

from core.docx_renderer import append_pdf_description, fill_disabled_spec_blocks, fill_product_table, replace_text
from core.excel_reader import parse_stulz_calc
from core.models import CalcData, OfferContext
from core.utils import MONTHS_RU, currency_label, currency_suffix, money, number_to_words_ru, sanitize_filename

BRAND_NAME = "Stulz"


def build_intro_text(calc: CalcData) -> str:
    if not calc.items:
        return "В ответ на Ваш запрос направляем коммерческое предложение на поставку оборудования STULZ."
    models = ", ".join(item.name for item in calc.items)
    total_qty = int(calc.quantity) if float(calc.quantity).is_integer() else calc.quantity
    if len(calc.items) == 1:
        word = "кондиционера" if calc.quantity == 1 else "кондиционеров"
        return (
            f"В ответ на Ваш запрос направляем коммерческое предложение на поставку "
            f"{total_qty} прецизионных {word} Stulz модели {models}. "
            "Опции, включенные в комплектацию и технические характеристики указаны в спецификации коммерческого предложения."
        )
    return (
        f"В ответ на Ваш запрос направляем коммерческое предложение на поставку "
        f"прецизионных кондиционеров Stulz в количестве {total_qty} шт.: {models}. "
        "Опции, включенные в комплектацию и технические характеристики указаны в спецификации коммерческого предложения."
    )


def build_total_price_block(calc: CalcData) -> str:
    amount = money(calc.total_price)
    words = number_to_words_ru(calc.total_price)
    currency = currency_suffix(calc.currency)
    vat_text = f"с учетом НДС {money(calc.vat_percent, 0)}%" if calc.vat_percent > 0 else "без учета НДС"
    return f"Итого, стоимость оборудования составляет: {words} ({amount}) {currency}, {vat_text}."


def build_currency_terms(calc: CalcData) -> str:
    if calc.currency == "KZT" and calc.exchange_rate and calc.exchange_rate > 1.01:
        return (
            f"Расчет был сделан в тенге согласно курса 1 EUR = {money(calc.exchange_rate, 0)} тенге, "
            "при изменении данного курса более чем на 3 тенге будет сделан перерасчет."
        )
    return "Взаиморасчет осуществляется в тенге по курсу БанкаЦентрКредит РК на день оплаты."


def build_options_title(calc: CalcData) -> str:
    if not calc.items:
        return "Опции, включенные в комплектацию оборудования:"
    if len(calc.items) == 1:
        return f"Опции, включенные в комплектацию кондиционера {calc.model}:"
    return "Опции, включенные в комплектацию оборудования:"


def normalize_version(version: str) -> str:
    text = str(version or "1").strip()
    if text.lower().startswith("version"):
        m = re.search(r"(\d+)", text)
        return m.group(1) if m else "1"
    text = text.replace("Версия", "").replace("№", "").strip()
    return text or "1"


def build_replacements(calc: CalcData, client_name: str, date_ru: str, version: str | None) -> dict[str, str]:
    # Later these values can move to config/managers.json and config/signers.json.
    signer_name = "Сания Санаткызы"
    signer_position = "Коммерческий директор"
    manager_name = "Асель Абжанова"
    manager_position = "Ведущий менеджер по продажам"
    manager_email = "assel@sam.kz"
    manager_phone = ""

    version_value = normalize_version(version or calc.version)
    replacements = {
        "{{offer_date}}": date_ru,
        "{{offer_version}}": version_value,
        "{{client_company_full}}": client_name,
        "{{intro_text}}": build_intro_text(calc),
        "{{unit_price_header}}": f"Цена за единицу, {currency_label(calc.currency)}",
        "{{total_price_header}}": f"Сумма, {currency_label(calc.currency)}",
        "{{total_label}}": "ИТОГО",
        "{{grand_total}}": money(calc.total_price),
        "{{total_price_words}}": number_to_words_ru(calc.total_price),
        "{{total_price_eur}}": money(calc.total_price),
        "{{equipment_table}}": "",
        "{{total_price_block}}": build_total_price_block(calc),
        "{{payment_terms}}": "70% предоплата, 30% после поставки оборудования.",
        "{{delivery_time}}": "в течение 13-14 недель после поступления предоплаты.",
        "{{delivery_terms}}": calc.delivery_basis,
        "{{installation_terms}}": "включены / не включены",
        "{{startup_terms}}": "включены / не включены",
        "{{offer_validity}}": "30 календарных дней",
        "{{currency_terms}}": build_currency_terms(calc),
        "{{signer_name}}": signer_name,
        "{{signer_position}}": signer_position,
        "{{director_name}}": signer_name,
        "{{director_position}}": signer_position,
        "{{manager_name}}": manager_name,
        "{{manager_position}}": manager_position,
        "{{manager_email}}": manager_email,
        "{{manager_phone}}": manager_phone,
        "{{options_title}}": build_options_title(calc),
        "{{options_table}}": "",
        "{{technical_specs_table}}": "",
        # Backward compatibility with older SAM templates.
        "ТОО «[Организация]»": client_name,
        "[Организация]": client_name,
        "г. Алматы, 15 апреля 2021 г.": f"г. Алматы, {date_ru}",
        "Версия №1": f"Версия №{version_value}",
        "DDP г. Алматы": calc.delivery_basis,
        "Итого, стоимость оборудования составляет:  ( ) EUR 00 eurocent, с учетом НДС 12%.": build_total_price_block(calc),
        "кондиционеров CCU121A": f"кондиционеров {calc.model}",
        "Взаиморасчет осуществляется в тенге по курсу БанкаЦентрКредит РК на день оплаты.": build_currency_terms(calc),
        "Расчет был сделан в тенге согласно курса 1 EUR = 530 тенге, при изменении данного курса более чем на 3 тенге будет сделан перерасчет.": "",
        "Коммерческий директор \t\t": signer_position,
        "Ведущий Менеджер по продажам": manager_position,
        "assel@sam.kz ": manager_email,
        "Тип модуля:\n\tASU 211 AL": f"Тип модуля:\n\t{calc.model}",
    }
    return replacements


def make_offer(context: OfferContext) -> Path:
    calc = parse_stulz_calc(context.calc_path, context.sheet_name)
    document = Document(context.template_path)
    now = datetime.now()
    date_ru = f"{now.day} {MONTHS_RU[now.month]} {now.year} г."

    replacements = build_replacements(calc, context.client_name, date_ru, context.version)
    replace_text(document, replacements)
    fill_product_table(document, calc)
    fill_disabled_spec_blocks(document)
    append_pdf_description(document, context.pdf_dir, calc.model)

    context.output_dir.mkdir(parents=True, exist_ok=True)
    main_model = sanitize_filename(calc.model)
    filename = f"КП_{sanitize_filename(context.client_name)}_{main_model}_{now:%Y-%m-%d}.docx"
    out = context.output_dir / filename
    document.save(out)
    return out


def preview(context: OfferContext) -> str:
    calc = parse_stulz_calc(context.calc_path, context.sheet_name)
    qty = int(calc.quantity) if calc.quantity.is_integer() else calc.quantity
    return "\n".join([
        f"Направление: {BRAND_NAME}",
        f"Лист: {calc.sheet_name}",
        f"Версия: {calc.version}",
        f"Модель: {calc.model}",
        f"Количество: {qty}",
        f"Условия поставки: {calc.delivery_basis}",
        f"Итого: {money(calc.total_price)} {calc.currency}",
        f"Опций найдено: {len(calc.options)}",
    ])
