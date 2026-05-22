from __future__ import annotations
from num2words import num2words
from datetime import datetime
from pathlib import Path
from typing import Any
import re

from core.docx_renderer import render_docx
from core.excel_reader import parse_stulz_calc
from core.models import CalcData, OfferContext, OfferItem
from config.stulz_series import AIRFLOW_TEXT, DEFAULT_STULZ_SERIES, STULZ_SERIES

BRAND_NAME = "Stulz"


MONTHS_RU = {
    1: "января",
    2: "февраля",
    3: "марта",
    4: "апреля",
    5: "мая",
    6: "июня",
    7: "июля",
    8: "августа",
    9: "сентября",
    10: "октября",
    11: "ноября",
    12: "декабря",
}


def format_offer_date(dt=None) -> str:
    dt = dt or datetime.now()
    return f"{dt.day} {MONTHS_RU[dt.month]} {dt.year} г."


def sanitize_filename(value: str) -> str:
    bad = '<>:"/\\|?*'
    for ch in bad:
        value = value.replace(ch, "")
    return value.strip() or "Client"

def format_money(value: float) -> str:
    try:
        s = f"{float(value):,.2f}"

        # 35,054.00
        s = s.replace(",", "TEMP")

        # 35TEMP054,00
        s = s.replace(".", ",")

        # 35 054,00
        s = s.replace("TEMP", "\u00A0")

        return s

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


def _series_profile(
    line: str,
    equipment_type_single: str = "прецизионного кондиционера",
    equipment_type_plural: str = "прецизионных кондиционеров",
    install_type: str = "напольного исполнения",
    airflow: str = "",
) -> dict[str, str]:
    return {
        "line": line,
        "equipment_type_single": equipment_type_single,
        "equipment_type_plural": equipment_type_plural,
        "install_type": install_type,
        "airflow": airflow,
    }


STULZ_KEYWORD_PROFILES: tuple[tuple[tuple[str, ...], dict[str, str]], ...] = (
    (("mini space", "minispace", "mini-space"), _series_profile("Stulz Mini Space EC")),
    (("cyberair mini", "cyber air mini"), _series_profile("Stulz CyberAir Mini")),
    (("cyberair 3pro", "cyberair 3 pro", "cyber air 3pro", "cyber air 3 pro", "cyberair 3"), _series_profile("Stulz CyberAir 3PRO")),
    (("cyberair", "cyber air"), _series_profile("Stulz CyberAir")),
    (("cyberrow", "cyber row", "crs", "crl"), _series_profile(
        "Stulz CyberRow",
        equipment_type_single="межрядного прецизионного кондиционера",
        equipment_type_plural="межрядных прецизионных кондиционеров",
        install_type="межрядного исполнения",
    )),
    (("cyberwall", "cyber wall"), _series_profile("Stulz CyberWall", install_type="настенного исполнения")),
    (("cyberlab", "cyber lab"), _series_profile("Stulz CyberLab", install_type="лабораторного исполнения")),
    (("wallair", "wall air"), _series_profile(
        "Stulz WallAir",
        equipment_type_single="телекоммуникационного кондиционера",
        equipment_type_plural="телекоммуникационных кондиционеров",
        install_type="настенного исполнения",
    )),
    (("splitair", "split air", "split-air"), _series_profile(
        "Stulz Split Air",
        equipment_type_single="телекоммуникационного кондиционера",
        equipment_type_plural="телекоммуникационных кондиционеров",
        install_type="сплит-исполнения",
    )),
    (("telair", "tel air", "tel-air"), _series_profile(
        "Stulz TelAir",
        equipment_type_single="телекоммуникационного кондиционера",
        equipment_type_plural="телекоммуникационных кондиционеров",
        install_type="шкафного исполнения",
    )),
    (("shelterair", "shelter air", "shelter-air"), _series_profile(
        "Stulz ShelterAir FC",
        equipment_type_single="телекоммуникационного кондиционера",
        equipment_type_plural="телекоммуникационных кондиционеров",
        install_type="наружного исполнения",
    )),
    (("cybercool", "cyber cool"), _series_profile(
        "Stulz CyberCool",
        equipment_type_single="чиллера",
        equipment_type_plural="чиллеров",
        install_type="наружного исполнения",
    )),
)


def _norm_model_text(value: str) -> str:
    value = (value or "").lower().replace("ё", "е")
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _calc_search_text(calc: CalcData) -> str:
    parts = [calc.model, calc.sheet_name, calc.delivery_basis]
    parts.extend(item.name for item in calc.items if item.name)
    parts.extend(option[0] for option in calc.options if option and option[0])
    return _norm_model_text(" ".join(parts))


def _extract_model_prefix(value: str) -> str:
    match = re.search(r"[A-Za-zА-Яа-я0-9]+", value or "")
    return match.group(0).upper() if match else ""


def detect_stulz_profile(calc: CalcData) -> dict[str, str]:
    text = _calc_search_text(calc)

    for raw_value in [calc.model, calc.sheet_name, *(item.name for item in calc.items)]:
        prefix = _extract_model_prefix(raw_value)
        if prefix in STULZ_SERIES:
            return STULZ_SERIES[prefix]

    for prefix, profile in STULZ_SERIES.items():
        if re.search(rf"\b{re.escape(prefix.lower())}[a-z0-9]*\b", text):
            return profile

    for keywords, profile in STULZ_KEYWORD_PROFILES:
        if any(_norm_model_text(keyword) in text for keyword in keywords):
            return profile

    return DEFAULT_STULZ_SERIES


def detect_airflow_text(calc: CalcData, profile: dict[str, str] | None = None) -> str:
    text = _calc_search_text(calc)

    if any(x in text for x in ("downflow", "down flow", "underfloor", "under floor", "false floor", "фальшпол", "фальш пол", "нижняя подача", "нижней подачей")):
        return "с нижней подачей охлажденного воздуха под фальшпол"

    if any(x in text for x in ("upflow", "up flow", "верхняя подача", "верхней подачей")):
        return "с верхней подачей охлажденного воздуха"

    if any(x in text for x in ("front", "фронтальная", "фронтальной")):
        if any(x in text for x in ("box", "duct", "короб", "воздуховод")):
            return "с фронтальной подачей охлажденного воздуха через короб"
        return "с фронтальной подачей охлажденного воздуха"

    # Common STULZ model-code hints. For Mini-Space/CyberAir old model codes:
    # CCU/ASU generally indicate upflow, CCD/ASD generally indicate downflow.
    if re.search(r"\b(ccd|asd)[a-z0-9]*\b", text):
        return "с нижней подачей охлажденного воздуха под фальшпол"
    if re.search(r"\b(ccu|asu)[a-z0-9]*\b", text):
        return "с верхней подачей охлажденного воздуха"

    airflow_key = (profile or {}).get("airflow", "")
    if airflow_key in AIRFLOW_TEXT:
        return AIRFLOW_TEXT[airflow_key]

    return "с подачей охлажденного воздуха согласно выбранной конфигурации"


def detect_cooling_type_text(calc: CalcData) -> str:
    text = f" {_calc_search_text(calc)} "

    if re.search(r"\b(cw|acw)\b", text) or any(
        x in text for x in (" chilled water ", " чиллер ", " водяное охлаждение ", " холодная вода ")
    ):
        return "с водяным охлаждением (CW)"

    if re.search(r"\b(dx|direct expansion)\b", text) or any(
        x in text for x in (" фреон ", " фреонов", " компрессор")
    ):
        return "с непосредственным испарением (DX)"

    return ""


def _equipment_type(calc: CalcData, profile: dict[str, str]) -> str:
    qty = calc.quantity
    try:
        key = "equipment_type_single" if float(qty) == 1 else "equipment_type_plural"
    except Exception:
        key = "equipment_type_plural"
    return profile.get(key) or DEFAULT_STULZ_SERIES[key]


def build_intro_text(calc: CalcData) -> str:
    profile = detect_stulz_profile(calc)
    equipment_type = _equipment_type(calc, profile)
    line = profile.get("line") or DEFAULT_STULZ_SERIES["line"]
    install_type = profile.get("install_type") or DEFAULT_STULZ_SERIES["install_type"]
    cooling_type = detect_cooling_type_text(calc)
    airflow = detect_airflow_text(calc, profile)
    quantity = format_qty(calc.quantity)

    details = [part for part in (cooling_type, install_type, airflow, f"в количестве {quantity} шт") if part]

    return (
        "В ответ на Ваш запрос направляем коммерческое предложение "
        f"на поставку {equipment_type} {line}, "
        f"{', '.join(details)}. "
        "Опции, включенные в комплектацию и технические характеристики "
        "указаны в спецификации коммерческого предложения."
    )


def build_total_price_block(calc: CalcData) -> str:
    return (
        f"{format_money(calc.total_price)} {currency_name(calc.currency)} "
        f"({money_in_words(calc.total_price, calc.currency)}), "
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
        "{{offer_date}}": format_offer_date(),
        "{{offer_version}}": context.version or calc.version or "1",
        "{{client_company_full}}": context.client_name,
        "{{intro_text}}": build_intro_text(calc),
        "{{unit_price_header}}": f"Цена за единицу, {cur_name}",
        "{{total_price_header}}": f"Сумма, {cur_name}",
        "{{total_label}}": "ИТОГО",
        "{{grand_total}}": format_money(calc.total_price),
        "{{total_price_block}}": build_total_price_block(calc),
        "{{payment_terms}}": "70% предоплата",
        "{{delivery_time}}": "Срок поставки уточняется после размещения заказа.",
        "{{delivery_terms}}": calc.delivery_basis,
        "{{installation_terms}}": "Монтажные работы не включены",
        "{{startup_terms}}": "Пусконаладочные работы не включены",
        "{{offer_validity}}": "Коммерческое предложение действительно в течение 30 календарных дней.",
        "{{currency_terms}}": build_currency_terms(calc),
        "{{signer_name}}": context.signer_name,
        "{{signer_position}}": context.signer_position,
        "{{manager_name}}": context.manager_name,
        "{{manager_position}}": context.manager_position,
        "{{manager_email}}": context.manager_email,
        "{{manager_phone}}": context.manager_phone,
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

def money_in_words(amount: float, currency: str) -> str:
    whole = int(round(amount))

    cur = currency.upper()

    if cur == "KZT":
        main = "тенге"
        minor = "тиын"
    elif cur == "USD":
        main = "долларов США"
        minor = "центов"
    elif cur == "EUR":
        main = "евро"
        minor = "eurocents"
    else:
        main = currency
        minor = ""

    words = num2words(whole, lang="ru")

    if minor:
        return f"{words} {main} 00 {minor}"

    return f"{words} {main}"

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
