from __future__ import annotations
from num2words import num2words
from datetime import datetime
from pathlib import Path
from typing import Any
import re

from core.docx_renderer import render_docx
from core.excel_reader import parse_stulz_calc
from core.models import CalcData, OfferContext, OfferItem
from core.stulz_specification import build_stulz_specification
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

def extract_revision_number(value: str) -> int | None:
    """Extract revision number from strings like V1, v 2, (V3), rev4."""
    text = value or ""
    match = re.search(r"(?:^|[\s_\-\(])(?:v|rev)\s*(\d+)(?:\)|\b|$)", text, re.IGNORECASE)
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def find_next_offer_version(output_dir: Path, client_name: str = "", sheet_name: str = "") -> int:
    """Return next КП revision based on existing DOCX files in the result folder.

    If no previous offer files with revN are found, use Excel sheet revision Vn when present,
    otherwise start from revision 1.
    """
    output_dir = Path(output_dir)
    max_version = 0

    if output_dir.exists():
        for file_path in output_dir.glob("*.docx"):
            name = file_path.name
            if name.startswith("~$"):
                continue

            match = re.search(r"(?:^|[\s_\-])rev\s*(\d+)(?:\.docx|[\s_\-]|$)", name, re.IGNORECASE)
            if match:
                max_version = max(max_version, int(match.group(1)))

    if max_version > 0:
        return max_version + 1

    return extract_revision_number(sheet_name) or 1


def build_offer_filename(client_name: str, offer_version: int, dt=None) -> str:
    dt = dt or datetime.now()
    client = sanitize_filename(client_name).replace(" ", "_")
    return f"offer_{client}_{dt:%d-%m-%y}_rev{offer_version}.docx"


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


def _model_airflow_group(model: str) -> str:
    text = _norm_model_text(model)

    # Common STULZ model-code hints. For Mini-Space/CyberAir old model codes:
    # CCU/ASU generally indicate upflow, CCD/ASD generally indicate downflow.
    if re.search(r"\b(ccd|asd)[a-z0-9]*\b", text):
        return "downflow"
    if re.search(r"\b(ccu|asu)[a-z0-9]*\b", text):
        return "upflow"

    return ""


def detect_offer_airflow_text(calc: CalcData, profile: dict[str, str] | None = None) -> str:
    groups = {_model_airflow_group(item.name) for item in calc.items if item.name}
    groups.discard("")

    # If the offer contains different unit types, for example ASD + ASU,
    # do not write a common air-supply direction in the intro text.
    if len(groups) > 1:
        return ""

    if groups == {"downflow"}:
        return "с нижней подачей охлажденного воздуха под фальшпол"
    if groups == {"upflow"}:
        return "с верхней подачей охлажденного воздуха"

    return detect_airflow_text(calc, profile)


def build_intro_text(calc: CalcData) -> str:
    profile = detect_stulz_profile(calc)
    equipment_type = _equipment_type(calc, profile)
    line = profile.get("line") or DEFAULT_STULZ_SERIES["line"]
    install_type = profile.get("install_type") or DEFAULT_STULZ_SERIES["install_type"]
    cooling_type = detect_cooling_type_text(calc)
    airflow = detect_offer_airflow_text(calc, profile)
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
    vat = 0.0
    try:
        vat = float(calc.vat_percent or 0)
    except Exception:
        pass

    vat_text = (
        "без учета НДС"
        if abs(vat) < 0.001
        else f"с учетом НДС {format_qty(vat)}%"
    )

    return (
        f"{format_money(calc.total_price)} {currency_name(calc.currency)} "
        f"({money_in_words(calc.total_price, calc.currency)}), "
        f"{vat_text}."
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


def build_installation_terms(calc: CalcData) -> tuple[str, str]:
    if getattr(calc, "installation_included", False):
        return "Монтажные работы включены", "Пусконаладочные работы включены"
    return "Монтажные работы не включены", "Пусконаладочные работы не включены"


def build_replacements(context: OfferContext, calc: CalcData, offer_version: int | None = None) -> dict[str, Any]:
    cur_name = currency_name(calc.currency)
    installation_terms, startup_terms = build_installation_terms(calc)
    return {
        "{{offer_date}}": format_offer_date(),
        "{{offer_version}}": str(offer_version or context.version or calc.version or "1"),
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
        "{{installation_terms}}": installation_terms,
        "{{startup_terms}}": startup_terms,
        "{{offer_validity}}": "Коммерческое предложение действительно в течение 30 календарных дней.",
        "{{currency_terms}}": build_currency_terms(calc),
        "{{signer_name}}": context.signer_name,
        "{{signer_position}}": context.signer_position,
        "{{manager_name}}": context.manager_name,
        "{{manager_position}}": context.manager_position,
        "{{manager_email}}": context.manager_email,
        "{{manager_phone}}": context.manager_phone,
    }


def load_calc(context: OfferContext) -> CalcData:
    calc = parse_stulz_calc(context.calc_path, context.sheet_name)
    if not calc.items:
        raise ValueError(
            "В выбранном листе Excel не найдены позиции оборудования. "
            "Проверьте лист Excel или структуру расчета."
        )
    return calc



def _spec_model_key(value: str) -> str:
    """Normalize model names so ASD211A and ASD 211 A are treated as one model."""
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def _selected_spec_models(context: OfferContext, calc: CalcData) -> list[dict[str, Any]]:
    """Return models selected in the GUI specification table.

    The specification preview and specification blocks must use only the models
    shown in the Specifications block. Excel calculation rows are not used here,
    otherwise stale/extra Excel models can appear in preview and generated KP.
    """
    selected: list[dict[str, Any]] = []
    seen: set[str] = set()

    for row in getattr(context, "spec_models", []) or []:
        if not row.get("enabled", True):
            continue

        model = str(row.get("model", "")).strip()
        if not model:
            continue

        key = _spec_model_key(model)
        if key in seen:
            continue

        qty_value = row.get("qty_value", row.get("qty", row.get("quantity", 1)))
        try:
            qty = float(str(qty_value).replace(",", "."))
        except Exception:
            qty = 1.0
        if qty <= 0:
            qty = 1.0

        selected.append({"model": model, "qty": qty})
        seen.add(key)

    return selected

def _calc_for_spec_model(calc: CalcData, model_name: str, qty: float) -> CalcData:
    source_item = next((item for item in calc.items if item.name == model_name), None)
    if source_item is None:
        source_item = OfferItem(no=1, name=model_name, qty=qty, unit_price=0.0, total_price=0.0)
    else:
        source_item = OfferItem(
            no=source_item.no,
            name=source_item.name,
            qty=qty,
            unit_price=source_item.unit_price,
            total_price=source_item.unit_price * qty,
        )

    return CalcData(
        sheet_name=calc.sheet_name,
        version=calc.version,
        currency=calc.currency,
        vat_percent=calc.vat_percent,
        exchange_rate=calc.exchange_rate,
        delivery_basis=calc.delivery_basis,
        items=[source_item],
        options=list(calc.options),
        installation_included=calc.installation_included,
    )


def build_specification_blocks(context: OfferContext, calc: CalcData) -> tuple[list[dict[str, Any]], list[str]]:
    blocks: list[dict[str, Any]] = []
    warnings: list[str] = []
    for selected in _selected_spec_models(context, calc):
        model = selected["model"]
        qty = selected["qty"]
        model_calc = _calc_for_spec_model(calc, model, qty)
        specification = build_stulz_specification(context.pdf_dir, model_calc)
        warnings.extend(f"{model}: {warning}" for warning in specification.warnings)
        totals = specification.totals
        blocks.append({
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
                {"description": option.description, "qty": option.qty, "code": option.code, "source_name": option.source_name, "translated": option.translated}
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
        })
    return blocks, warnings


DESCRIPTION_OPTION_DEFAULTS: dict[str, bool] = {
    "stulz_unit": True,
    "cooling_capacity": True,
    "unit_dimensions": True,
    "condenser": True,
}


def _description_options(context: OfferContext) -> dict[str, bool]:
    options = dict(DESCRIPTION_OPTION_DEFAULTS)
    raw = getattr(context, "description_options", None) or {}
    if isinstance(raw, dict):
        options.update({key: bool(value) for key, value in raw.items() if key in options})
    return options


def _text_key(value: str) -> str:
    value = (value or "").lower().replace("ё", "е")
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _row_section_name(value: str) -> str:
    text = _text_key(value)
    return text or "unit"


def _spec_block_keys(block: dict[str, Any]) -> set[str]:
    keys: set[str] = set()
    for candidate in (block.get("model"), block.get("calc_model")):
        key = _spec_model_key(str(candidate or ""))
        if key:
            keys.add(key)
    return keys


def _find_spec_block_for_item(
    item: OfferItem,
    spec_blocks: list[dict[str, Any]],
    item_index: int,
    used_block_indexes: set[int],
) -> tuple[int | None, dict[str, Any] | None]:
    """Find specification data for a commercial offer row.

    First try strict normalized model matching, for example:
    ASD211A == ASD 211 A == ASD-211-A.

    If Excel contains a typo in the model name, but the user has already selected
    the correct specifications in the same order as the offer rows, use a safe
    positional fallback. This keeps prices and quantities from Excel, but builds
    the extended description from the matched specification block.
    """
    item_key = _spec_model_key(item.name)

    if item_key:
        for index, block in enumerate(spec_blocks):
            if index in used_block_indexes:
                continue
            if item_key in _spec_block_keys(block):
                return index, block

    # Fallback for Excel/model-name typos: use the specification row at the same
    # position if it has not already been consumed by another offer row.
    if 0 <= item_index < len(spec_blocks) and item_index not in used_block_indexes:
        return item_index, spec_blocks[item_index]

    return None, None


def _tech_value(
    block: dict[str, Any],
    name_terms: tuple[str, ...],
    section_terms: tuple[str, ...] | None = None,
) -> str:
    current_section = "unit"
    wanted_names = tuple(_text_key(term) for term in name_terms)
    wanted_sections = tuple(_text_key(term) for term in (section_terms or ()))

    for row in block.get("technical_specs") or []:
        name = _text_key(str(row.get("name") or ""))
        value = str(row.get("value") or "").strip()
        if row.get("is_section"):
            current_section = _row_section_name(str(row.get("name") or ""))
            continue
        if not value:
            continue
        if wanted_sections:
            section_ok = any(term and term in current_section for term in wanted_sections)
            if not section_ok:
                continue
        if all(term and term in name for term in wanted_names):
            return value
    return ""


def _unit_dimensions_text(block: dict[str, Any]) -> str:
    direct = _tech_value(block, ("габарит",), ("unit",))
    if direct:
        return direct

    height = _tech_value(block, ("высота",), ("unit",))
    width = _tech_value(block, ("ширина",), ("unit",))
    depth = _tech_value(block, ("глубина",), ("unit",))
    values = [value for value in (height, width, depth) if value]
    if len(values) == 3:
        cleaned = [re.sub(r"\s*мм\b", "", value, flags=re.IGNORECASE).strip() for value in values]
        return "x".join(cleaned) + " мм"
    return ""


def _cooling_capacity_text(block: dict[str, Any]) -> str:
    return (
        _tech_value(block, ("холодопроизводительность", "общ"), ("unit",))
        or _tech_value(block, ("cooling capacity", "total"), ("unit",))
        or _tech_value(block, ("холодопроизводительность",), ("unit",))
    )


def _condenser_text(block: dict[str, Any]) -> str:
    return (
        _tech_value(block, ("тип модуля",), ("конденсор",))
        or _tech_value(block, ("unit type",), ("condenser",))
    )


def _build_offer_item_description(item: OfferItem, block: dict[str, Any] | None, options: dict[str, bool]) -> str:
    if not block:
        return item.name

    model = str(block.get("calc_model") or block.get("model") or item.name).strip() or item.name
    parts: list[str] = []

    if options.get("stulz_unit", True):
        parts.append(f"Прецизионный кондиционер Stulz {model}")
    else:
        parts.append(model)

    if options.get("cooling_capacity", True):
        cooling = _cooling_capacity_text(block)
        if cooling:
            parts.append(f"хладопроизводительность {cooling}")

    if options.get("unit_dimensions", True):
        dimensions = _unit_dimensions_text(block)
        if dimensions:
            parts.append(f"размеры внутреннего блока {dimensions}")

    if options.get("condenser", True):
        condenser = _condenser_text(block)
        if condenser:
            parts.append(f"конденсор {condenser}")

    return ", ".join(parts)


def build_offer_items(context: OfferContext, calc: CalcData, spec_blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    options = _description_options(context)
    result: list[dict[str, Any]] = []
    used_block_indexes: set[int] = set()

    for item_index, item in enumerate(calc.items):
        block_index, block = _find_spec_block_for_item(item, spec_blocks, item_index, used_block_indexes)
        if block_index is not None:
            used_block_indexes.add(block_index)

        item_name = _build_offer_item_description(item, block, options)
        result.append({
            "item_no": item.no,
            "item_name": item_name,
            "item_qty": format_qty(item.qty),
            "item_unit_price": format_money(item.unit_price),
            "item_total": format_money(item.total_price),
        })

    return result

def make_offer(context: OfferContext) -> Path:
    calc = load_calc(context)
    spec_blocks, _warnings = build_specification_blocks(context, calc)
    offer_version = find_next_offer_version(context.output_dir, context.client_name, calc.sheet_name)
    replacements = build_replacements(context, calc, offer_version=offer_version)
    items = build_offer_items(context, calc, spec_blocks)

    filename = build_offer_filename(context.client_name, offer_version)
    output_path = context.output_dir / filename

    return render_docx(
        template_path=context.template_path,
        output_path=output_path,
        replacements=replacements,
        items=items,
        stulz_spec_blocks=spec_blocks,
    )

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
    spec_blocks, spec_warnings = build_specification_blocks(context, calc)
    models = []
    for item in calc.items:
        if item.name and item.name not in models:
            models.append(item.name)

    lines = [
        f"Заказчик: {context.client_name}",
        f"Лист Excel: {calc.sheet_name}",
        f"Версия расчета: {calc.version}",
        f"Версия КП: {find_next_offer_version(context.output_dir, context.client_name, calc.sheet_name)}",
        f"Валюта: {calc.currency}",
        f"Курс: {format_money(calc.exchange_rate)}",
        f"НДС: {format_qty(calc.vat_percent)}%",
        f"Условия поставки: {calc.delivery_basis}",
        f"Монтаж/ПНР: {'включены' if getattr(calc, 'installation_included', False) else 'не включены'}",
        f"Моделей для спецификации: {len(spec_blocks)}",
        f"Опций для спецификации: {sum(len(block.get('options', [])) for block in spec_blocks)}",
        f"Строк тех. характеристик: {sum(len(block.get('technical_specs', [])) for block in spec_blocks)}",
        f"Модели: {', '.join(models) if models else '-'}",
        f"Количество: {format_qty(calc.quantity)}",
        f"Сумма: {format_money(calc.total_price)} {currency_name(calc.currency)}",
        "",
        "Предупреждения спецификации:",
    ]
    if spec_warnings:
        lines.extend(f"- {warning}" for warning in spec_warnings)
    else:
        lines.append("- нет")

    lines.extend([
        "",
        "Позиции:",
    ])
    for item in calc.items:
        lines.append(
            f"{item.no}. {item.name} | кол-во {format_qty(item.qty)} | "
            f"цена {format_money(item.unit_price)} | сумма {format_money(item.total_price)}"
        )
    return "\n".join(lines)


# Backward-compatible aliases for older code/tests.
generate_offer = make_offer
build_preview = preview
