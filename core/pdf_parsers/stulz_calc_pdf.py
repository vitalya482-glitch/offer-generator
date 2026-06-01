from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

from core.stulz_reference import load_stulz_options, load_missing_options, save_missing_options




@dataclass
class StulzCalcTotals:
    model: str = ""
    quantity: float | None = None
    total_list_price: float | None = None
    total_purchase_price: float | None = None
    unit_list_price: float | None = None
    unit_purchase_price: float | None = None
    currency: str = ""


@dataclass
class StulzOptionRow:
    code: str
    source_name: str
    description: str
    qty: str
    translated: bool = True


CODE_RE = r"(?:\d{6,8}|[A-ZА-Я]{1,3}\d{4,8})"
ROW_START_RE = re.compile(rf"^\s*(?P<qty>\d+(?:[\.,]\d+)?)\s+(?P<code>{CODE_RE})\s+(?P<name>.+?)\s*$", re.IGNORECASE)
WEIRD_ROW_RE = re.compile(
    rf"(?ms)^\s*(?P<qty>\d+(?:[\.,]\d+)?)\s+[^\n]*?\b(?:EUR|USD|KZT)\s*(?P<code>{CODE_RE})\s+(?P<name>.*?)(?=\n\s*\d+(?:[\.,]\d+)?\s+[^\n]*?\b(?:EUR|USD|KZT)\s*{CODE_RE}\s+|\nTotal per device:|\nQuantity:|\Z)",
    re.IGNORECASE,
)


def extract_pdf_text(path: str | Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def _clean_text(value: str) -> str:
    value = (value or "").replace("\r", "\n")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def _normalize(value: str) -> str:
    value = _clean_text(value).lower().replace("ё", "е")
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _strip_price_tail(value: str) -> str:
    value = _clean_text(value)
    value = re.sub(r"\s+\d[\d\s]*[\.,]\d{2}\s*(?:EUR|USD|KZT)?\s*$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+(?:EUR|USD|KZT)\s+\d[\d\s]*[\.,]\d{2}\s*(?:EUR|USD|KZT)?\s*$", "", value, flags=re.IGNORECASE)
    return _clean_text(value)


def _build_option_lookup() -> tuple[dict[str, dict[str, str]], list[dict[str, str]]]:
    rows = load_stulz_options()
    by_code = {str(row.get("code", "")).strip(): row for row in rows if str(row.get("code", "")).strip()}
    return by_code, rows


def _auto_translate_condenser_option(source_name: str) -> str | None:
    text = source_name or ""

    if "condenser" not in text.lower():
        return None

    match = re.search(r"\b(KSV[0-9A-Z]+p?)\b", text, re.IGNORECASE)
    if not match:
        return None

    condenser_model = match.group(1)

    return (
        f"Конденсатор воздушного охлаждения {condenser_model} (для одного контура хладагента).\n"
        "Предназначается для наружного монтажа. Может монтироваться как горизонтально, "
        "так и вертикально к плоскости монтажа. Каркас устойчив к коррозии, сделан из алюминия."
    )


def _lookup_description(
    code: str,
    source_name: str,
    by_code: dict[str, dict[str, str]],
    rows: list[dict[str, str]],
) -> tuple[str, bool]:
    row = by_code.get(code)
    if row:
        description = (row.get("ru_description") or "").strip()
        if description:
            return description, True

    source_norm = _normalize(source_name)
    for candidate in rows:
        candidate_name = candidate.get("source_name", "")
        candidate_norm = _normalize(candidate_name)
        if candidate_norm and (candidate_norm in source_norm or source_norm in candidate_norm):
            description = (candidate.get("ru_description") or "").strip()
            if description:
                return description, True

    condenser_description = _auto_translate_condenser_option(source_name)
    if condenser_description:
        return condenser_description, True

    return source_name, False


def _remember_missing_option(code: str, source_name: str) -> None:
    missing = load_missing_options()
    key = (code.strip(), _normalize(source_name))
    existing = {(row.get("code", "").strip(), _normalize(row.get("source_name", ""))) for row in missing}
    if key in existing:
        return
    missing.append({"code": code.strip(), "source_name": source_name.strip(), "ru_description": ""})
    save_missing_options(missing)


def _parse_weird_price_before_code_rows(text: str) -> Iterable[tuple[str, str, str]]:
    for match in WEIRD_ROW_RE.finditer(text):
        qty = match.group("qty").replace(",", ".").strip()
        code = match.group("code").strip()
        name = _strip_price_tail(match.group("name"))
        name = re.sub(r"\b(?:EUR|USD|KZT)\b.*$", "", name, flags=re.IGNORECASE).strip()
        if code and name and "total per device" not in name.lower():
            yield qty, code, name


def _parse_normal_rows(text: str) -> Iterable[tuple[str, str, str]]:
    rows: list[tuple[str, str, list[str]]] = []
    current: tuple[str, str, list[str]] | None = None

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        low = line.lower()
        if low.startswith("total per device") or low.startswith("quantity:") or low.startswith("basic unit"):
            if current:
                rows.append(current)
                current = None
            break

        match = ROW_START_RE.match(line)
        if match:
            if current:
                rows.append(current)
            current = (
                match.group("qty").replace(",", ".").strip(),
                match.group("code").strip(),
                [match.group("name").strip()],
            )
            continue

        if current and not re.match(r"^(?:EUR|USD|KZT)\b", line, re.IGNORECASE):
            current[2].append(line)

    if current:
        rows.append(current)

    for qty, code, parts in rows:
        name = _strip_price_tail(" ".join(parts))
        name = re.sub(r"\b(?:EUR|USD|KZT)\b.*$", "", name, flags=re.IGNORECASE).strip()
        if code and name and "total per device" not in name.lower():
            yield qty, code, name


def _parse_raw_options(text: str) -> Iterable[tuple[str, str, str]]:
    yielded: set[tuple[str, str, str]] = set()

    for parser in (_parse_normal_rows, _parse_weird_price_before_code_rows):
        for qty, code, name in parser(text):
            key = (qty, code, _normalize(name))
            if key in yielded:
                continue
            yielded.add(key)
            yield qty, code, name


def _format_qty(option_qty: str, equipment_qty: float | int) -> str:
    try:
        oq = float(str(option_qty).replace(",", "."))
        oq_text = str(int(oq)) if oq.is_integer() else str(oq).replace(".", ",")
    except Exception:
        oq_text = str(option_qty)

    try:
        eq = float(equipment_qty)
        eq_text = str(int(eq)) if eq.is_integer() else str(eq).replace(".", ",")
    except Exception:
        eq_text = str(equipment_qty or 1)

    return f"{oq_text} * {eq_text} шт"


def parse_stulz_calc_options(path: str | Path, equipment_qty: float | int = 1) -> list[StulzOptionRow]:
    text = extract_pdf_text(path)
    by_code, rows = _build_option_lookup()
    result: list[StulzOptionRow] = []

    seen: set[str] = set()

    for raw_qty, code, source_name in _parse_raw_options(text):
        code = code.strip()

        if code in seen:
            continue

        seen.add(code)

        description, translated = _lookup_description(code, source_name, by_code, rows)

        if not translated:
            _remember_missing_option(code, source_name)

        result.append(
            StulzOptionRow(
                code=code,
                source_name=source_name,
                description=description,
                qty=_format_qty(raw_qty, equipment_qty),
                translated=translated,
            )
        )

    return result

MONEY_RE = re.compile(r"(?P<amount>\d[\d\s\u00a0]*[\.,]\d{2})\s*(?P<currency>EUR|USD|KZT)", re.IGNORECASE)


def _parse_money_value(value: str) -> float | None:
    try:
        cleaned = (value or "").replace("\u00a0", " ").replace(" ", "").replace(",", ".")
        return float(cleaned)
    except Exception:
        return None


def _fmt_model_from_code_row(text: str) -> str:
    # Main equipment row can be extracted in different orders by PDF readers:
    # 1401862 ASD 211 A 26 860,00 EUR ...
    # ASD 211 A1401862 9 401,00 EUR ...
    code_re = r"\b\d{6,8}\b"
    money_re = r"\d[\d\s\u00a0]*[\.,]\d{2}\s*(?:EUR|USD|KZT)"
    patterns = (
        re.compile(rf"(?m)^\s*{code_re}\s+(?P<model>.+?)\s+{money_re}", re.IGNORECASE),
        re.compile(r"(?m)^\s*(?P<model>[A-Z]{2,5}\s*\d{2,4}\s*[A-Z]{0,3})\s*\d{6,8}\s+", re.IGNORECASE),
    )
    for pattern in patterns:
        for match in pattern.finditer(text):
            model = _clean_text(match.group("model"))
            if not model:
                continue
            low = model.lower()
            if any(skip in low for skip in ("controller", "condenser", "display", "cable", "power supply", "shutdown", "unit-base")):
                continue
            return model
    return ""


def _line_amounts_for_label(text: str, label: str) -> list[tuple[float, str]]:
    for line in text.splitlines():
        if label.lower() not in line.lower():
            continue
        result: list[tuple[float, str]] = []
        for match in MONEY_RE.finditer(line):
            value = _parse_money_value(match.group("amount"))
            if value is not None:
                result.append((value, match.group("currency").upper()))
        return result
    return []


def _final_star_amount(text: str) -> tuple[float, str] | None:
    matches = list(MONEY_RE.finditer(text))
    for match in reversed(matches):
        tail = text[match.end(): match.end() + 5]
        if "*" in tail:
            value = _parse_money_value(match.group("amount"))
            if value is not None:
                return value, match.group("currency").upper()
    return None



def parse_stulz_calc_totals(path: str | Path) -> StulzCalcTotals:
    """Parse final commercial totals from STULZ Calc.pdf.

    The important values for offer calculation are taken from the
    "Total per quantity" line, because it already includes the selected
    options, condensers and discounts.
    """
    text = extract_pdf_text(path)
    totals = StulzCalcTotals(model=_fmt_model_from_code_row(text))

    quantity_match = re.search(r"Quantity:\s*(\d+(?:[\.,]\d+)?)", text, re.IGNORECASE)
    if not quantity_match:
        # Some Calc PDFs are extracted as "79 960,002Quantity".
        quantity_match = re.search(r"(?:\d[\d\s\u00a0]*[\.,]\d{2})(\d+)Quantity", text, re.IGNORECASE)
    if quantity_match:
        try:
            totals.quantity = float(quantity_match.group(1).replace(",", "."))
            if float(totals.quantity).is_integer():
                totals.quantity = int(totals.quantity)
        except Exception:
            totals.quantity = None

    per_device_amounts = _line_amounts_for_label(text, "Total per device:")
    if per_device_amounts:
        totals.unit_list_price = per_device_amounts[0][0]
        totals.unit_purchase_price = per_device_amounts[-1][0]
        if not totals.currency:
            totals.currency = per_device_amounts[-1][1]

    per_quantity_amounts = _line_amounts_for_label(text, "Total per quantity:")
    if len(per_quantity_amounts) >= 2:
        totals.total_list_price = per_quantity_amounts[0][0]
        totals.total_purchase_price = per_quantity_amounts[-1][0]
        totals.currency = per_quantity_amounts[-1][1]
    elif len(per_quantity_amounts) == 1:
        totals.total_purchase_price = per_quantity_amounts[0][0]
        totals.currency = per_quantity_amounts[0][1]

    final_amount = _final_star_amount(text)
    if final_amount and totals.total_purchase_price is None:
        totals.total_purchase_price, totals.currency = final_amount

    if totals.quantity and totals.unit_list_price is not None and totals.total_list_price is None:
        totals.total_list_price = totals.unit_list_price * float(totals.quantity)
    if totals.quantity and totals.unit_purchase_price is not None and totals.total_purchase_price is None:
        totals.total_purchase_price = totals.unit_purchase_price * float(totals.quantity)
    if totals.quantity and totals.total_list_price is not None and totals.unit_list_price is None:
        totals.unit_list_price = totals.total_list_price / float(totals.quantity)
    if totals.quantity and totals.total_purchase_price is not None and totals.unit_purchase_price is None:
        totals.unit_purchase_price = totals.total_purchase_price / float(totals.quantity)

    return totals
