from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

from core.stulz_reference import load_stulz_options, load_missing_options, save_missing_options


@dataclass
class StulzOptionRow:
    code: str
    source_name: str
    description: str
    qty: str
    translated: bool = True


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


def _build_option_lookup() -> tuple[dict[str, dict[str, str]], list[dict[str, str]]]:
    rows = load_stulz_options()
    by_code = {str(row.get("code", "")).strip(): row for row in rows if str(row.get("code", "")).strip()}
    return by_code, rows


def _lookup_description(code: str, source_name: str, by_code: dict[str, dict[str, str]], rows: list[dict[str, str]]) -> tuple[str, bool]:
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

    return source_name, False


def _remember_missing_option(code: str, source_name: str) -> None:
    missing = load_missing_options()
    key = (code.strip(), _normalize(source_name))
    existing = {(row.get("code", "").strip(), _normalize(row.get("source_name", ""))) for row in missing}
    if key in existing:
        return
    missing.append({"code": code.strip(), "source_name": source_name.strip(), "ru_description": ""})
    save_missing_options(missing)


def _parse_raw_options(text: str) -> Iterable[tuple[str, str, str]]:
    # Stulz Calc PDF usually has rows like:
    #  1  339,50 970,00 EUR1402918 A - Three phase supervision ... EUR 339,50 EUR
    pattern = re.compile(
        r"(?ms)^\s*(?P<qty>\d+(?:[\.,]\d+)?)\s+.*?EUR\s*(?P<code>\d{7,8})\s+(?P<name>.*?)(?=\s*EUR|\n\s*\d+(?:[\.,]\d+)?\s+.*?EUR\s*\d{7,8}|\nTotal per device:|\nQuantity:)",
    )
    for match in pattern.finditer(text):
        qty = match.group("qty").replace(",", ".").strip()
        code = match.group("code").strip()
        name = _clean_text(match.group("name"))
        if not code or not name:
            continue
        # Basic unit is not an option and normally has no leading qty, but keep an extra guard.
        if "total per device" in name.lower():
            continue
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

    seen: set[tuple[str, str]] = set()
    for raw_qty, code, source_name in _parse_raw_options(text):
        key = (code, source_name)
        if key in seen:
            continue
        seen.add(key)
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
