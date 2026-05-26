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
    # Remove common trailing price fragments copied from PDF tables.
    value = re.sub(r"\s+\d[\d\s]*[\.,]\d{2}\s*(?:EUR|USD|KZT)?\s*$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+(?:EUR|USD|KZT)\s+\d[\d\s]*[\.,]\d{2}\s*(?:EUR|USD|KZT)?\s*$", "", value, flags=re.IGNORECASE)
    return _clean_text(value)


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
            current = (match.group("qty").replace(",", ".").strip(), match.group("code").strip(), [match.group("name").strip()])
            continue

        # continuation line for wrapped option name
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
    # Two PDF text-layer formats are common:
    # 1) qty code name price
    # 2) qty prices EURcode name price
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
