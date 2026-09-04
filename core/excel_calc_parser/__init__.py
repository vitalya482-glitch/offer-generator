from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from typing import Any


def _load_legacy_module():
    legacy_path = Path(__file__).resolve().parents[1] / "excel_calc_parser.py"
    spec = importlib.util.spec_from_file_location("_sam_legacy_excel_calc_parser", legacy_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Не удалось загрузить базовый Excel parser: {legacy_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_legacy = _load_legacy_module()

for _name, _value in vars(_legacy).items():
    if _name not in globals():
        globals()[_name] = _value


def _recover_qty(qty: Any, unit_price: Any, total_price: Any) -> Any:
    if qty not in (None, "", 0):
        return qty
    if unit_price in (None, "", 0) or total_price in (None, "", 0):
        return qty
    try:
        recovered = float(total_price) / float(unit_price)
    except (TypeError, ValueError, ZeroDivisionError):
        return qty
    if abs(recovered - round(recovered)) < 1e-6:
        return int(round(recovered))
    return recovered


def _extract_items(
    sheet,
    *,
    model_row: int,
    qty_row: int,
    unit_price_row: int | None,
    total_row: int | None,
) -> list[Any]:
    items: list[Any] = []
    for col in range(2, sheet.max_column + 1):
        name = _legacy._clean_display_text(sheet.value(model_row, col))
        if not name or _legacy._is_helper_header(name):
            continue

        qty = _legacy._to_number(sheet.value(qty_row, col))
        unit_price = _legacy._to_number(sheet.value(unit_price_row, col)) if unit_price_row else None
        total_price = _legacy._to_number(sheet.value(total_row, col)) if total_row else None
        qty = _recover_qty(qty, unit_price, total_price)

        if qty in (None, "", 0) and total_price in (None, "", 0):
            continue
        if total_price in (None, 0) and unit_price not in (None, 0) and qty not in (None, "", 0):
            total_price = float(unit_price) * float(qty)
        if unit_price in (None, 0) and total_price not in (None, 0) and qty not in (None, "", 0):
            unit_price = float(total_price) / float(qty)

        items.append(
            _legacy.CalcItem(
                key=_legacy._make_key(name, col),
                name=name,
                qty=qty,
                unit_price=unit_price,
                total_price=total_price,
                source_col=col,
            )
        )
    return items


def _detect_currency(sheet, exchange_rate: float | None, grand_total_row: int | None = None) -> str | None:
    scores: dict[str, int] = {currency: 0 for currency in _legacy._CURRENCY_LABELS}

    for currency, aliases in _legacy._CURRENCY_LABELS.items():
        for row, col in sheet.find(aliases):
            weight = 8 if row == grand_total_row else 2
            if _nearest_number_distance(sheet, row, col) is not None:
                weight += 4
            scores[currency] += weight

    for row, row_weight in _candidate_money_rows(sheet, grand_total_row):
        for col in range(1, sheet.max_column + 1):
            if _legacy._to_number(sheet.value(row, col)) is None:
                continue
            detected = _currency_from_number_format(sheet.number_format(row, col))
            if detected:
                scores[detected] += row_weight

    best_currency, best_score = max(scores.items(), key=lambda item: item[1])
    if best_score > 0:
        return best_currency
    if exchange_rate and exchange_rate > 1.01:
        return "KZT"
    return None


def _candidate_money_rows(sheet, grand_total_row: int | None) -> list[tuple[int, int]]:
    weighted: list[tuple[int, int]] = []
    seen: set[int] = set()

    def add(row: int | None, weight: int) -> None:
        if row and 1 <= row <= sheet.max_row and row not in seen:
            weighted.append((row, weight))
            seen.add(row)

    add(grand_total_row, 40)
    for aliases, weight in (
        (_legacy._GRAND_TOTAL_ALIASES, 35),
        (_legacy._TOTAL_PER_QTY_ALIASES, 30),
        (_legacy._UNIT_PRICE_ALIASES, 25),
    ):
        for row, _col in sheet.find(aliases):
            add(row, weight)
    for row in range(1, sheet.max_row + 1):
        add(row, 1)
    return weighted


def _currency_from_number_format(number_format: str) -> str | None:
    compact = str(number_format or "").casefold().replace(" ", "")
    markers = {
        "KZT": ("₸", "kzt", "тенге", "тг", "[$₸", "[$kzt", "-kk-kz"),
        "EUR": ("€", "eur", "euro", "евро", "[$€", "[$eur", "-euro"),
        "USD": ("$", "usd", "доллар", "[$$", "[$usd", "-en-us"),
        "RUB": ("₽", "rub", "руб", "[$₽", "[$rub", "-ru-ru"),
    }
    for currency, aliases in markers.items():
        if any(alias in compact for alias in aliases):
            return currency
    return None


def _nearest_number_distance(sheet, row: int, col: int) -> int | None:
    best: int | None = None
    for r in range(max(1, row - 1), min(sheet.max_row, row + 1) + 1):
        for c in range(max(1, col - 4), min(sheet.max_column, col + 4) + 1):
            if _legacy._to_number(sheet.value(r, c)) is None:
                continue
            distance = abs(r - row) + abs(c - col)
            if best is None or distance < best:
                best = distance
    return best


_legacy._extract_items = _extract_items
_legacy._detect_currency = _detect_currency
_legacy._candidate_money_rows = _candidate_money_rows
_legacy._currency_from_number_format = _currency_from_number_format
_legacy._nearest_number_distance = _nearest_number_distance

for _name, _value in vars(_legacy).items():
    if not (_name.startswith("__") and _name.endswith("__")):
        globals()[_name] = _value

parse_calculation = _legacy.parse_calculation
read_sheet_names = _legacy.read_sheet_names
read_hvac_positions = _legacy.read_hvac_positions
HVACPosition = _legacy.HVACPosition
