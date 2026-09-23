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
    """Detect the offer currency without letting helper conversion blocks win.

    HVAC calculations may contain two currency blocks at the same time, for
    example a working EUR block and a lower KZT conversion block.  The offer
    currency must follow the decisive price rows first.  Only if those rows do
    not carry any currency marker do we apply the SAM fallback: rate=1 means the
    sheet is already in EUR, while rate>1 normally means KZT.
    """

    # 1. The strongest source is the currency used directly in the decisive
    # money rows: TOTAL, Total per quantity, Total per unit, unit price rows.
    decisive_scores = _money_row_currency_scores(sheet, grand_total_row, include_all_rows=False)
    decisive_currency = _best_currency(decisive_scores)
    if decisive_currency:
        return decisive_currency

    # 2. Fallback from the common SAM calculation logic.  This must be stronger
    # than broad text labels because sheets often contain both "€, euro" and a
    # lower helper "kzt" block.  If the rate is 1, the active values are EUR.
    rate_currency = _currency_from_exchange_rate(exchange_rate)
    if rate_currency:
        return rate_currency

    # 3. Only now look at broad currency labels in the sheet.  Return a value
    # only when the signal is not ambiguous.
    label_currency = _best_currency(_label_currency_scores(sheet, grand_total_row), require_unique=True)
    if label_currency:
        return label_currency

    # 4. Last fallback: any remaining Excel number formats/cell text markers on
    # the sheet.  Keep it conservative so a mixed EUR/KZT sheet does not pick the
    # wrong one by accident.
    broad_money_currency = _best_currency(
        _money_row_currency_scores(sheet, grand_total_row, include_all_rows=True),
        require_unique=True,
    )
    if broad_money_currency:
        return broad_money_currency
    return None


def _currency_from_exchange_rate(exchange_rate: float | None) -> str | None:
    try:
        rate = float(exchange_rate) if exchange_rate is not None else None
    except (TypeError, ValueError):
        rate = None
    if rate is None:
        return None
    if 0.99 <= rate <= 1.01:
        return "EUR"
    if rate > 1.01:
        return "KZT"
    return None


def _money_row_currency_scores(
    sheet,
    grand_total_row: int | None,
    *,
    include_all_rows: bool,
) -> dict[str, int]:
    scores: dict[str, int] = {currency: 0 for currency in _legacy._CURRENCY_LABELS}
    rows = _priority_money_rows(sheet, grand_total_row)

    if include_all_rows:
        seen = {row for row, _weight in rows}
        for row in range(1, sheet.max_row + 1):
            if row not in seen:
                rows.append((row, 1))
                seen.add(row)

    for row, row_weight in rows:
        for col in range(1, sheet.max_column + 1):
            detected = _currency_from_number_format(sheet.number_format(row, col))
            if detected:
                scores[detected] += row_weight
                continue

            # Some calculations store money as text like "24 486,00 €" instead
            # of a numeric value with a currency format.  Treat the currency
            # marker in the cell text as a valid signal for the row.
            detected = _currency_from_cell_text(sheet.value(row, col))
            if detected:
                scores[detected] += row_weight
    return scores


def _label_currency_scores(sheet, grand_total_row: int | None) -> dict[str, int]:
    scores: dict[str, int] = {currency: 0 for currency in _legacy._CURRENCY_LABELS}
    priority_rows = {row for row, _weight in _priority_money_rows(sheet, grand_total_row)}

    for currency, aliases in _legacy._CURRENCY_LABELS.items():
        for row, col in sheet.find(aliases):
            weight = 3 if row in priority_rows else 1
            if row == grand_total_row:
                weight += 10
            if _nearest_number_distance(sheet, row, col) is not None:
                weight += 1
            scores[currency] += weight
    return scores


def _best_currency(scores: dict[str, int], *, require_unique: bool = False) -> str | None:
    ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    if not ranked or ranked[0][1] <= 0:
        return None
    if require_unique and len(ranked) > 1 and ranked[1][1] == ranked[0][1]:
        return None
    return ranked[0][0]


def _priority_money_rows(sheet, grand_total_row: int | None) -> list[tuple[int, int]]:
    weighted: list[tuple[int, int]] = []
    seen: set[int] = set()

    def add(row: int | None, weight: int) -> None:
        if row and 1 <= row <= sheet.max_row and row not in seen:
            weighted.append((row, weight))
            seen.add(row)

    add(grand_total_row, 50)
    for aliases, weight in (
        (_legacy._GRAND_TOTAL_ALIASES, 45),
        (_legacy._TOTAL_PER_QTY_ALIASES, 40),
        (_legacy._UNIT_PRICE_ALIASES, 35),
    ):
        for row, _col in sheet.find(aliases):
            add(row, weight)
    return weighted


def _candidate_money_rows(sheet, grand_total_row: int | None) -> list[tuple[int, int]]:
    """Compatibility helper retained for older hotfix code."""

    rows = _priority_money_rows(sheet, grand_total_row)
    seen = {row for row, _weight in rows}
    for row in range(1, sheet.max_row + 1):
        if row not in seen:
            rows.append((row, 1))
            seen.add(row)
    return rows


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


def _currency_from_cell_text(value: Any) -> str | None:
    compact = str(value or "").casefold().replace(" ", "")
    if not compact:
        return None
    markers = {
        "KZT": ("₸", "kzt", "тенге", "тг"),
        "EUR": ("€", "eur", "euro", "евро"),
        "USD": ("$", "usd", "доллар"),
        "RUB": ("₽", "rub", "руб"),
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
_legacy._priority_money_rows = _priority_money_rows
_legacy._currency_from_exchange_rate = _currency_from_exchange_rate
_legacy._money_row_currency_scores = _money_row_currency_scores
_legacy._label_currency_scores = _label_currency_scores
_legacy._best_currency = _best_currency
_legacy._currency_from_number_format = _currency_from_number_format
_legacy._currency_from_cell_text = _currency_from_cell_text
_legacy._nearest_number_distance = _nearest_number_distance

for _name, _value in vars(_legacy).items():
    if not (_name.startswith("__") and _name.endswith("__")):
        globals()[_name] = _value

parse_calculation = _legacy.parse_calculation
read_sheet_names = _legacy.read_sheet_names
read_hvac_positions = _legacy.read_hvac_positions
HVACPosition = _legacy.HVACPosition
