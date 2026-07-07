from __future__ import annotations

from pathlib import Path
from typing import Any

from brands import dc_eltek as _base


_ORIGINAL_DETECT_LAYOUT = _base.detect_offer_layout
_ORIGINAL_READ_ITEMS = _base.read_dc_eltek_offer_items
_ORIGINAL_CELL_NUMBER = _base._cell_number


def _row_text(matrix: dict[int, dict[int, Any]], row: int) -> str:
    return " ".join(_base._normalize_text(value) for value in matrix.get(row, {}).values()).strip()


def _positive_numeric_count(matrix: dict[int, dict[int, Any]], row: int, start_col: int = 1) -> int:
    count = 0
    for col, value in matrix.get(row, {}).items():
        if col < start_col:
            continue
        number = _base._to_number(value)
        if number is not None and number > 0:
            count += 1
    return count


def _find_exact_label_rows(
    matrix: dict[int, dict[int, Any]],
    labels: tuple[str, ...],
    *,
    min_row: int | None = None,
    max_row: int | None = None,
) -> list[int]:
    normalized = tuple(_base._normalize_text(label) for label in labels)
    rows: list[int] = []
    for row in sorted(matrix):
        if min_row is not None and row < min_row:
            continue
        if max_row is not None and row > max_row:
            continue
        for value in matrix.get(row, {}).values():
            text = _base._normalize_text(value)
            if text in normalized:
                rows.append(row)
                break
    return rows


def _find_total_row_after_unit(
    matrix: dict[int, dict[int, Any]],
    unit_price_row: int,
    max_row: int,
) -> int | None:
    # В калькуляциях Eltek итог с НДС часто находится в строке DDP/DAP/EXW
    # сразу под нижней строкой Total per unit.
    incoterm_rows: list[int] = []
    for row in range(unit_price_row + 1, max_row + 1):
        text = _row_text(matrix, row)
        if not text:
            continue
        if any(_base._alias_matches(text, alias) for alias in _base.INCOTERM_ALIASES):
            if _positive_numeric_count(matrix, row, start_col=2) >= 1:
                incoterm_rows.append(row)
    if incoterm_rows:
        return incoterm_rows[0]

    total_quantity_rows = _find_exact_label_rows(
        matrix,
        ("total per quantity", "итого за количество", "общая сумма"),
        min_row=unit_price_row + 1,
        max_row=max_row,
    )
    candidates = [row for row in total_quantity_rows if _positive_numeric_count(matrix, row, start_col=2) >= 1]
    return candidates[-1] if candidates else None


def _find_vat_row(
    matrix: dict[int, dict[int, Any]],
    *,
    quantity_row: int,
    unit_price_row: int,
) -> int | None:
    rows: list[int] = []
    for row in range(quantity_row, unit_price_row):
        text = _row_text(matrix, row)
        if not text:
            continue
        if _base._alias_matches(text, "vat") or _base._alias_matches(text, "ндс"):
            rows.append(row)
    return rows[-1] if rows else None


def _extract_single_vat_percent(matrix: dict[int, dict[int, Any]], vat_row: int | None) -> float:
    if not vat_row:
        return 0.0

    # Процент НДС находится в служебной левой части строки (обычно B29),
    # а в товарных колонках расположены суммы НДС. Берём первое значение 0..100.
    for col in sorted(matrix.get(vat_row, {})):
        number = _base._to_number(matrix[vat_row][col])
        if number is not None and 0 < number <= 100:
            return float(number)
    return 0.0


def detect_offer_layout(calc_path: str | Path, sheet_name: str) -> dict[str, int | None]:
    matrix, _formats, max_row, _max_col = _base._read_sheet_matrix_with_formats(calc_path, sheet_name)
    layout = dict(_ORIGINAL_DETECT_LAYOUT(calc_path, sheet_name))

    name_row = int(layout.get("name_row") or 0)
    quantity_row = int(layout.get("quantity_row") or 0)

    unit_rows = _find_exact_label_rows(
        matrix,
        ("total per unit", "итого за единицу", "цена за единицу", "стоимость за единицу"),
        min_row=quantity_row or name_row or 1,
        max_row=max_row,
    )
    unit_rows = [row for row in unit_rows if _positive_numeric_count(matrix, row, start_col=2) >= 1]
    if unit_rows:
        layout["unit_price_row"] = unit_rows[-1]

    unit_price_row = int(layout.get("unit_price_row") or 0)
    if unit_price_row:
        corrected_total_row = _find_total_row_after_unit(matrix, unit_price_row, max_row)
        if corrected_total_row:
            layout["total_row"] = corrected_total_row

        vat_row = _find_vat_row(
            matrix,
            quantity_row=quantity_row or 1,
            unit_price_row=unit_price_row,
        )
        if vat_row:
            layout["vat_percent_row"] = vat_row

    return layout


def read_dc_eltek_offer_items(
    calc_path: str | Path,
    sheet_name: str,
    currency_override: str | None = None,
) -> dict[str, Any]:
    matrix, _formats, _max_row, _max_col = _base._read_sheet_matrix_with_formats(calc_path, sheet_name)
    layout = detect_offer_layout(calc_path, sheet_name)
    vat_row = int(layout.get("vat_percent_row") or 0) or None
    vat_percent = _extract_single_vat_percent(matrix, vat_row)

    old_detect = _base.detect_offer_layout
    old_cell_number = _base._cell_number

    def patched_detect(_calc_path: str | Path, _sheet_name: str) -> dict[str, int | None]:
        return dict(layout)

    def patched_cell_number(
        source_matrix: dict[int, dict[int, Any]],
        row: int | None,
        col: int,
    ) -> float | None:
        if vat_row and row == vat_row and vat_percent > 0:
            return vat_percent
        return _ORIGINAL_CELL_NUMBER(source_matrix, row, col)

    try:
        _base.detect_offer_layout = patched_detect
        _base._cell_number = patched_cell_number
        result = _ORIGINAL_READ_ITEMS(calc_path, sheet_name, currency_override)
    finally:
        _base.detect_offer_layout = old_detect
        _base._cell_number = old_cell_number

    result["layout"] = dict(layout)
    return result


# Подменяем функции в исходном модуле. Его preview() и make_offer() обращаются
# к глобальному read_dc_eltek_offer_items во время выполнения, поэтому получают исправленную логику.
_base.detect_offer_layout = detect_offer_layout
_base.read_dc_eltek_offer_items = read_dc_eltek_offer_items

BRAND_NAME = _base.BRAND_NAME
DEFAULT_TERMS = _base.DEFAULT_TERMS
build_offer_filename = _base.build_offer_filename
detect_dc_eltek_currency = _base.detect_dc_eltek_currency
extract_client_from_project_path = _base.extract_client_from_project_path
find_default_dc_eltek_template = _base.find_default_dc_eltek_template
find_next_offer_version = _base.find_next_offer_version
format_money = _base.format_money
make_offer = _base.make_offer
preview = _base.preview
