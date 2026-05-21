from __future__ import annotations

from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

from core.models import CalcData, OfferItem
from core.utils import as_float, first_not_empty


def find_row_by_label(ws, label: str) -> Optional[int]:
    label_norm = label.lower().strip()
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, str) and val.lower().strip() == label_norm:
                return cell.row
    return None


def find_all_rows_by_label(ws, label: str) -> list[int]:
    label_norm = label.lower().strip()
    result = []
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, str) and val.lower().strip() == label_norm:
                result.append(cell.row)
                break
    return result


def detect_currency(exchange_rate: float, sheet_title: str) -> str:
    if exchange_rate and exchange_rate > 1.01:
        return "KZT"
    title = sheet_title.upper()
    if "KZT" in title or "ТЕНГ" in title:
        return "KZT"
    return "EUR"


def parse_model_groups(ws) -> list[tuple[int, int, str]]:
    groups: list[tuple[int, int, str]] = []
    for col in range(1, ws.max_column + 1):
        label = ws.cell(2, col).value
        if isinstance(label, str) and label.strip().lower() in {"q-ty", "qty", "quantity"}:
            model = first_not_empty(ws.cell(2, col + 1).value, ws.cell(2, col + 2).value)
            if model:
                groups.append((col, col + 1, str(model).strip()))
    if not groups and first_not_empty(ws.cell(2, 4).value):
        groups.append((3, 4, str(ws.cell(2, 4).value).strip()))
    return groups


def parse_stulz_calc(xlsx_path: Path, sheet_name: Optional[str] = None) -> CalcData:
    wb_values = load_workbook(xlsx_path, data_only=True)
    ws = wb_values[sheet_name] if sheet_name else wb_values[wb_values.sheetnames[0]]

    groups = parse_model_groups(ws)
    version = str(first_not_empty(ws["C1"].value, "Version 1"))
    delivery_basis = "DDP г. Алматы" if "DDP" in ws.title.upper() else "EXW Hamburg, Germany"

    total_row = find_row_by_label(ws, "TOTAL")
    total_per_unit_row = find_row_by_label(ws, "Total per unit")
    qty_row = find_row_by_label(ws, "Quantity") or 3
    rate_row = find_row_by_label(ws, "Rate of currency")
    vat_rows = find_all_rows_by_label(ws, "VAT, %")
    vat_row = vat_rows[-1] if vat_rows else None

    first_group_col = groups[0][1] if groups else 4
    exchange_rate = as_float(ws.cell(rate_row, first_group_col).value if rate_row else 1, 1)
    currency = detect_currency(exchange_rate, ws.title)
    vat_percent = as_float(ws.cell(vat_row, 2).value if vat_row else 0, 0)

    items: list[OfferItem] = []
    for qty_col, amount_col, model in groups:
        qty = as_float(ws.cell(qty_row, qty_col).value, 0)
        if qty <= 0:
            continue
        total = as_float(ws.cell(total_row, amount_col).value if total_row else None, 0)
        unit = as_float(ws.cell(total_per_unit_row, amount_col).value if total_per_unit_row else None, 0)
        if not unit and total and qty:
            unit = total / qty
        if not total:
            for label in ("DDP Almaty", "EXW - Hamburg, Germany", "Total per quantity"):
                r = find_row_by_label(ws, label)
                val = as_float(ws.cell(r, amount_col).value if r else None, 0)
                if val:
                    total = val
                    unit = total / qty if qty else total
                    break
        if total or model:
            items.append(OfferItem(len(items) + 1, model, qty, unit, total))

    if not items and ws.title.lower().startswith("spare"):
        total_row = find_row_by_label(ws, "TOTAL DDP - ALA") or find_row_by_label(ws, "TOTAL")
        for row_idx in range(2, min(ws.max_row, 200) + 1):
            name = ws.cell(row_idx, 1).value
            qty = as_float(ws.cell(row_idx, 3).value, 0)
            unit = as_float(ws.cell(row_idx, 6).value, 0)
            total = as_float(ws.cell(row_idx, 7).value, 0)
            if name and qty > 0 and (unit or total):
                items.append(OfferItem(len(items) + 1, str(name).strip(), qty, unit, total))
        if total_row and not exchange_rate:
            exchange_rate = as_float(ws.cell(total_row + 2, 5).value, 1)

    options: list[tuple[str, float]] = []
    start = find_row_by_label(ws, "Options:") or 6
    end = find_row_by_label(ws, "Total EXW for us") or min(start + 80, ws.max_row)
    for row_idx in range(start + 1, end):
        name = ws.cell(row_idx, 1).value
        if not isinstance(name, str) or not name.strip():
            continue
        qty_value = 0.0
        for qty_col, _amount_col, _model in groups or [(3, 4, "")]:
            qty_value += as_float(ws.cell(row_idx, qty_col).value, 0)
        if qty_value > 0:
            options.append((name.strip(), qty_value))

    return CalcData(
        sheet_name=ws.title,
        version=version,
        currency=currency,
        vat_percent=vat_percent,
        exchange_rate=exchange_rate,
        delivery_basis=delivery_basis,
        items=items,
        options=options,
    )


def list_sheets(xlsx_path: Path) -> list[str]:
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    return list(wb.sheetnames)


def read_calc_excel(xlsx_path: Path, sheet_name: Optional[str] = None) -> CalcData:
    """Backward-compatible wrapper used by older brand modules."""
    return parse_stulz_calc(Path(xlsx_path), sheet_name)
