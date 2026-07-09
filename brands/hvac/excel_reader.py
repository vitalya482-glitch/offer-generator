"""Compatibility layer.

The calculation logic now lives in core.excel_calc_parser and is shared by all brands.
This file is intentionally kept for the first migration iteration so existing imports in
brands.hvac.__init__ and offer_builder.py do not break. It contains no parsing logic.
"""

from core.excel_calc_parser import (
    CalcItem as HVACPosition,
    format_money,
    format_qty,
    parse_calculation,
    read_hvac_positions,
    read_sheet_names,
)

__all__ = [
    "HVACPosition",
    "format_money",
    "format_qty",
    "parse_calculation",
    "read_hvac_positions",
    "read_sheet_names",
]
