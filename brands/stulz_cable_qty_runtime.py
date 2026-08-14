from __future__ import annotations

import re

import brands.stulz_legend_runtime as _legend
import core.pdf_parsers.stulz_calc_pdf as _calc_pdf
import core.stulz_specification as _spec_core


# Preserve the complete public surface of the current STULZ runtime.
for _name in dir(_legend):
    if not _name.startswith("__"):
        globals()[_name] = getattr(_legend, _name)


_ORIGINAL_PARSE_STULZ_CALC_OPTIONS = _calc_pdf.parse_stulz_calc_options


def _is_can_bus_cable_option(row: object) -> bool:
    """Return True only for the linear CAN-Bus / RS485 cable option.

    Some STULZ options contain the word "cable" but are supplied as a discrete
    cable kit. Do not reinterpret all cable-related options as metres. The
    supplier option shown in Calc.pdf is specifically "Cable CAN-Bus / RS485",
    where the first quantity is cable length per conditioner.
    """

    source_name = str(getattr(row, "source_name", "") or "")
    description = str(getattr(row, "description", "") or "")
    text = f"{source_name} {description}".lower().replace("ё", "е")

    is_cable = "cable" in text or "кабель" in text
    is_can_bus = "can-bus" in text or "can bus" in text or "canbus" in text or "rs485" in text or "rs-485" in text
    return is_cable and is_can_bus


def _format_linear_cable_qty(value: object) -> str:
    """Convert '10 * 2 шт' to '10 м × 2 шт.' without changing the numbers."""

    text = str(value or "").replace("\u00a0", " ").strip()
    match = re.match(
        r"^\s*(?P<metres>\d+(?:[.,]\d+)?)\s*[\*×xх]\s*(?P<count>\d+(?:[.,]\d+)?)\s*шт\.?\s*$",
        text,
        re.IGNORECASE,
    )
    if not match:
        return text

    metres = match.group("metres")
    count = match.group("count")
    return f"{metres} м × {count} шт."


def parse_stulz_calc_options(path, equipment_qty: float | int = 1):
    rows = _ORIGINAL_PARSE_STULZ_CALC_OPTIONS(path, equipment_qty)
    for row in rows:
        if _is_can_bus_cable_option(row):
            row.qty = _format_linear_cable_qty(row.qty)
    return rows


# core.stulz_specification imported parse_stulz_calc_options by name, therefore
# patch both namespaces. The existing STULZ specification builder will then use
# the corrected unit formatting everywhere (preview and generated Word offer).
_calc_pdf.parse_stulz_calc_options = parse_stulz_calc_options
_spec_core.parse_stulz_calc_options = parse_stulz_calc_options

globals()["parse_stulz_calc_options"] = parse_stulz_calc_options
