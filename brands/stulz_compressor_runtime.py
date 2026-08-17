from __future__ import annotations

import re
from typing import Any

import brands.stulz_cable_qty_runtime as _cable
import core.pdf_parsers.stulz_winplan_pdf as _winplan
import core.stulz_specification as _spec_core


# Preserve the full public STULZ surface: legend logic + cable quantity fix.
for _name in dir(_cable):
    if not _name.startswith("__"):
        globals()[_name] = getattr(_cable, _name)


STULZ_WINPLAN_PARSER_VERSION = "compressor-blocks-v2"
_ORIGINAL_PARSE_STULZ_WINPLAN_SPECS = _spec_core.parse_stulz_winplan_specs


def _plain(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\u00a0", " ")).strip()


def _split_parallel_value(value: object) -> list[str]:
    """Split parallel WinPlan values without breaking units such as kW/kW.

    WinPlan/PDF extraction often produces one label followed by values for two
    compressors in the same text cell, for example:
      9,4 кВт / 11,8 кВт
      4,32 кВт/кВт / 3,65 кВт/кВт
    Only a slash surrounded by whitespace is considered a compressor separator.
    """

    text = _plain(value)
    if not text:
        return []
    parts = [_plain(part) for part in re.split(r"\s+/\s+", text) if _plain(part)]
    return parts or [text]


def _is_compressor_section(row: object) -> bool:
    return bool(getattr(row, "is_section", False)) and _plain(getattr(row, "name", "")).lower().startswith("компрессор")


def _is_number_row(row: object) -> bool:
    name = _plain(getattr(row, "name", "")).lower().replace("ё", "е")
    return "количество" in name or name.startswith("number")


def _expand_single_compressor_block(block: list[Any]) -> list[Any]:
    """Expand one generic Compressor block into Compressor 1/2/... blocks.

    The core parser may receive either repeated WinPlan labels or a single label
    whose value already contains two parallel values separated by `` / ``. The
    latter was the reason v128 still rendered one compressor block. Infer the
    compressor count from the widest value row and duplicate true singleton
    fields (notably ``Количество: 1``) into every compressor subsection.
    """

    if not block:
        return block

    title = block[0]
    title_name = _plain(getattr(title, "name", ""))
    if re.match(r"^Компрессор\s+\d+\s*:", title_name, flags=re.IGNORECASE):
        # Already split by the core parser.
        return block

    data_rows = [row for row in block[1:] if not getattr(row, "is_section", False)]
    parts_by_row = {id(row): _split_parallel_value(getattr(row, "value", "")) for row in data_rows}
    compressor_count = max((len(parts) for parts in parts_by_row.values()), default=1)
    if compressor_count <= 1:
        return block

    row_type = type(title)
    result: list[Any] = []
    for index in range(compressor_count):
        result.append(row_type(f"Компрессор {index + 1}:", "", True))
        for row in data_rows:
            parts = parts_by_row.get(id(row), [])
            if not parts:
                continue
            if len(parts) == 1:
                # WinPlan can print a shared/identical field once even though the
                # neighbouring compressor properties prove there are two units.
                # ``Количество: 1`` is the common case; duplicating a singleton
                # keeps each compressor block self-contained like WinPlan.
                value = parts[0]
            elif index < len(parts):
                value = parts[index]
            else:
                continue
            result.append(row_type(getattr(row, "name", ""), value, False))

    return result


def _expand_compressor_sections(rows: list[Any]) -> list[Any]:
    result: list[Any] = []
    index = 0
    while index < len(rows):
        row = rows[index]
        if not _is_compressor_section(row):
            result.append(row)
            index += 1
            continue

        block = [row]
        index += 1
        while index < len(rows) and not getattr(rows[index], "is_section", False):
            block.append(rows[index])
            index += 1

        result.extend(_expand_single_compressor_block(block))

    return result


def parse_stulz_winplan_specs(path):
    rows = _ORIGINAL_PARSE_STULZ_WINPLAN_SPECS(path)
    return _expand_compressor_sections(rows)


# core.stulz_specification imported the parser by name, therefore patch both
# namespaces. build_stulz_specification() will now use the final runtime parser
# for preview and generated Word specifications.
_winplan.parse_stulz_winplan_specs = parse_stulz_winplan_specs
_spec_core.parse_stulz_winplan_specs = parse_stulz_winplan_specs

globals()["parse_stulz_winplan_specs"] = parse_stulz_winplan_specs

# The GUI mapping layer is imported last so it patches the already constructed
# STULZ page after the mature runtime/legend/specification hooks are installed.
import gui.pages.stulz_calc_mapping_runtime as _calc_mapping_ui  # noqa: E402,F401
