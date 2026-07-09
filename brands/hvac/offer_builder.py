from __future__ import annotations

"""HVAC compatibility wrapper around the universal Word offer maker."""

from pathlib import Path
from typing import Any, Mapping, Sequence

from core.excel_calc_parser import CalcItem, format_money, format_qty
from core.final_offer_word_maker import OfferBuildResult, RepeatingTable, make_final_offer


def build_hvac_offer(
    template_path: str | Path,
    output_path: str | Path,
    fields: Mapping[str, Any],
    items: Sequence[CalcItem | Mapping[str, Any]],
) -> OfferBuildResult:
    """Build a one-table HVAC offer using the shared DOCX renderer.

    New code may call :func:`core.final_offer_word_maker.make_final_offer`
    directly.  This wrapper is kept so older imports do not break during the
    gradual migration of the other brands.
    """

    rows = []
    for index, item in enumerate(items, start=1):
        values = _item_dict(item)
        rows.append(
            {
                "item_no": str(index),
                "item_name": values.get("name", ""),
                "item_qty": format_qty(values.get("qty")),
                "item_unit_price": format_money(values.get("unit_price")),
                "item_total": format_money(values.get("total_price")),
            }
        )

    return make_final_offer(
        template_path=template_path,
        output_path=output_path,
        tags=fields,
        repeating_tables=[RepeatingTable("item_name", rows, required=True)],
        clear_unresolved=True,
    )


def _item_dict(item: CalcItem | Mapping[str, Any]) -> dict[str, Any]:
    if isinstance(item, Mapping):
        data = dict(item)
        return {
            "name": data.get("name", data.get("item_name", "")),
            "qty": data.get("qty", data.get("item_qty")),
            "unit_price": data.get("unit_price", data.get("item_unit_price")),
            "total_price": data.get(
                "total_price",
                data.get("item_total", data.get("amount")),
            ),
        }
    return {
        "name": item.name,
        "qty": item.qty,
        "unit_price": item.unit_price,
        "total_price": item.total_price,
    }


__all__ = ["build_hvac_offer"]
