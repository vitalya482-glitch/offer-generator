from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class OfferItem:
    no: int
    name: str
    qty: float
    unit_price: float
    total_price: float


@dataclass
class CalcData:
    sheet_name: str
    version: str
    currency: str
    vat_percent: float
    exchange_rate: float
    delivery_basis: str
    items: list[OfferItem] = field(default_factory=list)
    options: list[tuple[str, float]] = field(default_factory=list)

    @property
    def model(self) -> str:
        return self.items[0].name if self.items else "Equipment"

    @property
    def quantity(self) -> float:
        return sum(item.qty for item in self.items)

    @property
    def total_price(self) -> float:
        return sum(item.total_price for item in self.items)

    @property
    def unit_price(self) -> float:
        qty = self.quantity
        return self.total_price / qty if qty else self.total_price


@dataclass
class OfferContext:
    brand: str
    project_dir: Path
    template_path: Path
    calc_path: Path
    output_dir: Path
    client_name: str
    sheet_name: Optional[str] = None
    pdf_dir: Optional[Path] = None
    version: Optional[str] = None
    city: str = "г. Алматы"
