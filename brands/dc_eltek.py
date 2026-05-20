from __future__ import annotations

from pathlib import Path

from core.models import OfferContext

BRAND_NAME = "DC Eltek"


def make_offer(context: OfferContext) -> Path:
    raise NotImplementedError("Логика DC Eltek пока не подключена. Добавьте парсер в brands/dc_eltek.py.")


def preview(context: OfferContext) -> str:
    return "Направление: DC Eltek\nЛогика чтения калькуляции пока не подключена."
