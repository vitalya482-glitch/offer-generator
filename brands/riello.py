from __future__ import annotations

from pathlib import Path

from core.models import OfferContext

BRAND_NAME = "Riello"


def make_offer(context: OfferContext) -> Path:
    raise NotImplementedError("Логика Riello пока не подключена. Используйте Stulz или добавьте парсер Riello в brands/riello.py.")


def preview(context: OfferContext) -> str:
    return "Направление: Riello\nЛогика чтения калькуляции пока не подключена."
