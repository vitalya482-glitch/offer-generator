from __future__ import annotations

from pathlib import Path

from core.models import OfferContext

BRAND_NAME = "Generator"


def make_offer(context: OfferContext) -> Path:
    raise NotImplementedError("Логика Generator пока не подключена. Добавьте парсер в brands/generator.py.")


def preview(context: OfferContext) -> str:
    return "Направление: Generator\nЛогика чтения калькуляции пока не подключена."
