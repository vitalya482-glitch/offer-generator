from __future__ import annotations

from importlib import import_module
from types import ModuleType

BRANDS = {
    "Stulz": "brands.stulz",
    "Riello": "brands.riello",
    "DC Eltek": "brands.dc_eltek",
    "Generator": "brands.generator",
}


def get_brand_module(name: str) -> ModuleType:
    module_path = BRANDS.get(name)
    if not module_path:
        raise ValueError(f"Неизвестное направление: {name}")
    return import_module(module_path)
