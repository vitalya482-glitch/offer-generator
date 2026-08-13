from __future__ import annotations

import sys
from importlib.abc import Loader, MetaPathFinder
from importlib.machinery import PathFinder
from importlib.util import spec_from_loader
from types import ModuleType
from typing import Any


_TARGET = "gui.pages.stulz_page"
_INSTALLED = False
_LOADING = False


class _StulzPageLoader(Loader):
    def __init__(self, original_loader: Loader, origin: str | None = None) -> None:
        self.original_loader = original_loader
        self.origin = origin

    def create_module(self, spec):
        create = getattr(self.original_loader, "create_module", None)
        return create(spec) if create else None

    def exec_module(self, module: ModuleType) -> None:
        global _LOADING
        _LOADING = True
        try:
            self.original_loader.exec_module(module)
        finally:
            _LOADING = False

        # The original page is fully defined now, so the runtime subclass can
        # safely import it and override only the specification-list methods.
        from gui.pages.stulz_page_runtime import StulzPage

        module.StulzPage = StulzPage


class _StulzPageFinder(MetaPathFinder):
    def find_spec(self, fullname: str, path: Any = None, target: Any = None):
        if fullname != _TARGET or _LOADING:
            return None

        # Ask the normal path finder directly, bypassing sys.meta_path so this
        # finder does not recurse into itself.
        original_spec = PathFinder.find_spec(fullname, path)
        if original_spec is None or original_spec.loader is None:
            return None

        loader = _StulzPageLoader(original_spec.loader, getattr(original_spec, "origin", None))
        spec = spec_from_loader(fullname, loader, origin=getattr(original_spec, "origin", None))
        if spec is not None:
            spec.submodule_search_locations = original_spec.submodule_search_locations
            spec.cached = getattr(original_spec, "cached", None)
            spec.has_location = getattr(original_spec, "has_location", False)
        return spec


def install_stulz_page_hook() -> None:
    """Patch StulzPage lazily without making CLI startup depend on PySide6."""

    global _INSTALLED
    if _INSTALLED:
        return

    # If the page has already been imported, patch it immediately.
    existing = sys.modules.get(_TARGET)
    if existing is not None:
        from gui.pages.stulz_page_runtime import StulzPage

        existing.StulzPage = StulzPage
        _INSTALLED = True
        return

    sys.meta_path.insert(0, _StulzPageFinder())
    _INSTALLED = True
