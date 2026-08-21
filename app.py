from __future__ import annotations

import argparse
from importlib import import_module
import sys
from pathlib import Path

from brands.registry import BRANDS
from gui.main_window import run_gui


APP_DIR = Path(__file__).resolve().parent

SELF_CHECK_MODULES = (
    "brands.registry",
    "brands.stulz_full_content_runtime",
    "brands.stulz_position_selection_runtime",
    "brands.stulz_compressor_runtime",
    "gui.pages.stulz_currency_runtime",
)


def _cleanup_legacy_updater_files() -> None:
    """Удаляет остатки старой системы обновления после перехода на LVKUpdater."""
    legacy_paths = (
        APP_DIR / "updater.exe",
        APP_DIR / "updater.exe.old",
        APP_DIR / "config" / "update.json",
    )
    for path in legacy_paths:
        try:
            if path.is_file():
                path.unlink()
        except Exception:
            # Очистка не должна мешать запуску приложения.
            pass


def run_self_check() -> int:
    """Verify that the frozen build contains the runtime modules used by GUI."""

    required_paths = (
        APP_DIR / "SAM-Offer-Generator.exe",
        APP_DIR / "_internal",
    )
    for path in required_paths:
        if not path.exists():
            raise RuntimeError(f"Required runtime path is missing: {path}")

    for module_name in SELF_CHECK_MODULES:
        module = import_module(module_name)
        print(f"OK import {module_name}: {getattr(module, '__file__', '<built-in>')}")

    print("SAM Offer Generator self-check passed.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="SAM Offer Generator")
    parser.add_argument("--brand", choices=list(BRANDS.keys()), help="Open a specific brand tab")
    parser.add_argument("--project", help="Project folder path")
    parser.add_argument("--input", help="Input calculation file")
    parser.add_argument("--output", help="Output folder path")
    parser.add_argument(
        "--self-check",
        action="store_true",
        help="Verify frozen runtime imports and exit without opening the GUI",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    # Аргументы пока оставлены для совместимости с ярлыками/скриптами запуска.
    # Текущий gui.main_window управляет выбором вкладок через сохраненные настройки.
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    args, _unknown = parser.parse_known_args(argv)

    if args.self_check:
        return run_self_check()

    _cleanup_legacy_updater_files()
    run_gui()
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
