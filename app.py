from __future__ import annotations

import argparse
import sys
from pathlib import Path

from brands import registry
from gui.main_window import run_gui


APP_DIR = Path(__file__).resolve().parent


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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="SAM Offer Generator")
    parser.add_argument("--brand", choices=registry.available_brands(), help="Open a specific brand tab")
    parser.add_argument("--project", help="Project folder path")
    parser.add_argument("--input", help="Input calculation file")
    parser.add_argument("--output", help="Output folder path")
    return parser


def main(argv: list[str] | None = None) -> int:
    # Аргументы пока оставлены для совместимости с ярлыками/скриптами запуска.
    # Текущий gui.main_window управляет выбором вкладок через сохраненные настройки.
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    parser.parse_known_args(argv)

    _cleanup_legacy_updater_files()
    run_gui()
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
