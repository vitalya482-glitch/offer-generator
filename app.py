from __future__ import annotations

import argparse
import sys
from pathlib import Path

from brands import registry
from gui.main_window import MainWindow

try:
    from PySide6.QtCore import QLockFile, QStandardPaths
    from PySide6.QtWidgets import QApplication
except Exception as exc:  # pragma: no cover
    print("PySide6 is required to run the GUI:", exc)
    raise


APP_DIR = Path(__file__).resolve().parent
LOCK_BASENAME = "sam_offer_generator.lock"


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


def _single_instance_lock() -> QLockFile:
    runtime_dir = Path(QStandardPaths.writableLocation(QStandardPaths.TempLocation))
    runtime_dir.mkdir(parents=True, exist_ok=True)
    lock_path = runtime_dir / LOCK_BASENAME
    lock = QLockFile(str(lock_path))
    lock.setStaleLockTime(0)
    if not lock.tryLock(100):
        raise RuntimeError("another instance is already running")
    return lock


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="SAM Offer Generator")
    parser.add_argument("--brand", choices=registry.available_brands(), help="Open a specific brand tab")
    parser.add_argument("--project", help="Project folder path")
    parser.add_argument("--input", help="Input calculation file")
    parser.add_argument("--output", help="Output folder path")
    return parser


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    args = parser.parse_args(argv)

    _cleanup_legacy_updater_files()

    app = QApplication(sys.argv[:1])
    app.setApplicationName("SAM Offer Generator")
    app.setOrganizationName("SAM")

    try:
        lock = _single_instance_lock()
    except RuntimeError:
        print("SAM Offer Generator is already running.")
        return 0

    window = MainWindow()

    if args.brand and hasattr(window, "set_brand"):
        window.set_brand(args.brand)
    if args.project and hasattr(window, "set_project_path"):
        window.set_project_path(args.project)
    if args.input and hasattr(window, "set_input_file"):
        window.set_input_file(args.input)
    if args.output and hasattr(window, "set_output_path"):
        window.set_output_path(args.output)

    window.show()
    try:
        return app.exec()
    finally:
        lock.unlock()


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
