from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Iterable


def is_frozen() -> bool:
    """Return True when the app is running from a PyInstaller build."""
    return bool(getattr(sys, "frozen", False))


def app_root() -> Path:
    """Writable application root.

    Source mode:
        <repo root>

    PyInstaller one-dir mode:
        folder that contains SAM-Offer-Generator.exe

    Mutable files such as config/*.json must live here, not inside PyInstaller's
    internal bundle directory.
    """
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def bundled_root() -> Path:
    """Read-only bundled resources root.

    In a PyInstaller build this is normally the _internal/_MEIPASS directory.
    In source mode it is the repository root.
    """
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", app_root())).resolve()
    return app_root()


def user_config_dir() -> Path:
    return app_root() / "config"


def bundled_config_dir() -> Path:
    return bundled_root() / "config"


def ensure_editable_config(filenames: Iterable[str]) -> Path:
    """Ensure editable config files exist next to the executable.

    PyInstaller stores bundled data in an internal/read-only runtime location.
    This helper copies default JSON files from that bundle to <app>/config on
    the first run so the GUI settings dialog can safely edit them.
    """
    target_dir = user_config_dir()
    target_dir.mkdir(parents=True, exist_ok=True)

    source_dir = bundled_config_dir()
    for filename in filenames:
        target = target_dir / filename
        if target.exists():
            continue
        source = source_dir / filename
        if source.exists() and source.resolve() != target.resolve():
            shutil.copy2(source, target)

    return target_dir


def resource_path(relative_path: str | Path) -> Path:
    """Resolve a bundled resource path in both source and frozen modes."""
    relative = Path(relative_path)
    external = app_root() / relative
    if external.exists():
        return external
    return bundled_root() / relative


def app_icon_path() -> Path:
    return resource_path(Path("assets") / "app_icon.ico")
