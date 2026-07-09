from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


class LVKUpdaterError(RuntimeError):
    """Raised when external LVKUpdater cannot be started."""


def app_dir() -> Path:
    """Return folder that contains SAM-Offer-Generator.exe in frozen build."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def find_lvk_updater(root: Path | None = None) -> Path:
    root = root or app_dir()
    candidates = [
        root / "LVKUpdater.exe",
        root / "_internal" / "LVKUpdater.exe",
    ]
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return candidate

    checked = "\n".join(f"- {path}" for path in candidates)
    raise LVKUpdaterError(
        "Не найден внешний обновлятор LVKUpdater.exe.\n\n"
        f"Проверенные пути:\n{checked}\n\n"
        "Что сделать:\n"
        "1. Пересоберите Offer Generator через GitHub Actions.\n"
        "2. Проверьте, что LVKUpdater.exe попал рядом с SAM-Offer-Generator.exe."
    )


def find_update_config(root: Path | None = None) -> Path:
    root = root or app_dir()
    config = root / "app.update.json"
    if config.exists() and config.is_file():
        return config
    raise LVKUpdaterError(
        "Не найден файл app.update.json рядом с программой.\n\n"
        f"Ожидаемый путь:\n{config}\n\n"
        "Что сделать:\n"
        "Пересоберите Offer Generator, чтобы app.update.json попал в app-модуль."
    )


def start_lvk_update_check() -> None:
    """Start LVKUpdater without closing the running application.

    LVKUpdater checks the manifest while Offer Generator remains open. If the
    user confirms an available update, LVKUpdater closes this process
    gracefully, installs the verified packages, and starts the application
    again. If no update is available, this process is left untouched.
    """
    root = app_dir()
    updater = find_lvk_updater(root)
    config = find_update_config(root)

    cmd = [
        str(updater),
        "--check",
        "--app-dir",
        str(root),
        "--config",
        str(config),
        "--app-pid",
        str(os.getpid()),
    ]

    creationflags = 0
    if sys.platform == "win32":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    try:
        subprocess.Popen(
            cmd,
            cwd=str(root),
            close_fds=True,
            creationflags=creationflags,
        )
    except PermissionError as exc:
        raise LVKUpdaterError(
            "Windows не разрешил запустить LVKUpdater.exe.\n\n"
            f"Файл:\n{updater}\n\n"
            "Что проверить:\n"
            "1. Не заблокировал ли файл антивирус.\n"
            "2. Есть ли права на запуск.\n"
            "3. Не лежит ли программа в защищённой системной папке."
        ) from exc
    except Exception as exc:
        raise LVKUpdaterError(
            "Не удалось запустить LVKUpdater.exe.\n\n"
            f"Файл:\n{updater}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}"
        ) from exc

