from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from typing import Any


class LVKUpdaterError(RuntimeError):
    """Raised when external LVKUpdater cannot be started."""


_active_updater_process: subprocess.Popen | None = None
_BUNDLED_UPDATER_ARCHIVE = "LVKUpdater-win-x64.zip"
_BUNDLED_UPDATER_EXE = "LVKUpdater.exe"


def app_dir() -> Path:
    """Return folder that contains SAM-Offer-Generator.exe in frozen build."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _refresh_bundled_lvk_updater(root: Path) -> None:
    """Refresh LVKUpdater.exe from the bundled ZIP when it changed.

    The running updater cannot overwrite itself during an application update.
    The app package still receives the new LVKUpdater ZIP, so on the next app
    start/check we can safely replace the updater before launching it.

    This is intentionally best-effort: a missing/corrupt ZIP must not disable an
    otherwise working installed updater.
    """
    archive = root / _BUNDLED_UPDATER_ARCHIVE
    target = root / _BUNDLED_UPDATER_EXE
    if not archive.exists() or not archive.is_file():
        return

    temp_target = root / f"{_BUNDLED_UPDATER_EXE}.new"
    try:
        with zipfile.ZipFile(archive, "r") as zf:
            member = next(
                (
                    name
                    for name in zf.namelist()
                    if Path(name.replace("\\", "/")).name.lower() == _BUNDLED_UPDATER_EXE.lower()
                ),
                None,
            )
            if not member:
                return
            payload = zf.read(member)

        if target.exists() and target.is_file():
            try:
                if target.read_bytes() == payload:
                    return
            except Exception:
                pass

        temp_target.write_bytes(payload)
        os.replace(temp_target, target)
    except Exception:
        # Keep the already installed updater usable if the bundled copy cannot
        # be refreshed for any reason.
        try:
            temp_target.unlink(missing_ok=True)
        except Exception:
            pass


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


def _read_json(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise LVKUpdaterError(
            "Не удалось прочитать настройки LVKUpdater.\n\n"
            f"Файл:\n{path}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}"
        ) from exc
    if not isinstance(data, dict):
        raise LVKUpdaterError(f"Файл {path.name} должен содержать JSON-объект.")
    return data


def _repository_from_config(root: Path, config: dict[str, Any]) -> str:
    for key in ("githubRepository", "github_repository"):
        value = str(config.get(key, "") or "").strip()
        if re.fullmatch(r"[^/\s]+/[^/\s]+", value):
            return value

    release_info = root / "release_info.json"
    if release_info.exists():
        try:
            info = json.loads(release_info.read_text(encoding="utf-8-sig"))
            value = str(info.get("github_repository", "") or "").strip()
            if re.fullmatch(r"[^/\s]+/[^/\s]+", value):
                return value
        except Exception:
            pass

    manifest_url = str(config.get("manifestUrl", "") or "").strip()
    match = re.match(r"https://github\.com/([^/]+/[^/]+)/releases/", manifest_url, re.IGNORECASE)
    if match:
        return match.group(1)

    raise LVKUpdaterError(
        "Не удалось определить GitHub-репозиторий для обновления.\n\n"
        "Проверьте release_info.json и app.update.json рядом с программой."
    )


def _manifest_asset_name(config: dict[str, Any]) -> str:
    explicit = str(config.get("manifestAsset", "") or "").strip()
    if explicit:
        return explicit
    manifest_url = str(config.get("manifestUrl", "") or "").strip()
    if manifest_url:
        name = manifest_url.rsplit("/", 1)[-1].split("?", 1)[0].strip()
        if name:
            return name
    return "offer-generator.json"


def _resolve_latest_manifest_url(repository: str, asset_name: str) -> str:
    """Resolve a version-specific manifest URL via the GitHub Releases API.

    Do not use ``releases/latest/download/...`` directly. GitHub/CDN can cache a
    redirect for that alias longer than the signed release-assets URL remains
    valid. Resolving the current release through the API gives us the concrete
    browser_download_url for the latest manifest asset on every check.
    """

    api_url = f"https://api.github.com/repos/{repository}/releases/latest"
    request = urllib.request.Request(
        api_url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "SAM-Offer-Generator",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise LVKUpdaterError(
            "GitHub не дал получить информацию о последнем релизе.\n\n"
            f"HTTP {exc.code}\n{api_url}"
        ) from exc
    except Exception as exc:
        raise LVKUpdaterError(
            "Не удалось определить последний релиз через GitHub API.\n\n"
            f"URL:\n{api_url}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}"
        ) from exc

    assets = payload.get("assets", []) if isinstance(payload, dict) else []
    for asset in assets if isinstance(assets, list) else []:
        if not isinstance(asset, dict):
            continue
        if str(asset.get("name", "")) != asset_name:
            continue
        download_url = str(asset.get("browser_download_url", "") or "").strip()
        if download_url:
            return download_url

    tag = str(payload.get("tag_name", "") or "").strip() if isinstance(payload, dict) else ""
    available = ", ".join(
        str(asset.get("name", ""))
        for asset in assets
        if isinstance(asset, dict) and asset.get("name")
    )
    raise LVKUpdaterError(
        "В последнем GitHub Release не найден manifest обновления.\n\n"
        f"Release: {tag or 'не определён'}\n"
        f"Искали: {asset_name}\n"
        f"Доступно: {available or 'нет файлов'}"
    )


def _prepare_resolved_config(root: Path, config_path: Path) -> Path:
    config = _read_json(config_path)
    repository = _repository_from_config(root, config)
    asset_name = _manifest_asset_name(config)
    resolved_url = _resolve_latest_manifest_url(repository, asset_name)

    config["githubRepository"] = repository
    config["manifestAsset"] = asset_name
    config["manifestUrl"] = resolved_url

    # Updating the real file is intentional: it also repairs installations that
    # still contain the old stale releases/latest/download URL. A successful app
    # update will replace this file with the current package afterwards.
    try:
        config_path.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as exc:
        raise LVKUpdaterError(
            "Не удалось обновить app.update.json перед запуском обновлятора.\n\n"
            f"Файл:\n{config_path}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}"
        ) from exc
    return config_path


def _updater_is_running() -> bool:
    global _active_updater_process
    process = _active_updater_process
    if process is None:
        return False
    if process.poll() is None:
        return True
    _active_updater_process = None
    return False


def start_lvk_update_check() -> None:
    """Start one LVKUpdater process with a freshly resolved manifest URL.

    Offer Generator resolves the exact latest release asset through the GitHub
    API before starting LVKUpdater. This avoids stale ``latest/download`` CDN
    redirects. Repeated button clicks while the updater is still running are
    ignored so only one update dialog can be open at a time.
    """

    global _active_updater_process

    if _updater_is_running():
        return

    root = app_dir()
    _refresh_bundled_lvk_updater(root)
    updater = find_lvk_updater(root)
    config = _prepare_resolved_config(root, find_update_config(root))

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
        _active_updater_process = subprocess.Popen(
            cmd,
            cwd=str(root),
            close_fds=True,
            creationflags=creationflags,
        )
    except PermissionError as exc:
        _active_updater_process = None
        raise LVKUpdaterError(
            "Windows не разрешил запустить LVKUpdater.exe.\n\n"
            f"Файл:\n{updater}\n\n"
            "Что проверить:\n"
            "1. Не заблокировал ли файл антивирус.\n"
            "2. Есть ли права на запуск.\n"
            "3. Не лежит ли программа в защищённой системной папке."
        ) from exc
    except Exception as exc:
        _active_updater_process = None
        raise LVKUpdaterError(
            "Не удалось запустить LVKUpdater.exe.\n\n"
            f"Файл:\n{updater}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}"
        ) from exc
