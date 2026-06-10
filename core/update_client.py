from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

APP_MODULE_ASSET = "SAM-Offer-Generator-App-No-Runtime.zip"
RUNTIME_MODULE_ASSET = "SAM-Offer-Generator-Runtime-Win64.zip"
CONFIG_RELATIVE_PATH = Path("config") / "update.json"


class UpdateError(RuntimeError):
    """Raised when update check/download/start cannot be completed."""


@dataclass(frozen=True)
class ReleaseAsset:
    name: str
    download_url: str
    size: int


@dataclass(frozen=True)
class ReleaseInfo:
    tag_name: str
    name: str
    html_url: str
    app_asset: ReleaseAsset | None
    runtime_asset: ReleaseAsset | None


def app_dir() -> Path:
    """Return folder that contains SAM-Offer-Generator.exe in frozen build."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def current_version() -> str:
    release_info = app_dir() / "release_info.json"
    if release_info.exists():
        try:
            data = json.loads(release_info.read_text(encoding="utf-8"))
            for key in ("app_version", "version", "tag_name"):
                value = data.get(key)
                if value:
                    return normalize_version(str(value))
        except Exception:
            pass
    cfg = load_update_config()
    return normalize_version(str(cfg.get("current_version", "0.0.0")))


def load_update_config() -> dict[str, Any]:
    path = app_dir() / CONFIG_RELATIVE_PATH
    if not path.exists():
        raise UpdateError(f"Не найден файл настроек обновления: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise UpdateError(f"Не удалось прочитать {path}: {exc}") from exc
    if not isinstance(data, dict):
        raise UpdateError(f"Файл {path} должен содержать JSON-объект")
    return data


def normalize_version(text: str) -> str:
    text = text.strip()
    return text[1:] if text.lower().startswith("v") else text


def version_tuple(text: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", normalize_version(text))
    return tuple(int(x) for x in nums) if nums else (0,)


def is_newer(latest: str, current: str) -> bool:
    return version_tuple(latest) > version_tuple(current)


def fetch_latest_release() -> ReleaseInfo:
    cfg = load_update_config()
    if not bool(cfg.get("enabled", True)):
        raise UpdateError("Проверка обновлений выключена в config/update.json")

    owner = str(cfg.get("github_owner", "")).strip()
    repo = str(cfg.get("github_repo", "")).strip()
    if not owner or not repo or owner.upper() == "OWNER" or repo.upper() == "REPO":
        raise UpdateError(
            "Не настроен GitHub repository. Заполните github_owner и github_repo в config/update.json"
        )

    api_url = str(
        cfg.get("latest_release_api_url")
        or f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    )
    data = _read_json_url(api_url)

    assets = data.get("assets", [])
    app_asset = _find_asset(assets, str(cfg.get("app_asset_name", APP_MODULE_ASSET)))
    runtime_asset = _find_asset(assets, str(cfg.get("runtime_asset_name", RUNTIME_MODULE_ASSET)))

    return ReleaseInfo(
        tag_name=str(data.get("tag_name", "")),
        name=str(data.get("name", "")),
        html_url=str(data.get("html_url", "")),
        app_asset=app_asset,
        runtime_asset=runtime_asset,
    )


def _read_json_url(url: str) -> dict[str, Any]:
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "SAM-Offer-Generator-Updater",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw)
    except Exception as exc:
        raise UpdateError(f"Не удалось получить данные обновления: {exc}") from exc


def _find_asset(assets: list[Any], name: str) -> ReleaseAsset | None:
    for item in assets:
        if not isinstance(item, dict):
            continue
        if item.get("name") == name:
            url = item.get("browser_download_url")
            if not url:
                continue
            return ReleaseAsset(
                name=name,
                download_url=str(url),
                size=int(item.get("size") or 0),
            )
    return None


def download_asset(asset: ReleaseAsset, target_dir: Path | None = None) -> Path:
    if target_dir is None:
        target_dir = app_dir() / "updates"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / asset.name
    temp = target.with_suffix(target.suffix + ".download")

    req = urllib.request.Request(asset.download_url, headers={"User-Agent": "SAM-Offer-Generator-Updater"})
    try:
        with urllib.request.urlopen(req, timeout=60) as response, temp.open("wb") as out:
            shutil.copyfileobj(response, out)
        temp.replace(target)
        return target
    except Exception as exc:
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        raise UpdateError(f"Не удалось скачать обновление: {exc}") from exc


def updater_exe_path() -> Path:
    root = app_dir()
    candidates = [
        root / "updater.exe",
        root / "_internal" / "updater.exe",
        root / "updater.py",
    ]
    for path in candidates:
        if path.exists():
            return path
    raise UpdateError("Не найден updater.exe рядом с программой")


def start_updater(package_path: Path, restart: bool = True) -> None:
    root = app_dir()
    updater = updater_exe_path()
    app_exe = Path(sys.executable).resolve() if getattr(sys, "frozen", False) else root / "app.py"

    if updater.suffix.lower() == ".py":
        cmd = [sys.executable, str(updater)]
    else:
        cmd = [str(updater)]

    cmd += [
        "--package",
        str(package_path),
        "--app-dir",
        str(root),
        "--pid",
        str(os.getpid()),
    ]
    if restart:
        cmd += ["--restart", str(app_exe)]

    try:
        subprocess.Popen(cmd, cwd=str(root), close_fds=True)
    except Exception as exc:
        raise UpdateError(f"Не удалось запустить updater: {exc}") from exc


def check_app_update() -> tuple[bool, str, ReleaseInfo]:
    release = fetch_latest_release()
    current = current_version()
    latest = normalize_version(release.tag_name)
    return is_newer(latest, current), current, release
