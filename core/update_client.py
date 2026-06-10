from __future__ import annotations

import json
import os
import re
import shutil
import socket
import ssl
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

APP_MODULE_ASSET = "SAM-Offer-Generator-App-No-Runtime.zip"
RUNTIME_MODULE_ASSET = "SAM-Offer-Generator-Runtime-Win64.zip"
CONFIG_RELATIVE_PATH = Path("config") / "update.json"


class UpdateError(RuntimeError):
    """Raised when update check/download/start cannot be completed."""

    def __init__(self, message: str, *, code: str | None = None) -> None:
        self.code = code or "UPDATE_ERROR"
        super().__init__(message)


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
        raise UpdateError(
            "Не найден файл настроек обновления.\n\n"
            f"Ожидаемый путь:\n{path}\n\n"
            "Что сделать:\n"
            "1. Проверьте, что рядом с программой есть папка config.\n"
            "2. Проверьте, что внутри есть файл update.json.\n"
            "3. Если файла нет — скопируйте его из проекта или из последнего app-модуля.",
            code="CONFIG_NOT_FOUND",
        )
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise UpdateError(
            "Ошибка в формате config/update.json.\n\n"
            f"Файл:\n{path}\n\n"
            f"Строка {exc.lineno}, колонка {exc.colno}: {exc.msg}\n\n"
            "Что сделать:\n"
            "Откройте update.json и проверьте запятые, кавычки и фигурные скобки.",
            code="CONFIG_INVALID_JSON",
        ) from exc
    except PermissionError as exc:
        raise UpdateError(
            "Нет прав на чтение config/update.json.\n\n"
            f"Файл:\n{path}\n\n"
            "Что сделать:\n"
            "Проверьте права доступа к папке программы. Лучше хранить программу не в Program Files, "
            "а в папке пользователя или на диске, где есть права записи/чтения.",
            code="CONFIG_PERMISSION_DENIED",
        ) from exc
    except Exception as exc:
        raise UpdateError(
            "Не удалось прочитать config/update.json.\n\n"
            f"Файл:\n{path}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}",
            code="CONFIG_READ_ERROR",
        ) from exc
    if not isinstance(data, dict):
        raise UpdateError(
            "Файл config/update.json должен содержать JSON-объект.\n\n"
            "Пример:\n"
            "{\n"
            '  "github_owner": "vitalya482-glitch",\n'
            '  "github_repo": "offer-generator",\n'
            '  "current_version": "0.2.0"\n'
            "}",
            code="CONFIG_NOT_OBJECT",
        )
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
        raise UpdateError(
            "Проверка обновлений выключена в config/update.json.\n\n"
            "Что сделать:\n"
            "Поставьте:\n"
            '"enabled": true',
            code="UPDATE_DISABLED",
        )

    owner = str(cfg.get("github_owner", "")).strip()
    repo = str(cfg.get("github_repo", "")).strip()
    if not owner or not repo or owner.upper() == "OWNER" or repo.upper() == "REPO":
        raise UpdateError(
            "Не настроен GitHub repository.\n\n"
            "В файле config/update.json нужно заполнить:\n"
            '"github_owner": "vitalya482-glitch",\n'
            '"github_repo": "offer-generator"\n\n'
            "Где взять значения:\n"
            "Откройте репозиторий в браузере. В адресе github.com/OWNER/REPO возьмите OWNER и REPO.",
            code="GITHUB_REPO_NOT_CONFIGURED",
        )

    _validate_github_name(owner, "github_owner")
    _validate_github_name(repo, "github_repo")

    api_url = str(
        cfg.get("latest_release_api_url")
        or f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    ).strip()
    _validate_url(api_url)

    data = _read_json_url(api_url, context="latest_release", owner=owner, repo=repo)
    if not isinstance(data, dict):
        raise UpdateError(
            "GitHub вернул неожиданный формат ответа.\n\n"
            f"URL:\n{api_url}\n\n"
            "Ожидался JSON-объект latest release.",
            code="GITHUB_BAD_RESPONSE_FORMAT",
        )

    tag_name = str(data.get("tag_name", "")).strip()
    if not tag_name:
        raise UpdateError(
            "В latest release нет поля tag_name.\n\n"
            "Что проверить:\n"
            "1. Это действительно GitHub Releases, а не Actions artifacts.\n"
            "2. Release создан с тегом, например v0.2.1.",
            code="RELEASE_TAG_MISSING",
        )

    assets = data.get("assets", [])
    if not isinstance(assets, list):
        raise UpdateError(
            "В latest release поле assets имеет неожиданный формат.\n\n"
            "Ожидался список файлов релиза.",
            code="RELEASE_ASSETS_BAD_FORMAT",
        )
    if not assets:
        raise UpdateError(
            "Latest release найден, но в нём нет прикреплённых файлов.\n\n"
            f"Release: {tag_name}\n"
            f"Страница: {data.get('html_url', '')}\n\n"
            "Что сделать:\n"
            "Прикрепите к GitHub Release файл app-модуля, например:\n"
            f"{cfg.get('app_asset_name', APP_MODULE_ASSET)}",
            code="RELEASE_HAS_NO_ASSETS",
        )

    app_asset_name = str(cfg.get("app_asset_name", APP_MODULE_ASSET)).strip()
    runtime_asset_name = str(cfg.get("runtime_asset_name", RUNTIME_MODULE_ASSET)).strip()
    app_asset = _find_asset(assets, app_asset_name)
    runtime_asset = _find_asset(assets, runtime_asset_name)

    if app_asset is None:
        available = _format_available_assets(assets)
        raise UpdateError(
            "Latest release найден, но app-модуль не найден среди файлов релиза.\n\n"
            f"Искали файл:\n{app_asset_name}\n\n"
            f"Release: {tag_name}\n"
            f"Страница: {data.get('html_url', '')}\n\n"
            f"Файлы, которые есть в релизе:\n{available}\n\n"
            "Что сделать:\n"
            "1. Проверьте app_asset_name в config/update.json.\n"
            "2. Либо переименуйте asset в GitHub Release.\n"
            "3. Важно: Actions artifact и Release asset — это разные вещи. "
            "Для автообновления нужен файл именно в GitHub Release.",
            code="APP_ASSET_NOT_FOUND",
        )

    return ReleaseInfo(
        tag_name=tag_name,
        name=str(data.get("name", "")),
        html_url=str(data.get("html_url", "")),
        app_asset=app_asset,
        runtime_asset=runtime_asset,
    )


def _validate_github_name(value: str, field_name: str) -> None:
    if "/" in value or "\\" in value or " " in value:
        raise UpdateError(
            f"Некорректное значение {field_name} в config/update.json.\n\n"
            f"Сейчас указано:\n{value}\n\n"
            "Нужно указывать только одну часть адреса GitHub.\n"
            "Пример для https://github.com/vitalya482-glitch/offer-generator:\n"
            'github_owner = "vitalya482-glitch"\n'
            'github_repo = "offer-generator"',
            code="GITHUB_NAME_INVALID",
        )


def _validate_url(url: str) -> None:
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise UpdateError(
            "Некорректный latest_release_api_url в config/update.json.\n\n"
            f"Сейчас указано:\n{url}\n\n"
            "Если вы не используете специальный адрес, просто удалите latest_release_api_url из config/update.json.",
            code="INVALID_UPDATE_URL",
        )


def _read_json_url(url: str, *, context: str, owner: str = "", repo: str = "") -> dict[str, Any]:
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "SAM-Offer-Generator-Updater",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as response:
            raw = response.read().decode("utf-8")
            try:
                return json.loads(raw)
            except json.JSONDecodeError as exc:
                raise UpdateError(
                    "GitHub вернул ответ, но это не JSON.\n\n"
                    f"URL:\n{url}\n\n"
                    f"Строка {exc.lineno}, колонка {exc.colno}: {exc.msg}",
                    code="RESPONSE_INVALID_JSON",
                ) from exc
    except UpdateError:
        raise
    except urllib.error.HTTPError as exc:
        message = _format_http_error(exc, url=url, context=context, owner=owner, repo=repo)
        raise UpdateError(message, code=f"HTTP_{exc.code}") from exc
    except urllib.error.URLError as exc:
        reason = exc.reason
        if isinstance(reason, socket.timeout):
            raise UpdateError(
                "Превышено время ожидания ответа от GitHub.\n\n"
                f"URL:\n{url}\n\n"
                "Что проверить:\n"
                "1. Интернет-соединение.\n"
                "2. VPN/прокси/корпоративный firewall.\n"
                "3. Попробуйте открыть этот адрес в браузере.",
                code="NETWORK_TIMEOUT",
            ) from exc
        if isinstance(reason, ssl.SSLError):
            raise UpdateError(
                "Ошибка SSL-сертификата при подключении к GitHub.\n\n"
                f"URL:\n{url}\n\n"
                "Что проверить:\n"
                "1. Дату и время на компьютере.\n"
                "2. Корпоративный антивирус/прокси, который подменяет сертификаты.\n"
                "3. Доступ к https://github.com в браузере.",
                code="SSL_ERROR",
            ) from exc
        raise UpdateError(
            "Не удалось подключиться к GitHub.\n\n"
            f"URL:\n{url}\n\n"
            f"Причина:\n{reason}\n\n"
            "Что проверить:\n"
            "1. Есть ли интернет.\n"
            "2. Открывается ли GitHub в браузере.\n"
            "3. Не блокирует ли подключение firewall, proxy или антивирус.",
            code="NETWORK_ERROR",
        ) from exc
    except socket.timeout as exc:
        raise UpdateError(
            "Превышено время ожидания ответа от GitHub.\n\n"
            f"URL:\n{url}",
            code="NETWORK_TIMEOUT",
        ) from exc
    except Exception as exc:
        raise UpdateError(
            "Не удалось получить данные обновления.\n\n"
            f"URL:\n{url}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}",
            code="UPDATE_CHECK_UNKNOWN_ERROR",
        ) from exc


def _format_http_error(exc: urllib.error.HTTPError, *, url: str, context: str, owner: str, repo: str) -> str:
    body = ""
    try:
        body = exc.read().decode("utf-8", errors="replace")[:800]
    except Exception:
        body = ""

    base = f"GitHub вернул HTTP Error {exc.code}: {exc.reason}.\n\nURL:\n{url}\n"

    if exc.code == 404 and context == "latest_release":
        return (
            "GitHub не нашёл latest release.\n\n"
            f"Репозиторий:\n{owner}/{repo}\n\n"
            "Что это обычно означает:\n"
            "1. В репозитории ещё нет GitHub Release.\n"
            "2. Есть только Actions artifacts — они не подходят для этого updater.\n"
            "3. Неправильно указан github_owner или github_repo в config/update.json.\n"
            "4. Репозиторий private, а updater работает без GitHub token.\n\n"
            "Что сделать:\n"
            "1. Откройте GitHub → Releases.\n"
            "2. Создайте Release с тегом, например v0.2.1.\n"
            "3. Прикрепите app-модуль к Release.\n\n"
            f"Технически: HTTP 404 на {url}"
        )
    if exc.code == 404:
        return (
            base
            + "\nВозможные причины:\n"
            + "1. Неверный адрес GitHub API.\n"
            + "2. Репозиторий или файл не существует.\n"
            + "3. Репозиторий private и нет доступа без token.\n"
            + (f"\nОтвет GitHub:\n{body}" if body else "")
        )
    if exc.code == 401:
        return (
            base
            + "\nGitHub требует авторизацию.\n\n"
            + "Возможные причины:\n"
            + "1. Репозиторий private.\n"
            + "2. Нужен GitHub token.\n"
            + "3. Токен указан неверно или истёк."
        )
    if exc.code == 403:
        return (
            base
            + "\nДоступ запрещён или превышен лимит GitHub API.\n\n"
            + "Возможные причины:\n"
            + "1. Превышен лимит запросов GitHub API.\n"
            + "2. Репозиторий private.\n"
            + "3. Корпоративная сеть блокирует GitHub API.\n"
            + "4. GitHub требует авторизацию для этого ресурса.\n\n"
            + "Что попробовать:\n"
            + "Откройте URL в браузере. Если GitHub показывает rate limit или forbidden — нужно подождать или добавить token."
            + (f"\n\nОтвет GitHub:\n{body}" if body else "")
        )
    if exc.code == 429:
        return (
            base
            + "\nСлишком много запросов к GitHub.\n\n"
            + "Что сделать:\n"
            + "Подождите несколько минут и попробуйте снова."
        )
    if 500 <= exc.code <= 599:
        return (
            base
            + "\nПроблема на стороне GitHub или промежуточной сети.\n\n"
            + "Что сделать:\n"
            + "Попробуйте позже и проверьте, открывается ли github.com в браузере."
        )
    return base + (f"\nОтвет сервера:\n{body}" if body else "")


def _asset_names(assets: list[Any]) -> list[str]:
    names: list[str] = []
    for item in assets:
        if isinstance(item, dict) and item.get("name"):
            names.append(str(item["name"]))
    return names


def _format_available_assets(assets: list[Any]) -> str:
    names = _asset_names(assets)
    if not names:
        return "- файлов нет"
    return "\n".join(f"- {name}" for name in names[:30])


def _candidate_asset_names(name: str) -> list[str]:
    cleaned = name.strip()
    candidates = [cleaned]
    if cleaned and not cleaned.lower().endswith(".zip"):
        candidates.append(cleaned + ".zip")
    if cleaned.lower().endswith(".zip"):
        candidates.append(cleaned[:-4])
    # GitHub artifact names in Actions often use '-module', while Release assets may be exported as '.zip'.
    # Keep matching forgiving, but still show exact diagnostics if nothing is found.
    return list(dict.fromkeys(c for c in candidates if c))


def _find_asset(assets: list[Any], name: str) -> ReleaseAsset | None:
    candidates = _candidate_asset_names(name)
    lower_candidates = {c.lower() for c in candidates}

    for item in assets:
        if not isinstance(item, dict):
            continue
        item_name = str(item.get("name") or "")
        if item_name in candidates or item_name.lower() in lower_candidates:
            url = item.get("browser_download_url")
            if not url:
                raise UpdateError(
                    "Файл релиза найден, но у него нет browser_download_url.\n\n"
                    f"Asset:\n{item_name}\n\n"
                    "Что сделать:\n"
                    "Проверьте, что это обычный файл GitHub Release, а не служебная запись API.",
                    code="ASSET_DOWNLOAD_URL_MISSING",
                )
            return ReleaseAsset(
                name=item_name,
                download_url=str(url),
                size=int(item.get("size") or 0),
            )
    return None


def download_asset(asset: ReleaseAsset, target_dir: Path | None = None) -> Path:
    if target_dir is None:
        target_dir = app_dir() / "updates"
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except PermissionError as exc:
        raise UpdateError(
            "Нет прав на создание папки updates рядом с программой.\n\n"
            f"Папка:\n{target_dir}\n\n"
            "Что сделать:\n"
            "Перенесите программу из Program Files в папку пользователя, например:\n"
            r"C:\Users\<User>\Apps\SAM-Offer-Generator",
            code="UPDATES_DIR_PERMISSION_DENIED",
        ) from exc

    target = target_dir / asset.name
    temp = target.with_suffix(target.suffix + ".download")

    req = urllib.request.Request(asset.download_url, headers={"User-Agent": "SAM-Offer-Generator-Updater"})
    try:
        with urllib.request.urlopen(req, timeout=60) as response, temp.open("wb") as out:
            shutil.copyfileobj(response, out)
        if asset.size and temp.stat().st_size <= 0:
            raise UpdateError(
                "Скачанный файл обновления пустой.\n\n"
                f"Asset:\n{asset.name}",
                code="DOWNLOADED_FILE_EMPTY",
            )
        temp.replace(target)
        return target
    except UpdateError:
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        raise
    except urllib.error.HTTPError as exc:
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        raise UpdateError(
            _format_http_error(exc, url=asset.download_url, context="download", owner="", repo=""),
            code=f"DOWNLOAD_HTTP_{exc.code}",
        ) from exc
    except PermissionError as exc:
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        raise UpdateError(
            "Нет прав на сохранение файла обновления.\n\n"
            f"Папка:\n{target_dir}\n\n"
            "Что сделать:\n"
            "Проверьте, что программа лежит в папке, доступной для записи текущему пользователю.",
            code="DOWNLOAD_PERMISSION_DENIED",
        ) from exc
    except Exception as exc:
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        raise UpdateError(
            "Не удалось скачать обновление.\n\n"
            f"Файл:\n{asset.name}\n"
            f"URL:\n{asset.download_url}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}",
            code="DOWNLOAD_FAILED",
        ) from exc


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
    checked = "\n".join(f"- {p}" for p in candidates)
    raise UpdateError(
        "Не найден updater.exe рядом с программой.\n\n"
        "Проверенные пути:\n"
        f"{checked}\n\n"
        "Что сделать:\n"
        "1. Проверьте, что app-модуль содержит updater.exe.\n"
        "2. Если запускаете из исходников — должен быть updater.py.\n"
        "3. Пересоберите portable-релиз.",
        code="UPDATER_NOT_FOUND",
    )


def start_updater(package_path: Path, restart: bool = True) -> None:
    root = app_dir()
    updater = updater_exe_path()
    app_exe = Path(sys.executable).resolve() if getattr(sys, "frozen", False) else root / "app.py"

    if not package_path.exists():
        raise UpdateError(
            "Файл обновления не найден перед запуском updater.\n\n"
            f"Путь:\n{package_path}",
            code="PACKAGE_NOT_FOUND_BEFORE_APPLY",
        )

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
    except PermissionError as exc:
        raise UpdateError(
            "Windows не разрешил запустить updater.\n\n"
            f"Updater:\n{updater}\n\n"
            "Что проверить:\n"
            "1. Не заблокировал ли файл антивирус.\n"
            "2. Есть ли права на запуск файла.\n"
            "3. Не лежит ли программа в защищённой системной папке.",
            code="UPDATER_START_PERMISSION_DENIED",
        ) from exc
    except Exception as exc:
        raise UpdateError(
            "Не удалось запустить updater.\n\n"
            f"Updater:\n{updater}\n\n"
            f"Техническая ошибка:\n{type(exc).__name__}: {exc}",
            code="UPDATER_START_FAILED",
        ) from exc


def check_app_update() -> tuple[bool, str, ReleaseInfo]:
    release = fetch_latest_release()
    current = current_version()
    latest = normalize_version(release.tag_name)
    return is_newer(latest, current), current, release
