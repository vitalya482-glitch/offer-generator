from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
import zipfile
from pathlib import Path


def show_error(title: str, message: str) -> None:
    """Show a simple Windows error dialog; fallback to console on other systems."""
    try:
        if os.name == "nt":
            ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
            return
    except Exception:
        pass
    print(f"{title}\n{message}", file=sys.stderr)


def write_log(app_dir: Path, message: str) -> None:
    try:
        log_dir = app_dir / "updates"
        log_dir.mkdir(parents=True, exist_ok=True)
        (log_dir / "update_error.log").write_text(message, encoding="utf-8")
    except Exception:
        pass


def wait_for_process(pid: int, timeout_sec: int = 60) -> None:
    if pid <= 0:
        return
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if not is_process_running(pid):
            return
        time.sleep(0.5)
    raise TimeoutError(
        "Основная программа не закрылась за 60 секунд.\n\n"
        "Что сделать:\n"
        "1. Закройте SAM Offer Generator вручную.\n"
        "2. Запустите обновление ещё раз."
    )


def is_process_running(pid: int) -> bool:
    if os.name == "nt":
        try:
            result = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            return str(pid) in result.stdout
        except Exception:
            return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def find_payload_root(extract_dir: Path) -> Path:
    candidate = extract_dir / "SAM-Offer-Generator"
    if candidate.exists() and candidate.is_dir():
        return candidate
    children = [p for p in extract_dir.iterdir() if p.is_dir()]
    if len(children) == 1:
        return children[0]
    return extract_dir


def copy_tree_merge(src: Path, dst: Path, running_updater: Path) -> None:
    if not src.exists():
        raise FileNotFoundError(
            "В архиве обновления не найдена папка с файлами приложения.\n\n"
            f"Ожидали:\n{src}"
        )
    dst.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        target = dst / item.name

        # Windows cannot overwrite the currently running updater.exe.
        # Skip it; update it only through a full manual reinstall if needed.
        try:
            if target.resolve() == running_updater.resolve():
                continue
        except Exception:
            pass

        if item.is_dir():
            copy_tree_merge(item, target, running_updater)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                try:
                    target.unlink()
                except PermissionError:
                    backup = target.with_suffix(target.suffix + ".old")
                    try:
                        backup.unlink(missing_ok=True)
                    except Exception:
                        pass
                    try:
                        target.rename(backup)
                    except PermissionError as exc:
                        raise PermissionError(
                            "Не удалось заменить файл, потому что он занят или нет прав доступа.\n\n"
                            f"Файл:\n{target}\n\n"
                            "Что сделать:\n"
                            "1. Закройте SAM Offer Generator.\n"
                            "2. Закройте Excel/Word, если они открыли файлы из папки программы.\n"
                            "3. Проверьте, что программа не лежит в Program Files."
                        ) from exc
            shutil.copy2(item, target)


def apply_update(package: Path, app_dir: Path) -> None:
    if not package.exists():
        raise FileNotFoundError(
            "Файл обновления не найден.\n\n"
            f"Путь:\n{package}"
        )
    if not zipfile.is_zipfile(package):
        raise ValueError(
            "Файл обновления не является ZIP-архивом или повреждён.\n\n"
            f"Файл:\n{package}\n\n"
            "Что сделать:\n"
            "Скачайте обновление заново."
        )
    if not app_dir.exists():
        raise FileNotFoundError(
            "Папка приложения не найдена.\n\n"
            f"Путь:\n{app_dir}"
        )

    running_updater = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve()
    temp_dir = app_dir / "updates" / "_apply_temp"
    if temp_dir.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(package, "r") as zf:
            bad_file = zf.testzip()
            if bad_file:
                raise ValueError(
                    "ZIP-архив обновления повреждён.\n\n"
                    f"Повреждённый файл внутри архива:\n{bad_file}\n\n"
                    "Что сделать:\n"
                    "Скачайте обновление заново."
                )
            zf.extractall(temp_dir)
        payload_root = find_payload_root(temp_dir)
        copy_tree_merge(payload_root, app_dir, running_updater)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def write_update_state(app_dir: Path, asset_states: list[dict]) -> None:
    if not asset_states:
        return
    config_dir = app_dir / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    path = config_dir / "update_state.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    except Exception:
        data = {}
    if not isinstance(data, dict):
        data = {}
    assets = data.get("assets")
    if not isinstance(assets, dict):
        assets = {}
    for item in asset_states:
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        assets[name] = {
            "kind": str(item.get("kind") or ""),
            "sha256": str(item.get("sha256") or "").lower(),
            "installed_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        }
    data["assets"] = assets
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def read_asset_states(raw_states: list[str], state_file: str) -> list[dict]:
    asset_states: list[dict] = []
    if state_file:
        try:
            data = json.loads(Path(state_file).read_text(encoding="utf-8"))
            items = data.get("assets") if isinstance(data, dict) else None
            if isinstance(items, list):
                asset_states.extend(item for item in items if isinstance(item, dict))
        except Exception:
            pass
    for raw_state in raw_states:
        try:
            item = json.loads(raw_state)
            if isinstance(item, dict):
                asset_states.append(item)
        except Exception:
            pass
    return asset_states


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="SAM Offer Generator updater")
    parser.add_argument("--package", action="append", required=True, help="ZIP package to apply")
    parser.add_argument("--asset-state", action="append", default=[], help="Installed asset metadata as JSON")
    parser.add_argument("--state-file", default="", help="Installed asset metadata JSON file")
    parser.add_argument("--app-dir", required=True, help="Application folder")
    parser.add_argument("--pid", type=int, default=0, help="PID of running main app")
    parser.add_argument("--restart", default="", help="Path to executable to start after update")
    args = parser.parse_args(argv)

    packages = [Path(item).resolve() for item in args.package]
    app_dir = Path(args.app_dir).resolve()
    asset_states = read_asset_states(args.asset_state, args.state_file)

    try:
        wait_for_process(args.pid)
        time.sleep(0.7)
        for package in packages:
            apply_update(package, app_dir)
        write_update_state(app_dir, asset_states)

        if args.restart:
            restart = Path(args.restart)
            if restart.exists():
                subprocess.Popen([str(restart)], cwd=str(app_dir), close_fds=True)
            else:
                show_error(
                    "Обновление установлено",
                    "Обновление установлено, но не удалось автоматически запустить программу.\n\n"
                    f"Файл запуска не найден:\n{restart}\n\n"
                    "Запустите SAM-Offer-Generator.exe вручную.",
                )
        return 0
    except Exception as exc:
        message = (
            "Не удалось применить обновление.\n\n"
            f"Причина:\n{type(exc).__name__}: {exc}\n\n"
            f"Архивы обновления:\n" + "\n".join(str(p) for p in packages) + "\n\n"
            f"Папка приложения:\n{app_dir}\n\n"
            "Подробности записаны в updates/update_error.log"
        )
        details = message + "\n\n" + traceback.format_exc()
        write_log(app_dir, details)
        show_error("Ошибка обновления", message)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
