from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path


def wait_for_process(pid: int, timeout_sec: int = 60) -> None:
    if pid <= 0:
        return
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        if not is_process_running(pid):
            return
        time.sleep(0.5)


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
                    target.rename(backup)
            shutil.copy2(item, target)


def apply_update(package: Path, app_dir: Path) -> None:
    if not package.exists():
        raise FileNotFoundError(f"Update package not found: {package}")
    if not zipfile.is_zipfile(package):
        raise ValueError(f"Update package is not a ZIP file: {package}")

    running_updater = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve()
    temp_dir = app_dir / "updates" / "_apply_temp"
    if temp_dir.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(package, "r") as zf:
            zf.extractall(temp_dir)
        payload_root = find_payload_root(temp_dir)
        copy_tree_merge(payload_root, app_dir, running_updater)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="SAM Offer Generator updater")
    parser.add_argument("--package", required=True, help="ZIP package to apply")
    parser.add_argument("--app-dir", required=True, help="Application folder")
    parser.add_argument("--pid", type=int, default=0, help="PID of running main app")
    parser.add_argument("--restart", default="", help="Path to executable to start after update")
    args = parser.parse_args(argv)

    package = Path(args.package).resolve()
    app_dir = Path(args.app_dir).resolve()

    wait_for_process(args.pid)
    time.sleep(0.7)
    apply_update(package, app_dir)

    if args.restart:
        restart = Path(args.restart)
        if restart.exists():
            subprocess.Popen([str(restart)], cwd=str(app_dir), close_fds=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
