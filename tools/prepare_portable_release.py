from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from tools.update_layout_policy import (  # noqa: E402
    APP_CONFIG_DIR,
    APP_OPTIONAL_ROOT_DIRS,
    APP_ROOT_FILES,
    APP_SOURCE_MODULE_DIRS,
    EXCLUDED_DIR_NAMES,
    EXCLUDED_SUFFIXES,
    RELEASE_LAYOUT_NAME,
    RELEASE_LAYOUT_NOTE,
    validate_update_layout_policy,
)


DEFAULT_DIST_DIR = ROOT_DIR / "dist" / "SAM-Offer-Generator"


def ignore_patterns(directory: str, names: list[str]) -> set[str]:
    ignored: set[str] = set()
    for name in names:
        path = Path(directory) / name
        if name in EXCLUDED_DIR_NAMES:
            ignored.add(name)
        elif path.is_file() and path.suffix.lower() in EXCLUDED_SUFFIXES:
            ignored.add(name)
    return ignored


def copy_file_if_exists(relative_path: str, target_dir: Path) -> None:
    source = ROOT_DIR / relative_path
    if source.exists() and source.is_file():
        shutil.copy2(source, target_dir / source.name)


def copy_tree_if_exists(relative_path: str, target: Path) -> None:
    source = ROOT_DIR / relative_path
    if source.exists() and source.is_dir():
        if target.exists():
            shutil.rmtree(target)
        shutil.copytree(source, target, ignore=ignore_patterns)


def write_release_readme(dist_dir: Path) -> None:
    app_dirs = ", ".join((APP_CONFIG_DIR,) + APP_SOURCE_MODULE_DIRS + APP_OPTIONAL_ROOT_DIRS)
    text = f"""SAM Offer Generator - portable Windows release

Run / first install:
  1. Extract SAM-Offer-Generator-Runtime-Win64.zip.
  2. Extract SAM-Offer-Generator-App-No-Runtime.zip into the same folder with replacement.
  3. Start SAM-Offer-Generator.exe or run_gui.cmd.
  4. Keep _internal, config, brands, core, gui and other folders next to the EXE.

Folder layout:
  SAM-Offer-Generator.exe  - launcher
  _internal/               - stable PyInstaller runtime and third-party dependencies
  brands/                  - editable brand logic, updated by the small app module
  core/                    - editable common application logic
  gui/                     - editable application pages and UI logic
  config/                  - editable JSON configuration files
  prices/                  - optional reference price files, when present
  templates/               - Excel/Word templates used by brand modules

App-No-Runtime package contains:
  {', '.join(APP_ROOT_FILES)}
  {app_dirs}

Important:
  Do not move only the EXE to another folder. This is a one-dir build, so the
  EXE depends on the files and folders shipped with it.

Release modules:
  GitHub Release publishes three ZIP files:
  - SAM-Offer-Generator-Full-Win64.zip
  - SAM-Offer-Generator-Runtime-Win64.zip
  - SAM-Offer-Generator-App-No-Runtime.zip

Update model:
  Normal Python/UI fixes are delivered through App-No-Runtime.zip. Runtime is
  downloaded only when Python, dependencies, PyInstaller runtime layout or the
  manual runtime epoch changes.

Single source of truth:
  tools/update_layout_policy.py
"""
    (dist_dir / "README_RELEASE.txt").write_text(text, encoding="utf-8")


def write_run_cmd(dist_dir: Path) -> None:
    text = '@echo off\r\ncd /d "%~dp0"\r\nstart "" "%~dp0SAM-Offer-Generator.exe" --gui\r\n'
    (dist_dir / "run_gui.cmd").write_text(text, encoding="utf-8")


def write_release_info(dist_dir: Path) -> None:
    info = {
        "project": "SAM Offer Generator",
        "layout": RELEASE_LAYOUT_NAME,
        "generated_at_utc": dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat(),
        "github_ref": os.environ.get("GITHUB_REF_NAME", ""),
        "github_sha": os.environ.get("GITHUB_SHA", ""),
        "notes": RELEASE_LAYOUT_NOTE,
        "update_layout_policy": "tools/update_layout_policy.py",
    }
    (dist_dir / "release_info.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Add editable files and source modules to a PyInstaller one-dir build.")
    parser.add_argument("--dist", type=Path, default=DEFAULT_DIST_DIR)
    args = parser.parse_args()

    validate_update_layout_policy()
    dist_dir = args.dist.resolve()
    if not dist_dir.exists():
        raise SystemExit(f"Build folder does not exist: {dist_dir}")

    for relative_path in APP_ROOT_FILES:
        copy_file_if_exists(relative_path, dist_dir)

    copy_tree_if_exists(APP_CONFIG_DIR, dist_dir / APP_CONFIG_DIR)

    for source_dir in APP_SOURCE_MODULE_DIRS:
        copy_tree_if_exists(source_dir, dist_dir / source_dir)

    for optional_dir in APP_OPTIONAL_ROOT_DIRS:
        copy_tree_if_exists(optional_dir, dist_dir / optional_dir)

    write_release_readme(dist_dir)
    write_run_cmd(dist_dir)
    write_release_info(dist_dir)

    print(f"Prepared portable release folder: {dist_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
