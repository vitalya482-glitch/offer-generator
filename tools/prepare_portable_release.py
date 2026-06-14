from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_DIST_DIR = ROOT_DIR / "dist" / "SAM-Offer-Generator"

COPY_TO_ROOT = [
    "README.md",
    "GITHUB_RELEASES.md",
    "MODULES_MANIFEST.json",
    "config.example.json",
    "requirements.txt",
]
SOURCE_MODULE_DIRS = ["assets", "brands", "core", "gui", "config"]
OPTIONAL_ROOT_DIRS = ["assets", "prices"]
EXCLUDED_DIR_NAMES = {"__pycache__", ".git", ".pytest_cache", ".mypy_cache"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


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
    text = """SAM Offer Generator - portable Windows release

Run:
  1. Extract the whole folder from SAM-Offer-Generator-windows-portable.zip.
  2. Start SAM-Offer-Generator.exe or run_gui.cmd.
  3. Keep _internal, config and other folders next to the EXE.

Folder layout:
  SAM-Offer-Generator.exe  - launcher
  _internal/               - PyInstaller runtime files and Python dependencies
  config/                  - editable JSON configuration files
  modules/source/          - source copies of project modules for review/reuse
  prices/                  - optional reference price files, when present

Important:
  Do not move only the EXE to another folder. This is a one-dir build, so the
  EXE depends on the files and folders shipped with it.

For developers:
  GitHub Actions also publishes separate source module ZIPs. Use them when you
  need to download or replace only one part of the project.
"""
    (dist_dir / "README_RELEASE.txt").write_text(text, encoding="utf-8")


def write_run_cmd(dist_dir: Path) -> None:
    text = '@echo off\r\ncd /d "%~dp0"\r\nstart "" "%~dp0SAM-Offer-Generator.exe" --gui\r\n'
    (dist_dir / "run_gui.cmd").write_text(text, encoding="utf-8")


def write_release_info(dist_dir: Path) -> None:
    info = {
        "project": "SAM Offer Generator",
        "layout": "pyinstaller-onedir-portable",
        "generated_at_utc": dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat(),
        "github_ref": os.environ.get("GITHUB_REF_NAME", ""),
        "github_sha": os.environ.get("GITHUB_SHA", ""),
        "notes": "Keep the full folder together; this is not a one-file executable release.",
    }
    (dist_dir / "release_info.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Add editable files and source modules to a PyInstaller one-dir build.")
    parser.add_argument("--dist", type=Path, default=DEFAULT_DIST_DIR)
    args = parser.parse_args()

    dist_dir = args.dist.resolve()
    if not dist_dir.exists():
        raise SystemExit(f"Build folder does not exist: {dist_dir}")

    for relative_path in COPY_TO_ROOT:
        copy_file_if_exists(relative_path, dist_dir)

    copy_tree_if_exists("config", dist_dir / "config")

    source_modules_dir = dist_dir / "modules" / "source"
    source_modules_dir.mkdir(parents=True, exist_ok=True)
    for module_dir in SOURCE_MODULE_DIRS:
        copy_tree_if_exists(module_dir, source_modules_dir / module_dir)

    for optional_dir in OPTIONAL_ROOT_DIRS:
        copy_tree_if_exists(optional_dir, dist_dir / optional_dir)

    write_release_readme(dist_dir)
    write_run_cmd(dist_dir)
    write_release_info(dist_dir)

    print(f"Prepared portable release folder: {dist_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
