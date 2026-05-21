from __future__ import annotations

import os
from pathlib import Path

EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
WORD_EXTENSIONS = {".docx"}

SKIP_DIRS = {
    ".git",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".venv",
    "venv",
    "env",
    "node_modules",
    "backup",
    "backups",
    "archive",
    "old",
    "temp",
    "tmp",
}

_SCAN_CACHE: dict[tuple[str, int], dict[str, list[Path]]] = {}


def _root_mtime_ns(project_dir: Path) -> int:
    try:
        return project_dir.stat().st_mtime_ns
    except OSError:
        return 0


def clear_scan_cache() -> None:
    _SCAN_CACHE.clear()


def scan_project_files(project_dir: Path, *, use_cache: bool = True) -> dict[str, list[Path]]:
    project_dir = Path(project_dir)
    if not project_dir.exists():
        return {"excel": [], "word": [], "pdf_dirs": []}

    cache_key = (str(project_dir.resolve()), _root_mtime_ns(project_dir))
    if use_cache and cache_key in _SCAN_CACHE:
        return {key: list(value) for key, value in _SCAN_CACHE[cache_key].items()}

    excel: list[Path] = []
    word: list[Path] = []
    pdf_dirs: set[Path] = set()

    for root, dirs, files in os.walk(project_dir):
        dirs[:] = [
            d for d in dirs
            if d not in SKIP_DIRS and not d.startswith(".")
        ]
        root_path = Path(root)
        for filename in files:
            if filename.startswith((".", "~$")):
                continue
            path = root_path / filename
            suffix = path.suffix.lower()
            if suffix in EXCEL_EXTENSIONS:
                excel.append(path)
            elif suffix in WORD_EXTENSIONS:
                word.append(path)
            elif suffix == ".pdf":
                pdf_dirs.add(path.parent)

    result = {
        "excel": sorted(excel, key=lambda p: p.name.lower()),
        "word": sorted(word, key=lambda p: p.name.lower()),
        "pdf_dirs": sorted(pdf_dirs, key=lambda p: str(p).lower()),
    }
    _SCAN_CACHE.clear()
    _SCAN_CACHE[cache_key] = {key: list(value) for key, value in result.items()}
    return result
