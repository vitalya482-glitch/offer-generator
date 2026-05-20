from __future__ import annotations

from pathlib import Path

EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
WORD_EXTENSIONS = {".docx"}


def scan_project_files(project_dir: Path) -> dict[str, list[Path]]:
    project_dir = Path(project_dir)
    if not project_dir.exists():
        return {"excel": [], "word": [], "pdf_dirs": []}

    excel: list[Path] = []
    word: list[Path] = []
    pdf_dirs: set[Path] = set()

    for path in project_dir.rglob("*"):
        if not path.is_file():
            continue
        if any(part.startswith(".") for part in path.relative_to(project_dir).parts):
            continue
        suffix = path.suffix.lower()
        if suffix in EXCEL_EXTENSIONS and not path.name.startswith("~$"):
            excel.append(path)
        elif suffix in WORD_EXTENSIONS and not path.name.startswith("~$"):
            word.append(path)
        elif suffix == ".pdf":
            pdf_dirs.add(path.parent)

    return {
        "excel": sorted(excel, key=lambda p: p.name.lower()),
        "word": sorted(word, key=lambda p: p.name.lower()),
        "pdf_dirs": sorted(pdf_dirs, key=lambda p: str(p).lower()),
    }
