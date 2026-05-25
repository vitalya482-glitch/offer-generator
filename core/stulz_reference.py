from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT_DIR = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT_DIR / "config"
STULZ_OPTIONS_PATH = CONFIG_DIR / "stulz_options.json"
STULZ_WINPLAN_PATH = CONFIG_DIR / "stulz_winplan.json"
STULZ_MISSING_OPTIONS_PATH = CONFIG_DIR / "stulz_missing_options.json"

XLSX_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def load_json_table(path: Path, default: list[dict[str, str]] | None = None) -> list[dict[str, str]]:
    if not path.exists():
        return list(default or [])
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return list(default or [])
    if not isinstance(data, list):
        return list(default or [])
    return [row for row in data if isinstance(row, dict)]


def save_json_table(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


def load_stulz_options() -> list[dict[str, str]]:
    return load_json_table(STULZ_OPTIONS_PATH)


def save_stulz_options(rows: list[dict[str, str]]) -> None:
    save_json_table(STULZ_OPTIONS_PATH, rows)


def load_stulz_winplan() -> list[dict[str, str]]:
    return load_json_table(STULZ_WINPLAN_PATH)


def save_stulz_winplan(rows: list[dict[str, str]]) -> None:
    save_json_table(STULZ_WINPLAN_PATH, rows)


def load_missing_options() -> list[dict[str, str]]:
    return load_json_table(STULZ_MISSING_OPTIONS_PATH)


def save_missing_options(rows: list[dict[str, str]]) -> None:
    save_json_table(STULZ_MISSING_OPTIONS_PATH, rows)


def _column_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)([0-9]+)", cell_ref)
    if not match:
        return 1
    col = 0
    for char in match.group(1):
        col = col * 26 + ord(char) - 64
    return col


def _shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("a:si", XLSX_NS):
        parts = []
        for text_node in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            parts.append(text_node.text or "")
        strings.append("".join(parts).replace("_x000D_", "\n"))
    return strings


def _sheet_xml_path(zf: ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    for sheet in wb.find("a:sheets", XLSX_NS) or []:
        if sheet.attrib.get("name") == sheet_name:
            rid = sheet.attrib[REL_ID]
            return "xl/" + rel_map[rid]
    raise ValueError(f"Лист {sheet_name!r} не найден")


def read_xlsm_sheet(path: Path, sheet_name: str) -> list[list[str]]:
    """Read simple cell values from xlsx/xlsm without Excel dependency."""
    with ZipFile(path) as zf:
        shared = _shared_strings(zf)
        sheet_path = _sheet_xml_path(zf, sheet_name)
        sheet_root = ET.fromstring(zf.read(sheet_path))
        rows: list[list[str]] = []
        for row in sheet_root.findall(".//a:row", XLSX_NS):
            values: list[str] = []
            for cell in row.findall("a:c", XLSX_NS):
                col = _column_index(cell.attrib.get("r", "A1"))
                while len(values) < col - 1:
                    values.append("")
                value_node = cell.find("a:v", XLSX_NS)
                value = ""
                if value_node is not None:
                    value = value_node.text or ""
                    if cell.attrib.get("t") == "s":
                        value = shared[int(value)]
                values.append(str(value).strip())
            while values and values[-1] == "":
                values.pop()
            if values and any(str(item).strip() for item in values):
                rows.append(values)
        return rows


def import_options_from_xlsm(path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for row in read_xlsm_sheet(path, "Options"):
        if len(row) >= 3 and row[0].strip() and row[1].strip():
            rows.append({
                "code": row[0].strip(),
                "source_name": row[1].strip(),
                "ru_description": row[2].strip(),
            })
    return rows


def import_winplan_from_xlsm(path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for row in read_xlsm_sheet(path, "WinPlan"):
        if len(row) >= 2 and (row[0].strip() or row[1].strip()):
            rows.append({
                "source_name": row[0].strip(),
                "ru_name": row[1].strip(),
                "section": row[2].strip() if len(row) > 2 else "",
                "unit": row[3].strip() if len(row) > 3 else "",
            })
    return rows
