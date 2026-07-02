from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Any

try:
    from core.models import OfferContext
except Exception:  # pragma: no cover
    OfferContext = Any  # type: ignore

BRAND_NAME = "DC Eltek"
PROJECTS_MARKER = "02_Projects"

_XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")



HEADER_ALIASES: dict[str, list[tuple[str, int]]] = {
    # Наименование/модель обычно находится в строке с заголовком Model.
    "name": [
        ("model", 100),
        ("модель", 100),
        ("наименование", 95),
        ("equipment", 90),
        ("оборудование", 90),
        ("description", 80),
        ("описание", 80),
        ("item", 70),
        ("позиция", 70),
        ("part number", 55),
    ],
    "quantity": [
        ("quantity", 100),
        ("qty", 100),
        ("количество", 100),
        ("кол-во", 100),
        ("кол во", 100),
        ("шт", 75),
        ("pcs", 75),
        ("count", 60),
    ],
    # Важно: в калькуляциях бывает несколько цен. Для КП нужен итоговый прайс за единицу.
    "unit_price": [
        ("total per unit", 120),
        ("итого за единицу", 120),
        ("цена за единицу", 115),
        ("цена за ед", 115),
        ("стоимость за единицу", 110),
        ("стоимость за ед", 110),
        ("unit price", 100),
        ("price per unit", 95),
        ("price", 45),
        ("цена", 45),
    ],
    # Предпочитаем итоговую строку TOTAL/ИТОГО, а не промежуточные Total per quantity.
    "total": [
        ("total", 130),
        ("итого", 130),
        ("общая сумма", 120),
        ("сумма итого", 120),
        ("total with vat", 115),
        ("total per quantity", 90),
        ("сумма", 80),
        ("amount", 70),
        ("стоимость", 60),
    ],
    "vat_percent": [
        ("vat", 100),
        ("vat %", 100),
        ("vat, %", 100),
        ("ндс", 100),
        ("ндс %", 100),
        ("ндс, %", 100),
    ],
}

SERVICE_NAME_TOKENS = {
    "%",
    "total",
    "итого",
    "сумма",
    "quantity",
    "qty",
    "количество",
    "кол-во",
    "price",
    "цена",
    "vat",
    "ндс",
    "ddp",
    "kzt",
    "eur",
    "euro",
    "rate",
    "margin",
}


def extract_client_from_project_path(path_text: str) -> str:
    """Возвращает имя клиента из пути ...\\02_Projects\\КЛИЕНТ\\..."""
    if not path_text:
        return ""

    parts = [part for part in path_text.replace("/", "\\").split("\\") if part]
    for index, part in enumerate(parts):
        if part.lower() == PROJECTS_MARKER.lower() and index + 1 < len(parts):
            return parts[index + 1].strip()

    return ""


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip().lower()
    text = text.replace("ё", "е")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[,:;()\[\]{}]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text)


def _col_letters_to_index(letters: str) -> int:
    result = 0
    for char in letters:
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _cell_ref_to_row_col(ref: str) -> tuple[int, int]:
    match = _CELL_RE.match(ref.upper())
    if not match:
        raise ValueError(f"Некорректная ссылка Excel-ячейки: {ref}")
    col_letters, row_text = match.groups()
    return int(row_text), _col_letters_to_index(col_letters)


def _to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace("\xa0", " ")
    if not text:
        return None

    # Убираем валюты и пробелы-разделители, поддерживаем десятичную запятую.
    text = re.sub(r"[^0-9,\.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None

    if "," in text and "." in text:
        # 1 234,56 или 1,234.56 — оставляем последний разделитель десятичным.
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def _format_money(value: float | int | None) -> str:
    if value is None:
        return "0"
    rounded = round(float(value))
    return f"{rounded:,.0f}".replace(",", " ")


def _format_qty(value: float | int | None) -> str:
    if value is None:
        return "0"
    value = float(value)
    if value.is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        raw_xml = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ET.fromstring(raw_xml)
    result: list[str] = []
    for item in root.findall(f"{{{_XLSX_MAIN_NS}}}si"):
        chunks: list[str] = []
        for text_node in item.iter(f"{{{_XLSX_MAIN_NS}}}t"):
            chunks.append(text_node.text or "")
        result.append("".join(chunks))
    return result


def _read_workbook_sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rel_targets: dict[str, str] = {}
    for rel in rels_root.findall(f"{{{_PACKAGE_REL_NS}}}Relationship"):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if not rel_id or not target:
            continue
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        rel_targets[rel_id] = target

    result: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{_XLSX_MAIN_NS}}}sheet"):
        name = sheet.attrib.get("name", "").strip()
        rel_id = sheet.attrib.get(f"{{{_XLSX_REL_NS}}}id")
        if name and rel_id and rel_id in rel_targets:
            result[name] = rel_targets[rel_id]
    return result


def _read_sheet_matrix(calc_path: str | Path, sheet_name: str) -> tuple[dict[int, dict[int, Any]], int, int]:
    path = Path(calc_path)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("DC Eltek пока читает только .xlsx/.xlsm. Старый .xls нужно сохранить как .xlsx.")
    if not path.exists():
        raise FileNotFoundError(f"Excel-файл не найден: {path}")

    with zipfile.ZipFile(path) as archive:
        shared_strings = _read_shared_strings(archive)
        sheet_paths = _read_workbook_sheet_paths(archive)
        if sheet_name not in sheet_paths:
            available = ", ".join(sheet_paths.keys())
            raise ValueError(f"Лист '{sheet_name}' не найден. Доступные листы: {available}")

        root = ET.fromstring(archive.read(sheet_paths[sheet_name]))

    matrix: dict[int, dict[int, Any]] = {}
    max_row = 0
    max_col = 0

    for cell in root.findall(f".//{{{_XLSX_MAIN_NS}}}c"):
        ref = cell.attrib.get("r", "")
        if not ref:
            continue
        try:
            row, col = _cell_ref_to_row_col(ref)
        except ValueError:
            continue

        cell_type = cell.attrib.get("t")
        value: Any = None

        if cell_type == "inlineStr":
            chunks = [node.text or "" for node in cell.iter(f"{{{_XLSX_MAIN_NS}}}t")]
            value = "".join(chunks)
        else:
            value_node = cell.find(f"{{{_XLSX_MAIN_NS}}}v")
            if value_node is None or value_node.text is None:
                continue
            raw_value = value_node.text
            if cell_type == "s":
                index = int(raw_value)
                value = shared_strings[index] if 0 <= index < len(shared_strings) else raw_value
            elif cell_type == "str":
                value = raw_value
            else:
                try:
                    number = float(raw_value)
                    value = int(number) if number.is_integer() else number
                except ValueError:
                    value = raw_value

        if value is None or value == "":
            continue
        matrix.setdefault(row, {})[col] = value
        max_row = max(max_row, row)
        max_col = max(max_col, col)

    return matrix, max_row, max_col


def _cell_text(matrix: dict[int, dict[int, Any]], row: int, col: int) -> str:
    return _clean_text(matrix.get(row, {}).get(col))


def _cell_number(matrix: dict[int, dict[int, Any]], row: int | None, col: int) -> float | None:
    if row is None:
        return None
    return _to_number(matrix.get(row, {}).get(col))


def _alias_matches(text: str, alias: str) -> bool:
    normalized_alias = _normalize_text(alias)
    if not text or not normalized_alias:
        return False
    if text == normalized_alias:
        return True
    # Для коротких алиасов вроде qty/vat требуем границы слова.
    if len(normalized_alias) <= 4 and re.search(rf"(^|\s){re.escape(normalized_alias)}($|\s|%)", text):
        return True
    return normalized_alias in text


def _detect_best_row(
    matrix: dict[int, dict[int, Any]],
    field: str,
    *,
    min_row: int | None = None,
    max_row: int | None = None,
    prefer_bottom: bool = False,
) -> int | None:
    aliases = HEADER_ALIASES[field]
    candidates: list[tuple[int, int, int, str]] = []

    for row, cols in matrix.items():
        if min_row is not None and row < min_row:
            continue
        if max_row is not None and row > max_row:
            continue

        for col, value in cols.items():
            text = _normalize_text(value)
            if not text:
                continue
            for alias, base_score in aliases:
                if not _alias_matches(text, alias):
                    continue

                score = base_score
                alias_norm = _normalize_text(alias)
                if text == alias_norm:
                    score += 35
                # Заголовки чаще всего в левой части листа.
                score += max(0, 20 - col)
                # Для итоговых строк полезно выбирать нижний блок, для model/quantity — верхний.
                score += row if prefer_bottom else max(0, 200 - row)
                candidates.append((score, row, col, text))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    return candidates[0][1]


def detect_offer_layout(calc_path: str | Path, sheet_name: str) -> dict[str, int | None]:
    matrix, _max_row, _max_col = _read_sheet_matrix(calc_path, sheet_name)

    name_row = _detect_best_row(matrix, "name", prefer_bottom=False)
    quantity_row = _detect_best_row(matrix, "quantity", min_row=name_row, prefer_bottom=False) if name_row else _detect_best_row(matrix, "quantity")
    total_row = _detect_best_row(matrix, "total", min_row=quantity_row, prefer_bottom=True) if quantity_row else _detect_best_row(matrix, "total", prefer_bottom=True)

    # Итоговую цену за единицу ищем ближе к итоговому блоку.
    unit_max = total_row - 1 if total_row else None
    unit_price_row = _detect_best_row(matrix, "unit_price", min_row=quantity_row, max_row=unit_max, prefer_bottom=True)
    if unit_price_row is None:
        unit_price_row = _detect_best_row(matrix, "unit_price", prefer_bottom=True)

    vat_row = None
    if total_row:
        vat_row = _detect_best_row(matrix, "vat_percent", min_row=quantity_row, max_row=total_row - 1, prefer_bottom=True)
    if vat_row is None:
        vat_row = _detect_best_row(matrix, "vat_percent", prefer_bottom=True)

    return {
        "name_row": name_row,
        "quantity_row": quantity_row,
        "unit_price_row": unit_price_row,
        "total_row": total_row,
        "vat_percent_row": vat_row,
    }


def _is_service_name(name: str) -> bool:
    text = _normalize_text(name)
    if not text:
        return True
    if text in SERVICE_NAME_TOKENS:
        return True
    if text.replace(".", "", 1).isdigit():
        return True
    if len(text) <= 1:
        return True
    return False


def read_dc_eltek_offer_items(calc_path: str | Path, sheet_name: str) -> dict[str, Any]:
    """Читает выбранный лист DC Eltek и возвращает позиции + итоги.

    Парсер не привязан к конкретным строкам/колонкам. Он ищет строки по заголовкам
    Model/Quantity/Total per unit/TOTAL/VAT и затем проходит все колонки слева направо.
    """
    matrix, _max_row, max_col = _read_sheet_matrix(calc_path, sheet_name)
    layout = detect_offer_layout(calc_path, sheet_name)

    required = ["name_row", "quantity_row", "unit_price_row", "total_row"]
    missing = [key for key in required if not layout.get(key)]
    if missing:
        raise ValueError(
            "Не удалось определить структуру листа DC Eltek. Не найдены строки: "
            + ", ".join(missing)
        )

    name_row = int(layout["name_row"] or 0)
    quantity_row = int(layout["quantity_row"] or 0)
    unit_price_row = int(layout["unit_price_row"] or 0)
    total_row = int(layout["total_row"] or 0)
    vat_percent_row = layout.get("vat_percent_row")

    items: list[dict[str, Any]] = []

    for col in range(1, max_col + 1):
        name = _cell_text(matrix, name_row, col)
        if _is_service_name(name):
            continue

        quantity = _cell_number(matrix, quantity_row, col)
        unit_price = _cell_number(matrix, unit_price_row, col)
        total = _cell_number(matrix, total_row, col)
        vat_percent = _cell_number(matrix, int(vat_percent_row), col) if vat_percent_row else 0.0

        if quantity is None or quantity <= 0:
            continue
        if (unit_price is None or unit_price <= 0) and (total is None or total <= 0):
            continue

        if (total is None or total <= 0) and unit_price is not None:
            total = unit_price * quantity
        if (unit_price is None or unit_price <= 0) and total is not None and quantity:
            unit_price = total / quantity

        vat_percent = vat_percent or 0.0
        total_with_vat = float(total or 0.0)
        if vat_percent > 0:
            total_without_vat = total_with_vat / (1 + vat_percent / 100)
            vat_amount = total_with_vat - total_without_vat
        else:
            total_without_vat = total_with_vat
            vat_amount = 0.0

        items.append(
            {
                "number": len(items) + 1,
                "name": name,
                "quantity": quantity,
                "unit_price": float(unit_price or 0.0),
                "total_without_vat": total_without_vat,
                "vat_percent": vat_percent,
                "vat_amount": vat_amount,
                "total": total_with_vat,
                "source_column": col,
            }
        )

    summary = summarize_items(items)
    return {
        "layout": layout,
        "items": items,
        "summary": summary,
    }


def summarize_items(items: list[dict[str, Any]]) -> dict[str, Any]:
    quantity_total = sum(float(item.get("quantity") or 0) for item in items)
    total_without_vat = sum(float(item.get("total_without_vat") or 0) for item in items)
    vat_amount = sum(float(item.get("vat_amount") or 0) for item in items)
    total = sum(float(item.get("total") or 0) for item in items)

    vat_rates = sorted({round(float(item.get("vat_percent") or 0), 4) for item in items})
    vat_label = ", ".join(f"{rate:g}%" for rate in vat_rates) if vat_rates else "0%"

    return {
        "positions_count": len(items),
        "quantity_total": quantity_total,
        "total_without_vat": total_without_vat,
        "vat_amount": vat_amount,
        "total": total,
        "vat_rates": vat_rates,
        "vat_label": vat_label,
    }


def make_offer(context: OfferContext | dict[str, Any]) -> Path:
    raise NotImplementedError(
        "Парсер DC Eltek уже подключен. Генерация Word-КП будет подключена следующим этапом после согласования тегов шаблона."
    )


def _context_values(context: OfferContext | dict[str, Any]) -> tuple[str, str, str, str, str]:
    if isinstance(context, dict):
        project_dir = str(context.get("project_dir", ""))
        client = str(context.get("client", "")) or extract_client_from_project_path(project_dir)
        calc_path = str(context.get("calc_path", ""))
        sheet_name = str(context.get("sheet_name", ""))
        template_path = str(context.get("template_path", ""))
    else:
        project_dir = str(getattr(context, "project_dir", ""))
        client = extract_client_from_project_path(project_dir)
        calc_path = str(getattr(context, "calc_path", ""))
        sheet_name = str(getattr(context, "sheet_name", ""))
        template_path = str(getattr(context, "template_path", ""))
    return project_dir, client, calc_path, sheet_name, template_path


def preview(context: OfferContext | dict[str, Any]) -> str:
    project_dir, client, calc_path, sheet_name, template_path = _context_values(context)

    lines = [
        "Направление: DC Eltek",
        f"Папка проекта: {project_dir or 'не выбрана'}",
        f"Клиент: {client or 'не указан'}",
        f"Расчёт Excel: {calc_path or 'не выбран'}",
        f"Лист для КП: {sheet_name or 'не выбран'}",
        f"Шаблон КП: {template_path or 'не выбран'}",
    ]

    if not calc_path or not sheet_name:
        lines.extend(["", "Выберите Excel calc и лист для КП — после этого появится предпросмотр позиций."])
        return "\n".join(lines)

    try:
        parsed = read_dc_eltek_offer_items(calc_path, sheet_name)
    except Exception as exc:
        lines.extend(["", f"Предпросмотр позиций не построен: {exc}"])
        return "\n".join(lines)

    layout = parsed["layout"]
    items = parsed["items"]
    summary = parsed["summary"]

    lines.extend(
        [
            "",
            "Найденные строки:",
            f"  модель/наименование: {layout.get('name_row')}",
            f"  количество: {layout.get('quantity_row')}",
            f"  цена за единицу: {layout.get('unit_price_row')}",
            f"  НДС: {layout.get('vat_percent_row') or 'не найдено / 0%'}",
            f"  итоговая сумма: {layout.get('total_row')}",
            "",
            "Позиции для КП:",
        ]
    )

    if not items:
        lines.append("  Позиции не найдены. Проверьте заголовки и выбранный лист.")
    else:
        for item in items[:25]:
            lines.append(
                "  {number}. {name} — {qty} шт × {unit_price} = {total}".format(
                    number=item["number"],
                    name=item["name"],
                    qty=_format_qty(item["quantity"]),
                    unit_price=_format_money(item["unit_price"]),
                    total=_format_money(item["total"]),
                )
            )
        if len(items) > 25:
            lines.append(f"  ... ещё {len(items) - 25} позиций")

    lines.extend(
        [
            "",
            "Итого:",
            f"  Количество позиций: {summary['positions_count']}",
            f"  Общее количество, шт: {_format_qty(summary['quantity_total'])}",
            f"  Сумма без НДС: {_format_money(summary['total_without_vat'])}",
            f"  НДС ({summary['vat_label']}): {_format_money(summary['vat_amount'])}",
            f"  Общая сумма с НДС: {_format_money(summary['total'])}",
        ]
    )

    return "\n".join(lines)
