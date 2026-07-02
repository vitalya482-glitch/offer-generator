from __future__ import annotations

from pathlib import Path
from typing import Any

from core.models import OfferContext

BRAND_NAME = "DC Eltek"
PROJECTS_MARKER = "02_Projects"


def extract_client_from_project_path(path_text: str) -> str:
    """Возвращает имя клиента из пути ...\\02_Projects\\КЛИЕНТ\\..."""
    if not path_text:
        return ""

    parts = [part for part in path_text.replace("/", "\\").split("\\") if part]
    for index, part in enumerate(parts):
        if part.lower() == PROJECTS_MARKER.lower() and index + 1 < len(parts):
            return parts[index + 1].strip()

    return ""


def make_offer(context: OfferContext | dict[str, Any]) -> Path:
    raise NotImplementedError(
        "Логика генерации КП DC Eltek пока не подключена. "
        "На текущем этапе добавлена вкладка выбора папки, клиента, Excel, листа и шаблона."
    )


def preview(context: OfferContext | dict[str, Any]) -> str:
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

    return "\n".join(
        [
            "Направление: DC Eltek",
            f"Папка проекта: {project_dir or 'не выбрана'}",
            f"Клиент: {client or 'не указан'}",
            f"Расчёт Excel: {calc_path or 'не выбран'}",
            f"Лист для КП: {sheet_name or 'не выбран'}",
            f"Шаблон КП: {template_path or 'не выбран'}",
        ]
    )
