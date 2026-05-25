from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt


def to_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def replace_in_paragraph(paragraph, replacements: dict[str, Any]) -> None:
    """Replace tags while preserving Word formatting where possible."""
    if not paragraph.runs:
        return

    changed = False
    for run in paragraph.runs:
        original = run.text
        updated = original
        for key, value in replacements.items():
            updated = updated.replace(key, to_text(value))
        if updated != original:
            run.text = updated
            changed = True

    if changed:
        return

    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return

    updated_full = full_text
    for key, value in replacements.items():
        updated_full = updated_full.replace(key, to_text(value))

    if updated_full != full_text:
        paragraph.runs[0].text = updated_full
        for run in paragraph.runs[1:]:
            run.text = ""


def replace_tags(doc: Document, replacements: dict[str, Any]) -> None:
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, replacements)


def set_cell_keep_style(cell, value: Any) -> None:
    value = to_text(value)
    if cell.paragraphs:
        paragraph = cell.paragraphs[0]
    else:
        paragraph = cell.add_paragraph()

    if paragraph.runs:
        paragraph.runs[0].text = value
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(value)

    for extra_paragraph in cell.paragraphs[1:]:
        for run in extra_paragraph.runs:
            run.text = ""


def row_contains_tags(row, tags: list[str]) -> bool:
    row_text = "\n".join(cell.text for cell in row.cells)
    return any(tag in row_text for tag in tags)


def clone_row_after(table, source_row, after_row):
    new_tr = deepcopy(source_row._tr)
    after_row._tr.addnext(new_tr)
    return table.rows[after_row._index + 1]


def remove_row(table, row) -> None:
    table._tbl.remove(row._tr)


def fill_equipment_table(doc: Document, items: list[dict[str, Any]], total_label: str, grand_total: str) -> None:
    item_tags = ["{{item_no}}", "{{item_name}}", "{{item_qty}}", "{{item_unit_price}}", "{{item_total}}"]

    for table in doc.tables:
        template_row = None
        total_row = None
        for row in table.rows:
            if row_contains_tags(row, item_tags):
                template_row = row
            if row_contains_tags(row, ["{{total_label}}", "{{grand_total}}"]):
                total_row = row

        if template_row is None:
            continue

        insert_after = template_row
        for item in items:
            new_row = clone_row_after(table, template_row, insert_after)
            insert_after = new_row
            values = {
                "{{item_no}}": item.get("item_no", ""),
                "{{item_name}}": item.get("item_name", ""),
                "{{item_qty}}": item.get("item_qty", ""),
                "{{item_unit_price}}": item.get("item_unit_price", ""),
                "{{item_total}}": item.get("item_total", ""),
            }
            for cell in new_row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, values)
                remaining = cell.text
                for tag, value in values.items():
                    if tag in remaining:
                        set_cell_keep_style(cell, value)
                        break

        remove_row(table, template_row)

        if total_row is not None:
            values = {"{{total_label}}": total_label, "{{grand_total}}": grand_total}
            for cell in total_row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, values)
        return


def _paragraph_contains(paragraph, tag: str) -> bool:
    return tag in "".join(run.text for run in paragraph.runs) or tag in paragraph.text


def _clear_paragraph(paragraph) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = ""
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = ""


def _insert_table_after_paragraph(doc: Document, paragraph, rows: int, cols: int):
    table = doc.add_table(rows=rows, cols=cols)
    paragraph._p.addnext(table._tbl)
    return table


def _find_tag_paragraph(doc: Document, tag: str):
    for paragraph in doc.paragraphs:
        if _paragraph_contains(paragraph, tag):
            return paragraph
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if _paragraph_contains(paragraph, tag):
                        return paragraph
    return None


def _set_cell_text(cell, text: Any, bold: bool = False) -> None:
    text = to_text(text)
    cell.text = ""
    paragraph = cell.paragraphs[0]
    for idx, part in enumerate(text.split("\n")):
        if idx:
            paragraph.add_run().add_break(WD_BREAK.LINE)
        run = paragraph.add_run(part)
        run.bold = bold
        try:
            run.font.size = Pt(9)
        except Exception:
            pass


def _apply_table_style(table) -> None:
    try:
        table.style = "Table Grid"
    except Exception:
        pass
    table.autofit = True


def insert_options_table(doc: Document, options: list[dict[str, Any]]) -> None:
    tag = "{{options_table}}"
    paragraph = _find_tag_paragraph(doc, tag)
    if not options:
        if paragraph is not None:
            _clear_paragraph(paragraph)
        return

    if paragraph is None:
        paragraph = doc.add_paragraph()
        paragraph.add_run(tag)

    table = _insert_table_after_paragraph(doc, paragraph, rows=1, cols=3)
    _apply_table_style(table)
    headers = ["№", "Наименование опции", "Кол-во"]
    for idx, header in enumerate(headers):
        _set_cell_text(table.rows[0].cells[idx], header, bold=True)

    for no, option in enumerate(options, start=1):
        row = table.add_row()
        _set_cell_text(row.cells[0], str(no))
        _set_cell_text(row.cells[1], option.get("description", ""))
        _set_cell_text(row.cells[2], option.get("qty", ""))

    _clear_paragraph(paragraph)


def insert_technical_specs_table(doc: Document, specs: list[dict[str, Any]]) -> None:
    tag = "{{technical_specs_table}}"
    paragraph = _find_tag_paragraph(doc, tag)
    if not specs:
        if paragraph is not None:
            _clear_paragraph(paragraph)
        return

    if paragraph is None:
        paragraph = doc.add_paragraph()
        paragraph.add_run(tag)

    table = _insert_table_after_paragraph(doc, paragraph, rows=0, cols=2)
    _apply_table_style(table)

    for spec in specs:
        row = table.add_row()
        if spec.get("is_section"):
            merged = row.cells[0].merge(row.cells[1])
            _set_cell_text(merged, spec.get("name", ""), bold=True)
        else:
            _set_cell_text(row.cells[0], spec.get("name", ""), bold=True)
            _set_cell_text(row.cells[1], spec.get("value", ""))

    _clear_paragraph(paragraph)


def remove_empty_service_tags(doc: Document) -> None:
    replace_tags(doc, {"{{options_table}}": "", "{{technical_specs_table}}": ""})


def get_unique_output_path(path: Path) -> Path:
    if not path.exists():
        return path
    parent = path.parent
    stem = path.stem
    suffix = path.suffix
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def render_docx(
    template_path: str | Path,
    output_path: str | Path,
    replacements: dict[str, Any],
    items: list[dict[str, Any]] | None = None,
    options: list[dict[str, Any]] | None = None,
    technical_specs: list[dict[str, Any]] | None = None,
) -> Path:
    template_path = Path(template_path)
    output_path = Path(output_path)
    doc = Document(template_path)

    fill_equipment_table(
        doc=doc,
        items=items or [],
        total_label=to_text(replacements.get("{{total_label}}", "ИТОГО")),
        grand_total=to_text(replacements.get("{{grand_total}}", "")),
    )

    replace_tags(doc, replacements)
    insert_options_table(doc, options or [])
    insert_technical_specs_table(doc, technical_specs or [])
    remove_empty_service_tags(doc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path = get_unique_output_path(output_path)
    doc.save(output_path)
    return output_path
