from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document


def to_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def replace_in_paragraph(paragraph, replacements: dict[str, Any]) -> None:
    """
    Replaces tags while preserving Word formatting as much as possible.
    First tries run-by-run replacement.
    If Word split the tag across several runs, falls back to first-run replacement.
    """
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

    if updated_full != full_text and paragraph.runs:
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
    """
    Sets cell text without using cell.text, so Word formatting is preserved.
    """
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
    """Clone source_row and insert the clone after after_row.

    Important: for repeated item rows we must always clone the original
    template row, not the previously filled row. Otherwise the second and
    next rows inherit already-replaced text and become duplicates.
    """
    new_tr = deepcopy(source_row._tr)
    after_row._tr.addnext(new_tr)

    # python-docx row indexes are recalculated from the XML tree. The newly
    # inserted row is directly after after_row.
    return table.rows[after_row._index + 1]


def remove_row(table, row) -> None:
    tbl = table._tbl
    tbl.remove(row._tr)


def fill_equipment_table(doc: Document, items: list[dict[str, Any]], total_label: str, grand_total: str) -> None:
    item_tags = [
        "{{item_no}}",
        "{{item_name}}",
        "{{item_qty}}",
        "{{item_unit_price}}",
        "{{item_total}}",
    ]

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
                if "{{item_no}}" in remaining:
                    set_cell_keep_style(cell, values["{{item_no}}"])
                elif "{{item_name}}" in remaining:
                    set_cell_keep_style(cell, values["{{item_name}}"])
                elif "{{item_qty}}" in remaining:
                    set_cell_keep_style(cell, values["{{item_qty}}"])
                elif "{{item_unit_price}}" in remaining:
                    set_cell_keep_style(cell, values["{{item_unit_price}}"])
                elif "{{item_total}}" in remaining:
                    set_cell_keep_style(cell, values["{{item_total}}"])

        remove_row(table, template_row)

        if total_row is not None:
            values = {
                "{{total_label}}": total_label,
                "{{grand_total}}": grand_total,
            }

            for cell in total_row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, values)

        return


def remove_empty_service_tags(doc: Document) -> None:
    empty_tags = {
        "{{options_table}}": "",
        "{{technical_specs_table}}": "",
    }
    replace_tags(doc, empty_tags)


def get_unique_output_path(path: Path) -> Path:
    """
    Returns a free output path without overwriting an existing file.

    Example:
    KP.docx -> KP_1.docx -> KP_2.docx
    """
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
) -> Path:
    template_path = Path(template_path)
    output_path = Path(output_path)

    doc = Document(template_path)

    items = items or []

    fill_equipment_table(
        doc=doc,
        items=items,
        total_label=to_text(replacements.get("{{total_label}}", "ИТОГО")),
        grand_total=to_text(replacements.get("{{grand_total}}", "")),
    )

    replace_tags(doc, replacements)
    remove_empty_service_tags(doc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path = get_unique_output_path(output_path)
    doc.save(output_path)

    return output_path