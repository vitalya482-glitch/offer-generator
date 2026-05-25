from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.shared import Pt


def to_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def replace_in_paragraph(paragraph, replacements: dict[str, Any]) -> None:
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


def fill_equipment_table(
    doc: Document,
    items: list[dict[str, Any]],
    total_label: str,
    grand_total: str,
) -> None:
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


def _paragraph_contains(paragraph, tag: str) -> bool:
    return tag in "".join(run.text for run in paragraph.runs) or tag in paragraph.text


def _clear_paragraph(paragraph) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = ""
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.text = ""


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


def _insert_paragraph_before(paragraph):
    new_p = OxmlElement("w:p")
    paragraph._p.addprevious(new_p)
    return paragraph._parent.add_paragraph("")


def _insert_paragraph_after_element(element):
    new_p = OxmlElement("w:p")
    element.addnext(new_p)
    return new_p


def _insert_page_break_before_paragraph(paragraph) -> None:
    new_p = OxmlElement("w:p")
    paragraph._p.addprevious(new_p)

    parent = paragraph._p.getparent()
    index = parent.index(new_p)

    temp_doc = paragraph._parent
    new_paragraph = temp_doc.paragraphs[index] if index < len(temp_doc.paragraphs) else None

    if new_paragraph is not None:
        new_paragraph.add_run().add_break(WD_BREAK.PAGE)


def _insert_empty_paragraph_after_paragraph(paragraph) -> None:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)


def _insert_empty_paragraph_after_table(table) -> None:
    new_p = OxmlElement("w:p")
    table._tbl.addnext(new_p)


def _insert_page_break_after_table(table) -> None:
    new_p = OxmlElement("w:p")
    table._tbl.addnext(new_p)

    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "page")
    run.append(br)
    new_p.append(run)


def _set_cell_text(cell, text: Any, bold: bool = False) -> None:
    text = to_text(text)

    if cell.paragraphs:
        paragraph = cell.paragraphs[0]
    else:
        paragraph = cell.add_paragraph()

    for run in paragraph.runs:
        run.text = ""

    if not paragraph.runs:
        run = paragraph.add_run("")
    else:
        run = paragraph.runs[0]

    parts = text.split("\n")
    run.text = parts[0]
    run.bold = bold

    try:
        run.font.size = Pt(9)
    except Exception:
        pass

    for part in parts[1:]:
        br_run = paragraph.add_run()
        br_run.add_break(WD_BREAK.LINE)
        text_run = paragraph.add_run(part)
        text_run.bold = bold
        try:
            text_run.font.size = Pt(9)
        except Exception:
            pass

    for extra_paragraph in cell.paragraphs[1:]:
        for extra_run in extra_paragraph.runs:
            extra_run.text = ""


def _find_table_with_row_tags(doc: Document, tags: list[str]):
    for table in doc.tables:
        for row in table.rows:
            if row_contains_tags(row, tags):
                return table, row
    return None, None


def fill_options_template_table(
    doc: Document,
    options: list[dict[str, Any]],
    title: str,
) -> None:
    title_tag = "{{options_title}}"
    row_tags = ["{{opt_no}}", "{{opt_name}}", "{{opt_qty}}"]

    title_paragraph = _find_tag_paragraph(doc, title_tag)
    table, template_row = _find_table_with_row_tags(doc, row_tags)

    if title_paragraph is not None:
        _insert_page_break_before_paragraph(title_paragraph)
        replace_in_paragraph(title_paragraph, {title_tag: title})
        _insert_empty_paragraph_after_paragraph(title_paragraph)

    if table is None or template_row is None:
        return

    if not options:
        remove_row(table, template_row)
        return

    insert_after = template_row

    for no, option in enumerate(options, start=1):
        new_row = clone_row_after(table, template_row, insert_after)
        insert_after = new_row

        values = {
            "{{opt_no}}": str(no),
            "{{opt_name}}": option.get("description", option.get("name", "")),
            "{{opt_qty}}": option.get("qty", ""),
        }

        for cell in new_row.cells:
            for paragraph in cell.paragraphs:
                replace_in_paragraph(paragraph, values)

    remove_row(table, template_row)
    _insert_empty_paragraph_after_table(table)


def fill_technical_specs_template_table(
    doc: Document,
    specs: list[dict[str, Any]],
    title: str,
) -> None:
    title_tag = "{{technical_specs_title}}"
    row_tags = ["{{technical_specs_parameter}}", "{{technical_specs_value}}"]

    title_paragraph = _find_tag_paragraph(doc, title_tag)
    table, template_row = _find_table_with_row_tags(doc, row_tags)

    if title_paragraph is not None:
        _insert_empty_paragraph_after_paragraph(title_paragraph)
        replace_in_paragraph(title_paragraph, {title_tag: title})
        _insert_empty_paragraph_after_paragraph(title_paragraph)

    if table is None or template_row is None:
        return

    if not specs:
        remove_row(table, template_row)
        return

    insert_after = template_row

    for spec in specs:
        new_row = clone_row_after(table, template_row, insert_after)
        insert_after = new_row

        if spec.get("is_section"):
            values = {
                "{{technical_specs_parameter}}": spec.get("name", ""),
                "{{technical_specs_value}}": "",
            }
        else:
            values = {
                "{{technical_specs_parameter}}": spec.get("name", ""),
                "{{technical_specs_value}}": spec.get("value", ""),
            }

        for cell in new_row.cells:
            for paragraph in cell.paragraphs:
                replace_in_paragraph(paragraph, values)

    remove_row(table, template_row)
    _insert_empty_paragraph_after_table(table)
    _insert_page_break_after_table(table)


def remove_empty_service_tags(doc: Document) -> None:
    replace_tags(
        doc,
        {
            "{{options_table}}": "",
            "{{technical_specs_table}}": "",
            "{{technical_specifications}}": "",
            "{{options_title}}": "",
            "{{opt_no}}": "",
            "{{opt_name}}": "",
            "{{opt_qty}}": "",
            "{{technical_specs_title}}": "",
            "{{technical_specs_parameter}}": "",
            "{{technical_specs_value}}": "",
        },
    )


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

    options_title = to_text(
        replacements.get(
            "{{options_title}}",
            "Опции, включенные в комплектацию кондиционеров:",
        )
    )

    technical_specs_title = to_text(
        replacements.get(
            "{{technical_specs_title}}",
            "Технические характеристики кондиционеров:",
        )
    )

    replace_tags(doc, replacements)

    fill_options_template_table(
        doc=doc,
        options=options or [],
        title=options_title,
    )

    fill_technical_specs_template_table(
        doc=doc,
        specs=technical_specs or [],
        title=technical_specs_title,
    )

    remove_empty_service_tags(doc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path = get_unique_output_path(output_path)

    doc.save(output_path)
    return output_path