from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Iterable, Optional

from docx import Document

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

from core.models import CalcData
from core.utils import currency_label, money, qty_text


def iter_paragraphs(document: Document):
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph


def replace_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    text = paragraph.text
    new_text = text
    for key, value in replacements.items():
        new_text = new_text.replace(key, value)
    if new_text != text:
        paragraph.clear()
        paragraph.add_run(new_text)


def replace_text(document: Document, replacements: dict[str, str]) -> None:
    for p in iter_paragraphs(document):
        replace_in_paragraph(p, replacements)


def find_table_by_headers(document: Document, required_headers: Iterable[str]):
    headers = [h.lower() for h in required_headers]
    for table in document.tables:
        table_text = " | ".join(cell.text.lower() for row in table.rows[:2] for cell in row.cells)
        if all(h in table_text for h in headers):
            return table
    return None


def set_cell(cell, text: str, bold: bool = False) -> None:
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    p.clear()
    run = p.add_run(str(text))
    run.bold = bold


def clone_row(table, template_row):
    new_tr = deepcopy(template_row._tr)
    table._tbl.append(new_tr)
    return table.rows[-1]


def remove_row(table, row) -> None:
    table._tbl.remove(row._tr)


def fill_product_table(document: Document, calc: CalcData) -> None:
    table = find_table_by_headers(document, ["наименование", "кол-во"])
    if not table or len(table.rows) < 2:
        return

    for row in table.rows[:1]:
        for cell in row.cells:
            text = cell.text.strip()
            if "{{unit_price_header}}" in text or text.lower().startswith("цена за единицу"):
                set_cell(cell, f"Цена за единицу, {currency_label(calc.currency)}", bold=True)
            elif "{{total_price_header}}" in text or text.lower().startswith("сумма"):
                set_cell(cell, f"Сумма, {currency_label(calc.currency)}", bold=True)

    template_row = None
    for row in table.rows:
        row_text = " ".join(c.text for c in row.cells)
        if "{{item_" in row_text:
            template_row = row
            break
    if template_row is None:
        template_row = table.rows[1]

    total_row = None
    for row in list(table.rows[1:]):
        row_text = " ".join(c.text for c in row.cells)
        row_text_lower = row_text.lower()
        if "итого" in row_text_lower or "{{total_label}}" in row_text:
            total_row = row
            continue
        remove_row(table, row)

    if total_row is not None:
        remove_row(table, total_row)

    for item in calc.items:
        row = clone_row(table, template_row) if len(table.rows) > 1 else table.add_row()
        cells = row.cells
        if len(cells) >= 5:
            set_cell(cells[0], str(item.no))
            set_cell(cells[1], item.name)
            set_cell(cells[2], qty_text(item.qty))
            set_cell(cells[3], money(item.unit_price))
            set_cell(cells[4], money(item.total_price))

    for row in list(table.rows[1:]):
        if "{{item_" in " ".join(c.text for c in row.cells):
            remove_row(table, row)

    row = clone_row(table, total_row) if total_row is not None else table.add_row()
    cells = row.cells
    if len(cells) >= 5:
        set_cell(cells[0], "ИТОГО", bold=True)
        for c in cells[1:-1]:
            set_cell(c, "")
        set_cell(cells[-1], money(calc.total_price), bold=True)


def fill_disabled_spec_blocks(document: Document) -> None:
    replace_text(document, {"{{options_table}}": "", "{{technical_specs_table}}": ""})


def extract_pdf_text(pdf_path: Path, max_chars: int = 5000) -> str:
    if not pdf_path.exists() or fitz is None:
        return ""
    chunks = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            chunks.append(page.get_text("text"))
            if sum(len(x) for x in chunks) > max_chars:
                break
    return "\n".join(chunks).strip()[:max_chars]


def append_pdf_description(document: Document, pdf_dir: Optional[Path], model: str) -> None:
    if not pdf_dir or not pdf_dir.exists():
        return
    candidates = [pdf_dir / f"{model}.pdf"] + list(pdf_dir.glob(f"*{model}*.pdf"))
    pdf_path = next((p for p in candidates if p.exists()), None)
    if not pdf_path:
        return
    text = extract_pdf_text(pdf_path)
    if not text:
        return
    try:
        document.add_page_break()
        document.add_heading("Техническое описание", level=1)
    except Exception:
        document.add_page_break()
        document.add_paragraph("Техническое описание")
    document.add_paragraph(text)
