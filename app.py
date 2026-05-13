from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl import load_workbook

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

MONTHS_RU = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля", 5: "мая", 6: "июня",
    7: "июля", 8: "августа", 9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}

@dataclass
class CalcData:
    sheet_name: str
    version: str
    model: str
    quantity: float
    unit_price: float
    total_price: float
    delivery_basis: str
    options: list[tuple[str, float]]


def money(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def sanitize_filename(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value).strip()
    return value or "client"


def load_config(project_dir: Path) -> dict:
    cfg = project_dir / "config.json"
    if not cfg.exists():
        cfg = project_dir / "config.example.json"
    if cfg.exists():
        return json.loads(cfg.read_text(encoding="utf-8"))
    return {}


def first_not_empty(*values):
    for value in values:
        if value not in (None, ""):
            return value
    return ""


def as_float(value, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default


def find_row_by_label(ws, label: str) -> Optional[int]:
    label_norm = label.lower().strip()
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, str) and val.lower().strip() == label_norm:
                return cell.row
    return None


def parse_calc(xlsx_path: Path, sheet_name: Optional[str] = None) -> CalcData:
    wb_values = load_workbook(xlsx_path, data_only=True)
    wb_formulas = load_workbook(xlsx_path, data_only=False)
    ws = wb_values[sheet_name] if sheet_name else wb_values[wb_values.sheetnames[0]]
    wsf = wb_formulas[ws.title]

    version = str(first_not_empty(ws["C1"].value, "Version 1"))
    model = str(first_not_empty(ws["D2"].value, ws["E2"].value, "Equipment"))
    quantity = as_float(ws["C3"].value, 1.0)

    total_row = find_row_by_label(ws, "TOTAL")
    total_price = as_float(ws.cell(total_row, 4).value if total_row else None, 0.0)
    unit_price = total_price / quantity if quantity else total_price

    delivery_basis = "DDP г. Алматы" if "DDP" in ws.title.upper() else "EXW Hamburg, Germany"

    options: list[tuple[str, float]] = []
    start = find_row_by_label(ws, "Options:") or 6
    end = find_row_by_label(ws, "Total EXW for us") or min(start + 80, ws.max_row)
    for row_idx in range(start + 1, end):
        name = ws.cell(row_idx, 1).value
        qty = ws.cell(row_idx, 3).value
        if isinstance(name, str) and name.strip() and as_float(qty, 0) > 0:
            options.append((name.strip(), as_float(qty)))

    # fallback: if Excel was not calculated, use visible zero but keep model/qty
    if total_price == 0:
        # Try to take price from row named DDP/EXW if it has a cached value.
        for label in ("DDP Almaty", "EXW - Hamburg, Germany", "Total per quantity"):
            r = find_row_by_label(ws, label)
            if r and as_float(ws.cell(r, 4).value) > 0:
                total_price = as_float(ws.cell(r, 4).value)
                unit_price = total_price / quantity if quantity else total_price
                break

    return CalcData(
        sheet_name=ws.title,
        version=version,
        model=model,
        quantity=quantity,
        unit_price=unit_price,
        total_price=total_price,
        delivery_basis=delivery_basis,
        options=options,
    )


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
        # Preserve simple paragraph style; rebuild runs for reliable replacement across split runs.
        for run in paragraph.runs:
            run.text = ""
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
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
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(10)


def fill_product_table(document: Document, calc: CalcData, currency: str) -> None:
    table = find_table_by_headers(document, ["наименование", "кол-во", "цена"])
    if not table:
        return
    # Prefer first data row after headers. Original template has one blank row.
    row = table.rows[1] if len(table.rows) > 1 else table.add_row()
    while len(row.cells) < 5:
        table.add_row()
        row = table.rows[-1]
    set_cell(row.cells[0], "1")
    set_cell(row.cells[1], calc.model)
    set_cell(row.cells[2], str(int(calc.quantity) if calc.quantity.is_integer() else calc.quantity))
    set_cell(row.cells[3], money(calc.unit_price))
    set_cell(row.cells[4], money(calc.total_price))

    # Try to update total row if present.
    for r in table.rows:
        if "итого" in " ".join(c.text.lower() for c in r.cells):
            set_cell(r.cells[-1], money(calc.total_price), bold=True)
            break


def fill_spec_table(document: Document, calc: CalcData) -> None:
    table = find_table_by_headers(document, ["наименование опции", "кол-во"])
    if not table or not calc.options:
        return
    # Keep header row, remove most existing body content by blanking rows, then repopulate.
    while len(table.rows) > 1:
        tbl = table._tbl
        tbl.remove(table.rows[-1]._tr)
    for idx, (name, qty) in enumerate(calc.options, start=1):
        row = table.add_row()
        set_cell(row.cells[0], str(idx))
        set_cell(row.cells[1], name)
        qty_text = str(int(qty) if float(qty).is_integer() else qty)
        set_cell(row.cells[2], f"{qty_text} шт.")


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
    document.add_page_break()
    document.add_heading("Техническое описание", level=1)
    document.add_paragraph(text)


def make_offer(template_path: Path, calc_path: Path, output_dir: Path, client_name: str,
               sheet_name: Optional[str] = None, pdf_dir: Optional[Path] = None,
               version: Optional[str] = None, city: str = "г. Алматы") -> Path:
    calc = parse_calc(calc_path, sheet_name)
    document = Document(template_path)
    now = datetime.now()
    date_ru = f"{city}, {now.day} {MONTHS_RU[now.month]} {now.year} г."
    version_value = version or calc.version.replace("Version", "Версия №")

    total_words = ""
    replacements = {
        "ТОО «[Организация]»": f"ТОО «{client_name}»",
        "[Организация]": client_name,
        "г. Алматы, 15 апреля 2021 г.": date_ru,
        "Версия №1": version_value,
        "в течение 13-14 недель после поступления предоплаты.": "в течение 13-14 недель после поступления предоплаты.",
        "DDP г. Алматы": calc.delivery_basis,
        "Итого, стоимость оборудования составляет:  ( ) EUR 00 eurocent, с учетом НДС 12%.":
            f"Итого, стоимость оборудования составляет: {money(calc.total_price)} EUR, с учетом НДС 12%.",
        "кондиционеров CCU121A": f"кондиционеров {calc.model}",
        "Тип модуля:\n\tASU 211 AL": f"Тип модуля:\n\t{calc.model}",
    }
    replace_text(document, replacements)
    fill_product_table(document, calc, "EUR")
    fill_spec_table(document, calc)
    append_pdf_description(document, pdf_dir, calc.model)

    output_dir.mkdir(parents=True, exist_ok=True)
    filename = f"КП_{sanitize_filename(client_name)}_{calc.model}_{now:%Y-%m-%d}.docx"
    out = output_dir / filename
    document.save(out)
    return out


def run_gui(project_dir: Path) -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    cfg = load_config(project_dir)
    root = tk.Tk()
    root.title("Генератор КП")
    root.geometry("720x360")

    client = tk.StringVar(value="ТОО Example")
    template = tk.StringVar(value=str(project_dir / cfg.get("default_template", "templates/kp_template.docx")))
    calc = tk.StringVar(value=str(project_dir / cfg.get("default_calc", "samples/Calc_23-12-24 PAC.xlsx")))
    pdf_dir = tk.StringVar(value=str(project_dir / cfg.get("default_pdf_dir", "pdf")))
    output_dir = tk.StringVar(value=str(project_dir / cfg.get("default_output_dir", "output")))
    sheet = tk.StringVar(value="")

    def browse_file(var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def browse_dir(var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def generate():
        try:
            out = make_offer(
                template_path=Path(template.get()),
                calc_path=Path(calc.get()),
                output_dir=Path(output_dir.get()),
                client_name=client.get().strip() or "Client",
                sheet_name=sheet.get().strip() or None,
                pdf_dir=Path(pdf_dir.get()) if pdf_dir.get().strip() else None,
            )
            messagebox.showinfo("Готово", f"КП сформировано:\n{out}")
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))

    rows = [
        ("Клиент", client, None),
        ("Word шаблон", template, lambda: browse_file(template, [("Word", "*.docx")])) ,
        ("Excel расчет", calc, lambda: browse_file(calc, [("Excel", "*.xlsx")])) ,
        ("Папка PDF", pdf_dir, lambda: browse_dir(pdf_dir)),
        ("Папка результата", output_dir, lambda: browse_dir(output_dir)),
        ("Лист Excel (пусто = первый)", sheet, None),
    ]
    for i, (label, var, cmd) in enumerate(rows):
        ttk.Label(root, text=label).grid(row=i, column=0, padx=10, pady=8, sticky="w")
        ttk.Entry(root, textvariable=var, width=70).grid(row=i, column=1, padx=10, pady=8, sticky="we")
        if cmd:
            ttk.Button(root, text="...", command=cmd).grid(row=i, column=2, padx=10, pady=8)
    ttk.Button(root, text="Сформировать КП", command=generate).grid(row=len(rows), column=1, padx=10, pady=20, sticky="e")
    root.columnconfigure(1, weight=1)
    root.mainloop()


def main(argv=None) -> int:
    project_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    parser = argparse.ArgumentParser(description="Генератор коммерческого предложения в Word")
    parser.add_argument("--gui", action="store_true", help="Запустить окно Windows")
    parser.add_argument("--template", default=str(project_dir / "templates/kp_template.docx"))
    parser.add_argument("--calc", default=str(project_dir / "samples/Calc_23-12-24 PAC.xlsx"))
    parser.add_argument("--pdf-dir", default=str(project_dir / "pdf"))
    parser.add_argument("--output-dir", default=str(project_dir / "output"))
    parser.add_argument("--client", default="ТОО Example")
    parser.add_argument("--sheet", default=None)
    args = parser.parse_args(argv)

    if args.gui or len(sys.argv) == 1:
        run_gui(project_dir)
        return 0

    out = make_offer(
        template_path=Path(args.template),
        calc_path=Path(args.calc),
        output_dir=Path(args.output_dir),
        client_name=args.client,
        sheet_name=args.sheet,
        pdf_dir=Path(args.pdf_dir) if args.pdf_dir else None,
    )
    print(out)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
