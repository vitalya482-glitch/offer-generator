from __future__ import annotations

from copy import copy
from datetime import datetime
from pathlib import Path
import re
from typing import Any, Iterable, Mapping

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter


class ExcelCalcMaker:
    """Build a calculation workbook from the common parsed-price row format.

    The input contract is brand-neutral:
    ``name, brand, quantity, unit_price, total, currency, delivery_terms``.
    Brand template layouts are applied here so the GUI and parser stay shared.
    """

    def make_calc(
        self,
        rows: Iterable[Mapping[str, Any]],
        *,
        brand: str,
        client_name: str,
        template_path: str | Path,
        output_dir: str | Path,
        now: datetime | None = None,
    ) -> Path:
        normalized_rows = self._normalize_rows(rows)
        currency = self._single_currency(normalized_rows)
        template = Path(template_path)
        if not template.is_file():
            raise FileNotFoundError(f"Excel-шаблон расчёта не найден: {template}")
        destination_dir = Path(output_dir)
        destination_dir.mkdir(parents=True, exist_ok=True)
        output_path = self.next_output_path(destination_dir, brand, client_name, now=now)

        # Riello's template contains an external-link relationship referenced by
        # its defined names. Dropping it while preserving those names makes
        # Excel report that the generated workbook has damaged content.
        workbook = load_workbook(template, keep_links=True)
        try:
            if brand.strip().casefold() == "riello":
                self._write_riello(workbook, normalized_rows, currency)
            else:
                self._write_generic(workbook, normalized_rows, brand, currency)
            try:
                workbook.calculation.fullCalcOnLoad = True
                workbook.calculation.forceFullCalc = True
            except AttributeError:
                pass
            workbook.save(output_path)
        finally:
            workbook.close()
        return output_path

    @classmethod
    def next_output_path(
        cls,
        output_dir: str | Path,
        brand: str,
        client_name: str,
        *,
        now: datetime | None = None,
    ) -> Path:
        date = now or datetime.now()
        brand_part = cls._safe_filename(brand)
        client_part = cls._safe_filename(client_name)
        prefix = f"Calc_{brand_part}_{client_part}_{date:%d-%m-%y}"
        existing = Path(output_dir)
        revisions = []
        if existing.is_dir():
            for path in existing.glob(f"{prefix}_rev*.xlsx"):
                match = re.search(r"_rev(\d+)\.xlsx$", path.name, re.IGNORECASE)
                if match:
                    revisions.append(int(match.group(1)))
        revision = max(revisions, default=0) + 1
        return existing / f"{prefix}_rev{revision}.xlsx"

    @classmethod
    def _normalize_rows(cls, rows: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
        result = []
        for index, source in enumerate(rows, start=1):
            name = cls._text(source.get("name"))
            if not name:
                raise ValueError(f"У позиции {index} не указано наименование.")
            quantity = cls._number(source.get("quantity"), f"Количество позиции «{name}»")
            unit_price = cls._number(source.get("unit_price"), f"Цена позиции «{name}»")
            if quantity <= 0:
                raise ValueError(f"Количество позиции «{name}» должно быть больше нуля.")
            if unit_price < 0:
                raise ValueError(f"Цена позиции «{name}» не может быть отрицательной.")
            result.append({
                "name": name,
                "brand": cls._text(source.get("brand")),
                "quantity": quantity,
                "unit_price": unit_price,
                "total": quantity * unit_price,
                "currency": cls._text(source.get("currency")).upper(),
                "delivery_terms": cls._text(source.get("delivery_terms")),
            })
        if not result:
            raise ValueError("Выберите хотя бы одну позицию для расчёта.")
        return result

    @staticmethod
    def _single_currency(rows: list[dict[str, Any]]) -> str:
        currencies = {row["currency"] for row in rows if row["currency"]}
        if len(currencies) > 1:
            raise ValueError(
                "В выбранных позициях несколько валют. Сформируйте отдельный Calc для каждой валюты."
            )
        if not currencies:
            raise ValueError("Не удалось определить валюту цен. Проверьте источник прайса.")
        return next(iter(currencies))

    @classmethod
    def _write_riello(cls, workbook, rows: list[dict[str, Any]], currency: str) -> None:
        """Fill the Riello template's configuration, source-price and DDP sheets."""
        conf = workbook["conf"] if "conf" in workbook.sheetnames else workbook.create_sheet("conf")
        cls._clear_cells(conf, 3, max(conf.max_row, len(rows) + 3), 2, 8)
        start = 3
        for row_number, item in enumerate(rows, start=start):
            conf.cell(row_number, 2).value = item["name"]
            conf.cell(row_number, 3).value = item["brand"]
            conf.cell(row_number, 4).value = item["unit_price"]
            conf.cell(row_number, 5).value = item["quantity"]
            conf.cell(row_number, 6).value = f"=D{row_number}*E{row_number}"
            conf.cell(row_number, 7).value = currency
            conf.cell(row_number, 8).value = item["delivery_terms"]
            for column in (4, 6):
                conf.cell(row_number, column).number_format = "#,##0.00"
        conf_total = start + len(rows)
        conf.cell(conf_total, 2).value = "TOTAL"
        conf.cell(conf_total, 6).value = f"=SUM(F{start}:F{conf_total - 1})"
        conf.cell(conf_total, 2).font = copy(conf.cell(3, 2).font)
        conf.cell(conf_total, 2).font = Font(bold=True)
        conf.cell(conf_total, 6).font = Font(bold=True)

        config = workbook["UPS configuration"] if "UPS configuration" in workbook.sheetnames else workbook.create_sheet("UPS configuration")
        for merged in list(config.merged_cells.ranges):
            config.unmerge_cells(str(merged))
        cls._clear_cells(config, 1, max(config.max_row, len(rows) + 6), 1, 10)
        config["A2"] = "Calc"
        config["C2"] = "Selected supplier price items"
        headers = ["Article", "Description / brand", "Size", "Weight, kg", "per unit", "Quantity", f"TOTAL, {currency}", "Delivery terms"]
        for column, value in enumerate(headers, start=2):
            config.cell(3, column).value = value
        cls._style_header(config, 3, 2, 9)
        data_start = 4
        for offset, item in enumerate(rows):
            row_number = data_start + offset
            if row_number != data_start:
                cls._copy_row_style(config, data_start, row_number, 2, 9)
            config.cell(row_number, 2).value = item["name"]
            config.cell(row_number, 3).value = item["brand"]
            config.cell(row_number, 6).value = item["unit_price"]
            config.cell(row_number, 7).value = item["quantity"]
            config.cell(row_number, 8).value = f"=G{row_number}*F{row_number}"
            config.cell(row_number, 9).value = item["delivery_terms"]
            config.cell(row_number, 6).number_format = "#,##0.00"
            config.cell(row_number, 8).number_format = "#,##0.00"
            for column in range(2, 10):
                cell = config.cell(row_number, column)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = cls._body_border()
            config.row_dimensions[row_number].height = 30
        total_row = data_start + len(rows)
        cls._copy_row_style(config, 8, total_row, 2, 9)
        config.cell(total_row, 6).value = f"Gross Total {currency}"
        config.cell(total_row, 8).value = f"=SUM(H{data_start}:H{total_row - 1})"
        config.cell(total_row, 6).font = Font(bold=True)
        config.cell(total_row, 8).font = Font(bold=True)
        config.cell(total_row, 8).number_format = "#,##0.00"
        for column, width in {"B": 30, "C": 24, "D": 22, "E": 12, "F": 14, "G": 12, "H": 18, "I": 30}.items():
            config.column_dimensions[column].width = width

        ddp = next((sheet for sheet in workbook.worksheets if sheet.title.casefold().startswith("ddp")), None)
        if ddp is not None:
            cls._write_riello_ddp_columns(ddp, rows, currency, conf_start=start)

    @classmethod
    def _write_riello_ddp_columns(
        cls,
        worksheet,
        rows: list[dict[str, Any]],
        currency: str,
        *,
        conf_start: int,
    ) -> None:
        """Expand the template calculation column once for every selected item."""
        # In the source template quantity and unit price span C:D. Multiple
        # positions need one complete, independent calculation column each.
        for merged_range in ("C4:D4", "C6:D6"):
            if merged_range in {str(item) for item in worksheet.merged_cells.ranges}:
                worksheet.unmerge_cells(merged_range)

        cls._copy_cell_style(worksheet["C4"], worksheet["D4"])
        cls._copy_cell_style(worksheet["C6"], worksheet["D6"])
        worksheet["C4"] = None
        worksheet["C6"] = None

        # Convert the template's merged-cell formulas to a normal per-column
        # pattern before copying the calculation to E, F, G and so on.
        worksheet["D8"] = "=D4*D6*D7"
        worksheet["D36"] = "=ROUND(D35/D4,0.5)"
        worksheet["D37"] = "=D36*D4"

        max_template_row = worksheet.max_row
        for offset, item in enumerate(rows):
            column = 4 + offset
            if column > 4:
                cls._copy_formula_column(worksheet, 4, column, 1, max_template_row)

            column_letter = get_column_letter(column)
            conf_row = conf_start + offset
            worksheet.cell(2, column).value = item["name"]
            worksheet.cell(3, column).value = item["brand"]
            worksheet.cell(4, column).value = item["quantity"]
            worksheet.cell(5, column).value = 0  # The price source does not provide weight.
            worksheet.cell(6, column).value = f"=conf!D{conf_row}"
            worksheet.cell(7, column).value = 1
            worksheet.cell(8, column).value = f"={column_letter}4*{column_letter}6*{column_letter}7"
            worksheet.cell(12, column).value = item["delivery_terms"]
            worksheet.cell(36, column).value = f"=ROUND({column_letter}35/{column_letter}4,0.5)"
            worksheet.cell(37, column).value = f"={column_letter}36*{column_letter}4"

        worksheet["A11"] = currency.lower()

    @classmethod
    def _copy_formula_column(
        cls,
        worksheet,
        source_column: int,
        target_column: int,
        min_row: int,
        max_row: int,
    ) -> None:
        source_letter = get_column_letter(source_column)
        target_letter = get_column_letter(target_column)
        worksheet.column_dimensions[target_letter].width = worksheet.column_dimensions[source_letter].width
        worksheet.column_dimensions[target_letter].hidden = worksheet.column_dimensions[source_letter].hidden

        for row_number in range(min_row, max_row + 1):
            source = worksheet.cell(row_number, source_column)
            target = worksheet.cell(row_number, target_column)
            cls._copy_cell_style(source, target)
            value = source.value
            if isinstance(value, str) and value.startswith("="):
                value = Translator(value, origin=source.coordinate).translate_formula(target.coordinate)
            target.value = value

    @staticmethod
    def _copy_cell_style(source, target) -> None:
        if source.has_style:
            target._style = copy(source._style)
        target.number_format = source.number_format
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)

    @classmethod
    def _write_generic(cls, workbook, rows: list[dict[str, Any]], brand: str, currency: str) -> None:
        """Write a brand-neutral positions sheet for templates not mapped yet."""
        sheet_name = "Calc positions"
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
        ws = workbook.create_sheet(sheet_name)
        ws.append([f"Calculation — {brand}"])
        ws.append([])
        headers = ["Name", "Brand", "Quantity", "Unit price", "Total", "Currency", "Delivery terms"]
        ws.append(headers)
        cls._style_header(ws, 3, 1, len(headers))
        data_start = 4
        for offset, row in enumerate(rows):
            row_number = data_start + offset
            ws.append([row["name"], row["brand"], row["quantity"], row["unit_price"], f"=C{row_number}*D{row_number}", currency, row["delivery_terms"]])
        total_row = data_start + len(rows)
        ws.cell(total_row, 1).value = "TOTAL"
        ws.cell(total_row, 5).value = f"=SUM(E{data_start}:E{total_row - 1})"
        ws.cell(total_row, 1).font = Font(bold=True)
        ws.cell(total_row, 5).font = Font(bold=True)
        for row_number in range(data_start, total_row + 1):
            for column in range(1, 8):
                ws.cell(row_number, column).border = cls._body_border()
                ws.cell(row_number, column).alignment = Alignment(vertical="top", wrap_text=True)
        for column, width in enumerate((48, 22, 14, 18, 18, 14, 30), start=1):
            ws.column_dimensions[chr(64 + column)].width = width
        ws.freeze_panes = "A4"

    @staticmethod
    def _clear_cells(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                cell = ws.cell(row, column)
                if not isinstance(cell, MergedCell):
                    cell.value = None

    @staticmethod
    def _copy_row_style(ws, source_row: int, target_row: int, min_col: int, max_col: int) -> None:
        for column in range(min_col, max_col + 1):
            source = ws.cell(source_row, column)
            target = ws.cell(target_row, column)
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
            if source.alignment:
                target.alignment = copy(source.alignment)
        if ws.row_dimensions[source_row].height:
            ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

    @staticmethod
    def _style_header(ws, row: int, min_col: int, max_col: int) -> None:
        side = Side(style="thin", color="CBD5E1")
        border = Border(left=side, right=side, top=side, bottom=side)
        for column in range(min_col, max_col + 1):
            cell = ws.cell(row, column)
            cell.font = Font(bold=True)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    @staticmethod
    def _body_border() -> Border:
        side = Side(style="thin", color="E2E8F0")
        return Border(left=side, right=side, top=side, bottom=side)

    @staticmethod
    def _text(value: Any) -> str:
        if value is None:
            return ""
        return " ".join(str(value).replace("\xa0", " ").split()).strip()

    @classmethod
    def _number(cls, value: Any, label: str) -> float:
        if isinstance(value, bool) or value is None or value == "":
            raise ValueError(f"Не задано числовое значение: {label}.")
        text = str(value).replace("\xa0", " ").replace(" ", "").replace(",", ".")
        try:
            return float(text)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Некорректное числовое значение: {label} — {value}.") from exc

    @classmethod
    def _safe_filename(cls, value: str) -> str:
        text = re.sub(r'[<>:"/\\|?*]+', "", cls._text(value)).strip()
        text = re.sub(r"\s+", "_", text)
        return text or "Client"
