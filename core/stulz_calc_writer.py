from __future__ import annotations

from pathlib import Path
import re
from typing import Any, Iterable, Mapping

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


class StulzCalcWriter:
    """Заполняет один лист STULZ Calc колонками из отмеченных Calc.pdf."""

    def write(
        self,
        *,
        template_path: str | Path,
        sheet_name: str,
        output_path: str | Path,
        models: Iterable[Mapping[str, Any]],
        currency: str,
        delivery_terms: str,
        delivery_point: str,
        exchange_rate: str | float,
        vat_percent: str | float,
        margin_percent: str | float,
    ) -> Path:
        template = Path(template_path)
        destination = Path(output_path)
        if not template.is_file():
            raise FileNotFoundError(f"Шаблон Calc не найден: {template}")
        if template.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Шаблон должен быть файлом Excel .xlsx или .xlsm.")
        if destination.resolve() == template.resolve():
            raise ValueError("Нельзя сохранять результат поверх файла шаблона.")
        if destination.suffix.lower() != template.suffix.lower():
            raise ValueError("Формат результата должен совпадать с форматом шаблона.")
        destination.parent.mkdir(parents=True, exist_ok=True)

        normalized = self._normalize_models(models, currency)
        rate = self._number(exchange_rate, "Курс")
        vat = self._number(vat_percent, "НДС")
        margin = self._number(margin_percent, "Маржа")
        if rate <= 0:
            raise ValueError("Курс должен быть больше нуля.")

        keep_vba = template.suffix.lower() == ".xlsm"
        workbook = load_workbook(template, keep_links=True, keep_vba=keep_vba)
        try:
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"В шаблоне нет листа «{sheet_name}».")
            base_sheet = workbook[sheet_name]
            if sheet_name.strip().casefold() == "spare parts":
                self._write_spare_parts(
                    base_sheet, normalized, rate, vat, margin, delivery_terms, delivery_point
                )
            else:
                self._write_equipment_sheets(
                    workbook,
                    base_sheet,
                    normalized,
                    rate,
                    vat,
                    margin,
                    currency,
                    delivery_terms,
                    delivery_point,
                )

            try:
                workbook.calculation.fullCalcOnLoad = True
                workbook.calculation.forceFullCalc = True
            except AttributeError:
                pass
            workbook.save(destination)
        finally:
            workbook.close()
        return destination

    @classmethod
    def _normalize_models(
        cls,
        models: Iterable[Mapping[str, Any]],
        selected_currency: str,
    ) -> list[dict[str, Any]]:
        rows = []
        expected_currency = str(selected_currency or "").strip().upper()
        for index, source in enumerate(models, start=1):
            name = str(source.get("model") or "").strip()
            if not name:
                continue
            quantity = cls._number(source.get("qty_value"), f"Количество «{name}»")
            unit_price = cls._number(source.get("unit_purchase_price"), f"Цена покупки «{name}»")
            source_currency = str(source.get("currency") or "").strip().upper()
            if quantity <= 0:
                raise ValueError(f"Количество позиции «{name}» должно быть больше нуля.")
            if unit_price < 0:
                raise ValueError(f"Цена покупки «{name}» не может быть отрицательной.")
            if not source_currency:
                raise ValueError(f"В Calc.pdf для позиции «{name}» не определена валюта.")
            if source_currency != expected_currency:
                raise ValueError(
                    f"В Calc.pdf у позиции «{name}» валюта {source_currency}, "
                    f"а в параметрах выбрана {expected_currency}."
                )
            rows.append({
                "name": name,
                "quantity": quantity,
                "unit_price": unit_price,
                "currency": source_currency,
                "source": str(source.get("calc_pdf") or ""),
            })
        if not rows:
            raise ValueError("Отметьте хотя бы одну модель с ценой в таблице позиций.")
        return rows

    @classmethod
    def _write_equipment_sheets(
        cls,
        workbook,
        base_sheet,
        rows: list[dict[str, Any]],
        rate: float,
        vat: float,
        margin: float,
        currency: str,
        delivery_terms: str,
        delivery_point: str,
    ) -> None:
        if base_sheet["A2"].value != "Model" or base_sheet["A3"].value != "Quantity":
            raise ValueError(
                f"Лист «{base_sheet.title}» не распознан как лист расчёта оборудования. "
                "Выберите лист оборудования или Spare Parts."
            )

        # Повторяем исходный трехколоночный блок B:D вправо для каждой модели.
        # В формулах сдвигаются также абсолютные ссылки на параметры внутри блока.
        block_merges = [
            merged
            for merged in base_sheet.merged_cells.ranges
            if merged.min_col >= 2 and merged.max_col <= 4
        ]
        max_row = base_sheet.max_row
        for model_index in range(1, len(rows)):
            delta = model_index * 3
            for merged in block_merges:
                base_sheet.merge_cells(
                    start_row=merged.min_row,
                    start_column=merged.min_col + delta,
                    end_row=merged.max_row,
                    end_column=merged.max_col + delta,
                )

            for row_number in range(1, max_row + 1):
                for source_col in (2, 3, 4):
                    source = base_sheet.cell(row_number, source_col)
                    target = base_sheet.cell(row_number, source_col + delta)
                    if isinstance(target, MergedCell):
                        continue
                    cls._copy_cell_style(source, target)
                    value = source.value
                    if isinstance(value, str) and value.startswith("="):
                        value = cls._shift_group_formula(value, delta)
                    target.value = value

            for source_col in (2, 3, 4):
                source_letter = get_column_letter(source_col)
                target_letter = get_column_letter(source_col + delta)
                base_sheet.column_dimensions[target_letter].width = base_sheet.column_dimensions[source_letter].width
                base_sheet.column_dimensions[target_letter].hidden = base_sheet.column_dimensions[source_letter].hidden

        is_ddp = any(
            str(base_sheet.cell(row, 1).value or "").strip().casefold().startswith("ddp")
            for row in range(1, min(base_sheet.max_row, 60) + 1)
        )
        is_exw = any(
            str(base_sheet.cell(row, 1).value or "").strip().casefold().startswith("exw")
            for row in range(1, min(base_sheet.max_row, 60) + 1)
        )

        cls._set_currency_labels(base_sheet, currency)
        for model_index, item in enumerate(rows):
            delta = model_index * 3
            percent_col = 2 + delta
            quantity_col = 3 + delta
            model_col = 4 + delta
            base_sheet.cell(2, model_col).value = item["name"]
            base_sheet.cell(3, quantity_col).value = item["quantity"]
            base_sheet.cell(4, quantity_col).value = item["unit_price"]

            if is_ddp:
                base_sheet.cell(28, quantity_col).value = margin
                base_sheet.cell(30, percent_col).value = vat
                base_sheet.cell(33, model_col).value = rate
                base_sheet.cell(26, percent_col).value = 0  # Нет отдельной статьи финансирования.
                for row in (7, 8, 9, 10, 11, 12, 15, 16, 19, 20, 21, 22, 23):
                    base_sheet.cell(row, quantity_col if row in range(7, 13) else model_col).value = 0
                base_sheet.cell(17, quantity_col).value = 0
                base_sheet.cell(18, quantity_col).value = 0
                base_sheet.cell(38, quantity_col).value = 0
                base_sheet["A31"] = cls._delivery_label(delivery_terms, delivery_point)
            elif is_exw:
                base_sheet.cell(16, quantity_col).value = margin
                base_sheet.cell(20, model_col).value = rate
                base_sheet.cell(24, quantity_col).value = 0
                for row in range(7, 13):
                    base_sheet.cell(row, quantity_col).value = 0
                label = cls._delivery_label(delivery_terms, delivery_point)
                if label:
                    base_sheet["A18"] = label
            else:
                raise ValueError(
                    f"Не удалось распознать расчётные строки листа «{base_sheet.title}». "
                    "Поддерживаются листы DDP и EXW из шаблона STULZ."
                )

    @staticmethod
    def _copy_cell_style(source, target) -> None:
        if source.has_style:
            target._style = source._style
        target.number_format = source.number_format
        target.alignment = source.alignment.copy()
        target.protection = source.protection.copy()

    @staticmethod
    def _shift_group_formula(formula: str, delta: int) -> str:
        reference = re.compile(r"(?<![A-Z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)")

        def shift(match: re.Match[str]) -> str:
            column = match.group(2)
            row = match.group(4)
            if column == "A":
                return f"$A{match.group(3)}{row}"
            if column in {"B", "C", "D"}:
                source_index = {"B": 2, "C": 3, "D": 4}[column]
                return f"{match.group(1)}{get_column_letter(source_index + delta)}{match.group(3)}{row}"
            return match.group(0)

        return reference.sub(shift, formula)

    @classmethod
    def _write_spare_parts(
        cls,
        worksheet,
        rows: list[dict[str, Any]],
        rate: float,
        vat: float,
        margin: float,
        delivery_terms: str,
        delivery_point: str,
    ) -> None:
        capacity = 13
        if len(rows) > capacity:
            raise ValueError(
                f"Лист Spare Parts в шаблоне рассчитан на {capacity} строк. "
                f"Отмечено позиций: {len(rows)}."
            )
        for row_number in range(3, 16):
            worksheet.cell(row_number, 1).value = None
            worksheet.cell(row_number, 2).value = None
            worksheet.cell(row_number, 3).value = 0
            worksheet.cell(row_number, 4).value = 0
        for row_number, item in enumerate(rows, start=3):
            worksheet.cell(row_number, 1).value = item["name"]
            worksheet.cell(row_number, 3).value = item["quantity"]
            worksheet.cell(row_number, 4).value = item["unit_price"]

        worksheet["E18"] = 0
        worksheet["E19"] = 0
        worksheet["E20"] = 0
        worksheet["B21"] = 0
        worksheet["E22"] = 0
        worksheet["E23"] = 0
        worksheet["E24"] = 0
        worksheet["B27"] = margin
        worksheet["B29"] = vat
        worksheet["E32"] = rate
        label = cls._delivery_label(delivery_terms, delivery_point)
        if label:
            worksheet["A30"] = f"TOTAL {label}"
        cls._set_currency_labels(worksheet, rows[0]["currency"])

    @staticmethod
    def _set_currency_labels(worksheet, currency: str) -> None:
        labels = {"EUR": "€, euro", "USD": "$, USD", "KZT": "kzt"}
        label = labels.get(str(currency).upper(), str(currency).upper())
        rows = (14, 24) if worksheet["A2"].value == "Model" else (17, 25)
        for row in rows:
            value = str(worksheet.cell(row, 1).value or "").strip().casefold()
            if value in {"€, euro", "usd", "$, usd", "kzt"}:
                worksheet.cell(row, 1).value = label

    @staticmethod
    def _delivery_label(terms: str, point: str) -> str:
        basis = str(terms or "").strip()
        destination = str(point or "").strip()
        return " ".join(part for part in (basis, destination) if part)

    @staticmethod
    def _number(value: Any, label: str) -> float:
        if isinstance(value, bool) or value is None or value == "":
            raise ValueError(f"Не задано числовое значение: {label}.")
        try:
            return float(str(value).replace("\u00a0", " ").replace(" ", "").replace(",", "."))
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Некорректное числовое значение: {label} — {value}.") from exc
