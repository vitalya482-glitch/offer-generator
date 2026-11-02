from __future__ import annotations

from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook


class ExcelParserToCalc:
    """Collect normalized calculation positions from supplier Excel files.

    ``parse`` always returns the same shape. On failure ``ok`` is false and
    ``error`` contains a user-readable explanation; successful rows contain
    name, brand, quantity, unit price, total and delivery terms.
    """

    _NAME = {
        "наименование", "наименование оборудования", "описание", "описание оборудования",
        "модель", "model", "description", "equipment", "equipment name", "item", "позиция",
        "товар", "артикул",
    }
    _QTY = {"количество", "кол во", "кол-во", "колво", "qty", "quantity", "количество шт"}
    _UNIT_PRICE = {
        "цена за ед", "цена за единицу", "цена за шт", "стоимость за ед", "стоимость за единицу", "стоимость за шт",
        "цена за ед euro", "unit price", "unit price euro", "price per unit", "unit cost",
    }
    _TOTAL = {"сумма", "итого", "общая стоимость", "стоимость общая", "total", "total price", "amount"}
    _CURRENCY = {"валюта", "currency", "currency code"}
    _BRAND = {"бренд", "производитель", "brand", "manufacturer", "supplier", "поставщик"}
    _DELIVERY = {
        "условия поставки", "условия поставки и сроки", "условия поставки производства",
        "базис поставки", "логистика", "delivery terms", "delivery basis", "delivery", "incoterms",
    }
    _TERMS = re.compile(r"\b(DDP|DAP|EXW|FCA|CPT|CIP|CFR|CIF|FOB)\b(?:\s+[A-ZА-Я][\w .-]*)?", re.I)
    _SUPPORTED = {".xlsx", ".xlsm"}

    def list_price_sources(self, project_path: str | Path) -> list[Path]:
        """Return available supplier Excel workbooks for a project."""
        suppliers_dir = self._find_suppliers_dir(Path(project_path))
        if suppliers_dir is None:
            raise FileNotFoundError("В выбранной папке не найдена директория «02_Suppliers».")
        return sorted(
            (p for p in suppliers_dir.rglob("*") if p.is_file()
             and p.suffix.casefold() in self._SUPPORTED and not p.name.startswith("~$")),
            key=lambda p: str(p).casefold(),
        )

    @staticmethod
    def sheet_names(source_file: str | Path) -> list[str]:
        path = Path(source_file)
        if not path.is_file():
            raise FileNotFoundError(f"Excel-файл не найден: {path}")
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            return list(workbook.sheetnames)
        finally:
            workbook.close()

    def parse(
        self,
        project_path: str | Path,
        source_file: str | Path | None = None,
        sheet_name: str | None = None,
    ) -> dict[str, Any]:
        """Parse one selected supplier workbook and worksheet."""
        try:
            suppliers_dir = self._find_suppliers_dir(Path(project_path))
            if suppliers_dir is None:
                return self._failure("В выбранной папке не найдена директория «02_Suppliers».")
            if not source_file:
                return self._failure("Выберите Excel-файл с источником цен.")
            if not sheet_name:
                return self._failure("Выберите лист источника цен.")
            file_path = Path(source_file)
            if not file_path.is_file():
                return self._failure(f"Excel-файл не найден или недоступен: {file_path}")
            try:
                if not file_path.resolve().is_relative_to(suppliers_dir.resolve()):
                    return self._failure("Файл источника цен должен находиться в папке «02_Suppliers».")
            except OSError as exc:
                return self._failure(f"Не удалось проверить путь к Excel-файлу: {exc}")
            if file_path.suffix.casefold() not in self._SUPPORTED:
                return self._failure("Поддерживаются Excel-файлы .xlsx и .xlsm.")
            available_sheets = self.sheet_names(file_path)
            if sheet_name not in available_sheets:
                return self._failure(f"Лист «{sheet_name}» не найден в файле «{file_path.name}».")
            rows = self._parse_workbook(file_path, suppliers_dir, selected_sheet=sheet_name)
            if not rows:
                return self._failure(
                    f"На листе «{sheet_name}» файла «{file_path.name}» не найдены строки "
                    "с наименованием, количеством и ценой."
                )
            return {
                "ok": True, "error": "", "rows": rows, "files_scanned": 1,
                "source_file": str(file_path), "sheet_name": sheet_name, "warnings": [],
            }
        except Exception as exc:
            return self._failure(f"Не удалось прочитать исходные данные: {exc}")

    @staticmethod
    def _failure(message: str) -> dict[str, Any]:
        return {"ok": False, "error": message, "rows": [], "files_scanned": 0, "warnings": []}

    @staticmethod
    def _find_suppliers_dir(path: Path) -> Path | None:
        path = path.expanduser()
        if not path.exists() or not path.is_dir():
            raise FileNotFoundError(f"Папка проекта недоступна: {path}")
        if path.name.casefold() in {"02_suppliers", "suppliers"}:
            return path
        candidates = [child for child in path.iterdir() if child.is_dir()]
        for child in candidates:
            if child.name.casefold() in {"02_suppliers", "suppliers"}:
                return child
        return None

    def _parse_workbook(
        self,
        file_path: Path,
        suppliers_dir: Path,
        *,
        selected_sheet: str | None = None,
    ) -> list[dict[str, Any]]:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        try:
            workbook_text = " ".join(
                str(cell.value) for sheet in workbook.worksheets
                for row in sheet.iter_rows() for cell in row
                if cell.value is not None and isinstance(cell.value, str)
            )
            workbook_delivery = self._find_delivery(workbook_text)
            folder_brand = self._brand_from_path(file_path, suppliers_dir)
            rows: list[dict[str, Any]] = []
            for worksheet in workbook.worksheets:
                if selected_sheet is not None and worksheet.title != selected_sheet:
                    continue
                values = [tuple(cell.value for cell in row) for row in worksheet.iter_rows()]
                delivery = self._find_delivery(" ".join(
                    str(value) for row in values for value in row if value is not None
                )) or workbook_delivery
                sheet_brand = folder_brand or self._brand_from_text(worksheet.title)
                header = self._find_header(values)
                if header is None:
                    continue
                header_row, columns = header
                for values_row in values[header_row + 1:]:
                    name = self._row_name(values_row, columns)
                    qty = self._number(self._get(values_row, columns.get("qty")))
                    unit_price = self._number(self._get(values_row, columns.get("unit_price")))
                    total = self._number(self._get(values_row, columns.get("total")))
                    if not name or qty is None or qty == 0 or (unit_price is None and total is None):
                        continue
                    if unit_price is None and total is not None:
                        unit_price = total / qty
                    if total is None and unit_price is not None:
                        total = unit_price * qty
                    brand = self._text(self._get(values_row, columns.get("brand"))) or sheet_brand
                    terms = self._text(self._get(values_row, columns.get("delivery"))) or delivery
                    currency = (
                        self._currency_from_text(self._get(values_row, columns.get("currency")))
                        or str(columns.get("price_currency") or "")
                    )
                    rows.append({
                        "name": name,
                        "brand": brand or "",
                        "quantity": qty,
                        "unit_price": unit_price,
                        "total": total,
                        "currency": currency,
                        "delivery_terms": terms or "",
                    })
            return rows
        finally:
            workbook.close()

    def _find_header(self, values: list[tuple[Any, ...]]) -> tuple[int, dict[str, int]] | None:
        best: tuple[int, int, dict[str, int]] | None = None
        for row_index, row in enumerate(values[:100]):
            columns: dict[str, int] = {}
            for col_index, value in enumerate(row):
                header = self._normalize(value)
                if not header:
                    continue
                if self._is_name_header(header):
                    columns.setdefault("name_columns", []).append(col_index)
                    # Prefer an explicit model/article/description column to a generic name.
                    current = columns.get("name")
                    if current is None or self._name_priority(header) > self._name_priority(self._normalize(row[current])):
                        columns["name"] = col_index
                if header in {"model", "модель", "код", "code", "article", "артикул", "part number", "код товара"}:
                    columns.setdefault("identifier", col_index)
                    columns.setdefault("name_columns", []).append(col_index)
                if self._is_qty_header(header):
                    columns["qty"] = col_index
                if self._is_unit_price_header(header):
                    # Supplier price often follows list price/discount columns.
                    columns["unit_price"] = col_index
                if self._is_total_header(header):
                    columns["total"] = col_index
                    currency = self._currency_from_text(value)
                    if currency:
                        columns["price_currency"] = currency
                if header in self._CURRENCY:
                    columns["currency"] = col_index
                if header in self._BRAND or any(x in header for x in ("brand", "бренд", "производител")):
                    columns["brand"] = col_index
                if header in self._DELIVERY or any(x in header for x in ("условия поставки", "delivery terms", "incoterm")):
                    columns["delivery"] = col_index
            if "name" not in columns or "qty" not in columns or not ({"unit_price", "total"} & columns.keys()):
                continue
            score = 10 + ("unit_price" in columns) * 2 + ("total" in columns) * 2 + ("brand" in columns) + ("delivery" in columns)
            if best is None or score > best[0]:
                best = (score, row_index, columns)
        return (best[1], best[2]) if best else None

    def _row_name(self, row: tuple[Any, ...], columns: dict[str, int]) -> str:
        # Keep model/code and description together. Supplier sheets sometimes
        # put useful labels in a column titled "Code" or "Model" inconsistently.
        values = []
        for index in columns.get("name_columns", [columns.get("name")]):
            text = self._text(self._get(row, index))
            if text and all(self._normalize(text) != self._normalize(previous) for previous in values):
                values.append(text)
        return " — ".join(values)

    @classmethod
    def _is_name_header(cls, text: str) -> bool:
        return text in cls._NAME or any(x in text for x in ("наименование оборудования", "описание оборудования", "unit description", "item description"))

    @classmethod
    def _name_priority(cls, text: str) -> int:
        return 3 if any(x in text for x in ("описание", "description")) else 2 if any(x in text for x in ("модель", "model", "артикул")) else 1

    @classmethod
    def _is_qty_header(cls, text: str) -> bool:
        return text in cls._QTY or text.startswith(("кол во", "кол-во", "quantity")) or text == "qty"

    @classmethod
    def _is_unit_price_header(cls, text: str) -> bool:
        return text in cls._UNIT_PRICE or any(x in text for x in ("цена за ед", "цена за шт", "стоимость за ед", "стоимость за шт", "unit price", "price per unit"))

    @classmethod
    def _is_total_header(cls, text: str) -> bool:
        return text in cls._TOTAL or any(x in text for x in ("total price", "сумма", "итого", "стоимость общая"))

    @staticmethod
    def _get(row: tuple[Any, ...], index: int | None) -> Any:
        return row[index] if index is not None and index < len(row) else None

    @staticmethod
    def _text(value: Any) -> str:
        if value is None:
            return ""
        return " ".join(str(value).replace("\xa0", " ").replace("\n", " ").split()).strip()

    @staticmethod
    def _number(value: Any) -> float | int | None:
        if isinstance(value, bool) or value is None or value == "":
            return None
        if isinstance(value, (int, float)):
            return value
        text = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return None

    @staticmethod
    def _normalize(value: Any) -> str:
        text = str(value or "").casefold().replace("ё", "е").replace("№", "")
        text = re.sub(r"[\r\n\t_./%]+", " ", text)
        text = re.sub(r"[^\w -]+", " ", text, flags=re.UNICODE)
        return " ".join(text.split())

    @classmethod
    def _find_delivery(cls, text: str) -> str:
        match = cls._TERMS.search(text)
        return " ".join(match.group(0).split()) if match else ""

    @classmethod
    def _currency_from_text(cls, value: Any) -> str:
        text = cls._normalize(value)
        if not text:
            return ""
        labels = {
            "EUR": ("eur", "euro", "евро", "€"),
            "USD": ("usd", "доллар", "$"),
            "KZT": ("kzt", "тенге", "тг", "тнг"),
            "CNY": ("cny", "rmb", "юань", "юан"),
            "RUB": ("rub", "руб", "₽"),
        }
        for currency, aliases in labels.items():
            if any(alias in text for alias in aliases):
                return currency
        return ""

    @classmethod
    def _brand_from_text(cls, text: str) -> str:
        text = cls._text(text)
        if not text:
            return ""
        # Supplier workbooks commonly put the brand in parentheses or after a
        # generic title. Keep this rule brand-neutral instead of maintaining a
        # list that would need edits for every new manufacturer.
        parenthetical = re.search(r"\(([^()]*)\)", text)
        if parenthetical and parenthetical.group(1).strip():
            return parenthetical.group(1).strip()
        cleaned = re.sub(
            r"^(?:price\s+of\s+equipment|price\s+list|прайс(?:\s+лист)?(?:\s+оборудования)?)\s*[-:—]*\s*",
            "", text, flags=re.IGNORECASE,
        ).strip()
        cleaned = re.sub(r"^\d+[_. -]+", "", cleaned).strip()
        if cls._normalize(cleaned) and cls._normalize(cleaned) != cls._normalize(text):
            return cleaned
        # For supplier folder names, the whole folder is the vendor/brand name.
        if len(cleaned.split()) <= 3 and cls._normalize(cleaned) not in {"", "suppliers", "02 suppliers"}:
            return cleaned
        return ""

    @classmethod
    def _brand_from_path(cls, file_path: Path, suppliers_dir: Path) -> str:
        try:
            relative_parts = file_path.relative_to(suppliers_dir).parts
        except ValueError:
            relative_parts = file_path.parts
        for part in relative_parts[:-1]:
            if cls._normalize(part) in {"suppliers", "02 suppliers"}:
                continue
            brand = cls._brand_from_text(part)
            if brand:
                return brand
        return ""
