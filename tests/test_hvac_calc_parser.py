import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from core.excel_calc_parser import parse_calculation


class HvacCalcParserTest(unittest.TestCase):
    def test_horizontal_hvac_table_uses_final_price_block(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "horizontal_hvac.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Calc"
            sheet.append([])
            sheet.append(
                [
                    "# No",
                    "Model",
                    "Quantity",
                    "Price per unit",
                    "Total per quantity",
                    "Price without margin",
                    "%",
                    "Margin",
                    "Price + Margin",
                    "Rate of currency",
                    "Price + Margin",
                    "%",
                    "Financing",
                    "Total per quantity",
                    "Total per unit",
                    "TOTAL",
                ]
            )
            sheet.append(
                [
                    1,
                    "Horizontal HVAC unit",
                    2,
                    100,
                    200,
                    220,
                    20,
                    44,
                    264,
                    1,
                    264,
                    5,
                    13.2,
                    277.2,
                    138.6,
                    277.2,
                ]
            )
            workbook.save(path)

            result = parse_calculation(path, "Calc")

        self.assertEqual(len(result.items), 1)
        self.assertEqual(result.items[0].name, "Horizontal HVAC unit")
        self.assertEqual(result.items[0].qty, 2)
        self.assertEqual(result.items[0].unit_price, 138.6)
        self.assertEqual(result.items[0].total_price, 277.2)
        self.assertNotIn("Не найдена итоговая строка", result.warnings)


if __name__ == "__main__":
    unittest.main()
