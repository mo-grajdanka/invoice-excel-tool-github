from io import BytesIO
from datetime import datetime
from unittest import TestCase

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from excel_writer import SHEET_MAPS, add_items, number, possible_duplicates, read_workbook, validate_lines


class ExcelWriterTests(TestCase):
    @staticmethod
    def _file() -> bytes:
        wb = Workbook()
        wb.remove(wb.active)
        for name, mapping in SHEET_MAPS.items():
            ws = wb.create_sheet(name)
            ws[f"{mapping.name_col}2"] = "Старый товар"
            ws[f"{mapping.quantity_col}2"] = 2
            ws[f"{mapping.price_col}2"] = 100
            ws[f"{mapping.total_col}2"] = "=2*100"
            if mapping.date_col:
                ws[f"{mapping.date_col}2"] = datetime(2026, 1, 1)
                ws[f"{mapping.date_col}2"].number_format = "DD.MM.YYYY"
            ws[f"{mapping.name_col}2"].fill = PatternFill('solid', fgColor='FFFFCC')
            if name == "УК ПК":
                for row in range(3, 10):
                    ws[f"B{row}"] = "не оплачено"
                    ws[f"D{row}"] = "другой номер счёта"
        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_all_six_mappings(self):
        raw = self._file()
        for name, m in SHEET_MAPS.items():
            with self.subTest(sheet=name):
                updated, indices = add_items(raw, name, [{
                    "name": "Кабель 3х2.5", "unit": "м", "quantity": 15,
                    "unit_price": 22, "amount": 330, "include": True,
                }], supplier="Тестовый поставщик", invoice="123", invoice_date="19.11.2026",
                    purpose="Монтаж", pdf_filename="test.pdf")
                wb = read_workbook(updated)
                sh = wb[name]
                self.assertEqual(indices, [10] if name == "УК ПК" else [3])
                r = indices[0]
                self.assertEqual(sh[f"{m.name_col}{r}"].value, "Кабель 3х2.5")
                self.assertEqual(sh[f"{m.quantity_col}{r}"].value, 15)
                self.assertEqual(sh[f"{m.total_col}{r}"].value, 330)
                self.assertEqual(sh[f"{m.total_col}2"].value, "=2*100")
                self.assertEqual(sh[f"{m.name_col}2"].value, "Старый товар")
                self.assertEqual(sh[f"{m.name_col}{r}"].fill.fgColor.rgb, "00FFFFCC")
                if name == "УК ПК":
                    self.assertEqual(sh["D9"].value, "другой номер счёта")
                if m.invoice_col:
                    self.assertIn("123", sh[f"{m.invoice_col}{r}"].value)
                if m.date_col:
                    self.assertEqual(sh[f"{m.date_col}{r}"].value.year, 2026)

    def test_rejects_invalid_line(self):
        with self.assertRaises(ValueError):
            validate_lines([{"include": True, "name": "", "quantity": 3, "unit_price": 12}])

    def test_rejects_formula_like_text(self):
        values = validate_lines([{"include": True, "name": "=CMD()", "quantity": 1, "unit_price": 1}])
        self.assertTrue(values[0]["name"].startswith("'"))

    def test_numbers(self):
        self.assertEqual(number("1 234,55"), 1234.55)
        self.assertEqual(number("1.234,55"), 1234.55)
        self.assertEqual(number("1,234.55"), 1234.55)
        self.assertEqual(number(""), None)
        self.assertIsNone(number(float("nan")))

    def test_duplicates(self):
        raw = self._file()
        wb = read_workbook(raw)
        dup = possible_duplicates(wb, "ИМПЕРИАЛ", [
            {"name": "Старый товар", "quantity": 2, "unit_price": 100},
            {"name": "Новый товар", "quantity": 2, "unit_price": 100}
        ], supplier="", invoice="")
        self.assertEqual(dup, [True, False])
