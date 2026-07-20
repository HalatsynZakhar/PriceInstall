from __future__ import annotations

import io
import unittest
from decimal import Decimal
from pathlib import Path
import tempfile

from openpyxl import Workbook, load_workbook

from horoshop_prices import CatalogIndex, PriceRow, build_excel_template, load_settings, parse_excel_prices, plan_prices


class HoroshopPricesTests(unittest.TestCase):
    def settings(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "config.json"
            config.write_text('{"horoshop":{"domain":"https://shop.test"},"wholesale_defaults":{"recalculate_on_current_price":true,"rules":[{"tier":1,"minimal_threshold":2,"discount_percent":10},{"tier":2,"minimal_threshold":5,"discount_percent":15},{"tier":3,"minimal_threshold":10,"discount_percent":25}]}}', encoding="utf-8")
            return load_settings(config)

    def excel(self, rows):
        workbook = Workbook()
        for row in rows:
            workbook.active.append(row)
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()
        return output.getvalue()

    def catalog(self):
        return CatalogIndex.from_raw([{
            "article": "REAL-1", "article_for_display": "Display-1", "price": 1000, "price_old": 1200,
            "wholesale_prices": [{"minimal_threshold": 2, "price": 900}, {"minimal_threshold": 5, "price": 850}, {"minimal_threshold": 20, "price": 700}],
        }])

    def test_current_price_recalculates_default_wholesale_tiers(self):
        plan = plan_prices([PriceRow("Display-1", "price", Decimal("800"), None, 2)], self.catalog(), self.settings())[0]
        self.assertTrue(plan.ready)
        self.assertEqual(plan.payload["price"], 800.0)
        self.assertEqual(plan.payload["wholesale_prices"], [
            {"minimal_threshold": 2, "price": 720.0},
            {"minimal_threshold": 5, "price": 680.0},
            {"minimal_threshold": 10, "price": 600.0},
            {"minimal_threshold": 20, "price": 700.0},
        ])

    def test_invalid_wholesale_price_is_removed(self):
        plan = plan_prices([PriceRow("REAL-1", "wholesale_4", Decimal("1100"), 30, 2)], self.catalog(), self.settings())[0]
        self.assertTrue(plan.ready)
        self.assertEqual(plan.payload["wholesale_prices"], [
            {"minimal_threshold": 2, "price": 900.0},
            {"minimal_threshold": 5, "price": 850.0},
            {"minimal_threshold": 20, "price": 700.0},
        ])
        self.assertIn("Вимкнено", plan.warnings[0])

    def test_excel_accepts_all_supported_price_rows(self):
        rows = parse_excel_prices(self.excel([
            ("Артикул для відображення", "Тип ціни", "Значення", "Кількість від"),
            ("A", "Поточна ціна", "100,5", ""),
            ("A", "Опт 4", 80, 20),
        ]))
        self.assertEqual(rows[0].price_type, "price")
        self.assertEqual(rows[1].threshold, 20)

    def test_old_price_can_be_cleared_with_zero(self):
        rows = parse_excel_prices(self.excel([("Артикул", "Тип ціни", "Значення"), ("A", "РРЦ", 0)]))
        self.assertEqual(rows[0].value, Decimal("0.00"))

    def test_template_has_expected_columns(self):
        workbook = load_workbook(io.BytesIO(build_excel_template()), read_only=True)
        self.assertEqual([cell.value for cell in next(workbook["Ціни"].iter_rows(max_row=1))], ["Артикул для відображення", "Тип ціни", "Значення", "Кількість від (лише для Опт 4-5)"])
        workbook.close()


if __name__ == "__main__":
    unittest.main()
