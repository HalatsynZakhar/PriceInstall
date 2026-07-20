from __future__ import annotations

import io
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from horoshop_prices import CatalogIndex, FieldChange, PriceRow, WholesaleChange, build_excel_template, load_settings, parse_excel_prices, plan_prices


class HoroshopPricesTests(unittest.TestCase):
    def settings(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "config.json"
            config.write_text('{"horoshop":{"domain":"https://shop.test"},"wholesale_defaults":{"recalculate_from_rrc":true,"rules":[{"tier":1,"minimal_threshold":2,"discount_percent":10},{"tier":2,"minimal_threshold":5,"discount_percent":15},{"tier":3,"minimal_threshold":10,"discount_percent":25}]}}', encoding="utf-8")
            return load_settings(config)

    def excel(self, headers, rows):
        workbook = Workbook()
        workbook.active.append(headers)
        for row in rows:
            workbook.active.append(row)
        output = io.BytesIO(); workbook.save(output); workbook.close()
        return output.getvalue()

    def catalog(self):
        return CatalogIndex.from_raw([{"article":"REAL-1","article_for_display":"Display-1","price":1000,"price_old":1000,"wholesale_prices":[{"minimal_threshold":2,"price":900},{"minimal_threshold":20,"price":700}]}])

    def test_sale_price_disables_wholesale_price_above_it(self):
        row = PriceRow("Display-1", FieldChange(Decimal("750")), FieldChange(), False, (), 2)
        plan = plan_prices([row], self.catalog(), self.settings())[0]
        self.assertEqual(plan.payload["wholesale_prices"], [{"minimal_threshold": 20, "price": 700.0}])
        self.assertIn("Вимкнено", "; ".join(plan.warnings))

    def test_moves_existing_current_price_to_rrc(self):
        row = PriceRow("REAL-1", FieldChange(Decimal("800")), FieldChange(), True, (), 2)
        plan = plan_prices([row], self.catalog(), self.settings())[0]
        self.assertEqual(plan.payload["price"], 800.0)
        self.assertEqual(plan.payload["price_old"], 1000.0)

    def test_explicit_delete_is_different_from_blank(self):
        row = PriceRow("REAL-1", FieldChange(), FieldChange(delete=True), False, (WholesaleChange(1, FieldChange(delete=True)),), 2)
        plan = plan_prices([row], self.catalog(), self.settings())[0]
        self.assertEqual(plan.payload["price_old"], 0)
        self.assertEqual(plan.payload["wholesale_prices"], [{"minimal_threshold": 20, "price": 700.0}])

    def test_excel_supports_one_row_and_reordered_columns(self):
        rows = parse_excel_prices(self.excel(["Опт 1", "Артикул", "Поточна ціна", "Поточну ціну в РРЦ (Так)", "Опт 1 від"], [(800, "A", 900, "Так", 3)]))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].wholesale[0].threshold, 3)
        self.assertTrue(rows[0].move_current_to_old)

    def test_empty_cells_do_not_create_operation(self):
        with self.assertRaisesRegex(ValueError, "немає жодної команди"):
            parse_excel_prices(self.excel(["Артикул", "Поточна ціна"], [("A", "")]))

    def test_missing_required_article_header_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "обов'язковий стовпець"):
            parse_excel_prices(self.excel(["Поточна ціна"], [(100,)]))

    def test_duplicate_service_headers_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "повторюється"):
            parse_excel_prices(self.excel(["Артикул", "Опт 1", "Опт 1"], [("A", 100, 90)]))

    def test_template_contains_all_price_columns(self):
        workbook = load_workbook(io.BytesIO(build_excel_template()), read_only=True)
        headers = [cell.value for cell in next(workbook["Ціни"].iter_rows(max_row=1))]
        self.assertIn("Поточну ціну в РРЦ (Так)", headers)
        self.assertIn("Опт 5 від", headers)
        workbook.close()


if __name__ == "__main__":
    unittest.main()
