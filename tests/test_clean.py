"""Tests for the cleaning layer and the report builder.

Run:  python -m unittest discover -s tests -v
"""
import csv
import os
import sys
import tempfile
import unittest
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import build_report  # noqa: E402
import clean  # noqa: E402

HEADER = ["Order ID", "Order Date", "Region", "Sales Rep", "Product", "Category", "Units", "Unit Price", "Discount", "Amount"]


def write_csv(rows, header=HEADER, delim=","):
    d = tempfile.mkdtemp()
    p = os.path.join(d, "export.csv")
    with open(p, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=delim)
        w.writerow(header)
        w.writerows(rows)
    return p


class MoneyTests(unittest.TestCase):
    def test_amount_rounds_half_up_in_cents(self):
        self.assertEqual(clean.calc_amount_c(3, 3333, 0.05), 9499)   # 3 x 3333 = 9999; 9999 x 0.95 = 9499.05 -> 9499
        self.assertEqual(clean.calc_amount_c(1, 1005, 0.5), 503)     # 502.5 rounds half up to 503

    def test_formatted_text_numbers(self):
        self.assertEqual(clean.parse_money_c("$1,299.50"), (129950, True))
        self.assertEqual(clean.parse_money_c(12.5), (1250, False))
        self.assertEqual(clean.parse_discount("10%"), (Decimal("0.1"), True))
        self.assertEqual(clean.parse_discount(0.15), (Decimal("0.15"), False))


class CleanFileTests(unittest.TestCase):
    ROWS = [
        ["SO-1", "2025-03-05", " north ", "Ana", "Desk", "Furniture", "2", "$100.00", "10%", "$180.00"],
        ["SO-2", "2025-03-14", "S", "Ben", "Lamp", "Lighting", "1", "$40.00", "0%", ""],            # amount missing
        ["SO-1", "2025-03-05", " north ", "Ana", "Desk", "Furniture", "2", "$100.00", "10%", "$180.00"],  # duplicate
        ["", "", "", "", "", "", "", "", "", ""],                                                  # blank row
        ["", "2025-03-20", "East", "Cy", "Chair", "Furniture", "1", "$50.00", "0%", "$50.00"],     # no order id
        ["Total", "", "", "", "", "", "", "", "", "$220.00"],                                      # footer
    ]

    def setUp(self):
        self.records, self.log = clean.clean_file(write_csv(self.ROWS))

    def test_kept_orders(self):
        self.assertEqual([r["order_id"] for r in self.records], ["SO-1", "SO-2"])

    def test_every_discarded_row_is_counted_once(self):
        lg = self.log
        self.assertEqual((lg["blank_rows"], lg["footer_rows"], lg["duplicates_removed"], lg["missing_id_rejected"]),
                         (1, 1, 1, 1))
        self.assertEqual(lg["rows_read"], lg["blank_rows"] + lg["footer_rows"] + lg["duplicates_removed"]
                         + lg["missing_id_rejected"] + lg["rows_kept"])

    def test_region_labels_standardised(self):
        self.assertEqual([r["region"] for r in self.records], ["North", "South"])

    def test_missing_amount_is_recomputed(self):
        self.assertEqual(self.log["amounts_recomputed"], 1)
        self.assertEqual(self.records[1]["amount_c"], 4000)

    def test_dates_and_cents(self):
        self.assertEqual(self.records[0]["order_date"].isoformat(), "2025-03-05")
        self.assertEqual(self.records[0]["amount_c"], 18000)


class DateTests(unittest.TestCase):
    def test_day_first_is_detected_from_the_data(self):
        self.assertEqual(clean.detect_date_format(["03/04/2025", "14/04/2025"]), ("%d/%m/%Y", False))

    def test_month_first_is_detected_from_the_data(self):
        self.assertEqual(clean.detect_date_format(["03/04/2025", "04/14/2025"]), ("%m/%d/%Y", False))

    def test_truly_ambiguous_dates_are_flagged(self):
        self.assertTrue(clean.detect_date_format(["03/04/2025", "05/06/2025"])[1])

    def test_text_dates(self):
        self.assertEqual(clean.detect_date_format(["Mar 05, 2025"])[0], "%b %d, %Y")


class HeaderTests(unittest.TestCase):
    def test_three_header_spellings_map_to_one_schema(self):
        for header in ([h.lower().replace(" ", "_") for h in HEADER],
                       ["OrderID", "OrderDate", "Area", "Rep", "Item", "Cat", "Units", "UnitPrice", "Disc", "Total"],
                       ["order_id", "date", "region", "rep", "product", "category", "qty", "price", "disc_pct", "net_amount"]):
            self.assertEqual(set(clean.map_header(header)), set(clean.SCHEMA))

    def test_unknown_header_is_rejected(self):
        with self.assertRaises(ValueError):
            clean.map_header(["foo", "bar"])


class FormulaTests(unittest.TestCase):
    def test_monthly_formula_is_a_sumifs_over_the_clean_data_sheet(self):
        f = build_report.monthly_formula(11443, 5, "B")
        self.assertTrue(f.startswith("=SUMIFS('Clean Data'!$J$2:$J$11443,"))
        self.assertIn('"<"&EDATE($A5,1)', f)
        self.assertTrue(f.endswith("B$4)"))


if __name__ == "__main__":
    unittest.main()
