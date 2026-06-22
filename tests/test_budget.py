"""Tests for utkbudget.

Run with::

    python -m unittest discover -s tests

The tests build a tiny synthetic workbook with known values written into the
cells the extractor reads, so they do not depend on the large binary template.
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook

from utkbudget import extractor
from utkbudget.extractor import extract_budget
from utkbudget.merge import merge_budgets, write_merged_workbook
from utkbudget.texdefs import format_value, tex_prefix, write_defs
from utkbudget.justification import build_document


def make_workbook(path, total=10000.0, travel_dom=1000.0, travel_for=500.0):
    """Write a minimal but structurally valid UTK Budget workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = extractor.SHEET_NAME

    # metadata
    ws["D2"] = "Dr. Test PI"
    ws["D3"] = "A & B Physics"      # contains an ampersand -> tests escaping

    def line(row, base):
        tot = 0.0
        for i, col in enumerate(extractor.PERIOD_COLS):
            v = base + i
            ws[f"{col}{row}"] = v
            tot += v
        ws[f"{extractor.TOTAL_COL}{row}"] = tot
        return tot

    line(extractor.SENIOR_SUBTOTAL_ROW, 100)
    line(extractor.WAGES_TOTAL_ROW, 200)
    line(extractor.FRINGE_TOTAL_ROW, 50)
    line(extractor.SALARY_BENEFITS_TOTAL_ROW, 250)
    line(extractor.DOMESTIC_TRAVEL_ROW, travel_dom)
    line(extractor.FOREIGN_TRAVEL_ROW, travel_for)
    line(extractor.TRAVEL_TOTAL_ROW, travel_dom + travel_for)
    line(extractor.DIRECT_TOTAL_ROW, 5000)
    line(extractor.INDIRECT_TOTAL_ROW, 2000)
    ws[f"{extractor.TOTAL_COL}{extractor.TOTAL_ROW}"] = total
    for col in extractor.PERIOD_COLS:
        ws[f"{col}{extractor.TOTAL_ROW}"] = total / len(extractor.PERIOD_COLS)

    # F&A rate + type
    for col in extractor.PERIOD_COLS:
        ws[f"{col}{extractor.FANDA_RATE_ROW}"] = 0.535
    ws[extractor.FANDA_RATE_TYPE_CELL] = "Research ON-Campus"

    wb.save(path)


class ExtractorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, "budget.xlsx")
        make_workbook(self.path, total=10000.0)

    def test_extracts_metadata_and_totals(self):
        b = extract_budget(self.path)
        self.assertEqual(b.get("PINames"), "Dr. Test PI")
        self.assertAlmostEqual(b.get("GrandTotal"), 10000.0)
        self.assertAlmostEqual(b.get("DomesticTravelTotal"), 5 * 1000.0 + (0 + 1 + 2 + 3 + 4))
        self.assertEqual(b.get("FandARateType"), "Research ON-Campus")
        # rate stored as a percentage
        self.assertAlmostEqual(b.get("OverheadRateYearOne"), 53.5)

    def test_no_doubled_total_macros(self):
        b = extract_budget(self.path)
        # The grand total uses base "Grand" so the Q macro is GrandTotal, never
        # "TotalTotal"; and aggregate bases never end in "Total".
        self.assertIn("GrandTotal", b.fields)
        self.assertIn("WagesTotal", b.fields)
        self.assertNotIn("TotalTotal", b.fields)
        self.assertNotIn("WagesTotalTotal", b.fields)


class MergeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.a = os.path.join(self.tmp, "a.xlsx")
        self.b = os.path.join(self.tmp, "b.xlsx")
        make_workbook(self.a, total=10000.0)
        make_workbook(self.b, total=25000.0)

    def test_sum_of_grand_totals(self):
        ba = extract_budget(self.a)
        bb = extract_budget(self.b)
        merged = merge_budgets([ba, bb])
        self.assertAlmostEqual(merged.get("GrandTotal"), 35000.0)

    def test_non_summable_carried_forward(self):
        merged = merge_budgets([extract_budget(self.a), extract_budget(self.b)])
        # rates are not summed (would be 107%); first value carried instead
        self.assertAlmostEqual(merged.get("OverheadRateYearOne"), 53.5)
        self.assertEqual(merged.get("FandARateType"), "Research ON-Campus")

    def test_write_merged_workbook(self):
        ba, bb = extract_budget(self.a), extract_budget(self.b)
        merged = merge_budgets([ba, bb])
        out = os.path.join(self.tmp, "merged.xlsx")
        write_merged_workbook(merged, [("a", ba), ("b", bb)], out)
        wb = load_workbook(out)
        self.assertEqual(wb.sheetnames, ["Merged", "a", "b"])


class TexTests(unittest.TestCase):
    def test_prefix_sanitization(self):
        self.assertEqual(tex_prefix("Proposal_Budget 2.xlsx"), "ProposalBudgetTwo")
        self.assertEqual(tex_prefix("123.xlsx"), "OneTwoThree")
        self.assertEqual(tex_prefix("!!!.xlsx"), "Budget")

    def test_format_value(self):
        from utkbudget.extractor import Field, KIND_MONEY, KIND_RATE, KIND_TEXT
        self.assertEqual(format_value(Field("x", 1234.5, KIND_MONEY)), "1,234.50")
        self.assertEqual(format_value(Field("x", None, KIND_MONEY)), "0.00")
        self.assertEqual(format_value(Field("x", 53.5, KIND_RATE)), "53.5")
        self.assertEqual(format_value(Field("x", "A & B", KIND_TEXT)), r"A \& B")

    def test_defs_and_document_macros_match(self):
        tmp = tempfile.mkdtemp()
        path = os.path.join(tmp, "budget.xlsx")
        make_workbook(path)
        b = extract_budget(path)
        defs = os.path.join(tmp, "b.tex")
        write_defs(b, "Test", defs)
        import re
        with open(defs) as fh:
            defs_text = fh.read()
        defined = set(re.findall(r"\\newcommand\{\\([A-Za-z]+)\}", defs_text))
        doc = build_document("T", ["b.tex"], [("Test", "Test Budget", False)])
        body = re.sub(r"\\input\{[^}]+\}", "", doc)
        used = set(re.findall(r"\\([A-Za-z]+)\{\}", body))
        undefined = used - defined
        self.assertEqual(undefined, set(), f"undefined macros: {undefined}")


if __name__ == "__main__":
    unittest.main()
