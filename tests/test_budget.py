"""Tests for utkbudget.

Run with::

    python -m unittest discover -s tests

Extractor/LaTeX tests use a tiny synthetic workbook with known values written
into the cells the extractor reads.  Workbook-merge tests copy the shipped
template (``examples/Proposal_Budget_Basic.xlsx``) so the real formulas are
present and formula-safety can be asserted for real.
"""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook

from utkbudget import extractor
from utkbudget.extractor import extract_budget
from utkbudget.merge import ValueConflict, merge_budgets, write_merged_workbook
from utkbudget.texdefs import format_value, tex_prefix, write_defs
from utkbudget.justification import build_document


def make_workbook(path, total=10000.0, travel_dom=1000.0, travel_for=500.0,
                  seniors=None):
    """Write a minimal but structurally valid UTK Budget workbook.

    ``seniors`` is an optional list of names written into the senior-personnel
    rows (used to exercise the concatenating merge and its row-overflow path).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = extractor.SHEET_NAME

    # metadata
    ws["D2"] = "Dr. Test PI"
    ws["D3"] = "A & B Physics"      # contains an ampersand -> tests escaping

    for i, name in enumerate(seniors or []):
        row = list(extractor.SENIOR_ROWS)[i]
        ws[f"B{row}"] = name
        ws[f"D{row}"] = 100000
        for col in extractor.PERIOD_COLS:
            ws[f"{col}{row}"] = 50000

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


TEMPLATE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "examples", "Proposal_Budget_Basic.xlsx")


def make_real_input(path, seniors=(), domestic_airfares=(), subcontracts=(),
                    gras=0, tuition=None):
    """Copy the shipped template (real formulas intact) and set user inputs.

    ``seniors`` is a list of ``(name, base_annual, person_months)``; each is
    written into a senior-personnel row with the same months in every period.
    ``domestic_airfares`` go into the Period-1 domestic travel rows on the
    TRAVEL sheet.  ``subcontracts`` is a list of ``(institution, amount)``.
    ``gras`` fills that many GRA rows; ``tuition`` sets the annual per-GRA
    tuition cost.
    """
    shutil.copy(TEMPLATE, path)
    wb = load_workbook(path)  # keep formulas
    ws = wb[extractor.SHEET_NAME]
    senior_rows = list(extractor.SENIOR_ROWS)
    for i, (name, base, months) in enumerate(seniors):
        r = senior_rows[i]
        ws[f"B{r}"] = name
        ws[f"C{r}"] = "UT"
        ws[f"D{r}"] = base
        ws[f"E{r}"] = 9
        for col in ["F", "G", "H", "I", "J"]:   # person-months, periods 1-5
            ws[f"{col}{r}"] = months
    for i in range(gras):
        r = list(extractor.GRA_ROWS)[i]
        ws[f"D{r}"] = 2500
        ws[f"E{r}"] = 1
        for col in ["F", "G", "H", "I", "J"]:
            ws[f"{col}{r}"] = 12
    if tuition is not None:
        ws[f"F{extractor.TUITION_ROW}"] = tuition
    if domestic_airfares:
        tr = wb["TRAVEL"]
        for i, air in enumerate(domestic_airfares):
            row = 4 + i  # Period-1 domestic entry rows start at 4
            tr[f"E{row}"] = 1     # travelers
            tr[f"G{row}"] = air   # airfare
    if subcontracts:
        sc = wb["SUBCONTRACTS"]
        for i, (inst, amount) in enumerate(subcontracts):
            row = 3 + i
            sc[f"B{row}"] = inst
            sc[f"E{row}"] = amount
            sc[f"K{row}"] = "Y"
    wb.save(path)
    return path


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

    def test_merged_workbook_is_formula_safe(self):
        # Build two realistic inputs from the shipped template (which has the
        # real formulas) and confirm the merge never clobbers a formula.
        a = make_real_input(os.path.join(self.tmp, "ra.xlsx"),
                            seniors=[("Alice", 100000, 3)])
        b = make_real_input(os.path.join(self.tmp, "rb.xlsx"),
                            seniors=[("Bob", 120000, 2)])
        out = os.path.join(self.tmp, "rmerged.xlsx")
        write_merged_workbook([a, b], out)
        wb = load_workbook(out, data_only=False)
        ws = wb[extractor.SHEET_NAME]
        # Every one of these is a formula in the template and must stay a formula.
        for coord in ["L11", "Q11", "L23", "L44", "L70", "L71",
                      "L81", "L88", "L94", "L99", "L105", "L108", "L110", "L111"]:
            v = ws[coord].value
            self.assertTrue(isinstance(v, str) and v.startswith("="),
                            f"{coord} should still be a formula, got {v!r}")

    def test_personnel_user_inputs_concatenated(self):
        a = make_real_input(os.path.join(self.tmp, "pa.xlsx"),
                            seniors=[("Alice", 100000, 3), ("Bob", 120000, 2)])
        b = make_real_input(os.path.join(self.tmp, "pb.xlsx"),
                            seniors=[("Carol", 90000, 1)])
        out = os.path.join(self.tmp, "pmerged.xlsx")
        self.assertEqual(write_merged_workbook([a, b], out), [])
        ws = load_workbook(out)[extractor.SHEET_NAME]
        rows = list(extractor.SENIOR_ROWS)
        self.assertEqual([ws[f"B{rows[i]}"].value for i in range(3)],
                         ["Alice", "Bob", "Carol"])
        # The base-salary *input* is copied; the salary cell stays a formula.
        self.assertEqual([ws[f"D{rows[i]}"].value for i in range(3)],
                         [100000, 120000, 90000])
        self.assertTrue(str(ws[f"L{rows[0]}"].value).startswith("="))
        # Person-months are copied for ALL five periods (the period 2-5 salary
        # formulas read columns G-J), and the raise flag in column A survives.
        for col in ["F", "G", "H", "I", "J"]:
            self.assertEqual([ws[f"{col}{rows[i]}"].value for i in range(3)],
                             [3, 2, 1], f"person-months column {col}")
        self.assertEqual(ws[f"A{rows[0]}"].value, "Yes")

    def test_subcontracts_concatenated_with_flag(self):
        # Column K (Y/N) must travel with each subcontract row or the sheet's
        # own validation breaks the MTDC base.
        a = make_real_input(os.path.join(self.tmp, "sa.xlsx"),
                            subcontracts=[("ORNL", 50000)])
        b = make_real_input(os.path.join(self.tmp, "sb.xlsx"),
                            subcontracts=[("FNAL", 30000)])
        out = os.path.join(self.tmp, "smerged.xlsx")
        write_merged_workbook([a, b], out)
        sc = load_workbook(out)["SUBCONTRACTS"]
        self.assertEqual([sc["B3"].value, sc["B4"].value], ["ORNL", "FNAL"])
        self.assertEqual([sc["E3"].value, sc["E4"].value], [50000, 30000])
        self.assertEqual([sc["K3"].value, sc["K4"].value], ["Y", "Y"])

    def test_conflicting_tuition_reported(self):
        # Two inputs both budget GRAs but disagree on the per-GRA tuition cost.
        # The formulas apply one cost to every merged GRA, so the merge must
        # flag the conflict (and carry a non-zero value).
        a = make_real_input(os.path.join(self.tmp, "ca.xlsx"), gras=1, tuition=12000)
        b = make_real_input(os.path.join(self.tmp, "cb.xlsx"), gras=1, tuition=0)
        out = os.path.join(self.tmp, "cmerged.xlsx")
        issues = write_merged_workbook([a, b], out)
        conflicts = [i for i in issues if isinstance(i, ValueConflict)]
        self.assertEqual(len(conflicts), 1)
        self.assertIn("Tuition", conflicts[0].section)
        ws = load_workbook(out)[extractor.SHEET_NAME]
        self.assertEqual(ws[f"F{extractor.TUITION_ROW}"].value, 12000)

    def test_consistent_tuition_no_conflict(self):
        a = make_real_input(os.path.join(self.tmp, "na.xlsx"), gras=2, tuition=12000)
        b = make_real_input(os.path.join(self.tmp, "nb.xlsx"), gras=1, tuition=12000)
        out = os.path.join(self.tmp, "nmerged.xlsx")
        issues = write_merged_workbook([a, b], out)
        self.assertEqual([i for i in issues if isinstance(i, ValueConflict)], [])
        # GRA rows concatenated: 3 of 4 slots filled
        ws = load_workbook(out)[extractor.SHEET_NAME]
        gra_rows = list(extractor.GRA_ROWS)
        self.assertEqual([ws[f"D{r}"].value for r in gra_rows],
                         [2500, 2500, 2500, None])

    def test_no_formula_cell_modified_anywhere(self):
        # The rock-solid invariant: after a merge, every formula cell on every
        # sheet is byte-for-byte identical to the template's.
        a = make_real_input(os.path.join(self.tmp, "ia.xlsx"),
                            seniors=[("Alice", 100000, 3)],
                            domestic_airfares=[500],
                            subcontracts=[("ORNL", 50000)])
        b = make_real_input(os.path.join(self.tmp, "ib.xlsx"),
                            seniors=[("Bob", 120000, 2)],
                            domestic_airfares=[700])
        out = os.path.join(self.tmp, "imerged.xlsx")
        write_merged_workbook([a, b], out)
        template = load_workbook(a)  # first input is the structural template
        merged = load_workbook(out)
        checked = 0
        for sheet in template.sheetnames:
            ws_t, ws_m = template[sheet], merged[sheet]
            for row in ws_t.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v.startswith("="):
                        self.assertEqual(
                            ws_m[cell.coordinate].value, v,
                            f"formula clobbered at {sheet}!{cell.coordinate}")
                        checked += 1
        self.assertGreater(checked, 1000)  # sanity: we really scanned formulas

    def test_fringe_rates_untouched(self):
        # Fringe rates (column F on the fringe rows) ship with the template and
        # must not be modified by the merge.
        a = make_real_input(os.path.join(self.tmp, "fa.xlsx"),
                            seniors=[("Alice", 100000, 3)])
        b = make_real_input(os.path.join(self.tmp, "fb.xlsx"),
                            seniors=[("Bob", 120000, 2)])
        # Put a sentinel fringe rate in the template (the first input).
        wb = load_workbook(a)
        wb[extractor.SHEET_NAME]["F44"] = 0.42
        wb.save(a)
        out = os.path.join(self.tmp, "fmerged.xlsx")
        write_merged_workbook([a, b], out)
        merged = load_workbook(out)[extractor.SHEET_NAME]
        template = load_workbook(a)[extractor.SHEET_NAME]
        # The sentinel survives, and every senior fringe rate equals the
        # template's (i.e. the merge left them all untouched).
        self.assertEqual(merged["F44"].value, 0.42)
        for r in range(44, 56):
            self.assertEqual(merged[f"F{r}"].value, template[f"F{r}"].value)

    def test_travel_entries_concatenated(self):
        a = make_real_input(os.path.join(self.tmp, "ta.xlsx"), domestic_airfares=[500])
        b = make_real_input(os.path.join(self.tmp, "tb.xlsx"), domestic_airfares=[700])
        out = os.path.join(self.tmp, "tmerged.xlsx")
        write_merged_workbook([a, b], out)
        wb = load_workbook(out)
        tr = wb["TRAVEL"]
        # Period-1 domestic entries start at row 4; both files' airfares stack.
        self.assertEqual(tr["G4"].value, 500)
        self.assertEqual(tr["G5"].value, 700)
        # The per-row total stays a formula.
        self.assertTrue(str(tr["K4"].value).startswith("="))

    def test_row_overflow_warns(self):
        cap = len(list(extractor.SENIOR_ROWS))
        a = make_real_input(os.path.join(self.tmp, "big.xlsx"),
                            seniors=[(f"PI{i}", 80000, 1) for i in range(cap)])
        b = make_real_input(os.path.join(self.tmp, "one.xlsx"),
                            seniors=[("Overflow", 80000, 1)])
        out = os.path.join(self.tmp, "omerged.xlsx")
        overflows = write_merged_workbook([a, b], out)
        senior = [o for o in overflows if o.section == "Senior Personnel"]
        self.assertEqual(len(senior), 1)
        self.assertEqual(senior[0].capacity, cap)
        self.assertEqual(senior[0].needed, cap + 1)
        self.assertEqual(senior[0].dropped, ["Overflow"])


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


class ProvenanceTests(unittest.TestCase):
    def test_defs_header_has_timestamp_and_commit(self):
        tmp = tempfile.mkdtemp()
        path = os.path.join(tmp, "budget.xlsx")
        make_workbook(path)
        defs = os.path.join(tmp, "b.tex")
        write_defs(extract_budget(path), "Test", defs)
        with open(defs) as fh:
            head = fh.read(400)
        self.assertIn("Auto-generated by UTKBudgetExtractor", head)
        self.assertIn("Generator commit:", head)

    def test_document_header_before_documentclass(self):
        doc = build_document("T", [], [])
        # Provenance comments must precede \documentclass (valid LaTeX).
        self.assertTrue(doc.lstrip().startswith("%"))
        self.assertLess(doc.index("Generator commit:"), doc.index("\\documentclass"))


if __name__ == "__main__":
    unittest.main()
