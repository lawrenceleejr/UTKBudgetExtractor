"""Merge several :class:`~utkbudget.extractor.Budget` objects.

Two products are produced from a list of budgets:

* :func:`merge_budgets` -- a new :class:`Budget` whose summable (dollar /
  person-month) fields are the element-wise sum across the inputs.  Aggregate
  totals (salaries, travel, indirect, grand total, ...) are exactly what a
  combined proposal needs.
* :func:`write_merged_workbook` -- a clean ``.xlsx`` with one "Merged" summary
  sheet plus one sheet per input budget, so the combined request and its
  provenance are both visible in a single file.
"""

from __future__ import annotations

from collections import OrderedDict
from typing import Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .extractor import Budget, Field, PERIOD_WORDS

# The category rows shown in the merged workbook, as
# (display label, macro base).  ``<base>YearOne..Five`` and ``<base>Total``
# are pulled for each.
SUMMARY_ROWS: "OrderedDict[str, str]" = OrderedDict([
    ("Senior Personnel", "SeniorSubtotal"),
    ("Other Personnel", "OtherPersonnelSubtotal"),
    ("Total Salaries & Wages", "Wages"),
    ("Fringe Benefits", "Fringe"),
    ("Total Salaries & Benefits", "SalaryAndBenefits"),
    ("Equipment", "Equipment"),
    ("Domestic Travel", "DomesticTravel"),
    ("Foreign Travel", "ForeignTravel"),
    ("Total Travel", "Travel"),
    ("Participant Support", "ParticipantSupport"),
    ("Supplies", "Supplies"),
    ("Publication", "Publication"),
    ("Subcontracts", "Subcontracts"),
    ("Tuition & Fees", "TuitionSubtotal"),
    ("Total Other Direct Costs", "OtherDirect"),
    ("Total Direct Costs", "Direct"),
    ("Modified Total Direct Costs", "ModifiedTotalDirectCosts"),
    ("Indirect (F&A) Costs", "Indirect"),
    ("Total Costs", "Grand"),
])


def merge_budgets(budgets: Sequence[Budget], source: str = "merged") -> Budget:
    """Return a new Budget summing the summable fields of ``budgets``."""
    merged = Budget(source=source)
    if not budgets:
        return merged

    # Preserve the field order/kinds of the first budget.
    for name, field in budgets[0].fields.items():
        merged.fields[name] = Field(name, None, field.kind)

    for budget in budgets:
        for name, field in budget.fields.items():
            target = merged.fields.get(name)
            if target is None:
                target = Field(name, None, field.kind)
                merged.fields[name] = target
            if not field.summable:
                # Non-summable fields (rates, F&A type, dates, names) cannot be
                # added; carry the first non-empty value forward so things like
                # the F&A rate and rate type survive into the merged budget.
                if target.value in (None, "") and field.value not in (None, ""):
                    target.value = field.value
                continue
            try:
                addend = float(field.value) if field.value not in (None, "") else 0.0
            except (TypeError, ValueError):
                continue
            target.value = (target.value or 0.0) + addend

    return merged


def _money(budget: Budget, base: str, word: str):
    val = budget.get(f"{base}Year{word}")
    try:
        return float(val) if val not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _total(budget: Budget, base: str):
    val = budget.get(f"{base}Total")
    try:
        return float(val) if val not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _write_summary_sheet(ws, title: str, budget: Budget) -> None:
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF")
    bold = Font(bold=True)
    money_fmt = "#,##0.00"
    thin = Side(style="thin", color="BFBFBF")
    border = Border(bottom=thin)
    right = Alignment(horizontal="right")

    ws.title = title[:31]
    ws["A1"] = "Category"
    headers = ["Category"] + [f"Period {i + 1}" for i in range(len(PERIOD_WORDS))] + ["Total"]
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = right if col > 1 else Alignment(horizontal="left")

    grand_total_labels = {"Total Costs", "Total Direct Costs", "Total Salaries & Benefits"}
    for r, (label, base) in enumerate(SUMMARY_ROWS.items(), start=2):
        ws.cell(row=r, column=1, value=label)
        if label in grand_total_labels:
            ws.cell(row=r, column=1).font = bold
        for c, word in enumerate(PERIOD_WORDS, start=2):
            cell = ws.cell(row=r, column=c, value=_money(budget, base, word))
            cell.number_format = money_fmt
            cell.alignment = right
            if label in grand_total_labels:
                cell.font = bold
        total_cell = ws.cell(row=r, column=2 + len(PERIOD_WORDS), value=_total(budget, base))
        total_cell.number_format = money_fmt
        total_cell.alignment = right
        total_cell.font = bold
        if label in grand_total_labels:
            for c in range(1, 3 + len(PERIOD_WORDS)):
                ws.cell(row=r, column=c).border = border

    # Column widths
    ws.column_dimensions["A"].width = 32
    for col in range(2, 3 + len(PERIOD_WORDS)):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.freeze_panes = "B2"


def write_merged_workbook(
    merged: Budget,
    per_input: Sequence[Tuple[str, Budget]],
    out_path: str,
) -> str:
    """Write the merged summary workbook.

    ``per_input`` is a list of ``(sheet_label, budget)`` for the contributing
    files, each given its own sheet after the merged summary.
    """
    wb = Workbook()
    _write_summary_sheet(wb.active, "Merged", merged)

    used = {"Merged"}
    for label, budget in per_input:
        # Excel sheet names must be unique and <= 31 chars.
        name = (label or "Budget")[:31]
        suffix = 1
        while name in used:
            suffix += 1
            name = f"{label[:28]}_{suffix}"
        used.add(name)
        ws = wb.create_sheet(title=name)
        _write_summary_sheet(ws, name, budget)

    wb.save(out_path)
    return out_path
