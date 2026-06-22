"""Merge several UTK budget spreadsheets.

Two products are produced from a set of budgets:

* :func:`merge_budgets` -- a new :class:`Budget` whose summable (dollar /
  person-month) fields are the element-wise sum across the inputs.  This drives
  the merged LaTeX definitions and justification.
* :func:`write_merged_workbook` -- a real ``.xlsx`` *in the same format as the
  input spreadsheets*.  Individual line items (senior personnel, postdocs,
  GRAs, equipment, ...) from every input are **concatenated** into the template
  sections; fixed categories (travel, supplies, tuition, ...) are **summed**.
  The template's subtotal/total/fringe/indirect formulas are left untouched so
  Excel recomputes the correct combined budget on open.  If a section has more
  line items than it has rows, the overflow is dropped and reported so the
  caller can raise a prominent warning.
"""

from __future__ import annotations

import warnings as _warnings
from dataclasses import dataclass
from typing import List, Optional, Sequence

from openpyxl import load_workbook

from . import extractor as X
from .extractor import Budget, Field

PERIOD_COLS = X.PERIOD_COLS
SHEET_NAME = X.SHEET_NAME


# ---------------------------------------------------------------------------
# Budget summing (used for the merged LaTeX definitions)
# ---------------------------------------------------------------------------

def merge_budgets(budgets: Sequence[Budget], source: str = "merged") -> Budget:
    """Return a new Budget summing the summable fields of ``budgets``."""
    merged = Budget(source=source)
    if not budgets:
        return merged

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
                # added; carry the first non-empty value forward.
                if target.value in (None, "") and field.value not in (None, ""):
                    target.value = field.value
                continue
            try:
                addend = float(field.value) if field.value not in (None, "") else 0.0
            except (TypeError, ValueError):
                continue
            target.value = (target.value or 0.0) + addend

    return merged


# ---------------------------------------------------------------------------
# Format-preserving workbook merge
# ---------------------------------------------------------------------------

# Personnel groups: (label, personnel rows, matching fringe rows, name_in_b).
# Fringe rows auto-reference their personnel row (B44=B11, L44=ROUND(L11*F44,0)),
# so only the per-person fringe *rate* (column F) needs to be written for fringe.
#
# ``name_in_b`` marks groups whose column B holds an actual person's name (the
# senior-personnel section).  In the "Other Personnel" sections column B instead
# holds a fixed role label ("Post Doc(s)", "GRA(s)", ...) that ships with the
# template, so it must NOT be treated as a name nor used to detect a real entry.
PERSONNEL_GROUPS = [
    ("Senior Personnel", X.SENIOR_ROWS, X.SENIOR_FRINGE_ROWS, True),
    ("Post Docs", X.POSTDOC_ROWS, X.POSTDOC_FRINGE_ROWS, False),
    ("Other Professionals", X.OTHER_PROF_ROWS, X.OTHER_PROF_FRINGE_ROWS, False),
    ("Graduate Research Assistants", X.GRA_ROWS, X.GRA_FRINGE_ROWS, False),
    ("Undergraduate Researchers", [X.UNDERGRAD_ROW], [X.UNDERGRAD_FRINGE_ROW], False),
    ("Admin/Clerical", [X.ADMIN_ROW], [X.ADMIN_FRINGE_ROW], False),
    ("Other Personnel", X.OTHER_STAFF_ROWS, X.OTHER_STAFF_FRINGE_ROWS, False),
]

# Single-line categories that are summed cell-by-cell across inputs.
LEAF_ROWS = (
    [X.DOMESTIC_TRAVEL_ROW, X.FOREIGN_TRAVEL_ROW, X.PARTICIPANT_SUPPORT_ROW]
    + list(X.OTHER_DIRECT_ROWS.values())
    + [X.TUITION_ROW, X.DIFFERENTIAL_TUITION_ROW, X.MANDATORY_FEES_ROW, X.MTDC_ROW]
)

# Columns copied for each personnel entry (besides the period columns).  The
# row total Q is left as its =SUM(L:P) formula and recomputes.  Column C
# (UT/JFO/GRA) is written on used rows but left untouched on unused ones so the
# template's dropdown defaults survive.
PERSONNEL_DATA_COLS = ["C", "D", "E", "F"]


@dataclass
class RowOverflow:
    """Reported when a section has more line items than it has rows."""
    section: str
    capacity: int
    needed: int
    dropped: List[str]


def _num(value) -> float:
    try:
        return float(value) if value not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _present_name(value) -> bool:
    return value not in (None, "") and bool(str(value).strip())


def _is_present(periods, base, total) -> bool:
    """A line item is "real" only if it carries money (not just a label)."""
    return any(_num(p) for p in periods) or _num(base) or _num(total)


def _read_personnel(ws, pers_rows, fringe_rows, name_in_b) -> List[dict]:
    entries = []
    for prow, frow in zip(pers_rows, fringe_rows):
        periods = [ws[f"{c}{prow}"].value for c in PERIOD_COLS]
        base = ws[f"D{prow}"].value
        total = ws[f"{X.TOTAL_COL}{prow}"].value
        if not _is_present(periods, base, total):
            continue
        entry = {col: ws[f"{col}{prow}"].value for col in PERSONNEL_DATA_COLS}
        entry["B"] = ws[f"B{prow}"].value if name_in_b else None
        entry["periods"] = periods
        entry["fringe_rate"] = ws[f"F{frow}"].value
        entries.append(entry)
    return entries


def _read_equipment(ws, rows) -> List[dict]:
    entries = []
    for r in rows:
        periods = [ws[f"{c}{r}"].value for c in PERIOD_COLS]
        name = ws[f"B{r}"].value
        total = ws[f"{X.TOTAL_COL}{r}"].value
        if _is_present(periods, None, total) or _present_name(name):
            entries.append({"B": name, "periods": periods})
    return entries


def _drop_label(entry, fallback) -> str:
    name = entry.get("B")
    return str(name) if _present_name(name) else fallback


def _write_personnel(ws, entries, pers_rows, fringe_rows, label,
                     name_in_b) -> Optional[RowOverflow]:
    capacity = len(pers_rows)
    clear_cols = (["B"] if name_in_b else []) + ["D", "E", "F"]

    # Clear each section row's money/input columns first (drops the per-person
    # salary formulas and any leftover template numbers); role labels in B for
    # the non-name groups, and column C dropdowns, are left intact.
    for prow, frow in zip(pers_rows, fringe_rows):
        for col in clear_cols:
            ws[f"{col}{prow}"] = None
        for c in PERIOD_COLS:
            ws[f"{c}{prow}"] = None
        ws[f"F{frow}"] = None  # fringe rate

    for k, entry in enumerate(entries[:capacity]):
        prow, frow = pers_rows[k], fringe_rows[k]
        if name_in_b:
            ws[f"B{prow}"] = entry["B"]
        for col in PERSONNEL_DATA_COLS:
            ws[f"{col}{prow}"] = entry[col]
        for c, value in zip(PERIOD_COLS, entry["periods"]):
            ws[f"{c}{prow}"] = value
        ws[f"F{frow}"] = entry["fringe_rate"]

    if len(entries) > capacity:
        dropped = [_drop_label(e, f"{label} entry {capacity + i + 1}")
                   for i, e in enumerate(entries[capacity:])]
        return RowOverflow(label, capacity, len(entries), dropped)
    return None


def _write_equipment(ws, entries, rows) -> Optional[RowOverflow]:
    capacity = len(rows)
    for r in rows:
        ws[f"B{r}"] = None
        for c in PERIOD_COLS:
            ws[f"{c}{r}"] = None
    for k, entry in enumerate(entries[:capacity]):
        r = rows[k]
        ws[f"B{r}"] = entry["B"]
        for c, value in zip(PERIOD_COLS, entry["periods"]):
            ws[f"{c}{r}"] = value
    if len(entries) > capacity:
        dropped = [_drop_label(e, f"Equipment item {capacity + i + 1}")
                   for i, e in enumerate(entries[capacity:])]
        return RowOverflow("Equipment", capacity, len(entries), dropped)
    return None


def _sum_leaf_rows(ws, input_wss, rows) -> None:
    for r in rows:
        for c in PERIOD_COLS:
            ws[f"{c}{r}"] = sum(_num(w[f"{c}{r}"].value) for w in input_wss)


def _set_merged_metadata(ws, input_wss, n_inputs) -> None:
    names = []
    for w in input_wss:
        v = w["D2"].value
        if _present_name(v):
            names.append(str(v).strip())
    if names:
        ws["D2"] = "; ".join(names)
    ws["D3"] = f"MERGED budget ({n_inputs} proposal(s)) -- detail tabs reflect template only"


def write_merged_workbook(
    input_paths: Sequence[str],
    out_path: str,
    template_path: Optional[str] = None,
) -> List[RowOverflow]:
    """Merge ``input_paths`` into a single budget workbook at ``out_path``.

    Returns a list of :class:`RowOverflow` for any section that ran out of
    rows (empty when everything fit).
    """
    input_paths = list(input_paths)
    if not input_paths:
        raise ValueError("write_merged_workbook: no input files")
    template_path = template_path or input_paths[0]

    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore")
        template_wb = load_workbook(template_path, data_only=False)
        value_wbs = [load_workbook(p, data_only=True) for p in input_paths]

    if SHEET_NAME not in template_wb.sheetnames:
        raise ValueError(f"{template_path!r}: missing worksheet {SHEET_NAME!r}")
    ws = template_wb[SHEET_NAME]
    value_sheets = [wb[SHEET_NAME] for wb in value_wbs]

    overflows: List[RowOverflow] = []

    for label, pers_rows, fringe_rows, name_in_b in PERSONNEL_GROUPS:
        pers_rows = list(pers_rows)
        fringe_rows = list(fringe_rows)
        entries: List[dict] = []
        for vws in value_sheets:
            entries.extend(_read_personnel(vws, pers_rows, fringe_rows, name_in_b))
        of = _write_personnel(ws, entries, pers_rows, fringe_rows, label, name_in_b)
        if of:
            overflows.append(of)

    eq_rows = list(X.EQUIPMENT_ROWS)
    eq_entries: List[dict] = []
    for vws in value_sheets:
        eq_entries.extend(_read_equipment(vws, eq_rows))
    of = _write_equipment(ws, eq_entries, eq_rows)
    if of:
        overflows.append(of)

    _sum_leaf_rows(ws, value_sheets, LEAF_ROWS)
    _set_merged_metadata(ws, value_sheets, len(input_paths))

    template_wb.save(out_path)
    return overflows
