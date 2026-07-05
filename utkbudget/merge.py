"""Merge several UTK budget spreadsheets.

Two products are produced from a set of budgets:

* :func:`merge_budgets` -- a new :class:`Budget` whose summable (dollar /
  person-month) fields are the element-wise sum across the inputs.  This drives
  the merged LaTeX definitions and justification.
* :func:`write_merged_workbook` -- a real ``.xlsx`` *in the same format as the
  input spreadsheets*.

The workbook merge is **formula-safe**: it never writes into a cell that holds a
formula.  Only genuine *user-input* cells are copied -- on the main sheet the
personnel inputs (name, raise flag, UT/JFO, base salary, appointment, per-period
person-months, tenure flags), equipment descriptions/costs, the manually-entered
other-direct amounts, and the per-GRA tuition costs; and on the TRAVEL /
SUPPLIES / SUBCONTRACTS / PARTICIPANT SUPPORT detail sheets the per-line
entries.  The fringe-rate cells are institutional constants that ship with the
template and are never touched.  Every subtotal, total, salary, fringe, tuition,
F&A, and roll-up formula is left untouched so Excel recomputes the correct
combined budget on open.

Line items are *concatenated* into the template's fixed sections.  If a section
has more line items than rows, the overflow is dropped and reported so the
caller can raise a prominent warning.
"""

from __future__ import annotations

import os
import warnings as _warnings
from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple, Union

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
# Formula-safe workbook merge
# ---------------------------------------------------------------------------

@dataclass
class RowOverflow:
    """Reported when a section has more line items than it has rows."""
    section: str
    capacity: int
    needed: int
    dropped: List[str]


@dataclass
class ValueConflict:
    """Reported when inputs disagree on a workbook-wide (non-summable) input."""
    section: str
    detail: str


MergeIssue = Union[RowOverflow, ValueConflict]


def _num(value) -> float:
    try:
        return float(value) if value not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _present_name(value) -> bool:
    return value not in (None, "") and bool(str(value).strip())


def _is_formula(cell) -> bool:
    return cell.data_type == "f" or (
        isinstance(cell.value, str) and cell.value.startswith("=")
    )


def _set(ws, row: int, col: str, value) -> bool:
    """Set ``ws[col][row]`` to ``value`` -- unless it holds a formula.

    Returns ``True`` if the cell was written.  This is the single guarantee
    that no formula is ever clobbered by the merge.
    """
    cell = ws[f"{col}{row}"]
    if _is_formula(cell):
        return False
    cell.value = value
    return True


def _resolve_sheet(wb, name: str):
    """Return the worksheet matching ``name`` (tolerant of trailing spaces)."""
    if name in wb.sheetnames:
        return wb[name]
    for sn in wb.sheetnames:
        if sn.strip() == name.strip():
            return wb[sn]
    return None


# -- Personnel ---------------------------------------------------------------

# (label, personnel rows, name_in_b).  ``name_in_b`` marks groups whose column B
# holds an actual person's name (the senior-personnel section); elsewhere column
# B is a fixed role label ("Post Doc(s)", "GRA(s)", ...) that ships with the
# template and is left alone.  The fringe rows (44-69) are never touched: their
# names and amounts are formulas, and the rates are the institution's standard
# rates that ship with the template.
PERSONNEL_GROUPS = [
    ("Senior Personnel", X.SENIOR_ROWS, True),
    ("Post Docs", X.POSTDOC_ROWS, False),
    ("Other Professionals", X.OTHER_PROF_ROWS, False),
    ("Graduate Research Assistants", X.GRA_ROWS, False),
    ("Undergraduate Researchers", [X.UNDERGRAD_ROW], False),
    ("Admin/Clerical", [X.ADMIN_ROW], False),
    ("Other Personnel", X.OTHER_STAFF_ROWS, False),
]

# User-input columns feeding the salary formulas, per the template:
#   A = apply-annual-raise flag ("Yes"/"No")
#   C = UT / JFO / GRA (selects the inflation rate)
#   D = base salary,  E = appointment divisor or headcount
#   F..J = person-months for periods 1-5
# The period salary columns L-P are formulas over these and recompute.
PERSONNEL_INPUT_COLS = ["A", "C", "D", "E", "F", "G", "H", "I", "J"]
# Tenure inputs referenced only by the senior-personnel salary formulas.
SENIOR_TENURE_COLS = ["S", "T", "U"]
PERSON_MONTH_COLS = ["F", "G", "H", "I", "J"]


def _merge_personnel(ws_t, value_sheets) -> List[RowOverflow]:
    overflows = []
    for label, pers_rows, name_in_b in PERSONNEL_GROUPS:
        pers_rows = list(pers_rows)
        cols = list(PERSONNEL_INPUT_COLS)
        if name_in_b:
            cols = ["B"] + cols + SENIOR_TENURE_COLS

        # An entry is "real" if it carries a base salary, any person-months, or
        # any cached period salary.
        entries = []
        for vs in value_sheets:
            for prow in pers_rows:
                signal = ([vs[f"D{prow}"].value]
                          + [vs[f"{c}{prow}"].value for c in PERSON_MONTH_COLS]
                          + [vs[f"{c}{prow}"].value for c in PERIOD_COLS])
                if not any(_num(v) for v in signal):
                    continue
                entry = {c: vs[f"{c}{prow}"].value for c in cols}
                entry["_name"] = vs[f"B{prow}"].value
                entries.append(entry)

        # Clear each slot's input columns, then stack the entries (formula-safe
        # throughout: _set never writes into a formula cell).
        for prow in pers_rows:
            for c in cols:
                _set(ws_t, prow, c, None)

        capacity = len(pers_rows)
        for k, entry in enumerate(entries[:capacity]):
            prow = pers_rows[k]
            for c in cols:
                _set(ws_t, prow, c, entry[c])

        if len(entries) > capacity:
            dropped = []
            for i, e in enumerate(entries[capacity:]):
                nm = e.get("_name")
                dropped.append(str(nm) if _present_name(nm) else f"{label} entry {capacity + i + 1}")
            overflows.append(RowOverflow(label, capacity, len(entries), dropped))
    return overflows


# -- Generic stacking of line-item entry rows --------------------------------

@dataclass
class StackSpec:
    """A set of repeating, concatenable entry slots on a sheet.

    ``anchors`` are the base rows of each slot; ``cells`` are ``(row_delta, col)``
    user-input cells to copy per entry; ``present`` is the subset used to decide
    whether a slot is a real (funded) entry.
    """
    sheet: str
    anchors: List[int]
    cells: List[Tuple[int, str]]
    present: List[Tuple[int, str]]
    label: str
    name_cell: Optional[Tuple[int, str]] = None


def _stack(ws_t, value_sheets, spec: StackSpec) -> Optional[RowOverflow]:
    entries = []
    for vs in value_sheets:
        for a in spec.anchors:
            if any(_num(vs[f"{c}{a + d}"].value) for d, c in spec.present):
                entry = {(d, c): vs[f"{c}{a + d}"].value for d, c in spec.cells}
                if spec.name_cell:
                    nd, nc = spec.name_cell
                    entry["_name"] = vs[f"{nc}{a + nd}"].value
                entries.append(entry)

    capacity = len(spec.anchors)
    for a in spec.anchors:                       # clear slots (formula-safe)
        for d, c in spec.cells:
            _set(ws_t, a + d, c, None)
    for k, entry in enumerate(entries[:capacity]):
        a = spec.anchors[k]
        for key, value in entry.items():
            if key == "_name":
                continue
            d, c = key
            _set(ws_t, a + d, c, value)
    if len(entries) > capacity:
        dropped = []
        for i, e in enumerate(entries[capacity:]):
            nm = e.get("_name")
            dropped.append(str(nm) if _present_name(nm) else f"{spec.label} entry {capacity + i + 1}")
        return RowOverflow(spec.label, capacity, len(entries), dropped)
    return None


def _main_sheet_specs() -> List[StackSpec]:
    """Concatenable sections that live on the main UTK Budget sheet."""
    periods = PERIOD_COLS
    equip_cells = [(0, "B")] + [(0, c) for c in periods]
    return [
        StackSpec(SHEET_NAME, list(X.EQUIPMENT_ROWS), equip_cells,
                  [(0, c) for c in periods], "Equipment", name_cell=(0, "B")),
    ]


def _travel_specs(sheet: str) -> List[StackSpec]:
    """TRAVEL: 5 period blocks, each with 10 domestic + 5 foreign entry rows."""
    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "L"]
    cells = [(0, c) for c in cols]
    money = [(0, c) for c in ["D", "E", "F", "G", "H", "I", "J"]]
    specs = []
    for p in range(5):
        base = 1 + p * 22
        dom = list(range(base + 3, base + 13))       # 10 domestic rows
        foreign = list(range(base + 15, base + 20))   # 5 foreign rows
        specs.append(StackSpec(sheet, dom, cells, money,
                               f"TRAVEL Period {p + 1} Domestic", name_cell=(0, "A")))
        specs.append(StackSpec(sheet, foreign, cells, money,
                               f"TRAVEL Period {p + 1} Foreign", name_cell=(0, "A")))
    return specs


def _supplies_specs(sheet: str) -> List[StackSpec]:
    cols = ["A", "B", "C", "D", "E", "F"]            # A=desc, B-F=periods
    return [StackSpec(sheet, list(range(3, 38)), [(0, c) for c in cols],
                      [(0, c) for c in ["B", "C", "D", "E", "F"]],
                      "Supplies", name_cell=(0, "A"))]


def _subcontracts_specs(sheet: str) -> List[StackSpec]:
    # B/C = institution & lead, E-I = per-period amounts, K = the required
    # Y/N flag (the sheet's own validation errors out -- breaking the MTDC
    # base -- whenever a row has an institution but no K value).
    cols = ["B", "C", "E", "F", "G", "H", "I", "K"]
    return [StackSpec(sheet, list(range(3, 18)), [(0, c) for c in cols],
                      [(0, c) for c in ["E", "F", "G", "H", "I"]],
                      "Subcontracts", name_cell=(0, "B"))]


def _participant_specs(sheet, ws) -> List[StackSpec]:
    """PARTICIPANT SUPPORT: a fixed number of participant-cost blocks (each a
    description, a "Number of Participants" row, and per-participant cost
    categories).

    Blocks are located by their header label rather than assumed to repeat on a
    fixed stride, so we never write past the real blocks (the sheet ends them
    with a GRAND TOTALS row)."""
    anchors = [r for r in range(1, ws.max_row + 1)
               if isinstance(ws[f"A{r}"].value, str)
               and ws[f"A{r}"].value.strip().startswith("PARTICIPANT SUPPORT")]
    desc = [(0, "D")]                                   # workshop description
    num = [(2, c) for c in ["F", "G", "H", "I", "J"]]   # participants per period
    costs = [(d, "B") for d in range(6, 12)]            # cost per participant
    return [StackSpec(sheet, anchors, desc + num + costs, num,
                      "Participant Support", name_cell=(0, "D"))]


# -- Single-value leaves -----------------------------------------------------

# Manually-entered other-direct rows (Publication, Shipping, ...).  Their period
# cells are plain numbers and are summed across inputs; the Supplies and
# Subcontracts rows are formulas (driven by their detail sheets) and left alone.
MANUAL_OTHER_DIRECT_ROWS = [row for name, row in X.OTHER_DIRECT_ROWS.items()
                            if name not in ("Supplies", "Subcontracts")]
# Per-GRA tuition / fee costs (column F).  These are *global* inputs: the
# tuition formulas apply them to every GRA row, so they can only be carried --
# and only honestly when every input that budgets GRAs uses the same values.
TUITION_COST_CELLS = [(X.TUITION_ROW, "F", "Tuition (annual cost per GRA)"),
                      (X.DIFFERENTIAL_TUITION_ROW, "F", "Differential tuition per GRA"),
                      (X.MANDATORY_FEES_ROW, "F", "Mandatory fees per GRA")]

# Other workbook-wide inputs that cannot be summed, only carried from the first
# input.  If the inputs disagree, the merged totals cannot equal the sum of the
# input totals, so a conflict is reported for the big warning.
GLOBAL_INPUT_CELLS = [
    ("D6", "Salary inflation rate (UT)"),
    ("D7", "Salary inflation rate (JFO)"),
    ("D8", "Salary inflation rate (GRA)"),
    ("D9", "Tuition/fees inflation rate"),
    ("B108", "F&A base type"),
    ("D110", "F&A rate type"),
]


def _sum_cells(ws_t, value_sheets, rows, cols) -> None:
    for r in rows:
        for c in cols:
            total = sum(_num(vs[f"{c}{r}"].value) for vs in value_sheets)
            _set(ws_t, r, c, total)


def _has_gra_months(vs) -> bool:
    return any(_num(vs[f"{c}{r}"].value)
               for r in X.GRA_ROWS for c in PERSON_MONTH_COLS)


def _merge_tuition(ws_t, value_sheets, paths) -> List["ValueConflict"]:
    """Carry the per-GRA tuition/fee costs; report conflicts.

    Only inputs that actually budget GRA months matter -- the costs apply per
    GRA.  When those inputs disagree (including one budgeting tuition and
    another not), the merged workbook cannot reproduce the sum of the inputs
    (the formulas apply one cost to every merged GRA), so a conflict is
    returned for the caller's warning banner.
    """
    conflicts = []
    relevant = [(vs, p) for vs, p in zip(value_sheets, paths) if _has_gra_months(vs)]
    for r, c, label in TUITION_COST_CELLS:
        seen = {}
        for vs, p in relevant:
            seen.setdefault(_num(vs[f"{c}{r}"].value), []).append(os.path.basename(p))
        nonzero = [v for v in seen if v]
        if nonzero:
            _set(ws_t, r, c, nonzero[0])
        if len(seen) > 1:
            detail = "; ".join(f"{v:,.0f} in {', '.join(fs)}" for v, fs in seen.items())
            conflicts.append(ValueConflict(
                label, f"inputs with GRAs disagree ({detail}); the merged sheet "
                       f"applies one value to every GRA -- fix cell {c}{r} by hand"))
    return conflicts


def _check_global_conflicts(value_sheets, paths) -> List["ValueConflict"]:
    conflicts = []
    for cell, label in GLOBAL_INPUT_CELLS:
        seen = {}
        for vs, p in zip(value_sheets, paths):
            v = vs[cell].value
            if v not in (None, ""):
                seen.setdefault(str(v), []).append(os.path.basename(p))
        if len(seen) > 1:
            detail = "; ".join(f"{v!r} in {', '.join(fs)}" for v, fs in seen.items())
            conflicts.append(ValueConflict(
                f"{label} [{cell}]",
                f"inputs disagree ({detail}); merged uses the first input's value"))
    return conflicts


def _set_metadata(ws_t, value_sheets, n_inputs) -> None:
    names = [str(vs["D2"].value).strip() for vs in value_sheets
             if _present_name(vs["D2"].value)]
    if names:
        _set(ws_t, 2, "D", "; ".join(names))
    _set(ws_t, 3, "D",
         f"MERGED budget ({n_inputs} proposal(s)) -- combined by UTKBudgetExtractor")


def write_merged_workbook(
    input_paths: Sequence[str],
    out_path: str,
    template_path: Optional[str] = None,
) -> List[MergeIssue]:
    """Merge ``input_paths`` into a single budget workbook at ``out_path``.

    Only user-input cells are copied; formulas are never overwritten and
    recompute when the workbook is opened in Excel.  Returns a list of
    :class:`RowOverflow` (a section ran out of rows; overflow entries dropped)
    and :class:`ValueConflict` (inputs disagree on a workbook-wide input such
    as an inflation rate, the F&A type, or the per-GRA tuition cost) records --
    empty when the merge is exact.
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
    ws_t = template_wb[SHEET_NAME]
    main_values = [wb[SHEET_NAME] for wb in value_wbs]

    issues: List[MergeIssue] = []

    # --- Main sheet: personnel, equipment, manual leaves, tuition, metadata
    issues.extend(_merge_personnel(ws_t, main_values))
    for spec in _main_sheet_specs():
        of = _stack(ws_t, main_values, spec)
        if of:
            issues.append(of)
    _sum_cells(ws_t, main_values, MANUAL_OTHER_DIRECT_ROWS, PERIOD_COLS)
    issues.extend(_merge_tuition(ws_t, main_values, input_paths))
    issues.extend(_check_global_conflicts(main_values, input_paths))
    _set_metadata(ws_t, main_values, len(input_paths))

    # --- Detail sheets (so the main-sheet formula leaves recompute) --------
    detail_builders = [
        ("TRAVEL", _travel_specs),
        ("SUPPLIES", _supplies_specs),
        ("SUBCONTRACTS", _subcontracts_specs),
    ]
    for sheet_name, builder in detail_builders:
        ws_detail = _resolve_sheet(template_wb, sheet_name)
        if ws_detail is None:
            continue
        detail_values = [vs for vs in
                         (_resolve_sheet(wb, sheet_name) for wb in value_wbs)
                         if vs is not None]
        for spec in builder(ws_detail.title):
            of = _stack(ws_detail, detail_values, spec)
            if of:
                issues.append(of)

    ws_part = _resolve_sheet(template_wb, "PARTICIPANT SUPPORT COSTS")
    if ws_part is not None:
        part_values = [vs for vs in
                       (_resolve_sheet(wb, "PARTICIPANT SUPPORT COSTS") for wb in value_wbs)
                       if vs is not None]
        for spec in _participant_specs(ws_part.title, ws_part):
            of = _stack(ws_part, part_values, spec)
            if of:
                issues.append(of)

    template_wb.save(out_path)
    return issues
