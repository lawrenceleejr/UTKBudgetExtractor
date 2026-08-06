"""Merge several UTK budget spreadsheets.

Two products are produced from a set of budgets:

* :func:`merge_budgets` -- a new :class:`Budget` whose summable (dollar /
  person-month) fields are the element-wise sum across the inputs.  This drives
  the merged LaTeX definitions and justification.
* :func:`write_merged_workbook` -- a real ``.xlsx`` *in the same format as the
  input spreadsheets*.

The workbook merge is **formula-safe**: it never writes into a cell that holds a
formula.  Only genuine *user-input* cells are copied, and every subtotal, total,
salary, fringe, tuition, F&A, and roll-up formula is left untouched so Excel
recomputes the correct combined budget on open.  The fringe-rate cells are
institutional constants that ship with the template and are never touched.

How each section is combined:

* **Concatenated** (one row per distinct entry): senior personnel, other
  professionals, admin, other personnel, equipment, subcontracts, and
  participant-support blocks.
* **Consolidated by base salary** (one row per distinct base; headcount x months
  summed per period, which is cost-exact because the salary formulas are linear
  in headcount x months): post-docs (<=3 rows) and GRAs (<=4 rows).
* **Collapsed to a single row**: undergraduate researchers.
* **Consolidated per period into one domestic + one foreign summary row**:
  travel (reproduces each period's subtotal exactly).
* **Grouped by description and summed**: supplies.

If a section still has more distinct entries than it has rows, the overflow is
dropped and reported so the caller can raise a prominent warning.
"""

from __future__ import annotations

import os
import warnings as _warnings
from collections import OrderedDict
from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple, Union

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

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
    """Set ``ws[col][row]`` to ``value`` -- unless it holds a formula or is the
    read-only member of a merged range.

    Returns ``True`` if the cell was written.  This is the single guarantee that
    the merge never clobbers a formula, and never crashes on a merged cell (only
    the top-left anchor of a merged range is writable; the rest are read-only
    ``MergedCell`` proxies that carry no independent value).
    """
    cell = ws[f"{col}{row}"]
    if isinstance(cell, MergedCell) or _is_formula(cell):
        return False
    cell.value = value
    return True


def _resolve_sheet(wb, name: str):
    """Return the worksheet matching ``name`` (tolerant of case and trailing
    spaces, e.g. a re-save that renames 'UTK Budget' to 'UTK BUDGET')."""
    if name in wb.sheetnames:
        return wb[name]
    for sn in wb.sheetnames:
        if sn.strip().upper() == name.strip().upper():
            return wb[sn]
    return None


# -- Personnel ---------------------------------------------------------------

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

# Groups that CONCATENATE -- each entry is a distinct named person kept on its
# own row.  ``name_in_b`` marks the senior section, whose column B holds a real
# name (elsewhere B is a fixed role label that ships with the template).
CONCAT_GROUPS = [
    ("Senior Personnel", X.SENIOR_ROWS, True),
    ("Other Professionals", X.OTHER_PROF_ROWS, False),
    ("Admin/Clerical", [X.ADMIN_ROW], False),
    ("Other Personnel", X.OTHER_STAFF_ROWS, False),
]

# Groups that CONSOLIDATE.  Postdocs and GRAs are grouped by base salary (one
# row per distinct base); undergraduates collapse to a single row.  Because the
# salary formulas are linear in headcount x months, consolidating with E=1 and
# months = sum(headcount_i x months_i) reproduces the cost exactly.
# (label, rows, collapse_all)
CONSOLIDATE_GROUPS = [
    ("Post Docs", X.POSTDOC_ROWS, False),
    ("Graduate Research Assistants", X.GRA_ROWS, False),
    ("Undergraduate Researchers", [X.UNDERGRAD_ROW], True),
]

# Columns cleared/written for a consolidated line.  Column B (the role label)
# and the fringe rows are left untouched.
CONSOLIDATE_COLS = ["A", "C", "D", "E"] + PERSON_MONTH_COLS


def _concat_personnel(ws_t, value_sheets) -> List[RowOverflow]:
    """Stack distinct named personnel (senior, other prof, admin, other)."""
    overflows = []
    for label, pers_rows, name_in_b in CONCAT_GROUPS:
        pers_rows = list(pers_rows)
        cols = list(PERSONNEL_INPUT_COLS)
        if name_in_b:
            cols = ["B"] + cols + SENIOR_TENURE_COLS

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

        for prow in pers_rows:
            for c in cols:
                _set(ws_t, prow, c, None)

        capacity = len(pers_rows)
        for k, entry in enumerate(entries[:capacity]):
            prow = pers_rows[k]
            for c in cols:
                _set(ws_t, prow, c, entry[c])

        if len(entries) > capacity:
            dropped = [str(e["_name"]) if _present_name(e["_name"])
                       else f"{label} entry {capacity + i + 1}"
                       for i, e in enumerate(entries[capacity:])]
            overflows.append(RowOverflow(label, capacity, len(entries), dropped))
    return overflows


def _collect_consolidatable(value_sheets, rows):
    """Return the funded entries in ``rows`` as dicts of base/headcount/months."""
    entries = []
    for vs in value_sheets:
        for r in rows:
            base = _num(vs[f"D{r}"].value)
            head = _num(vs[f"E{r}"].value)
            months = {c: _num(vs[f"{c}{r}"].value) for c in PERSON_MONTH_COLS}
            # Funded only if base, headcount, and at least one month are set --
            # matching the template's cost formula (base x headcount x months).
            if base > 0 and head > 0 and any(months.values()):
                entries.append({"D": base, "E": head, "months": months,
                                "A": vs[f"A{r}"].value, "C": vs[f"C{r}"].value})
    return entries


def _agg_months(group, weight_by_base=False):
    out = {}
    for c in PERSON_MONTH_COLS:
        total = sum((e["D"] if weight_by_base else 1.0) * e["E"] * e["months"][c]
                    for e in group)
        # weight_by_base sums are dollars (base x headcount x months against a
        # $1 base), so round them like every other derived dollar figure; plain
        # month sums stay fractional (2.5 months is meaningful).
        out[c] = X.round_dollar(total) if weight_by_base else total
    return out


def _escalation_key(e):
    """Group only entries that share a base salary AND the same escalation
    inputs -- the personnel type in column C (UT/JFO/GRA selects the inflation
    rate) and the raise flag in column A.  Two $5,000 post-docs, one UT and one
    JFO, escalate at different rates, so merging them onto one line would make
    years 2-5 wrong; keying on all three keeps every consolidated line
    cost-exact in every period."""
    return (round(_num(e["D"]), 2), str(e["C"]), str(e["A"]))


def _group_by_base(entries):
    """One consolidated line per (base salary, type, raise flag); largest base
    first."""
    groups = OrderedDict()
    for e in entries:
        groups.setdefault(_escalation_key(e), []).append(e)
    lines = []
    for key in sorted(groups, key=lambda k: k[0], reverse=True):
        grp = groups[key]
        lines.append({"D": grp[0]["D"], "E": 1, "A": grp[0]["A"], "C": grp[0]["C"],
                      "months": _agg_months(grp)})
    return lines


def _collapse_all(entries):
    """Collapse every entry into a single line.

    If all entries share a base salary the base is kept and months are summed;
    otherwise the line is normalised to base=1 with months = sum(base x
    headcount x months), which is still cost-exact when the escalation rate is
    uniform across the entries (true for undergraduates, all UT-rate)."""
    bases = {round(e["D"], 2) for e in entries}
    first = entries[0]
    if len(bases) == 1:
        line = {"D": first["D"], "E": 1, "months": _agg_months(entries)}
    else:
        line = {"D": 1, "E": 1, "months": _agg_months(entries, weight_by_base=True)}
    line["A"], line["C"] = first["A"], first["C"]
    return [line]


def _consolidate_personnel(ws_t, value_sheets) -> List[RowOverflow]:
    overflows = []
    for label, rows, collapse_all in CONSOLIDATE_GROUPS:
        rows = list(rows)
        entries = _collect_consolidatable(value_sheets, rows)

        # Clear the section's input columns (formula-safe; B/role labels and the
        # fringe rows are left untouched).
        for r in rows:
            for c in CONSOLIDATE_COLS:
                _set(ws_t, r, c, None)
        if not entries:
            continue

        lines = _collapse_all(entries) if collapse_all else _group_by_base(entries)

        capacity = len(rows)
        for k, line in enumerate(lines[:capacity]):
            r = rows[k]
            _set(ws_t, r, "A", line["A"])
            _set(ws_t, r, "C", line["C"])
            _set(ws_t, r, "D", line["D"])
            _set(ws_t, r, "E", line["E"])
            for c in PERSON_MONTH_COLS:
                _set(ws_t, r, c, line["months"][c] or None)

        if len(lines) > capacity:
            dropped = [f"{label} @ base ${l['D']:,.2f}" for l in lines[capacity:]]
            overflows.append(RowOverflow(label, capacity, len(lines), dropped))
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


# TRAVEL: 5 period blocks (22 rows each) with 10 domestic + 5 foreign rows.
TRAVEL_INPUT_COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "L"]
TRAVEL_BLOCKS = [(1 + p * 22, p + 1) for p in range(5)]


def _consolidate_travel(ws_t, value_sheets) -> None:
    """Collapse every trip in a period/kind into one summary row.

    The per-trip and subtotal formulas compute
        cost = travelers*(regfee+airfare+additional) + travelers*days*(lodging+perdiem)
    so a single row with days=1, travelers=1 and the pre-multiplied per-column
    sums reproduces the subtotal exactly.  days>0 is required or the sheet
    reports "# Days Missing" and the total breaks."""
    for base, period in TRAVEL_BLOCKS:
        for kind, first_row, span in (("domestic", base + 3, 10),
                                      ("foreign", base + 15, 5)):
            rows = list(range(first_row, first_row + span))
            reg = air = add = lodg = perd = 0.0
            ntrips = 0
            for vs in value_sheets:
                for r in rows:
                    F = _num(vs[f"F{r}"].value); G = _num(vs[f"G{r}"].value)
                    H = _num(vs[f"H{r}"].value); I = _num(vs[f"I{r}"].value)
                    J = _num(vs[f"J{r}"].value)
                    if F + G + H + I + J <= 0:
                        continue
                    D = _num(vs[f"D{r}"].value); E = _num(vs[f"E{r}"].value)
                    ntrips += 1
                    reg += F * E; air += G * E; add += H * E
                    lodg += I * D * E; perd += J * D * E

            for r in rows:                       # clear all trip rows (formula-safe)
                for c in TRAVEL_INPUT_COLS:
                    _set(ws_t, r, c, None)

            if reg + air + add + lodg + perd <= 0:
                continue
            r = first_row
            _set(ws_t, r, "A", f"Merged {kind} travel ({ntrips} trip(s))")
            _set(ws_t, r, "C", "Various")
            _set(ws_t, r, "D", 1)   # days   (>0 so the row's total is not blocked)
            _set(ws_t, r, "E", 1)   # travelers
            _set(ws_t, r, "F", X.round_dollar(reg) or None)
            _set(ws_t, r, "G", X.round_dollar(air) or None)
            _set(ws_t, r, "H", X.round_dollar(add) or None)
            _set(ws_t, r, "I", X.round_dollar(lodg) or None)
            _set(ws_t, r, "J", X.round_dollar(perd) or None)


# SUPPLIES: entry rows 3-37 (A=description, B-F = per-period dollar amounts).
SUPPLIES_ROWS = list(range(3, 38))
SUPPLIES_PERIOD_COLS = ["B", "C", "D", "E", "F"]


def _consolidate_supplies(ws_t, value_sheets) -> Optional[RowOverflow]:
    """Group supply lines by description and sum the per-period amounts.

    Supply amounts are plain dollars, so summing is exact.  Blank descriptions
    all collapse into a single line; distinct descriptions stay separate."""
    groups = OrderedDict()
    for vs in value_sheets:
        for r in SUPPLIES_ROWS:
            amounts = {c: _num(vs[f"{c}{r}"].value) for c in SUPPLIES_PERIOD_COLS}
            if not any(amounts.values()):
                continue
            desc = vs[f"A{r}"].value
            key = str(desc).strip().lower() if _present_name(desc) else ""
            if key not in groups:
                groups[key] = {"A": desc if _present_name(desc) else None,
                               "sums": {c: 0.0 for c in SUPPLIES_PERIOD_COLS}}
            for c in SUPPLIES_PERIOD_COLS:
                groups[key]["sums"][c] += amounts[c]

    lines = list(groups.values())
    for r in SUPPLIES_ROWS:                      # clear (formula-safe)
        for c in ["A"] + SUPPLIES_PERIOD_COLS:
            _set(ws_t, r, c, None)

    capacity = len(SUPPLIES_ROWS)
    for k, line in enumerate(lines[:capacity]):
        r = SUPPLIES_ROWS[k]
        if line["A"] is not None:
            _set(ws_t, r, "A", line["A"])
        for c in SUPPLIES_PERIOD_COLS:
            _set(ws_t, r, c, X.round_dollar(line["sums"][c]) or None)

    if len(lines) > capacity:
        dropped = [str(l["A"] or "(unlabeled)") for l in lines[capacity:]]
        return RowOverflow("Supplies", capacity, len(lines), dropped)
    return None


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
    """Sum manually-entered dollar cells across the inputs (derived figures, so
    rounded to the nearest dollar like all other final numbers)."""
    for r in rows:
        for c in cols:
            total = sum(_num(vs[f"{c}{r}"].value) for vs in value_sheets)
            _set(ws_t, r, c, X.round_dollar(total))


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

    ws_t = _resolve_sheet(template_wb, SHEET_NAME)
    if ws_t is None:
        raise ValueError(f"{template_path!r}: missing worksheet {SHEET_NAME!r}")
    main_values = [_resolve_sheet(wb, SHEET_NAME) for wb in value_wbs]

    issues: List[MergeIssue] = []

    # --- Main sheet: personnel, equipment, manual leaves, tuition, metadata
    issues.extend(_concat_personnel(ws_t, main_values))       # named people
    issues.extend(_consolidate_personnel(ws_t, main_values))  # postdoc/GRA/undergrad
    for spec in _main_sheet_specs():
        of = _stack(ws_t, main_values, spec)
        if of:
            issues.append(of)
    _sum_cells(ws_t, main_values, MANUAL_OTHER_DIRECT_ROWS, PERIOD_COLS)
    issues.extend(_merge_tuition(ws_t, main_values, input_paths))
    issues.extend(_check_global_conflicts(main_values, input_paths))
    _set_metadata(ws_t, main_values, len(input_paths))

    # --- Detail sheets (so the main-sheet formula leaves recompute) --------
    # Travel and supplies are consolidated; subcontracts and participant support
    # are concatenated (distinct entities).
    def detail_values(sheet_name):
        return [vs for vs in (_resolve_sheet(wb, sheet_name) for wb in value_wbs)
                if vs is not None]

    ws_travel = _resolve_sheet(template_wb, "TRAVEL")
    if ws_travel is not None:
        _consolidate_travel(ws_travel, detail_values("TRAVEL"))

    ws_supplies = _resolve_sheet(template_wb, "SUPPLIES")
    if ws_supplies is not None:
        of = _consolidate_supplies(ws_supplies, detail_values("SUPPLIES"))
        if of:
            issues.append(of)

    ws_sub = _resolve_sheet(template_wb, "SUBCONTRACTS")
    if ws_sub is not None:
        for spec in _subcontracts_specs(ws_sub.title):
            of = _stack(ws_sub, detail_values("SUBCONTRACTS"), spec)
            if of:
                issues.append(of)

    ws_part = _resolve_sheet(template_wb, "PARTICIPANT SUPPORT COSTS")
    if ws_part is not None:
        for spec in _participant_specs(ws_part.title, ws_part):
            of = _stack(ws_part, detail_values("PARTICIPANT SUPPORT COSTS"), spec)
            if of:
                issues.append(of)

    template_wb.save(out_path)
    return issues
