"""Generate DOE-style budget-justification LaTeX documents.

The generated document is a *template baseline* justification suitable for a
U.S. Department of Energy (Office of Science) grant proposal.  All dollar
figures are pulled from the ``\\newcommand`` macros emitted by
:mod:`utkbudget.texdefs`, so editing the spreadsheet and re-running the tool
keeps the prose and the numbers in sync.

A document contains one ``\\section`` of justification per contributing budget
(one per input spreadsheet) followed by a justification of the combined / summed
budget -- exactly the structure DOE expects for a joint, multi-institution or
multi-thrust request.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Sequence, Tuple

from .extractor import PERIOD_WORDS
from .provenance import provenance_comment
from .texdefs import escape_tex


def _m(prefix: str, field: str) -> str:
    """A LaTeX macro reference that is safe to follow with a space."""
    return f"\\{prefix}{field}{{}}"


def _usd(prefix: str, field: str) -> str:
    return f"\\${_m(prefix, field)}"


PREAMBLE = r"""\documentclass[11pt]{article}
\usepackage[margin=1in]{geometry}
\usepackage{booktabs}
\usepackage{hyperref}
\setlength{\parskip}{0.5em}
\setlength{\parindent}{0pt}
"""


def _travel_table(prefix: str) -> str:
    """A per-period domestic/foreign travel table built from the macros."""
    period_cols = " & ".join(f"Period {i + 1}" for i in range(len(PERIOD_WORDS)))
    dom = " & ".join(_usd(prefix, f"DomesticTravelYear{w}") for w in PERIOD_WORDS)
    foreign = " & ".join(_usd(prefix, f"ForeignTravelYear{w}") for w in PERIOD_WORDS)
    total = " & ".join(_usd(prefix, f"TravelYear{w}") for w in PERIOD_WORDS)
    colspec = "l" + "r" * len(PERIOD_WORDS) + "r"
    return (
        "\\begin{center}\n"
        f"\\begin{{tabular}}{{{colspec}}}\n"
        "\\toprule\n"
        f"Travel & {period_cols} & Total \\\\\n"
        "\\midrule\n"
        f"Domestic & {dom} & {_usd(prefix, 'DomesticTravelTotal')} \\\\\n"
        f"Foreign & {foreign} & {_usd(prefix, 'ForeignTravelTotal')} \\\\\n"
        "\\midrule\n"
        f"\\textbf{{Total Travel}} & {total} & \\textbf{{{_usd(prefix, 'TravelTotal')}}} \\\\\n"
        "\\bottomrule\n"
        "\\end{tabular}\n"
        "\\end{center}\n"
    )


def _num(budget, name) -> float:
    """Numeric value of a field from the Budget (0.0 if missing/blank)."""
    value = budget.get(name) if budget is not None else None
    try:
        return float(value) if value not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


# (table label, macro base) for the salary categories shown on their own rows.
SALARY_CATEGORIES = [
    ("Senior Personnel", "SeniorSubtotal"),
    ("Post-docs", "PostdocSubtotal"),
    ("Other Professionals", "OtherProfSubtotal"),
    ("Graduate Research Assistants", "GRASubtotal"),
    ("Undergraduate Researchers", "UndergradSubtotal"),
    ("Admin/Clerical", "AdminSubtotal"),
    ("Other Personnel", "OtherStaffSubtotal"),
]

# (table label, macro base) for the non-salary direct-cost rows.
OTHER_SUMMARY_ROWS = [
    ("Fringe Benefits", "Fringe"),
    ("Equipment", "Equipment"),
    ("Travel", "Travel"),
    ("Participant Support", "ParticipantSupport"),
    ("Other Direct Costs", "OtherDirect"),
]

# (fringe-category label, representative rate macro, subtotal macro base)
FRINGE_CATEGORIES = [
    ("Senior personnel and faculty", "SeniorAFringeRate", "SeniorSubtotal"),
    ("Post-docs", "PostdocAFringeRate", "PostdocSubtotal"),
    ("Other professionals", "OtherProfAFringeRate", "OtherProfSubtotal"),
    ("Graduate research assistants", "GRAAFringeRate", "GRASubtotal"),
    ("Undergraduate researchers", "UndergradAFringeRate", "UndergradSubtotal"),
    ("Administrative/clerical staff", "AdminAFringeRate", "AdminSubtotal"),
    ("Other personnel", "OtherStaffAFringeRate", "OtherStaffSubtotal"),
]


def _has_jfo(budget) -> bool:
    """True if any senior-personnel line is a jointly-appointed (JFO) faculty."""
    if budget is None:
        return False
    for i in range(12):
        t = budget.get(f"Senior{chr(65 + i)}Type")
        if t and str(t).strip().upper() == "JFO":
            return True
    return False


def _present_seniors(budget):
    """Indices (letters) of senior rows that carry a name, salary, or base."""
    out = []
    if budget is None:
        return out
    for i in range(12):
        L = chr(65 + i)
        name = budget.get(f"Senior{L}Name")
        if (name and str(name).strip()) or _num(budget, f"Senior{L}Total") \
                or _num(budget, f"Senior{L}BaseAnnual"):
            out.append(L)
    return out


def _summary_table(prefix: str, budget) -> str:
    """By-category cost summary; rows with no money are omitted, salaries are
    split by category, and the F&A rate is named in the indirect-cost row."""
    rows = [(lbl, base) for lbl, base in SALARY_CATEGORIES + OTHER_SUMMARY_ROWS
            if _num(budget, base + "Total") != 0]
    lines = [
        "\\begin{center}",
        "\\begin{tabular}{lr}",
        "\\toprule",
        "Cost Category & Total Requested \\\\",
        "\\midrule",
    ]
    for label, base in rows:
        lines.append(f"{label} & {_usd(prefix, base + 'Total')} \\\\")
    lines.append("\\midrule")
    lines.append(f"Total Direct Costs & {_usd(prefix, 'DirectTotal')} \\\\")
    if _num(budget, "IndirectTotal") != 0:
        fa = f"Indirect Costs (F\\&A, {_m(prefix, 'OverheadRateYearOne')}\\%)"
        lines.append(f"{fa} & {_usd(prefix, 'IndirectTotal')} \\\\")
    lines.append("\\midrule")
    lines.append(f"\\textbf{{Total Project Cost}} & "
                 f"\\textbf{{{_usd(prefix, 'GrandTotal')}}} \\\\")
    lines += ["\\bottomrule", "\\end{tabular}", "\\end{center}"]
    return "\n".join(lines) + "\n"


def _escalation_table(prefix: str, budget) -> str:
    """Small table of the assumed annual escalation rates that apply."""
    rows = [("University of Tennessee (UT) personnel", "SalaryInflationUT")]
    if _has_jfo(budget):
        rows.append(("Jointly-appointed faculty (JFO)", "SalaryInflationJFO"))
    if _num(budget, "GRASubtotalTotal") != 0:
        rows.append(("Graduate research assistants", "SalaryInflationGRA"))
        rows.append(("Graduate tuition \\& mandatory fees", "TuitionInflation"))
    lines = [
        "\\begin{center}",
        "\\begin{tabular}{lr}",
        "\\toprule",
        "Category & Annual escalation \\\\",
        "\\midrule",
    ]
    for label, rate in rows:
        lines.append(f"{label} & {_m(prefix, rate)}\\% \\\\")
    lines += ["\\bottomrule", "\\end{tabular}", "\\end{center}"]
    return "\n".join(lines) + "\n"


def _fringe_rate_list(prefix: str, budget) -> str:
    """Itemised fringe rates for each personnel category present."""
    present = [(lbl, rate) for lbl, rate, sub in FRINGE_CATEGORIES
               if _num(budget, sub + "Total") != 0]
    if not present:
        return ""
    items = "\n".join(f"\\item {lbl}: {_m(prefix, rate)}\\%" for lbl, rate in present)
    return "\\begin{itemize}\n" + items + "\n\\end{itemize}"


def _senior_list(prefix: str, budget) -> str:
    """Itemise each senior person's requested months and base salary."""
    present = _present_seniors(budget)
    if not present:
        return ""
    items = []
    for L in present:
        items.append(
            f"\\item {_m(prefix, f'Senior{L}Name')}: "
            f"{_m(prefix, f'Senior{L}PersonMonths')} person-months at a base "
            f"{_m(prefix, f'Senior{L}ApptType')}-month salary of "
            f"\\${_m(prefix, f'Senior{L}BaseAnnual')}."
        )
    return "\\begin{itemize}\n" + "\n".join(items) + "\n\\end{itemize}"


def render_section(prefix: str, budget=None, heading=None,
                   is_sum: bool = False, detailed: bool = False) -> str:
    """Return the LaTeX for one budget's justification section.

    ``budget`` is the :class:`~utkbudget.extractor.Budget` behind ``prefix`` and
    is used only for *structural* decisions (which rows/sections/rates to show);
    the displayed numbers are always the macros so editing the defs keeps them in
    sync.  ``heading`` is optional ``\\section`` text (LaTeX-escaped; ``None``
    emits none, keeping a single-budget document free of any source filename).
    ``detailed`` lists individual senior personnel (only meaningful for a single
    real budget, not a merged one)."""
    if is_sum:
        opener = (
            "This justifies the combined budget -- the sum of the contributing "
            "budgets -- requested from the U.S. Department of Energy. Costs are "
            "organized following the standard DOE Office of Science budget "
            "categories."
        )
    else:
        opener = (
            "The following justifies the funds requested from the U.S. Department "
            "of Energy. Costs are organized following the standard DOE Office of "
            "Science budget categories."
        )

    def fy(base):
        """First-year figure phrase: '$X in the first year'."""
        return f"{_usd(prefix, base + 'YearOne')} in the first year"

    parts: List[str] = []
    if heading:
        parts.append(f"\\section{{{escape_tex(heading)}}}")
    parts.append(opener)
    parts.append(_summary_table(prefix, budget))

    # --- Basis of estimate / assumed raise structure -------------------
    parts.append("\\subsection*{Basis of Estimate and Escalation}")
    parts.append(
        "Salaries and wages are budgeted at current institutional rates and "
        "escalated annually at the rates below. Fringe-benefit and "
        "facilities-and-administrative (F\\&A) rates follow the institution's "
        "current federally negotiated rate agreements."
    )
    parts.append(_escalation_table(prefix, budget))

    # --- A. Senior Personnel -------------------------------------------
    parts.append("\\subsection*{A. Senior Personnel}")
    parts.append(
        "Funds are requested for the academic-year and/or summer effort of the "
        "senior personnel, escalated at the assumed raise rate for each "
        f"appointment type. Senior-personnel salaries total "
        f"{_usd(prefix, 'SeniorSubtotalTotal')} over the project period "
        f"({fy('SeniorSubtotal')})."
    )
    if detailed and not is_sum:
        senior_items = _senior_list(prefix, budget)
        if senior_items:
            parts.append("The following senior personnel are supported:")
            parts.append(senior_items)

    # --- B. Other Personnel --------------------------------------------
    parts.append("\\subsection*{B. Other Personnel}")
    parts.append(
        "Funds are requested for postdoctoral researchers, graduate research "
        "assistants (GRAs), undergraduate researchers, and other personnel "
        "essential to the proposed research. Postdoctoral and student effort "
        "drives the technical work of the project; their salaries escalate at the "
        "assumed annual rates above. The other-personnel request is "
        f"{_usd(prefix, 'OtherPersonnelSubtotalTotal')} over the project period "
        f"({fy('OtherPersonnelSubtotal')}), bringing total salaries and wages to "
        f"{_usd(prefix, 'WagesTotal')}."
    )

    # --- C. Fringe Benefits --------------------------------------------
    parts.append("\\subsection*{C. Fringe Benefits}")
    parts.append(
        "Fringe benefits are calculated using the institution's federally "
        "negotiated fringe-benefit rates for each personnel category, at the rates "
        f"below. Fringe benefits total {_usd(prefix, 'FringeTotal')} over the "
        f"project period, for total salaries and benefits of "
        f"{_usd(prefix, 'SalaryAndBenefitsTotal')}."
    )
    fringe_rates = _fringe_rate_list(prefix, budget)
    if fringe_rates:
        parts.append(fringe_rates)

    # --- D. Equipment ---------------------------------------------------
    parts.append("\\subsection*{D. Equipment}")
    equipment_lead = (
        "Equipment is defined as items of tangible personal property with a "
        "useful life of more than one year and a unit acquisition cost of "
        "\\$5,000 or more."
    )
    if _num(budget, "EquipmentTotal") == 0:
        parts.append(equipment_lead + " No equipment is requested in this proposal.")
    else:
        parts.append(
            equipment_lead + f" Equipment is {fy('Equipment')}, for a total request "
            f"of {_usd(prefix, 'EquipmentTotal')}. Each item is itemized in the "
            "accompanying budget spreadsheet."
        )

    # --- E. Travel (emphasized) ----------------------------------------
    parts.append("\\subsection*{E. Travel}")
    parts.append(
        f"A total of {_usd(prefix, 'TravelTotal')} is requested for travel "
        f"({fy('Travel')}). Travel is essential to disseminate results, participate "
        "in DOE program activities, and sustain the scientific collaborations on "
        "which this research depends. The request is broken down by period and by "
        "domestic vs.\\ foreign travel below."
    )
    parts.append(_travel_table(prefix))
    parts.append(
        "\\textbf{Domestic travel} "
        f"({_usd(prefix, 'DomesticTravelTotal')} total, {fy('DomesticTravel')}) "
        "supports attendance at topical workshops and major conferences (e.g., APS "
        "April Meeting and DPF), DOE-sponsored principal-investigator and program "
        "review meetings, and collaboration meetings at partner national "
        "laboratories and universities. Costs are estimated per traveler and "
        "include airfare, lodging at prevailing per-diem rates, ground "
        "transportation, and conference registration."
    )
    parts.append(
        "\\textbf{Foreign travel} "
        f"({_usd(prefix, 'ForeignTravelTotal')} total, {fy('ForeignTravel')}) "
        "supports participation in international collaboration meetings, experiment "
        "shifts, and conferences central to the proposed program. All foreign "
        "travel will comply with the Fly America Act and DOE foreign-travel "
        "approval requirements."
    )

    # --- F. Participant Support ----------------------------------------
    parts.append("\\subsection*{F. Participant Support Costs}")
    if _num(budget, "ParticipantSupportTotal") == 0:
        parts.append("No participant support costs are requested in this proposal.")
    else:
        parts.append(
            f"A total of {_usd(prefix, 'ParticipantSupportTotal')} is requested for "
            f"participant support costs ({fy('ParticipantSupport')}): stipends, "
            "travel, and subsistence for participants in workshops, schools, or "
            "training activities associated with the project. These funds are "
            "budgeted separately and are excluded from the indirect-cost base."
        )

    # --- G. Other Direct Costs -----------------------------------------
    parts.append("\\subsection*{G. Other Direct Costs}")
    parts.append(
        f"Other direct costs are {fy('OtherDirect')} and {_usd(prefix, 'OtherDirectTotal')} "
        f"over five years, and include materials and supplies ({_usd(prefix, 'SuppliesTotal')}, "
        f"{fy('Supplies')}), publication and page charges, and other direct project "
        "expenses. Graduate-student tuition and mandatory fees "
        f"({_usd(prefix, 'TuitionSubtotalTotal')} total, {fy('TuitionSubtotal')}) are "
        "requested for the GRAs supported on this project and escalate at the "
        "assumed tuition rate above, consistent with institutional policy."
    )

    # --- Indirect Costs -------------------------------------------------
    parts.append("\\subsection*{Indirect (F\\&A) Costs}")
    parts.append(
        "Indirect costs are computed on the Modified Total Direct Cost (MTDC) base "
        f"using the institution's federally negotiated rate of "
        f"{_m(prefix, 'OverheadRateYearOne')}\\% ({_m(prefix, 'FandARateType')}), "
        "which is fixed for all periods of the proposal. The total indirect-cost "
        f"request is {_usd(prefix, 'IndirectTotal')}."
    )

    # --- Total ----------------------------------------------------------
    parts.append("\\subsection*{Total Requested}")
    parts.append(
        "The total funds requested from the Department of Energy are "
        f"{fy('Grand')} and \\textbf{{{_usd(prefix, 'GrandTotal')}}} over the project "
        f"period (direct costs {_usd(prefix, 'DirectTotal')} plus indirect costs "
        f"{_usd(prefix, 'IndirectTotal')})."
    )

    return "\n\n".join(parts)


def build_document(
    title: str,
    defs_inputs: Sequence[str],
    sections: Sequence[Tuple[str, str, bool]],
    intro: str = "",
    defs_inline: str = None,
) -> str:
    """Assemble a complete justification document.

    Definitions come from EITHER ``defs_inline`` (a ``\\newcommand`` block
    emitted verbatim, so the document is self-contained and names no external
    file) OR ``defs_inputs`` (relative paths to ``\\input``).  ``sections`` is a
    list of ``(prefix, budget, heading_or_None, is_sum, detailed)`` (see
    :func:`render_section`).  ``title`` is human text and is LaTeX-escaped here;
    ``intro`` is authored LaTeX emitted verbatim (callers must escape any names
    they weave in).
    """
    today = _dt.date.today()
    out: List[str] = [provenance_comment(), PREAMBLE]
    out.append(f"\\title{{{escape_tex(title)}}}")
    out.append(f"\\date{{{today:%B} {today.day}, {today:%Y}}}")
    out.append("\\begin{document}")
    out.append("\\maketitle")

    if defs_inline:
        out.append("% --- budget definitions (inlined; self-contained) ---")
        out.append(defs_inline)
    else:
        out.append("% --- generated budget definitions ---")
        for rel in defs_inputs:
            base = rel[:-4] if rel.endswith(".tex") else rel  # strip .tex for \input
            out.append(f"\\input{{{base}}}")
        out.append("")

    if intro:
        out.append(intro)

    for prefix, budget, heading, is_sum, detailed in sections:
        out.append(render_section(prefix, budget, heading, is_sum, detailed))

    out.append("\\end{document}")
    return "\n\n".join(out) + "\n"


def write_document(path: str, *args, **kwargs) -> str:
    content = build_document(*args, **kwargs)
    with open(path, "w") as fh:
        fh.write(content)
    return path
