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

from .extractor import PERIOD_WORDS, round_dollar
from .provenance import provenance_comment
from .texdefs import escape_tex, tex_name


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


def _travel_table(prefix: str, periods) -> str:
    """A per-period domestic/foreign travel table built from the macros.

    ``periods`` is the list of active period words (e.g. ``["One", "Two"]``);
    only those period columns are shown so empty later years are not printed."""
    nums = [PERIOD_WORDS.index(w) + 1 for w in periods]
    period_cols = " & ".join(f"Period {i}" for i in nums)
    dom = " & ".join(_usd(prefix, f"DomesticTravelYear{w}") for w in periods)
    foreign = " & ".join(_usd(prefix, f"ForeignTravelYear{w}") for w in periods)
    total = " & ".join(_usd(prefix, f"TravelYear{w}") for w in periods)
    colspec = "l" + "r" * len(periods) + "r"
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


def _active_periods(budget):
    """Period words that carry any money (fall back to all five if unknown)."""
    active = [w for w in PERIOD_WORDS if _num(budget, f"GrandYear{w}") != 0]
    return active or list(PERIOD_WORDS)


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
    """Itemise each senior person's base salary and requested person-months in
    every year that has effort (not just the first)."""
    present = _present_seniors(budget)
    if not present:
        return ""
    items = []
    for L in present:
        year_parts = [
            f"{_m(prefix, f'Senior{L}MonthsYear{w}')} in year {i}"
            for i, w in enumerate(PERIOD_WORDS, start=1)
            if _num(budget, f"Senior{L}MonthsYear{w}") > 0
        ]
        months = ", ".join(year_parts) if year_parts else "effort as budgeted"
        items.append(
            f"\\item {_m(prefix, f'Senior{L}Name')} "
            f"(base {_m(prefix, f'Senior{L}ApptType')}-month salary "
            f"\\${_m(prefix, f'Senior{L}BaseAnnual')}): person-months of {months}."
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
    parts.append(_travel_table(prefix, _active_periods(budget)))
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


INCLUDE_GUARD = r"\budgetjustificationincluded"

# Prefix prepended to every generated \input.  Empty by default (all generated
# files live in one flat directory, so bare names resolve when the document is
# compiled from there).  A parent document in a *different* directory can set it
# once -- \def\budgetjustificationpath{output/} -- and every nested \input
# follows, because TeX resolves relative \input paths against the main
# document's directory rather than the included file's.
PATH_MACRO = r"\budgetjustificationpath"

# Set only by a file that opened \begin{document} itself, so it knows to close
# it.  The include guard cannot serve double duty here: a standalone driver
# *defines* the guard (so the justifications it pulls in skip their preambles),
# which would otherwise make it look included to itself.
STANDALONE_MARKER = r"\budgetjustificationstandalone"


def _path_preamble() -> str:
    """Declare the \\input path prefix (harmless if the parent already set it)."""
    return f"\\providecommand{{{PATH_MACRO}}}{{}}"


def _input(rel: str) -> str:
    """An ``\\input`` of a generated file, honouring the path prefix macro."""
    base = rel[:-4] if rel.endswith(".tex") else rel   # strip .tex for \input
    return f"\\input{{{PATH_MACRO} {base}}}"


def defs_input(name: str) -> str:
    """The output-root-relative ``\\input`` path of a defs file in ``defs/``."""
    return f"{DEFS_DIR}/{name}"


def build_document(
    title: str,
    defs_inputs: Sequence[str],
    sections: Sequence[Tuple[str, str, bool]],
    intro: str = "",
    subtitle: str = None,
) -> str:
    """Assemble a justification document.

    The definitions are pulled in with ``\\input`` from ``defs_inputs`` (the
    dedicated per-budget defs files, each with its own unique macro prefix), so
    the numbers stay in one place and several justifications can be combined
    without macro clashes.  ``sections`` is a list of ``(prefix, budget,
    heading_or_None, is_sum, detailed)`` (see :func:`render_section`).

    The file both compiles on its own AND ``\\input``s cleanly into a larger
    proposal.  The ``\\documentclass`` .. ``\\begin{document}`` preamble and the
    closing ``\\end{document}`` are wrapped in an ``\\ifdefined`` guard: define
    ``\\budgetjustificationincluded`` in the parent document's preamble to skip
    them (see :func:`build_pdf_driver`).  ``title`` / ``subtitle`` are human text
    and LaTeX-escaped; ``intro`` is authored LaTeX emitted verbatim.
    """
    today = _dt.date.today()
    title_tex = escape_tex(title)
    subtitle_tex = escape_tex(subtitle) if subtitle else ""

    out: List[str] = [provenance_comment()]
    out.append(
        "% Self-contained: this file \\input{}s its own budget definitions and\n"
        "% compiles on its own (e.g. `pdflatex thisfile.tex`).  To \\input it into\n"
        "% a larger proposal instead, put  \\def\\budgetjustificationincluded{}  in\n"
        "% that document's preamble first; the preamble and\n"
        "% \\begin{document}/\\end{document} below are then skipped automatically.\n"
        "% If the parent document lives in a different directory, also set\n"
        "% \\def\\budgetjustificationpath{output/} so the \\input below resolves.")
    out.append(_path_preamble())

    titled = title_tex + (r"\\[4pt]{\large " + subtitle_tex + "}" if subtitle_tex else "")
    out.append(
        f"\\ifdefined{INCLUDE_GUARD}\\else\n"
        + PREAMBLE
        + f"\\title{{{titled}}}\n"
        + f"\\date{{{today:%B} {today.day}, {today:%Y}}}\n"
        + "\\begin{document}\n"
        + "\\maketitle\n"
        + "\\fi")

    out.append("% --- budget definitions (the dedicated defs file) ---")
    for rel in defs_inputs:
        out.append(_input(rel))

    # When included there is no \maketitle, so print a heading identifying this
    # budget (kept out of standalone output, where the title block already shows).
    heading = title_tex + (" --- " + subtitle_tex if subtitle_tex else "")
    out.append(f"\\ifdefined{INCLUDE_GUARD}\\section*{{{heading}}}\\fi")

    if intro:
        out.append(intro)

    for prefix, budget, sec_heading, is_sum, detailed in sections:
        out.append(render_section(prefix, budget, sec_heading, is_sum, detailed))

    out.append(f"\\ifdefined{INCLUDE_GUARD}\\else\n\\end{{document}}\n\\fi")
    return "\n\n".join(out) + "\n"


def build_pdf_driver(justifications: Sequence[str],
                     title: str = "Budget Justifications",
                     summaries: Sequence[Tuple[str, str]] = ()) -> str:
    """A driver that pulls every justification into one document.

    ``summaries`` are ``(heading, filename)`` pairs ``\\input`` at the very top,
    ahead of everything else -- the at-a-glance request tables.  They get a
    heading here because the summary files deliberately emit none of their own
    (so they can also drop into a section of the user's choosing).
    ``justifications`` are names of generated files in the same output directory,
    emitted in the order given (the combined/all-programs one first).

    Like the justifications and the summary tables, this file works BOTH ways:

    * on its own -- ``latexmk -pdf all_justifications.tex`` -- it emits the
      preamble, defines the include guard so each justification it pulls in skips
      its own preamble, and closes the document;
    * ``\\input`` into a larger proposal (which defines the guard in its own
      preamble) -- it then emits no preamble and no ``\\end{document}``, just the
      justifications one after another.

    It therefore tracks whether *it* opened the document with its own marker
    rather than reusing the include guard, which it may have defined itself."""
    out = [provenance_comment()]
    out.append(
        "% Every justification, one after another.  Compiles on its own\n"
        "% (`latexmk -pdf all_justifications.tex`).  To \\input it into a larger\n"
        "% proposal, put  \\def\\budgetjustificationincluded{}  in that document's\n"
        "% preamble; if it lives in a different directory, also set\n"
        "% \\def\\budgetjustificationpath{output/} so the \\input{}s below resolve.")
    out.append(_path_preamble())
    out.append(
        f"\\ifdefined{INCLUDE_GUARD}\\else\n"
        + PREAMBLE
        # Defined for the justifications pulled in below, so they skip their own
        # preambles; the separate marker records that WE opened the document.
        + f"\\def{INCLUDE_GUARD}{{}}\n"
        + f"\\def{STANDALONE_MARKER}{{}}\n"
        + f"\\title{{{escape_tex(title)}}}\n"
        + "\\begin{document}\n"
        + "\\maketitle\n"
        + "\\fi")
    body = []
    for heading, rel in summaries:
        body.append(f"\\section*{{{escape_tex(heading)}}}")
        body.append(_input(rel))
    if summaries:
        body.append("\\clearpage")
    for rel in justifications:
        body.append(_input(rel))
        body.append("\\clearpage")
    out.append("\n".join(body))
    out.append(f"\\ifdefined{STANDALONE_MARKER}\n\\end{{document}}\n\\fi")
    return "\n\n".join(out) + "\n"


# Every \newcommand file goes in this sub-directory of the output root, so a
# re-run's numbers can be dropped in wholesale (copy over `defs/`) without
# touching the justification/table text that lives beside it.
DEFS_DIR = "defs"
SUMMARY_DEFS_FILE = "faculty_summary_defs.tex"
SUMMARY_MACRO_PREFIX = "FacultySummary"
# Guards the defs file against being loaded twice (both summary tables \input
# it, so a document including both would otherwise redefine every \newcommand).
SUMMARY_DEFS_GUARD = r"\facultysummarydefsloaded"


def _stem(*parts) -> str:
    """A letters-only macro stem from arbitrary name parts."""
    return SUMMARY_MACRO_PREFIX + "".join(tex_name(p) for p in parts if p)


def _summary_stems(groups):
    """Assign each row a unique macro stem: ``[(gname, [(label, stem)],
    subtotal_stem)]``.

    Derived only from the group and faculty names, so the by-faculty and the
    by-year table independently arrive at the SAME stem for the same person and
    can share one defs file."""
    used = set()

    def uniq(cand):
        stem, i = cand, 1
        while stem in used:
            i += 1
            stem = cand + tex_name(str(i))     # ...Two, ...Three (letters only)
        used.add(stem)
        return stem

    out = []
    for gname, es in groups:
        rows = [(label, uniq(_stem(gname, label))) for label, *_ in es]
        out.append((gname, rows, uniq(_stem(gname, "Subtotal")) if gname else None))
    return out


def _money_macro(stem: str, suffix: str) -> str:
    """A dollar figure pulled from the summary defs file."""
    return f"\\${{}}\\{stem}{suffix}{{}}"


def build_summary_defs(by_faculty=None, by_year=None) -> str:
    """The ``\\newcommand`` definitions behind the summary tables.

    Both tables reference these macros instead of hard-coding their figures, so
    re-running the tool updates the numbers while any surrounding text -- in the
    generated tables or in prose the user writes -- stays put.

    ``by_faculty`` is the grouped ``(name, [(label, direct, indirect, total)])``
    data and ``by_year`` the grouped ``(name, [(label, [year, ...])])`` data;
    each contributes its own suffixes under a shared per-person stem.

    ROUNDING RULE: figures arrive already rounded per year and every total is a
    sum of those (see :func:`utkbudget.cli._summed`), so the two tables always
    agree -- a PI's ``Total`` here is by construction their ``YearsTotal``.
    Rounding again below is therefore a no-op on the values the CLI passes; it
    only guards against a caller handing in raw floats."""
    lines = [provenance_comment()]
    lines.append(
        "% Every summed figure in the two faculty summary tables.  The tables\n"
        "% \\input this file and reference these macros, so regenerating the\n"
        "% budgets updates the numbers and leaves the surrounding text alone.\n"
        "% You can reference them in your own prose too, e.g.\n"
        "%   The total request is \\${}\\FacultySummaryGrandTotal{}.")
    lines.append(f"\\ifdefined{SUMMARY_DEFS_GUARD}\\else")
    lines.append(f"\\def{SUMMARY_DEFS_GUARD}{{}}")

    def define(stem, suffix, value):
        lines.append("\\newcommand{\\%s%s}{%s}" % (stem, suffix, f"{value:,}"))

    if by_faculty:
        groups = [(g, [(n, round_dollar(d), round_dollar(i), round_dollar(t))
                       for n, d, i, t in es]) for g, es in by_faculty]
        stems = _summary_stems(groups)
        lines.append("% --- by faculty: direct / indirect / total ---")
        gt = [0, 0, 0]
        for (gname, rows, sub_stem), (_, es) in zip(stems, groups):
            st = [0, 0, 0]
            for (label, stem), (_, d, i, t) in zip(rows, es):
                lines.append(f"% {label}" + (f"  ({gname})" if gname else ""))
                for suffix, v in (("Direct", d), ("Indirect", i), ("Total", t)):
                    define(stem, suffix, v)
                st = [a + b for a, b in zip(st, (d, i, t))]
            if sub_stem:
                lines.append(f"% {gname} subtotal")
                for suffix, v in zip(("Direct", "Indirect", "Total"), st):
                    define(sub_stem, suffix, v)
            gt = [a + b for a, b in zip(gt, st)]
        lines.append("% grand total across every faculty member")
        for suffix, v in zip(("Direct", "Indirect", "Total"), gt):
            define(_stem("Grand"), suffix, v)

    if by_year:
        groups = [(g, [(n, [round_dollar(v) for v in years]) for n, years in es])
                  for g, es in by_year]
        stems = _summary_stems(groups)
        width = max((len(y) for _, es in groups for _, y in es),
                    default=len(PERIOD_WORDS))
        words = list(PERIOD_WORDS)[:width]
        lines.append("% --- by faculty and year ---")
        gt = [0] * width
        for (gname, rows, sub_stem), (_, es) in zip(stems, groups):
            st = [0] * width
            for (label, stem), (_, years) in zip(rows, es):
                years = list(years) + [0] * (width - len(years))
                lines.append(f"% {label}" + (f"  ({gname})" if gname else ""))
                for w, v in zip(words, years):
                    define(stem, f"Year{w}", v)
                define(stem, "YearsTotal", sum(years))
                st = [a + b for a, b in zip(st, years)]
            if sub_stem:
                lines.append(f"% {gname} subtotal by year")
                for w, v in zip(words, st):
                    define(sub_stem, f"Year{w}", v)
                define(sub_stem, "YearsTotal", sum(st))
            gt = [a + b for a, b in zip(gt, st)]
        lines.append("% grand total by year")
        for w, v in zip(words, gt):
            define(_stem("Grand"), f"Year{w}", v)
        define(_stem("Grand"), "YearsTotal", sum(gt))

    lines.append("\\fi")
    return "\n".join(lines) + "\n"


def build_faculty_summary(entries=None, title: str = "DOE Budget Request by Faculty",
                          groups=None) -> str:
    """A summary table -- one row per faculty with their final DOE ask -- meant
    to be ``\\input`` into a larger document.

    Flat form: ``entries`` is a list of ``(faculty_name, direct, indirect,
    total)`` with the dollar figures as numbers.

    Grouped form (multi-thrust proposals): ``groups`` is a list of
    ``(group_name, entries)`` -- one block per input sub-folder (e.g. ``Energy
    Frontier``, ``Intensity Frontier``, ``Theory Frontier``).  Each block gets a
    bold group heading, its PIs indented beneath it, and a subtotal row summed
    over that sub-folder; a grand total closes the table.

    Every figure is a macro reference resolved from ``faculty_summary_defs.tex``
    (see :func:`build_summary_defs`), which this file ``\\input``s -- so
    regenerating the budgets updates the numbers and leaves the text alone.

    Uses only plain ``\\hline`` rules so it drops into any document with no
    extra packages, and the same ``\\ifdefined\\budgetjustificationincluded``
    guard as the justifications so it also compiles on its own."""
    if groups is None:
        groups = [(None, list(entries))]
    stems = _summary_stems(groups)

    out = [provenance_comment()]
    out.append(
        "% Summary table: one row per faculty with their DOE request.  Compiles\n"
        "% on its own; to \\input it into a larger document, put\n"
        "% \\def\\budgetjustificationincluded{} in that document's preamble first.\n"
        "% The figures come from " + defs_input(SUMMARY_DEFS_FILE) + ",\n"
        "% \\input just below.")
    out.append(_path_preamble())
    out.append(
        f"\\ifdefined{INCLUDE_GUARD}\\else\n"
        + PREAMBLE
        + f"\\title{{{escape_tex(title)}}}\n"
        + "\\begin{document}\n\\maketitle\n\\fi")
    out.append(_input(defs_input(SUMMARY_DEFS_FILE)))

    lines = ["\\begin{center}", "\\begin{tabular}{lrrr}", "\\hline",
             "Faculty & Direct Costs & Indirect (F\\&A) & Total Requested \\\\",
             "\\hline"]
    for gname, rows, sub_stem in stems:
        indent = "\\quad " if gname is not None else ""
        if gname is not None:
            lines.append(f"\\multicolumn{{4}}{{l}}{{\\textbf{{{escape_tex(str(gname))}}}}} \\\\")
        for label, stem in rows:
            cells = " & ".join(_money_macro(stem, s)
                               for s in ("Direct", "Indirect", "Total"))
            lines.append(f"{indent}{escape_tex(str(label))} & {cells} \\\\")
        if sub_stem is not None:
            cells = " & ".join(f"\\textit{{{_money_macro(sub_stem, s)}}}"
                               for s in ("Direct", "Indirect", "Total"))
            lines.append(f"\\textit{{{escape_tex(str(gname))} subtotal}} & {cells} \\\\")
            lines.append("\\hline")
    if stems[-1][0] is None:
        lines.append("\\hline")
    grand = " & ".join(f"\\textbf{{{_money_macro(_stem('Grand'), s)}}}"
                       for s in ("Direct", "Indirect", "Total"))
    lines.append(f"\\textbf{{Total}} & {grand} \\\\")
    lines.append("\\hline")
    lines += ["\\end{tabular}", "\\end{center}"]
    out.append("\n".join(lines))

    out.append(f"\\ifdefined{INCLUDE_GUARD}\\else\n\\end{{document}}\n\\fi")
    return "\n\n".join(out) + "\n"


def build_faculty_summary_by_year(
        entries=None, title: str = "DOE Budget Request by Faculty and Year",
        groups=None) -> str:
    """A summary table -- one row per faculty, one column per year -- so a program
    manager can see each PI's ask per year at a glance.

    Flat form: ``entries`` is a list of ``(faculty_name, [year1, year2, ...])``
    with the per-period dollar figures as numbers.  Grouped form: ``groups`` is a
    list of ``(group_name, entries)``, one block per input sub-folder, each with a
    per-thrust subtotal row.  A "Total" row and a "Total Requested" column close
    the table.

    Only periods that carry money anywhere in the request get a column, so an
    unfunded year 4/5 is not printed.  Every figure is a macro reference resolved
    from ``faculty_summary_defs.tex``.  Same plain-``\\hline`` styling and
    ``\\ifdefined`` guard as :func:`build_faculty_summary`."""
    if groups is None:
        groups = [(None, list(entries))]
    rounded = [(g, [(n, [round_dollar(v) for v in years]) for n, years in es])
               for g, es in groups]
    all_entries = [e for _, es in rounded for e in es]
    stems = _summary_stems(groups)

    width = max((len(years) for _, years in all_entries), default=len(PERIOD_WORDS))

    def pad(years):
        return list(years) + [0] * (width - len(years))

    # Drop periods that are unfunded across the whole request.  (Which years get
    # a column is decided from the values; the values themselves are printed as
    # macros, so editing the defs file changes the figures, not the layout.)
    active = [i for i in range(width)
              if any(pad(years)[i] for _, years in all_entries)]
    if not active:
        active = [0]
    ncols = len(active)
    words = list(PERIOD_WORDS)[:width]

    def row(label, stem, fmt="{}"):
        cells = [fmt.format(_money_macro(stem, f"Year{words[i]}")) for i in active]
        cells.append(fmt.format(_money_macro(stem, "YearsTotal")))
        return f"{label} & " + " & ".join(cells) + " \\\\"

    out = [provenance_comment()]
    out.append(
        "% Summary table: one row per faculty, one column per year, so the ask\n"
        "% per year per PI is visible at a glance.  Compiles on its own; to\n"
        "% \\input it into a larger document, put\n"
        "% \\def\\budgetjustificationincluded{} in that document's preamble first.\n"
        "% The figures come from " + defs_input(SUMMARY_DEFS_FILE) + ",\n"
        "% \\input just below.")
    out.append(_path_preamble())
    out.append(
        f"\\ifdefined{INCLUDE_GUARD}\\else\n"
        + PREAMBLE
        + f"\\title{{{escape_tex(title)}}}\n"
        + "\\begin{document}\n\\maketitle\n\\fi")
    out.append(_input(defs_input(SUMMARY_DEFS_FILE)))

    header = " & ".join(f"Year {i + 1}" for i in active)
    lines = ["\\begin{center}",
             "\\begin{tabular}{l" + "r" * (ncols + 1) + "}", "\\hline",
             f"Faculty & {header} & Total Requested \\\\",
             "\\hline"]
    for gname, rows, sub_stem in stems:
        indent = "\\quad " if gname is not None else ""
        if gname is not None:
            lines.append(f"\\multicolumn{{{ncols + 2}}}{{l}}"
                         f"{{\\textbf{{{escape_tex(str(gname))}}}}} \\\\")
        for label, stem in rows:
            lines.append(row(f"{indent}{escape_tex(str(label))}", stem))
        if sub_stem is not None:
            lines.append(row(f"\\textit{{{escape_tex(str(gname))} subtotal}}",
                             sub_stem, "\\textit{{{}}}"))
            lines.append("\\hline")
    if stems[-1][0] is None:
        lines.append("\\hline")
    lines.append(row("\\textbf{Total}", _stem("Grand"), "\\textbf{{{}}}"))
    lines.append("\\hline")
    lines += ["\\end{tabular}", "\\end{center}"]
    out.append("\n".join(lines))

    out.append(f"\\ifdefined{INCLUDE_GUARD}\\else\n\\end{{document}}\n\\fi")
    return "\n\n".join(out) + "\n"


def write_document(path: str, *args, **kwargs) -> str:
    content = build_document(*args, **kwargs)
    with open(path, "w") as fh:
        fh.write(content)
    return path
