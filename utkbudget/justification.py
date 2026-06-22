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


def _m(prefix: str, field: str) -> str:
    """A LaTeX macro reference that is safe to follow with a space."""
    return f"\\{prefix}{field}{{}}"


def _usd(prefix: str, field: str) -> str:
    return f"\\${_m(prefix, field)}"


PREAMBLE = r"""\documentclass[11pt]{article}
\usepackage[margin=1in]{geometry}
\usepackage{booktabs}
\usepackage{array}
\usepackage{longtable}
\usepackage{enumitem}
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


def _summary_table(prefix: str) -> str:
    """A by-category cost summary table built from the macros."""
    rows = [
        ("Salaries \\& Wages", "WagesTotal"),
        ("Fringe Benefits", "FringeTotal"),
        ("Equipment", "EquipmentTotal"),
        ("Travel", "TravelTotal"),
        ("Participant Support", "ParticipantSupportTotal"),
        ("Other Direct Costs", "OtherDirectTotal"),
        ("Total Direct Costs", "DirectTotal"),
        ("Indirect (F\\&A) Costs", "IndirectTotal"),
    ]
    lines = [
        "\\begin{center}",
        "\\begin{tabular}{lr}",
        "\\toprule",
        "Cost Category & Total Requested \\\\",
        "\\midrule",
    ]
    for label, base in rows:
        lines.append(f"{label} & {_usd(prefix, base)} \\\\")
    lines.append("\\midrule")
    lines.append(f"\\textbf{{Total Project Cost}} & \\textbf{{{_usd(prefix, 'GrandTotal')}}} \\\\")
    lines.append("\\bottomrule")
    lines.append("\\end{tabular}")
    lines.append("\\end{center}")
    return "\n".join(lines) + "\n"


def render_section(prefix: str, title: str, is_sum: bool = False) -> str:
    """Return the LaTeX for one budget's justification section."""
    if is_sum:
        opener = (
            f"This section justifies the {title} request, which is the sum of the "
            "individual budgets justified above. It represents the total funds "
            "requested from the Department of Energy across all contributing "
            "programs and personnel."
        )
    else:
        opener = (
            f"The following justifies the funds requested in the {title} budget. "
            "Costs are organized following the standard DOE Office of Science "
            "budget categories."
        )

    parts: List[str] = []
    parts.append(f"\\section{{{title}}}")
    parts.append(opener)
    parts.append(_summary_table(prefix))

    # --- A. Senior Personnel -------------------------------------------
    parts.append("\\subsection*{A. Senior Personnel}")
    parts.append(
        "Funds are requested for the academic-year and/or summer effort of the "
        "senior personnel listed in the budget. Salaries are based on current "
        "institutional rates and escalated at the approved annual rate. The "
        f"total request for senior-personnel salaries is {_usd(prefix, 'SeniorSubtotalTotal')}."
    )

    # --- B. Other Personnel --------------------------------------------
    parts.append("\\subsection*{B. Other Personnel}")
    parts.append(
        "Funds are requested for postdoctoral researchers, graduate research "
        "assistants (GRAs), undergraduate researchers, and other personnel "
        "essential to the proposed research. Postdoctoral and student effort "
        "drives the technical work of the project. The total request for other "
        f"personnel is {_usd(prefix, 'OtherPersonnelSubtotalTotal')}, bringing total "
        f"salaries and wages to {_usd(prefix, 'WagesTotal')}."
    )

    # --- C. Fringe Benefits --------------------------------------------
    parts.append("\\subsection*{C. Fringe Benefits}")
    parts.append(
        "Fringe benefits are calculated using the institution's federally "
        "negotiated fringe-benefit rates applicable to each personnel category "
        "(faculty, postdoctoral, student, and staff). The total fringe-benefit "
        f"request is {_usd(prefix, 'FringeTotal')}, for total salaries and benefits "
        f"of {_usd(prefix, 'SalaryAndBenefitsTotal')}."
    )

    # --- D. Equipment ---------------------------------------------------
    parts.append("\\subsection*{D. Equipment}")
    parts.append(
        "Equipment is defined as items of tangible personal property with a "
        "useful life of more than one year and a unit acquisition cost of "
        f"\\$5,000 or more. The total equipment request is {_usd(prefix, 'EquipmentTotal')}. "
        "Each item, where requested, is itemized in the accompanying budget "
        "spreadsheet."
    )

    # --- E. Travel (emphasized) ----------------------------------------
    parts.append("\\subsection*{E. Travel}")
    parts.append(
        f"A total of {_usd(prefix, 'TravelTotal')} is requested for travel. Travel is "
        "essential to disseminate results, participate in DOE program activities, "
        "and sustain the scientific collaborations on which this research depends. "
        "The request is broken down by period and by domestic vs.\\ foreign travel "
        "below."
    )
    parts.append(_travel_table(prefix))
    parts.append(
        "\\textbf{Domestic travel} "
        f"({_usd(prefix, 'DomesticTravelTotal')} total) supports attendance at "
        "topical workshops and major conferences (e.g., APS April Meeting and "
        "DPF), DOE-sponsored principal-investigator and program review meetings, "
        "and collaboration meetings at partner national laboratories and "
        "universities. Costs are estimated per traveler and include airfare, "
        "lodging at prevailing per-diem rates, ground transportation, and "
        "conference registration."
    )
    parts.append(
        "\\textbf{Foreign travel} "
        f"({_usd(prefix, 'ForeignTravelTotal')} total) supports participation in "
        "international collaboration meetings, experiment shifts, and conferences "
        "central to the proposed program. All foreign travel will comply with the "
        "Fly America Act and DOE foreign-travel approval requirements."
    )

    # --- F. Participant Support ----------------------------------------
    parts.append("\\subsection*{F. Participant Support Costs}")
    parts.append(
        f"A total of {_usd(prefix, 'ParticipantSupportTotal')} is requested for "
        "participant support costs (stipends, travel, and subsistence for "
        "participants in workshops, schools, or training activities associated "
        "with the project). These funds are budgeted and accounted for separately "
        "and are excluded from the indirect-cost base."
    )

    # --- G. Other Direct Costs -----------------------------------------
    parts.append("\\subsection*{G. Other Direct Costs}")
    parts.append(
        f"Other direct costs total {_usd(prefix, 'OtherDirectTotal')} and include "
        f"materials and supplies ({_usd(prefix, 'SuppliesTotal')}), publication and "
        "page charges, and other direct project expenses. Graduate-student "
        "tuition and mandatory fees "
        f"({_usd(prefix, 'TuitionSubtotalTotal')}) are requested for the GRAs "
        "supported on this project, consistent with institutional policy."
    )

    # --- Indirect Costs -------------------------------------------------
    parts.append("\\subsection*{Indirect (F\\&A) Costs}")
    parts.append(
        "Indirect costs are computed on the Modified Total Direct Cost (MTDC) base "
        "using the institution's federally negotiated rate "
        f"({_m(prefix, 'FandARateType')}; {_m(prefix, 'OverheadRateYearOne')}\\% in "
        f"the first period). The total indirect-cost request is "
        f"{_usd(prefix, 'IndirectTotal')}."
    )

    # --- Total ----------------------------------------------------------
    parts.append("\\subsection*{Total Requested}")
    parts.append(
        f"The total funds requested from the Department of Energy for {title} are "
        f"\\textbf{{{_usd(prefix, 'GrandTotal')}}} over the project period "
        f"(direct costs {_usd(prefix, 'DirectTotal')} plus indirect costs "
        f"{_usd(prefix, 'IndirectTotal')})."
    )

    return "\n\n".join(parts)


def build_document(
    title: str,
    defs_inputs: Sequence[str],
    sections: Sequence[Tuple[str, str, bool]],
    intro: str = "",
) -> str:
    """Assemble a complete justification document.

    ``defs_inputs`` are relative paths to ``\\input`` (the generated defs
    files).  ``sections`` is a list of ``(prefix, heading, is_sum)``.
    """
    out: List[str] = [PREAMBLE]
    out.append(f"\\title{{{title}}}")
    out.append(f"\\date{{{_dt.date.today():%B %-d, %Y}}}")
    out.append("\\begin{document}")
    out.append("\\maketitle")

    out.append("% --- generated budget definitions ---")
    for rel in defs_inputs:
        # strip the .tex extension for \input
        base = rel[:-4] if rel.endswith(".tex") else rel
        out.append(f"\\input{{{base}}}")
    out.append("")

    if intro:
        out.append(intro)

    for prefix, heading, is_sum in sections:
        out.append(render_section(prefix, heading, is_sum))

    out.append("\\end{document}")
    return "\n\n".join(out) + "\n"


def write_document(path: str, *args, **kwargs) -> str:
    content = build_document(*args, **kwargs)
    with open(path, "w") as fh:
        fh.write(content)
    return path
