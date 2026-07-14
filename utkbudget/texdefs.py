"""Write a :class:`~utkbudget.extractor.Budget` to a LaTeX ``\\newcommand`` file.

Every numeric/text field becomes a macro named ``\\<prefix><FieldName>`` so the
values can be dropped straight into a proposal or budget-justification document
with ``\\input{...}``.  Because LaTeX command names may only contain letters,
prefixes are sanitised by :func:`tex_prefix` (digits are spelled out, other
characters are dropped).
"""

from __future__ import annotations

import datetime as _dt
import os

from .extractor import (
    Budget,
    KIND_DATE,
    KIND_MONEY,
    KIND_MONTHS,
    KIND_RATE,
)
from .provenance import provenance_lines

_DIGIT_WORDS = {
    "0": "Zero", "1": "One", "2": "Two", "3": "Three", "4": "Four",
    "5": "Five", "6": "Six", "7": "Seven", "8": "Eight", "9": "Nine",
}

# LaTeX characters that must be escaped inside ordinary text.
_TEX_ESCAPES = {
    "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
    "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def tex_prefix(name: str) -> str:
    """Turn an arbitrary string into a valid (letters-only) LaTeX macro prefix.

    ``"Proposal_Budget 2.xlsx"`` -> ``"ProposalBudgetTwo"``.
    """
    base = os.path.splitext(os.path.basename(name))[0]
    out = []
    for ch in base:
        if ch.isalpha():
            out.append(ch)
        elif ch.isdigit():
            out.append(_DIGIT_WORDS[ch])
    return "".join(out) or "Budget"


def escape_tex(text: str) -> str:
    return "".join(_TEX_ESCAPES.get(ch, ch) for ch in text)


def format_value(field) -> str:
    """Render one :class:`~utkbudget.extractor.Field` value as a LaTeX string."""
    value = field.value
    kind = field.kind

    if value is None or value == "":
        # Empty money/rate cells read more naturally as zero in prose.
        return "0.00" if kind in (KIND_MONEY,) else ""

    if kind == KIND_MONEY:
        try:
            return "{:,.2f}".format(float(value))
        except (TypeError, ValueError):
            return escape_tex(str(value))

    if kind == KIND_RATE:
        try:
            return "{:.1f}".format(float(value))
        except (TypeError, ValueError):
            return escape_tex(str(value))

    if kind == KIND_MONTHS:
        try:
            return "{:g}".format(float(value))
        except (TypeError, ValueError):
            return escape_tex(str(value))

    if kind == KIND_DATE:
        if isinstance(value, (_dt.datetime, _dt.date)):
            # value.day avoids strftime's platform-specific %-d / %#d.
            return f"{value:%B} {value.day}, {value:%Y}"
        return escape_tex(str(value))

    return escape_tex(str(value))


def render_defs(budget: Budget, prefix: str, header_note: str = "",
                include_source: bool = True, bare: bool = False) -> str:
    """Return ``budget`` as ``\\newcommand`` definitions text.

    ``bare`` drops the comment header entirely (used when the definitions are
    inlined into a document that already carries its own provenance and must not
    reveal the source filename).  ``include_source`` keeps/drops just the
    ``% Source:`` line.
    """
    lines = []
    if not bare:
        lines.extend(f"% {line}" for line in provenance_lines())
        if include_source and budget.source:
            lines.append(f"% Source: {budget.source}")
        if header_note:
            lines.append(f"% {header_note}")
        lines.append(f"% Macro prefix: {prefix}")
        lines.append("")
    for field in budget:
        lines.append("\\newcommand{\\%s%s}{%s}" % (prefix, field.name, format_value(field)))
    return "\n".join(lines) + "\n"


def write_defs(budget: Budget, prefix: str, out_path: str,
               header_note: str = "") -> str:
    """Write ``budget`` to ``out_path`` as ``\\newcommand`` definitions.

    Returns ``out_path`` for convenience.
    """
    with open(out_path, "w") as fh:
        fh.write(render_defs(budget, prefix, header_note))
    return out_path
