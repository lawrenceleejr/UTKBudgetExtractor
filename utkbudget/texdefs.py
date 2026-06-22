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
from typing import Iterable

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
            return value.strftime("%B %-d, %Y")
        return escape_tex(str(value))

    return escape_tex(str(value))


def write_defs(budget: Budget, prefix: str, out_path: str,
               header_note: str = "") -> str:
    """Write ``budget`` to ``out_path`` as ``\\newcommand`` definitions.

    Returns ``out_path`` for convenience.
    """
    with open(out_path, "w") as fh:
        for line in provenance_lines():
            fh.write(f"% {line}\n")
        if budget.source:
            fh.write(f"% Source: {budget.source}\n")
        if header_note:
            fh.write(f"% {header_note}\n")
        fh.write(f"% Macro prefix: {prefix}\n\n")
        for field in budget:
            fh.write("\\newcommand{\\%s%s}{%s}\n" % (prefix, field.name, format_value(field)))
    return out_path


def write_defs_for_files(budgets: Iterable, out_dir: str):
    """Write one defs file per (prefix, budget, filename) tuple.

    ``budgets`` is an iterable of ``(prefix, budget, tex_filename)``.
    """
    written = []
    for prefix, budget, tex_name in budgets:
        out_path = os.path.join(out_dir, tex_name)
        write_defs(budget, prefix, out_path)
        written.append(out_path)
    return written
