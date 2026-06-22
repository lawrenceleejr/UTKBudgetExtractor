"""UTKBudgetExtractor -- parse UTK proposal-budget spreadsheets into LaTeX.

See :mod:`utkbudget.cli` for the command-line entry point, or import the
building blocks directly:

    from utkbudget.extractor import extract_budget
    from utkbudget.merge import merge_budgets, write_merged_workbook
    from utkbudget.texdefs import write_defs
    from utkbudget.justification import write_document
"""

from .extractor import Budget, Field, extract_budget
from .merge import merge_budgets, write_merged_workbook
from .texdefs import tex_prefix, write_defs
from .justification import build_document, write_document

__all__ = [
    "Budget",
    "Field",
    "extract_budget",
    "merge_budgets",
    "write_merged_workbook",
    "tex_prefix",
    "write_defs",
    "build_document",
    "write_document",
]

__version__ = "2.0.0"
