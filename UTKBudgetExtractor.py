#!/usr/bin/env python
"""Convenience launcher: ``python UTKBudgetExtractor.py <folder> [-o out]``.

The implementation lives in the :mod:`utkbudget` package; this thin wrapper
keeps the historical script name working.
"""

from utkbudget.cli import main

if __name__ == "__main__":
    main()
