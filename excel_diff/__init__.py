"""Excel difference checker.

Compare two ``.xlsx`` workbooks and find out which sheets, rows, columns and
cells differ.

Example:
    >>> from excel_diff import compare_files, render_text
    >>> diff = compare_files("old.xlsx", "new.xlsx")
    >>> print(render_text(diff))
"""

from .diff import compare_files
from .model import CellDiff, SheetDiff, WorkbookDiff
from .report import render_text

__all__ = [
    "compare_files",
    "render_text",
    "CellDiff",
    "SheetDiff",
    "WorkbookDiff",
]

__version__ = "0.1.0"
