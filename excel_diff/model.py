"""Data model for Excel diff results.

These dataclasses describe the outcome of comparing two workbooks. They are
plain data holders so the results can be serialized (see ``to_dict``) and
rendered by any frontend we build later.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class CellDiff:
    """A single cell whose value differs between the two files."""

    sheet: str
    row: int  # 1-indexed, matches the Excel row number
    column: int  # 1-indexed
    column_letter: str  # e.g. "A", "B", "AB"
    old_value: Any
    new_value: Any
    # Optional human-friendly label for the row/column when headers or a key
    # column are used to align the sheets.
    row_key: Any = None
    column_name: Any = None

    def to_dict(self) -> dict:
        return {
            "sheet": self.sheet,
            "row": self.row,
            "column": self.column,
            "column_letter": self.column_letter,
            "old_value": _jsonable(self.old_value),
            "new_value": _jsonable(self.new_value),
            "row_key": _jsonable(self.row_key),
            "column_name": _jsonable(self.column_name),
        }


@dataclass
class SheetDiff:
    """The differences found within a single sheet."""

    name: str
    # One of: "modified", "identical", "added", "removed".
    status: str
    cell_diffs: list[CellDiff] = field(default_factory=list)
    changed_rows: set[int] = field(default_factory=set)
    changed_columns: set[int] = field(default_factory=set)
    added_rows: list[int] = field(default_factory=list)
    removed_rows: list[int] = field(default_factory=list)
    added_columns: list[int] = field(default_factory=list)
    removed_columns: list[int] = field(default_factory=list)
    dims_old: tuple[int, int] = (0, 0)  # (max_row, max_col) in file1
    dims_new: tuple[int, int] = (0, 0)  # (max_row, max_col) in file2

    @property
    def has_differences(self) -> bool:
        return self.status != "identical"

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "status": self.status,
            "cell_diffs": [c.to_dict() for c in self.cell_diffs],
            "changed_rows": sorted(self.changed_rows),
            "changed_columns": sorted(self.changed_columns),
            "added_rows": self.added_rows,
            "removed_rows": self.removed_rows,
            "added_columns": self.added_columns,
            "removed_columns": self.removed_columns,
            "dims_old": list(self.dims_old),
            "dims_new": list(self.dims_new),
        }


@dataclass
class WorkbookDiff:
    """Top-level result of comparing two workbooks."""

    file1: str
    file2: str
    sheet_diffs: list[SheetDiff] = field(default_factory=list)
    added_sheets: list[str] = field(default_factory=list)
    removed_sheets: list[str] = field(default_factory=list)

    @property
    def has_differences(self) -> bool:
        return (
            bool(self.added_sheets)
            or bool(self.removed_sheets)
            or any(s.has_differences for s in self.sheet_diffs)
        )

    def to_dict(self) -> dict:
        return {
            "file1": self.file1,
            "file2": self.file2,
            "has_differences": self.has_differences,
            "added_sheets": self.added_sheets,
            "removed_sheets": self.removed_sheets,
            "sheets": [s.to_dict() for s in self.sheet_diffs],
        }


def _jsonable(value: Any) -> Any:
    """Convert values (e.g. datetimes) into something JSON-serializable."""
    import datetime

    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    return value
