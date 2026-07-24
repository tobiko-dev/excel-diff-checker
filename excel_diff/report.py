"""Human-readable rendering of a :class:`WorkbookDiff`."""

from __future__ import annotations

from .model import SheetDiff, WorkbookDiff


def _fmt(value) -> str:
    if value is None:
        return "(empty)"
    return repr(value) if isinstance(value, str) else str(value)


def _summarize_ranges(numbers: list[int]) -> str:
    """Collapse a sorted list of ints into compact range notation."""
    if not numbers:
        return ""
    nums = sorted(numbers)
    parts: list[str] = []
    start = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        parts.append(str(start) if start == prev else f"{start}-{prev}")
        start = prev = n
    parts.append(str(start) if start == prev else f"{start}-{prev}")
    return ", ".join(parts)


def render_text(diff: WorkbookDiff, *, max_cells: int = 100) -> str:
    """Render the diff as a plain-text report."""
    lines: list[str] = []
    lines.append("=" * 70)
    lines.append("Excel Difference Report")
    lines.append("=" * 70)
    lines.append(f"File 1 (baseline):   {diff.file1}")
    lines.append(f"File 2 (comparison): {diff.file2}")
    lines.append("")

    if not diff.has_differences:
        lines.append("No differences found. The files are identical.")
        return "\n".join(lines)

    if diff.added_sheets:
        lines.append(f"Sheets only in file 2: {', '.join(diff.added_sheets)}")
    if diff.removed_sheets:
        lines.append(f"Sheets only in file 1: {', '.join(diff.removed_sheets)}")
    if diff.added_sheets or diff.removed_sheets:
        lines.append("")

    for sheet in diff.sheet_diffs:
        lines.extend(_render_sheet(sheet, max_cells))

    return "\n".join(lines)


def _render_sheet(sheet: SheetDiff, max_cells: int) -> list[str]:
    lines: list[str] = []
    lines.append("-" * 70)
    if not sheet.has_differences:
        lines.append(f"Sheet '{sheet.name}': identical")
        lines.append("")
        return lines

    lines.append(f"Sheet '{sheet.name}': {len(sheet.cell_diffs)} changed cell(s)")

    changed_rows = _summarize_ranges(sorted(sheet.changed_rows))
    changed_cols = _summarize_ranges(sorted(sheet.changed_columns))
    if changed_rows:
        lines.append(f"  Rows with differences:    {changed_rows}")
    if changed_cols:
        lines.append(f"  Columns with differences: {changed_cols}")

    if sheet.added_rows:
        lines.append(
            f"  Rows only in file 2:      {_summarize_ranges(sheet.added_rows)}"
        )
    if sheet.removed_rows:
        lines.append(
            f"  Rows only in file 1:      {_summarize_ranges(sheet.removed_rows)}"
        )
    if sheet.added_columns:
        lines.append(
            f"  Columns only in file 2:   {_summarize_ranges(sheet.added_columns)}"
        )
    if sheet.removed_columns:
        lines.append(
            f"  Columns only in file 1:   {_summarize_ranges(sheet.removed_columns)}"
        )

    if sheet.cell_diffs:
        lines.append("")
        lines.append("  Cell changes:")
        for cd in sheet.cell_diffs[:max_cells]:
            loc = f"{cd.column_letter}{cd.row}"
            label = ""
            if cd.row_key is not None or cd.column_name is not None:
                label = f" [{_fmt(cd.row_key)} / {cd.column_name}]"
            lines.append(
                f"    {loc}{label}: {_fmt(cd.old_value)} -> {_fmt(cd.new_value)}"
            )
        if len(sheet.cell_diffs) > max_cells:
            lines.append(
                f"    ... and {len(sheet.cell_diffs) - max_cells} more"
            )
    lines.append("")
    return lines
