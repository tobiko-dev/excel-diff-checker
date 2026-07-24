# Excel Difference Checker

A backend tool to compare two Excel (`.xlsx`) workbooks and report **which
sheets, rows, columns, and cells differ**. Use it as a library or from the
command line. A frontend can be layered on top later — the engine already
returns fully structured, JSON-serializable results.

## Features

- **Sheet-aware**: compares every sheet by name; reports sheets added/removed.
- **Row & column awareness**: tells you exactly which rows and which columns
  contain differences, plus rows/columns that exist in only one file.
- **Cell-level detail**: every changed cell with its location (e.g. `B3`) and
  its old → new values.
- **Two matching strategies**:
  - *Positional* (default): compares cell `(row, col)` to cell `(row, col)`.
  - *Keyed* (`--key`): matches rows by a key column's value, so inserted,
    deleted, or reordered rows don't create false differences. This answers
    "which records changed?" rather than "which cells shifted?".
- **Blank handling**: blank cells and empty strings are treated as equal by
  default (toggle with `--count-empty`).
- **JSON output** for easy consumption by any frontend or pipeline.

## Installation

```bash
pip install -r requirements.txt
```

Requires Python 3.9+ and `openpyxl`.

## Command-line usage

```bash
# Positional comparison (default)
python -m excel_diff.cli old.xlsx new.xlsx

# Match rows by a key column called "id" (recommended for data tables)
python -m excel_diff.cli old.xlsx new.xlsx --key id

# Restrict to specific sheets
python -m excel_diff.cli old.xlsx new.xlsx --sheet Inventory --sheet Orders

# Machine-readable output
python -m excel_diff.cli old.xlsx new.xlsx --key id --json
```

Options:

| Flag | Description |
| --- | --- |
| `--sheet NAME` | Only compare this sheet (repeatable). Default: all sheets. |
| `--key COLUMN` | Header name of the key column for row matching. |
| `--header-row N` | 1-indexed header row (default `1`); used with `--key`. |
| `--count-empty` | Treat a blank cell as different from an empty string. |
| `--json` | Emit the full diff as JSON. |

**Exit codes**: `0` = no differences, `1` = differences found, `2` = error.
This makes the tool easy to drop into scripts and CI.

## Library usage

```python
from excel_diff import compare_files, render_text

diff = compare_files("old.xlsx", "new.xlsx", key="id")

print(diff.has_differences)                 # True / False
for sheet in diff.sheet_diffs:
    print(sheet.name, sorted(sheet.changed_rows), sorted(sheet.changed_columns))
    for cell in sheet.cell_diffs:
        print(cell.column_letter, cell.row, cell.old_value, "->", cell.new_value)

# Human-readable report, or JSON for a frontend
print(render_text(diff))
data = diff.to_dict()
```

## Result structure

`compare_files(...)` returns a `WorkbookDiff`:

- `has_differences: bool`
- `added_sheets` / `removed_sheets: list[str]`
- `sheet_diffs: list[SheetDiff]`, each with:
  - `status`: `"identical"` or `"modified"`
  - `changed_rows` / `changed_columns`: sets of 1-indexed positions
  - `added_rows` / `removed_rows`, `added_columns` / `removed_columns`
  - `cell_diffs: list[CellDiff]` — location + `old_value`/`new_value`
    (plus `row_key`/`column_name` in keyed mode)

Every object has a `to_dict()` method for serialization.

## Running the tests

```bash
pip install pytest
python -m pytest tests/ -q
```

## Roadmap

- Frontend (upload two files, view highlighted diff).
- Support for `.xls` and CSV inputs.
- Cell-formatting / formula-aware comparison.
