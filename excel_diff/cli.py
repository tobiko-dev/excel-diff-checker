"""Command-line interface for the Excel diff checker."""

from __future__ import annotations

import argparse
import json
import sys

from .diff import compare_files
from .report import render_text


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-diff",
        description="Compare two Excel (.xlsx) files and report which rows, "
        "columns and cells differ.",
    )
    parser.add_argument("file1", help="Path to the baseline .xlsx file")
    parser.add_argument("file2", help="Path to the comparison .xlsx file")
    parser.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        metavar="NAME",
        help="Only compare this sheet (repeatable). Default: all sheets.",
    )
    parser.add_argument(
        "--key",
        metavar="COLUMN",
        help="Header name of a key column. Rows are matched by this column's "
        "value instead of by position (survives inserted/reordered rows).",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="1-indexed row containing column headers (default: 1). "
        "Used with --key.",
    )
    parser.add_argument(
        "--count-empty",
        action="store_true",
        help="Treat a blank cell as different from an empty string. By "
        "default blanks are considered equal.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit the full diff as JSON instead of a text report.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        result = compare_files(
            args.file1,
            args.file2,
            sheets=args.sheets,
            key=args.key,
            header_row=args.header_row,
            ignore_empty=not args.count_empty,
        )
    except FileNotFoundError as exc:
        print(f"Error: file not found: {exc.filename}", file=sys.stderr)
        return 2
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps(result.to_dict(), indent=2, default=str))
    else:
        print(render_text(result))

    # Exit code 1 signals "differences found" so the tool composes in scripts.
    return 1 if result.has_differences else 0


if __name__ == "__main__":
    sys.exit(main())
