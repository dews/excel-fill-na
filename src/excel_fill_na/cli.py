from __future__ import annotations

import argparse

from .core import DEFAULT_FILL_VALUE, process_workbook


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="fna",
        description="Fill empty Excel cells inside a selected range.",
    )
    parser.add_argument("input_path", help="Path to the source .xlsx or .xlsm workbook.")
    parser.add_argument(
        "-r",
        "--range",
        dest="target_range",
        required=True,
        help="Target cell range to process, for example A1:C20.",
    )
    parser.add_argument(
        "-x",
        "--exclude-range",
        dest="excluded_ranges",
        action="append",
        default=[],
        help="Cell range to skip. Repeat the flag or pass comma-separated ranges.",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        dest="sheet_name",
        help="Worksheet name. Defaults to the active sheet.",
    )
    parser.add_argument(
        "-t",
        "--fill-text",
        default=DEFAULT_FILL_VALUE,
        help="Replacement text for empty cells. Defaults to NA.",
    )
    parser.add_argument(
        "--merge-empty-runs",
        action="store_true",
        help="Merge contiguous empty cells vertically within each column before filling them.",
    )
    parser.add_argument(
        "-o",
        "--output",
        dest="output_path",
        help="Destination workbook path. Defaults to <input>.filled<suffix>.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        result = process_workbook(
            args.input_path,
            target_range=args.target_range,
            excluded_ranges=args.excluded_ranges,
            fill_value=args.fill_text,
            merge_empty_runs=args.merge_empty_runs,
            sheet_name=args.sheet_name,
            output_path=args.output_path,
        )
    except Exception as exc:
        parser.exit(2, f"error: {exc}\n")

    merged_count = len(result.merged_ranges)
    merged_label = "range" if merged_count == 1 else "ranges"
    cell_label = "cell" if result.filled_cells == 1 else "cells"
    print(
        f"Wrote {result.output_path} | "
        f"filled {result.filled_cells} empty {cell_label} | "
        f"created {merged_count} merged {merged_label}."
    )
    if result.merged_ranges:
        print("Merged ranges: " + ", ".join(result.merged_ranges))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
