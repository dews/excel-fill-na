from __future__ import annotations

from typing import Iterable

from openpyxl.worksheet.cell_range import CellRange


def parse_ranges(range_strings: Iterable[str] | None) -> tuple[CellRange, ...]:
    if not range_strings:
        return ()

    parsed_ranges: list[CellRange] = []
    for range_string in range_strings:
        for candidate in str(range_string).split(","):
            cleaned = candidate.strip()
            if cleaned:
                parsed_ranges.append(parse_range(cleaned))
    return tuple(parsed_ranges)


def parse_range(range_string: str) -> CellRange:
    try:
        return CellRange(range_string)
    except ValueError as exc:
        raise ValueError(f"Invalid cell range: {range_string!r}") from exc


def contains_cell(cell_range: CellRange, row: int, column: int) -> bool:
    return (
        cell_range.min_row <= row <= cell_range.max_row
        and cell_range.min_col <= column <= cell_range.max_col
    )


def is_excluded(exclusions: tuple[CellRange, ...], row: int, column: int) -> bool:
    return any(contains_cell(cell_range, row, column) for cell_range in exclusions)
