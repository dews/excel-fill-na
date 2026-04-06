from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from ._archive import (
    find_value_metadata_cells,
    persist_workbook_changes,
    resolve_worksheet_archive_path,
)
from ._models import DEFAULT_FILL_VALUE, FillResult
from ._planning import apply_fill_plan_to_worksheet, build_fill_plan, resolve_worksheet

FillResult.__module__ = __name__


def process_workbook(
    input_path: str | Path,
    *,
    target_range: str,
    excluded_ranges: Iterable[str] | None = None,
    fill_value: str = DEFAULT_FILL_VALUE,
    merge_empty_runs: bool = False,
    sheet_name: str | None = None,
    output_path: str | Path | None = None,
) -> FillResult:
    """Load a workbook, fill empty cells, and save the result."""
    source = Path(input_path)
    if not source.exists():
        raise FileNotFoundError(f"Workbook not found: {source}")

    destination = (
        Path(output_path)
        if output_path is not None
        else source.with_name(f"{source.stem}.filled{source.suffix}")
    )

    workbook = load_workbook(
        filename=source,
        keep_vba=source.suffix.lower() == ".xlsm",
    )
    try:
        worksheet = resolve_worksheet(workbook, sheet_name)
        worksheet_path = resolve_worksheet_archive_path(source, worksheet.title)
        plan = build_fill_plan(
            worksheet,
            target_range=target_range,
            excluded_ranges=excluded_ranges,
            fill_value=fill_value,
            merge_empty_runs=merge_empty_runs,
            preserved_coordinates=find_value_metadata_cells(source, worksheet_path),
        )
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            close()

    persist_workbook_changes(
        source=source,
        destination=destination,
        worksheet_path=worksheet_path,
        plan=plan,
    )

    return FillResult(
        sheet_name=plan.sheet_name,
        target_range=plan.target_range,
        fill_value=plan.fill_value,
        filled_cells=plan.filled_cells,
        merged_ranges=plan.merged_ranges,
        output_path=destination,
    )


def fill_empty_cells(
    worksheet: Worksheet,
    *,
    target_range: str,
    excluded_ranges: Iterable[str] | None = None,
    fill_value: str = DEFAULT_FILL_VALUE,
    merge_empty_runs: bool = False,
) -> FillResult:
    """Fill empty cells inside a worksheet range."""
    plan = build_fill_plan(
        worksheet,
        target_range=target_range,
        excluded_ranges=excluded_ranges,
        fill_value=fill_value,
        merge_empty_runs=merge_empty_runs,
    )
    apply_fill_plan_to_worksheet(worksheet, plan)
    return FillResult(
        sheet_name=plan.sheet_name,
        target_range=plan.target_range,
        fill_value=plan.fill_value,
        filled_cells=plan.filled_cells,
        merged_ranges=plan.merged_ranges,
    )
