from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

DEFAULT_FILL_VALUE = "NA"


@dataclass(frozen=True, slots=True)
class FillResult:
    sheet_name: str
    target_range: str
    fill_value: str
    filled_cells: int
    merged_ranges: tuple[str, ...]
    output_path: Path | None = None
    deleted_rows: int = 0


@dataclass(frozen=True, slots=True)
class CellWrite:
    row: int
    column: int
    value: str


@dataclass(frozen=True, slots=True)
class FillPlan:
    sheet_name: str
    target_range: str
    fill_value: str
    filled_cells: int
    merged_ranges: tuple[str, ...]
    cell_writes: tuple[CellWrite, ...]
    deleted_row_indices: tuple[int, ...] = ()
