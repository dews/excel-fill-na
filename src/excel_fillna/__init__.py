"""Public package interface for excel-fillna."""

from .core import DEFAULT_FILL_VALUE, FillResult, fill_empty_cells, process_workbook

__all__ = ["DEFAULT_FILL_VALUE", "FillResult", "fill_empty_cells", "process_workbook"]
__version__ = "0.1.0"

