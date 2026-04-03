# excel-fillna

`excel-fillna` is a small pip-installable library with an `fna` CLI for filling empty Excel cells inside a selected range.

It supports:

- filling empty cells with `NA` or custom text
- skipping one or more excluded ranges
- optionally merging contiguous empty cells vertically within a column before filling them
- working against the active worksheet or a named sheet

The implementation uses `openpyxl`, so it supports `.xlsx` and `.xlsm` workbooks. `.xls` is not supported.

## Install

```bash
pip install .
```

For development:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -e ".[dev]"
```

## CLI Usage

```bash
fna workbook.xlsx \
  --sheet Sheet1 \
  --range A1:C20 \
  --exclude-range B2:B4 \
  --exclude-range C10 \
  --fill-text MISSING \
  --merge-empty-runs \
  --output workbook.filled.xlsx
```

### CLI options

- `input_path`: source workbook path
- `--sheet`: worksheet name, defaults to the active sheet
- `--range`: target range to inspect, required
- `--exclude-range`: ranges to leave untouched; repeat the flag or pass comma-separated ranges
- `--fill-text`: replacement text, defaults to `NA`
- `--merge-empty-runs`: merge contiguous empty cells vertically within each column before filling
- `--output`: destination path, defaults to `<input>.filled<suffix>`

## Python Usage

```python
from excel_fillna import process_workbook

result = process_workbook(
    "workbook.xlsx",
    target_range="A1:C20",
    excluded_ranges=["B2:B4", "C10"],
    fill_value="NA",
    merge_empty_runs=True,
)

print(result.output_path)
print(result.filled_cells)
print(result.merged_ranges)
```

## Merge behavior

When `merge_empty_runs=True`, the library scans each column in the selected range and looks for contiguous empty cells. Runs of length 2 or more are merged vertically and the top cell is filled with the requested value.

Example:

- `A1`, `A2`, and `A3` are empty
- `A4` contains data
- target range is `A1:A4`
- merge mode creates `A1:A3` and fills `A1` with `NA`

Existing merged cells are supported. If an existing merged range has an empty anchor cell, the anchor is filled, but the library does not attempt to re-merge or resize that existing merged range.

## Specs

- Functional spec: [`docs/FUNCTIONAL_SPEC.md`](docs/FUNCTIONAL_SPEC.md)
- CLI spec: [`docs/CLI_SPEC.md`](docs/CLI_SPEC.md)

## Tests

```bash
.venv/bin/pytest
```
