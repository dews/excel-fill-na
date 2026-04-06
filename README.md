# excel-fill-na

`excel-fill-na` is a small pip-installable library with an `fna` CLI for filling empty Excel cells inside a selected range.

It supports:

- filling empty cells with `NA` or custom text
- skipping one or more excluded ranges
- optionally merging contiguous empty cells vertically within a column before filling them
- preserving comment-only cells without filling or merging them
- working against the active worksheet or a named sheet

The implementation uses `openpyxl` to inspect workbook contents and compute fill operations, then patches only the target worksheet XML inside the workbook archive when saving. That keeps existing drawings, charts, images, and other unsupported workbook parts intact instead of round-tripping the whole file through `openpyxl.save()`. `.xlsx` and `.xlsm` are supported. `.xls` is not supported.

Pillow is not required to preserve existing workbook images. It is only relevant if you want to create new images through `openpyxl` itself. The committed test fixture under `tests/fixtures/` does not require Pillow at test runtime.

## Install

From PyPI:

```bash
pip install excel-fill-na
```

From a local checkout:

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
from excel_fill_na import process_workbook

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

Existing merged cells are supported. If an existing merged range has an eligible empty anchor cell, the anchor is filled, but the library does not attempt to re-merge or resize that existing merged range.

## Comment behavior

Only truly empty cells are modified. If a cell has no value but does have a comment, the library preserves that cell as-is:

- it is not filled with `NA` or custom text
- it is not included in a generated merge run
- it breaks a vertical empty run when `merge_empty_runs=True`

Cells that already contain a value are untouched regardless of whether they also have a comment.

## Specs

- Functional spec: [`docs/FUNCTIONAL_SPEC.md`](docs/FUNCTIONAL_SPEC.md)
- CLI spec: [`docs/CLI_SPEC.md`](docs/CLI_SPEC.md)

## Tests

```bash
.venv/bin/pytest
```
