# CLI Spec

## Command

```bash
fna INPUT --range TARGET_RANGE [options]
```

## Required Arguments

- `INPUT`: path to an `.xlsx` or `.xlsm` workbook
- `--range`: target cell range to process

## Optional Arguments

- `--sheet SHEET_NAME`: worksheet name, defaults to the active sheet
- `--exclude-range RANGE`: excluded range; may be repeated and may include comma-separated values
- `--fill-text TEXT`: replacement text, defaults to `NA`
- `--merge-empty-runs`: merge contiguous empty cells vertically within each column before filling
- `--output PATH`: output workbook path; defaults to `<input>.filled<suffix>`

## Behavior

1. Load the workbook from `INPUT`.
2. Resolve the worksheet using `--sheet` or the active sheet.
3. Inspect the requested target range.
4. Skip any cells contained by excluded ranges.
5. Preserve comment-only empty cells. They are not filled and are not included in merge runs.
6. Fill the remaining eligible empty cells with `--fill-text`.
7. If `--merge-empty-runs` is enabled, merge each contiguous vertical eligible empty run before filling it.
8. Save the updated workbook to the output path by patching only the target worksheet XML and copying the rest of the workbook archive unchanged.
9. Print a short processing summary to stdout.

## Exit Semantics

- Exit `0` on success
- Exit `2` for argument or runtime validation failures surfaced through the CLI
