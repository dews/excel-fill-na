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
- `--delete-empty-rows`: delete rows whose selected-range cells are all empty
- `--output PATH`: output workbook path; defaults to `<input>.filled<suffix>`

## Behavior

1. Load the workbook from `INPUT`.
2. Resolve the worksheet using `--sheet` or the active sheet.
3. Inspect the requested target range.
4. Skip any cells contained by excluded ranges.
5. Preserve comment-only empty cells. They are not filled and are not included in merge runs.
6. Fill the remaining eligible empty cells with `--fill-text`.
7. If `--merge-empty-runs` is enabled, merge each contiguous vertical eligible empty run before filling it.
8. If `--delete-empty-rows` is enabled, treat the command as delete mode instead of fill mode:
   - candidate rows come only from the selected `--range`
   - a row is deleted when every selected-range cell on that row is logically empty
   - cells with comments count as non-empty
   - merged cells only count for the row that owns the merged anchor; covered cells anchored on another row are ignored
   - any row intersecting `--exclude-range` is protected from deletion
9. Save the updated workbook to the output path by patching the target worksheet XML and any directly related sheet-owned XML parts that must shift with deleted rows.
10. Print a short processing summary to stdout.

## Mode Constraints

- `--delete-empty-rows` cannot be combined with `--merge-empty-runs`
- `--delete-empty-rows` cannot be combined with an explicit `--fill-text`

## Exit Semantics

- Exit `0` on success
- Exit `2` for argument or runtime validation failures surfaced through the CLI
