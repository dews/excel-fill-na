# Functional Spec

## Goal

Provide a reusable Python library that can:

- open an Excel workbook
- inspect a user-selected rectangular range
- replace empty cells with a default value of `NA` or a user-provided string
- optionally delete rows whose selected-range cells are all empty
- optionally merge contiguous empty cells vertically within a column before filling them
- save the processed workbook to a new file while preserving non-target workbook artifacts

## Definitions

- Empty cell: a cell whose value is `None` or a string containing only whitespace
- Comment-only empty cell: an empty cell that also has an attached cell comment
- Eligible empty cell: an empty cell that is not excluded and does not have a comment
- Deletable row: a worksheet row inside the target range whose selected-range cells are all logically empty and that does not intersect an excluded range
- Target range: the rectangular range selected by the caller, for example `A1:C20`
- Excluded range: one or more rectangular ranges that should not be modified, even if they are inside the target range
- Empty run: two or more contiguous eligible empty cells in the same column inside the target range

## Processing Rules

1. The library loads the workbook from disk and selects either the requested worksheet or the active sheet.
2. Only cells inside the target range are considered.
3. Any cell inside an excluded range is skipped.
4. Comment-only empty cells are preserved as intentional content:
   - they are not filled
   - they are not merged
   - they terminate an empty run during merge planning
5. If merge mode is disabled, each eligible empty cell is filled individually.
6. If merge mode is enabled, each column is scanned top to bottom:
   - an empty run of length 1 is filled normally
   - an empty run of length 2 or more is merged vertically and filled via the top cell
7. Existing merged ranges are respected:
   - if the anchor cell of an existing merged range is inside the target range and is an eligible empty cell, it is filled
   - the library does not expand, shrink, or replace an existing merged range
8. When saving to disk, only the selected worksheet XML is rewritten. Other workbook archive members are copied unchanged so existing drawings, charts, images, metadata, comments, and VBA payloads are preserved.
9. Cells backed by SpreadsheetML value metadata, such as Excel in-cell picture anchors, are preserved during the save-to-disk flow even when `openpyxl` surfaces them as empty or as `#VALUE!`.

## Delete-Row Rules

1. Delete mode is opt-in via `delete_empty_rows=True` in the library or `--delete-empty-rows` in the CLI.
2. Candidate rows come only from the row numbers covered by the target range.
3. A candidate row is deleted when every cell position inside that row slice of the target range is logically empty.
4. Logical emptiness for delete mode follows these rules:
   - `None` and whitespace-only strings are empty
   - cells with comments are non-empty
   - merged cells count only for the row that owns the merged anchor cell; covered cells whose anchor is on another row are ignored
   - value-metadata-backed cells are preserved as non-empty
5. Any candidate row intersecting an excluded range is protected and is not deleted.
6. Deleting a row removes the entire worksheet row, including cells outside the selected range on that row.
7. During save-to-disk delete mode, the archive patching flow also shifts directly related sheet-owned XML parts that carry row-based anchors or refs, including comments, threaded comments, legacy VML note anchors, drawings, hyperlinks, and merged ranges for the target sheet.

## Output Rules

- If no output path is supplied, the default output path is `<stem>.filled<suffix>`
- The function returns a summary containing:
  - sheet name
  - target range
  - fill value
  - number of logical empty cells processed
  - any merged ranges created by the library
  - number of deleted rows
  - output path when saving to disk

## Non-Goals

- formula evaluation
- `.xls` support
- preserving user intent for partially selected existing merged ranges where the anchor lies outside the target range
- horizontal merge generation
