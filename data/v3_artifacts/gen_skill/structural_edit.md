## When to use
Use this skill when you need to delete rows, remove blank cells, or restructure spreadsheet data based on specific conditions using Python and openpyxl.

## Key principles

1. **Always iterate backwards when deleting rows** — use `for row_idx in range(max_row, min_row - 1, -1)` to avoid index shifting issues that cause rows to be skipped after deletion.

2. **Load the entire dataset first, then process** — read all relevant cell values into memory before making structural changes. Use `ws.iter_rows()` or direct cell access to collect data, then perform deletions in a separate pass.

3. **Identify block boundaries before sorting or processing** — when working with separated data blocks, scan the sheet once to find blank row positions, then process each block independently using `ws.delete_rows()` and data manipulation.

4. **Match conditions across sheets carefully** — when deleting based on cross-sheet comparisons, build lookup sets or dictionaries from one sheet, then check against them while iterating backwards through the target sheet.

5. **Handle empty cells explicitly** — use `cell.value is None or cell.value == ''` to detect blanks. Don't assume empty cells will evaluate to False in all contexts.

## Common mistakes

- Iterating forward while deleting rows causes skipped rows due to index shifting
- Modifying the sheet structure while iterating through `iter_rows()` leads to unpredictable behavior
- Forgetting to save the workbook with `wb.save()` after making structural changes