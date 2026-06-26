## When to use
Use this skill when filtering or extracting rows from a spreadsheet based on conditional logic (equality, inequality, text matching, or exclusion criteria) and writing results to a specific location.

## Key principles

1. **Read source data completely**: Load all rows from the source range using `iter_rows()` with `values_only=True` to access cell values efficiently. Separate headers from data rows for proper processing.

2. **Implement conditional logic in Python**: Use standard Python operators (`==`, `!=`, `in`, `not in`, `startswith()`, `endswith()`) to evaluate conditions. Combine multiple conditions with `and`/`or` as specified in the instruction. Remember Python uses 0-based indexing for row tuples.

3. **Write to target location precisely**: Clear existing content in the destination range first if needed. Write headers and filtered rows to the exact starting position specified (e.g., "starting at A3" means row 3, column 1). Use `cell()` or direct assignment to write values.

4. **Handle text matching carefully**: Use `.strip()` to remove whitespace, `.lower()` or `.upper()` for case-insensitive comparisons when appropriate, and check for partial matches with `in` operator when the instruction requires finding text within cells.

## Common mistakes

- Forgetting that openpyxl uses 1-based indexing for cells but row tuples from `iter_rows(values_only=True)` are 0-based (row[0] is first column)
- Not preserving the original row order when filtering unless explicitly told to sort or reorder
- Overwriting headers when the instruction specifies they should be placed separately from data rows