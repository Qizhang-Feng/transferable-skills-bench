## When to use
Use this skill when you need to group rows by one or more columns and calculate aggregated values (sum, count, etc.) for each group, then output the consolidated results.

## Key principles

1. **Use dictionaries for grouping**: Create a dictionary where keys are tuples of the grouping column values and values store the aggregated data. This efficiently handles duplicate detection and aggregation in one pass.

2. **Read all data first, then aggregate**: Load the relevant columns into memory, iterate through rows to build your aggregation dictionary, then write results back. Don't try to aggregate while reading.

3. **Handle the first occurrence strategy**: When consolidating duplicates, decide whether to keep the first row's non-aggregated values, create a new row, or merge data. Store complete row data in your dictionary value if you need to preserve non-grouped columns.

4. **Count occurrences with counters**: For counting tasks (like occurrence order), use a separate counter dictionary or the `collections.Counter` class to track how many times each value has appeared.

5. **Clear and rewrite output ranges**: When replacing data with aggregated results, delete old rows (except headers) and write the consolidated data cleanly to avoid leftover duplicate rows.

## Common mistakes

- Trying to aggregate by modifying rows in place while iterating—this causes skipped rows and incorrect results. Always build aggregated data separately first.
- Forgetting to convert grouping values to a hashable type (like tuples) when grouping by multiple columns—lists can't be dictionary keys.
- Not preserving the header row when clearing data ranges for output.