## When to use
Use this skill when writing Python code with openpyxl to populate cells with formulas or computed values derived from date parsing, text manipulation, or conditional logic based on existing spreadsheet data.

## Key principles

1. **Write values, not formulas**: When the instruction asks for date formatting, text extraction, or calculations, compute the result in Python and write the final value to the cell. Don't write Excel formula strings unless explicitly required.

2. **Parse dates carefully**: Use `datetime.strptime()` or `dateutil.parser` to convert date strings or Excel serial numbers to Python datetime objects. Handle multiple date formats (e.g., 'YYYYMMDD', 'YYYY-MM-DD', Excel datetime objects) by checking the data type first.

3. **Handle text manipulation in Python**: For splitting, extracting, or reformatting text (names, dates embedded in strings), use Python string methods (`.split()`, `.strip()`, slicing, regex) rather than attempting Excel TEXT functions.

4. **Iterate row-by-row for conditional logic**: When output depends on matching conditions (date ranges, keyword searches, holiday lookups), loop through source rows, evaluate conditions in Python, and write results to the target range.

5. **Preserve data types**: When writing dates back to cells, use Python datetime objects so openpyxl formats them correctly. For text that looks like numbers, ensure you're writing strings if needed.

## Common mistakes

- Writing Excel formula syntax (e.g., `"=TEXT(A1, 'MMM')"`) instead of computing values in Python
- Failing to handle mixed data types in date columns (strings vs. datetime objects)
- Not accounting for empty cells or None values when parsing dates or text, causing exceptions