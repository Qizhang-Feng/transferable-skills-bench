## When to use
Use this skill when you need to find and retrieve values from one part of a spreadsheet based on matching criteria from another part, similar to VLOOKUP, INDEX/MATCH, or cross-referencing operations.

## Key principles

1. **Use dictionary mapping for simple lookups**: Build a dictionary from the lookup table with keys as search values and values as return data. This handles one-to-one matches efficiently and allows quick retrieval.

2. **Iterate through target cells for complex matching**: When applying lookups to multiple rows, loop through the target range, extract the search value from each row, perform the lookup, and write the result back to the appropriate cell.

3. **Handle multiple criteria with composite keys**: For multi-column matching, create tuple keys combining all criteria values (e.g., `(col1_val, col2_val)`) in both the lookup dictionary and search logic.

4. **Implement fuzzy matching for partial text**: When searching for substrings within cells (like finding "Red" in "Red Jumper"), iterate through lookup values and use Python's `in` operator or string methods to check for partial matches.

5. **Provide default values for no-match cases**: Always handle scenarios where no match is found by setting a default value (empty string, "N/A", or keeping original value) to prevent errors.

## Common mistakes

- Building the lookup dictionary with values as keys instead of the search criteria column, causing failed matches
- Not converting data types (dates, numbers) to comparable formats before matching, leading to false negatives
- Forgetting to handle case sensitivity when matching text strings