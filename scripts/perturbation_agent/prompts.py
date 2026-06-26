"""
All LLM prompts for the perturbation agent pipeline.
"""

FAMILY_CLASSIFY_PROMPT = """You are classifying a spreadsheet task into a workflow family and identifying perturbation risks.

## Task
Instruction: {instruction}

## Schema Summary
Sheets: {sheet_names}
{schema_summary}

## Families
- formula_writeback: Read columns, compute via formula/logic, write back results
- lookup_match_merge: Match keys across tables/sheets, combine/transfer data
- conditional_partition: Filter/split rows based on conditions
- grouped_aggregation: Group by key, aggregate (sum/count/avg), write summary
- structural_edit: Delete/insert/sort/transpose rows/columns

## Output JSON
{{
  "family": "one of the 5 families above",
  "confidence": "high|medium|low",
  "risk_flags": ["list of applicable flags from: instruction_references_absolute_rows, task_deletes_or_inserts_rows, task_depends_on_blank_row_positions, task_creates_new_sheet, header_not_in_row_1, task_references_specific_cell_addresses, multi_sheet_references, data_values_match_headers"],
  "reasoning": "brief explanation"
}}"""


DESIGN_PROMPT = """You are designing a structural perturbation for a spreadsheet task. The perturbation must change surface anchors WITHOUT changing the task's underlying logic or difficulty.

## Family Guide
{family_guide}

## Task
Instruction: {instruction}
Answer Position: {answer_position}

## Schema
{schema_json}

## Constraints
Safe anchor types for this task: {safe_anchors}
Risk flags: {risk_flags}

{error_section}

## Rules
1. Choose 2-3 anchor types from the safe list
2. Rename: same descriptiveness level (abbreviation to abbreviation, full word to full word)
3. If a header name also appears as data values, include data_value_renames with sheet and range
4. Layout: respect risk_flags. NO title_row_insert if task deletes/inserts rows
5. Distractor column name should be plausibly similar but clearly different from real columns
6. Do NOT rename a column if its name (or a close variant) appears as a common English word in the instruction body. Example: do NOT rename "Performance" if the instruction says "lowest performing students". Only rename columns whose names are used purely as column references, not as general English.
7. Set instruction_paraphrase to false. Paraphrasing adds risk of changing meaning. Only rename and layout.
8. For data_value_renames, use the NEW sheet name (after rename) in the "sheet" field

## Output JSON
{{
  "anchor_types_used": ["lexical", ...],
  "rename": {{
    "column_rename_map": {{"old_name": "new_name"}},
    "sheet_rename_map": {{"old_name": "new_name"}},
    "data_value_renames": [{{"sheet": "SheetName", "range": "A2:A100", "find": "old", "replace": "new"}}]
  }},
  "layout": {{
    "type": "title_row_insert|blank_row_insert|distractor_column|skip",
    "target_sheets": ["Sheet1"],
    "params": {{}}
  }},
  "distractor": {{
    "type": "column|none",
    "name": "column name if applicable",
    "target_sheet": "Sheet1",
    "position": "after_last_column"
  }},
  "instruction_paraphrase": false,
  "reasoning": "why these choices"
}}"""


RENAME_SAFETY_PROMPT = """Check if these renames are safe and neutral for a benchmark perturbation.

## Rename Map
Columns: {column_rename_map}
Sheets: {sheet_rename_map}
Data values: {data_value_renames}

## Original Instruction (first 500 chars)
{instruction}

## All Headers in Workbook
{all_headers}

## Checks — ONLY reject (approved=false) for these hard failures:
1. Duplicate headers after rename within the same sheet
2. Sheet name conflicts (two sheets with the same name after rename)
3. New name contains illegal Excel sheet characters (\\/*?:[])
4. New sheet name exceeds 31 characters

The following are WARNINGS, not rejections (still set approved=true):
- Descriptiveness level mismatch (abbreviation vs full word)
- Potential collision with English words in instruction
- Missing data-value renames for header/data overlaps

## Output JSON
{{
  "approved": true or false,
  "issues": [{{"problem": "description", "fix": "suggestion"}}]
}}"""


LAYOUT_IMPACT_PROMPT = """Check if this layout perturbation changes the task difficulty or breaks the task logic.

## Rewritten Instruction (already updated with renames and row shifts)
{instruction}

## Layout Applied
Type: {layout_type}
Target sheets: {target_sheets}

## Perturbed File Preview (first 6 rows per sheet)
{perturbed_preview}

## Risk Flags
{risk_flags}

## Context
The instruction above has ALREADY been rewritten to account for renames and row shifts.
Cell references and row numbers in the instruction should already reflect the perturbed file.

## Checks — ONLY flag as P0 (unsafe) if:
1. The instruction references specific row numbers or cell addresses that do NOT match the perturbed file (rewrite failed to update them)
2. The task logic explicitly depends on being the last column or last row, and that changed
3. An inserted row/column would be processed as data by the task (e.g., blank row treated as delimiter, distractor column included in a range)

The following are NOT P0 issues:
- Row numbers that were correctly updated in the instruction (e.g., "row 4" matching data now in row 4)
- Distractor column added after the last data column (this is always safe)
- Sheet names changed (handled by rename)

## Output JSON
{{
  "safe": true or false,
  "issues": [{{"problem": "description", "severity": "P0 or P1"}}],
  "recommendation": "proceed|skip_layout|use_alternative"
}}"""


REWRITE_INSTRUCTION_PROMPT = """Rewrite this spreadsheet task instruction to account for structural perturbations. The rewritten instruction must be semantically equivalent: a human should solve the task with EXACTLY the same logic.

## Perturbations Applied
Column renames: {column_rename_map}
Sheet renames: {sheet_rename_map}
Row shift: {row_shift_description}
Distractor: {distractor_description}
Paraphrase: {paraphrase}

## Original Instruction
{instruction}

## Rules
1. Replace ALL old column/sheet names with new names
2. Update ALL row numbers if rows were shifted (A1 becomes A2, row 1 becomes row 2, etc.)
3. Keep natural and fluent English. Fix grammar (a/an, capitalization)
4. Do NOT change task logic, constraints, or expected behavior
5. Do NOT add or remove ANY information. Do NOT add column names, sheet names, or descriptions not in the original
6. Common English words (name, date, type, rate, data, group, sum, total, performance) that are NOT column/sheet references must NOT be replaced
7. Preserve the EXACT same level of detail. If the original is vague, keep it vague. If truncated, keep the same truncation
8. NEVER mention perturbation implementation details: do NOT say rows were shifted, do NOT mention distractor columns, do NOT say names were renamed, do NOT add notes about structural changes. The rewritten instruction must read as if it was always written this way.
{paraphrase_instruction}

## Output JSON
{{
  "replacements_applied": [
    {{"old": "exact old text", "new": "exact new text", "reason": "column_rename|sheet_rename|row_shift|cell_ref_shift"}}
  ],
  "final_instruction": "the complete rewritten instruction text"
}}"""


INSTRUCTION_QUALITY_PROMPT = """Compare original and perturbed instructions for a spreadsheet task.

## Original
{original_instruction}

## Perturbed
{perturbed_instruction}

## Rename Map
Columns: {column_rename_map}
Sheets: {sheet_rename_map}

## Checks
1. Natural language: grammar, articles (a/an), capitalization, singular/plural
2. Completeness: all old names replaced, no leftovers
3. Solvability: same logic to solve, no new ambiguity, no lost constraints
4. No extra information added or removed

## Output JSON
{{
  "pass": true or false,
  "issues": [{{"type": "naturalness|completeness|solvability", "evidence": "the problematic text", "fix": "suggested fix"}}]
}}"""


FINAL_SANITY_PROMPT = """Final review of a perturbed spreadsheet task. You are the last checkpoint before this goes into the benchmark.

## Perturbed Instruction
{perturbed_instruction}

## Perturbed File Preview
{perturbed_preview}

## Answer Position
Original: {original_ap}
Perturbed: {perturbed_ap}

## Code Validation Results
{code_validation_summary}

## Perturbation Applied
Anchor types: {anchor_types}
Renames: {rename_summary}
Layout: {layout_summary}

## Question
If you were an LLM seeing this task for the first time, could you solve it with the same approach as the original? Is anything confusing, inconsistent, or broken?

## Output JSON
{{
  "verdict": "pass|warn|fail",
  "issues": [{{"severity": "P0|P1|P2", "description": "what is wrong"}}],
  "confidence": "high|medium|low"
}}"""
