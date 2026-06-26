#!/usr/bin/env python3
"""
Automated review of expansion perturbation specs.
Sends each task to an LLM for structured review.
Outputs: scratch/expansion-review-results.json
"""

import json
import os
import sys
import time

VERIFIED_DIR = "vendor/SpreadsheetBench/data/spreadsheetbench_verified_400"

REVIEW_PROMPT_TEMPLATE = """You are a semantic quality reviewer for a benchmark perturbation. A deterministic code validator has already checked hard rules (namespace validity, answer range mapping, formula baking, rewrite collisions). Your job is to check the things code cannot: semantic equivalence, naturalness, difficulty neutrality, and solvability.

## Context

We perturb spreadsheet tasks by (1) renaming column/sheet names to semantic equivalents and rewriting the instruction, and (2) inserting a title row at row 1 (+1 row shift). The goal: change only surface anchors, NOT task logic or difficulty.

## Your 4 review groups

### Group 1: Rename validity
- **Semantic equivalence**: Are new names semantically equivalent? (e.g. Amount->Total_Amount OK; Amount->Date NOT OK)
- **Descriptiveness neutrality**: Are new names at the SAME clarity level as originals? If old name was abbreviated (QTY), new name should also be abbreviated (AMT), not spelled out (QUANTITY). A more descriptive name makes the task easier for an LLM, which is a confound.
- **Scope correctness**: Did the rename only change what it should (headers, sheet tabs), not unrelated cell values or natural language?

### Group 2: Instruction rewrite validity
- **Naturalness**: Does the perturbed instruction read like natural English? Check grammar, articles (a/an), capitalization, singular/plural.
- **Completeness**: Are there leftover old names that should have been updated?
- **Solvability equivalence**: Can a human solve the perturbed task with EXACTLY the same logic as the original? No new ambiguity, no lost constraints?

### Group 3: Layout / answer mapping validity
- **Layout shift neutrality**: Does the instruction reference "row 1", "first row", "top row", "A1", or "starting at row 1"? If so, inserting a title row may change the task difficulty (not just surface anchors).
- **Answer position**: Is the perturbed answer position consistent with the instruction and schema?

### Group 4: Difficulty neutrality (beyond semantics)
- Could any rename make the LLM's job easier or harder? (e.g. Data->Records sounds more like a database table; Summary->Overview changes output expectations)
- Could the title row insertion change how the LLM interprets the spreadsheet structure?

## Code validator already checked (do NOT re-check these):
- Sheet/column name legality and uniqueness
- Answer range row +1 correctness and shape preservation
- Formula baking coverage
- Rewrite rule collision with common English words
- Old name leftovers in instruction (exact match)

## Code validator P1 findings for this task:
{code_validator_findings}

---

## Task details

**Task ID:** {task_id}
**Family:** {family}
**Instruction type:** {instruction_type}

### Original instruction
{original_instruction}

### Rename map
Columns: {column_rename_map}
Sheets: {sheet_rename_map}

### Instruction rewrite rules
{rewrite_rules}

### Perturbed instruction
{perturbed_instruction}

### Answer position
- Original: {original_ap}
- Perturbed: {perturbed_ap}

### Spreadsheet schema
{schema_info}

---

## Output format

Respond ONLY with a JSON object:
```json
{{
  "task_id": "{task_id}",
  "verdict": "PASS or WARN or FAIL",
  "confidence": "high or medium or low",
  "issues": [
    {{
      "severity": "P0 or P1 or P2",
      "criterion": "rename_equivalence|descriptiveness_neutrality|rename_scope|instruction_naturalness|instruction_completeness|solvability_equivalence|layout_neutrality|answer_position|difficulty_neutrality",
      "evidence": "the specific text or name causing the issue",
      "description": "what is wrong",
      "suggested_fix": "how to fix it"
    }}
  ],
  "final_recommendation": "ready_for_experiment or revise_before_experiment or exclude_task"
}}
```

Verdict rules:
- PASS: No P0 or P1 issues. Ready for experiment.
- WARN: P1 issues only (neutrality/naturalness concerns, but no correctness bugs). Usable with caveats.
- FAIL: P0 issues (semantic change, solvability broken, critical ambiguity). Must fix before experiment.
"""


def build_review_prompt(task_id, spec, meta, schema, perturbed_instruction, perturbed_ap, code_findings=None):
    """Build the review prompt for one task."""
    schema_lines = []
    for sh in schema.get("sheets", []):
        headers = [h["value"] for h in sh.get("headers", [])]
        schema_lines.append(
            f"  Sheet '{sh['name']}': {sh['rows']} rows, {sh['cols']} cols, "
            f"header_row={sh.get('header_row', '?')}, "
            f"formulas={sh.get('formula_count', 0)}, "
            f"headers={headers[:8]}"
        )
    schema_info = "\n".join(schema_lines) if schema_lines else "(no schema info)"

    rules = spec["rename"].get("instruction_rewrite_rules", [])
    rules_lines = []
    for r in rules:
        ci = " (case-insensitive)" if r.get("case_insensitive") else ""
        rules_lines.append(f'  "{r["find"]}" -> "{r["replace"]}"{ci}')
    rewrite_rules = "\n".join(rules_lines) if rules_lines else "(no rewrite rules)"

    # Format code validator findings
    if code_findings:
        findings_lines = []
        for f in code_findings:
            findings_lines.append(f"- [{f['severity']}] {f['criterion']}: {f['description']}")
        code_validator_findings = "\n".join(findings_lines)
    else:
        code_validator_findings = "(no issues found by code validator)"

    return REVIEW_PROMPT_TEMPLATE.format(
        task_id=task_id,
        family=spec.get("family", "?"),
        instruction_type=meta.get("instruction_type", "?"),
        original_instruction=meta.get("instruction", ""),
        column_rename_map=json.dumps(spec["rename"].get("column_rename_map", {})),
        sheet_rename_map=json.dumps(spec["rename"].get("sheet_rename_map", {})),
        rewrite_rules=rewrite_rules,
        perturbed_instruction=perturbed_instruction,
        original_ap=meta.get("answer_position", ""),
        perturbed_ap=perturbed_ap,
        code_validator_findings=code_validator_findings,
        schema_info=schema_info,
    )


def call_reviewer(prompt, model="us.anthropic.claude-sonnet-4-5-20250929-v1:0"):
    """Call LLM to review a task. Returns parsed JSON response."""
    try:
        sys.path.insert(0, "vendor/PowerComputeCustomImage/configuration/rewards/Inference/APIInference")
        from api_invocation import ClaudeInference

        if not hasattr(call_reviewer, '_claude'):
            call_reviewer._claude = ClaudeInference()

        prompt_dict = {
            "system": "You are a quality reviewer. Respond only with valid JSON.",
            "user": prompt,
        }
        text = call_reviewer._claude.inference(
            prompt_dict=prompt_dict,
            retry=3,
            guardrails=[],
            model_id=model,
            max_tokens=2000,
            temperature=0,
        )

        if "```json" in text:
            json_str = text.split("```json")[1].split("```")[0].strip()
        elif "```" in text:
            json_str = text.split("```")[1].split("```")[0].strip()
        else:
            json_str = text.strip()

        return json.loads(json_str)

    except Exception as e:
        return {
            "task_id": "?",
            "verdict": "ERROR",
            "issues": [{"criterion": "system", "severity": "critical",
                        "description": str(e)[:200]}],
            "suggested_fixes": [],
            "notes": f"LLM call failed: {e}",
        }


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--task_id", default=None, help="Review single task")
    parser.add_argument("--dry_run", action="store_true", help="Print prompts only")
    parser.add_argument("--model",
                        default="us.anthropic.claude-opus-4-6-v1")
    args = parser.parse_args()

    with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
        dataset = json.load(f)
    meta_map = {str(d["id"]): d for d in dataset}

    with open("scratch/expansion-perturbation-specs.json") as f:
        specs = json.load(f)

    # Filter out excluded tasks
    specs = {tid: s for tid, s in specs.items()
             if s.get("review_status") != "excluded"}

    with open("scratch/expansion-task-schemas.json") as f:
        schemas_list = json.load(f)
    schema_map = {s["task_id"]: s for s in schemas_list}

    # Load code validator results
    code_results = {}
    code_results_path = "scratch/expansion-code-validation-results.json"
    if os.path.exists(code_results_path):
        with open(code_results_path) as f:
            for r in json.load(f):
                code_results[r["task_id"]] = r.get("issues", [])

    tasks_to_review = sorted(specs.keys())
    if args.task_id:
        tasks_to_review = [args.task_id]

    results = []
    for i, tid in enumerate(tasks_to_review, 1):
        spec = specs[tid]
        meta = meta_map.get(tid, {})
        schema = schema_map.get(tid, {})

        pert_inst_path = f"scratch/expansion_perturbed/{tid}/instruction_perturbed.txt"
        pert_inst = ""
        if os.path.exists(pert_inst_path):
            with open(pert_inst_path) as f:
                pert_inst = f.read()

        pert_meta_path = f"scratch/expansion_perturbed/{tid}/task_meta.json"
        pert_ap = ""
        if os.path.exists(pert_meta_path):
            with open(pert_meta_path) as f:
                pert_ap = json.load(f).get("answer_position", "")

        prompt = build_review_prompt(tid, spec, meta, schema, pert_inst, pert_ap,
                                    code_findings=code_results.get(tid, []))

        if args.dry_run:
            print(f"\n{'='*60}")
            print(f"[{i}/{len(tasks_to_review)}] Task {tid}")
            print(f"{'='*60}")
            print(prompt[:500] + "...")
            continue

        print(f"[{i}/{len(tasks_to_review)}] Reviewing {tid}...", end=" ",
              flush=True)
        result = call_reviewer(prompt, model=args.model)
        result["task_id"] = tid
        results.append(result)

        verdict = result.get("verdict", "?")
        n_issues = len(result.get("issues", []))
        symbol = {"PASS": "V", "WARN": "W", "FAIL": "X"}.get(verdict, "?")
        print(f"{symbol} {verdict} ({n_issues} issues)")

        if result.get("issues"):
            for issue in result["issues"]:
                sev = issue.get("severity", "?")
                crit = issue.get("criterion", "?")
                desc = issue.get("description", "")[:100]
                print(f"    [{sev}] {crit}: {desc}")

        time.sleep(0.5)

    if not args.dry_run:
        output_path = "scratch/expansion-review-results.json"
        with open(output_path, "w") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"\n{'='*60}")
        print("REVIEW SUMMARY")
        print(f"{'='*60}")
        verdicts = [r["verdict"] for r in results]
        for v in ["PASS", "WARN", "FAIL", "ERROR"]:
            print(f"  {v}: {verdicts.count(v)}")
        print(f"\nResults saved to {output_path}")


if __name__ == "__main__":
    main()
