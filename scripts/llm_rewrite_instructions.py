#!/usr/bin/env python3
"""
Use LLM to rewrite perturbed instructions instead of brittle find-replace.
For each task: give LLM the original instruction, rename map, and layout shift info,
ask it to produce a natural, complete, semantically equivalent perturbed instruction.
"""

import json
import os
import sys
import time

VERIFIED_DIR = "vendor/SpreadsheetBench/data/spreadsheetbench_verified_400"

REWRITE_PROMPT = """You are rewriting a spreadsheet task instruction to account for two structural perturbations. Your rewrite must be semantically equivalent — a human should solve the perturbed task with EXACTLY the same logic as the original.

## Perturbations applied

### 1. Rename
The following names have been changed in the spreadsheet file:

Column header renames: {column_rename_map}
Sheet name renames: {sheet_rename_map}

### 2. Layout shift
A title row ("Report") has been inserted at row 1 of every sheet. All data has shifted down by 1 row. This means:
- What was row 1 is now row 2
- What was row 2 is now row 3
- Cell A1 is now A2, B2 is now B3, etc.
- Any row number referenced in the instruction must be incremented by 1

## Original instruction
{original_instruction}

## Rules
1. Replace ALL occurrences of old column/sheet names with their new names.
2. Increment ALL row numbers by 1 (cell references like A1->A2, B2->B3, row 1->row 2, etc.)
3. Keep the instruction natural and fluent. Do NOT produce awkward phrases like "yearly Score" or "on a Overview".
4. Do NOT change the task logic, constraints, or expected behavior.
5. Do NOT add or remove information.
6. Preserve the original tone and style.
7. If a word like "name", "date", "type", "rate", "data", "group", "sum", "total" appears as a common English word (not as a column/sheet reference), do NOT replace it.
8. If a renamed column/sheet name would read unnaturally in the sentence, adjust the surrounding grammar (e.g., "a Overview" -> "an Overview", "yearly Score" -> "yearly score").
9. Keep the same level of descriptiveness — if the original used an abbreviation, use a similar abbreviation in the replacement.

## Output
Return ONLY the rewritten instruction text. No explanation, no markdown, no quotes around it. Just the instruction."""


def rewrite_instruction(original_instruction, spec, model="us.anthropic.claude-opus-4-6-v1"):
    """Call LLM to rewrite instruction."""
    sys.path.insert(0, "vendor/PowerComputeCustomImage/configuration/rewards/Inference/APIInference")
    from api_invocation import ClaudeInference

    if not hasattr(rewrite_instruction, '_claude'):
        rewrite_instruction._claude = ClaudeInference()

    col_map = spec["rename"].get("column_rename_map", {})
    sheet_map = spec["rename"].get("sheet_rename_map", {})

    prompt = REWRITE_PROMPT.format(
        column_rename_map=json.dumps(col_map, indent=2) if col_map else "(none)",
        sheet_rename_map=json.dumps(sheet_map, indent=2) if sheet_map else "(none)",
        original_instruction=original_instruction,
    )

    prompt_dict = {
        "system": "You are a precise text rewriter. Follow the rules exactly.",
        "user": prompt,
    }

    text = rewrite_instruction._claude.inference(
        prompt_dict=prompt_dict,
        retry=3,
        guardrails=[],
        model_id=model,
        max_tokens=4096,
        temperature=0,
    )

    return text.strip()


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--task_id", default=None)
    parser.add_argument("--dry_run", action="store_true")
    parser.add_argument("--model", default="us.anthropic.claude-opus-4-6-v1")
    args = parser.parse_args()

    with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
        dataset = json.load(f)
    meta_map = {str(d["id"]): d for d in dataset}

    with open("scratch/expansion-perturbation-specs.json") as f:
        specs = json.load(f)

    specs = {tid: s for tid, s in specs.items()
             if s.get("review_status") != "excluded"}

    tasks = sorted(specs.keys())
    if args.task_id:
        tasks = [args.task_id]

    results = {}
    for i, tid in enumerate(tasks, 1):
        spec = specs[tid]
        meta = meta_map.get(tid, {})
        original_instruction = meta.get("instruction", "")

        if args.dry_run:
            print(f"[{i}/{len(tasks)}] {tid} (dry run)")
            continue

        print(f"[{i}/{len(tasks)}] Rewriting {tid}...", end=" ", flush=True)

        try:
            rewritten = rewrite_instruction(original_instruction, spec, model=args.model)

            # Save to perturbed directory
            out_dir = os.path.join("scratch/expansion_perturbed", str(tid))
            os.makedirs(out_dir, exist_ok=True)

            inst_path = os.path.join(out_dir, "instruction_perturbed.txt")
            with open(inst_path, "w") as f:
                f.write(rewritten)

            # Also update task_meta.json
            meta_path = os.path.join(out_dir, "task_meta.json")
            if os.path.exists(meta_path):
                with open(meta_path) as f:
                    task_meta = json.load(f)
                task_meta["instruction"] = rewritten
                with open(meta_path, "w") as f:
                    json.dump(task_meta, f, indent=2, ensure_ascii=False)

            # Quick validation: check old names don't appear
            col_map = spec["rename"].get("column_rename_map", {})
            sheet_map = spec["rename"].get("sheet_rename_map", {})
            leftovers = []
            for old in list(col_map.keys()) + list(sheet_map.keys()):
                if len(old) > 3 and old in rewritten:
                    new = col_map.get(old, sheet_map.get(old, ""))
                    if old not in new:
                        leftovers.append(old)

            status = "OK" if not leftovers else f"WARN: leftovers {leftovers}"
            print(f"{status} ({len(rewritten)} chars)")

            results[tid] = {
                "status": "ok" if not leftovers else "warn",
                "leftovers": leftovers,
                "length_original": len(original_instruction),
                "length_rewritten": len(rewritten),
            }

        except Exception as e:
            print(f"ERROR: {e}")
            results[tid] = {"status": "error", "error": str(e)[:200]}

        time.sleep(0.5)

    if not args.dry_run:
        with open("scratch/expansion-llm-rewrite-results.json", "w") as f:
            json.dump(results, f, indent=2)

        ok = sum(1 for r in results.values() if r["status"] == "ok")
        warn = sum(1 for r in results.values() if r["status"] == "warn")
        err = sum(1 for r in results.values() if r["status"] == "error")
        print(f"\nSUMMARY: OK={ok} WARN={warn} ERROR={err}")


if __name__ == "__main__":
    main()
