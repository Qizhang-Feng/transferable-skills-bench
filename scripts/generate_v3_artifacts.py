#!/usr/bin/env python3
"""
Generate artifacts for v3 experiment:
1. gen_skill: one per family, LLM-generated from 5-10 task examples
2. local_patch: one per task, LLM-generated from a paired task's solution trace
"""

import json
import os
import sys
import random
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.path.insert(0, ".")
from scripts.perturbation_agent.llm_client import call_llm
from scripts.run_pilot_experiment import gen_spreadsheet_content

VERIFIED_DIR = "vendor/SpreadsheetBench/data/spreadsheetbench_verified_400"
V3_DIR = "data/perturbed_v3"
ARTIFACT_DIR = "data/v3_artifacts"
LLM_MODEL = "us.anthropic.claude-sonnet-4-5-20250929-v1:0"  # Sonnet for speed


def load_v3_tasks():
    """Load all pass tasks grouped by family."""
    with open(os.path.join(V3_DIR, "all_manifests.json")) as f:
        manifests = json.load(f)
    with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
        dataset = json.load(f)
    meta_map = {str(d["id"]): d for d in dataset}

    by_family = defaultdict(list)
    for tid, m in manifests.items():
        if m["status"] != "pass" or tid not in meta_map:
            continue
        task = meta_map[tid]
        task["_family"] = m.get("family", "unknown")
        task["_anchor_types"] = m.get("anchor_types_used", [])
        by_family[m.get("family", "unknown")].append(task)
    return by_family


def get_task_context(task, max_rows=5):
    """Get instruction + spreadsheet preview for a task."""
    tid = str(task["id"])
    init_path = os.path.join(VERIFIED_DIR, "spreadsheet", tid, f"1_{tid}_init.xlsx")
    preview = gen_spreadsheet_content(init_path, max_rows=max_rows) if os.path.exists(init_path) else ""
    return {
        "id": tid,
        "instruction": task["instruction"],
        "answer_position": task["answer_position"],
        "preview": preview[:1500],
    }


# ============================================================
# 1. Generate gen_skill per family
# ============================================================

GEN_SKILL_PROMPT = """You are creating a concise skill document for an LLM that solves spreadsheet tasks BY WRITING PYTHON CODE using openpyxl.

Below are {n} example tasks from the "{family}" workflow family.

## Example Tasks
{examples}

## Rules for the skill document:
1. Start with "## When to use" — one sentence
2. Write 3-5 numbered "## Key principles" — the most important things to get right
3. Include "## Common mistakes" — 2-3 specific mistakes to avoid
4. CRITICAL: The LLM will write PYTHON CODE using openpyxl. Do NOT mention Excel formulas, VBA, or Excel functions. Focus on Python/openpyxl patterns.
5. NEVER include specific column names, sheet names, cell ranges, or task IDs
6. Use parameterized references: "the column mentioned in the instruction"
7. Keep it SHORT: 150-250 words maximum. Every word must earn its place.
8. Focus on the 3-4 things that matter most, not a comprehensive guide

## Output
Write ONLY the skill document in markdown. No preamble."""


def generate_gen_skill(family, tasks):
    """Generate one gen_skill document for a family."""
    # Sample 5-8 tasks for examples
    sample = random.sample(tasks, min(8, len(tasks)))
    examples = ""
    for i, task in enumerate(sample):
        ctx = get_task_context(task)
        examples += f"\n### Example {i+1} (task {ctx['id']})\n"
        examples += f"Instruction: {ctx['instruction'][:300]}\n"
        examples += f"Answer position: {ctx['answer_position']}\n"
        examples += f"Spreadsheet preview:\n{ctx['preview'][:500]}\n"

    prompt = GEN_SKILL_PROMPT.format(
        n=len(sample), family=family, examples=examples)

    result = call_llm(prompt, model=LLM_MODEL, max_tokens=2000, temperature=0.3)
    return result.strip()


# ============================================================
# 2. Generate local_patch per task (round-robin pairing)
# ============================================================

LOCAL_PATCH_PROMPT = """You previously solved this exact spreadsheet task. Write a brief solution note for future reference.

## The task:
Instruction: {instruction}
Answer position: {answer_position}
Spreadsheet structure:
{preview}

## Write a solution note that:
1. Names the EXACT sheet names, column names, and cell ranges from this spreadsheet
2. Describes the Python/openpyxl approach: which columns to read, what logic to apply, where to write
3. Be SPECIFIC: say "read column 'Amount' from sheet 'Sales'" not "read the relevant column"
4. Mention the header row position and data start row
5. Keep it 150-250 words
6. DO NOT include actual answer values or computed results
7. Write as past tense: "The data was in sheet 'Sales', column 'Amount' (column D)..."

## Output
Write ONLY the note in markdown. No preamble."""


def generate_local_patch(task):
    """Generate a local_patch from the task's OWN context (self-referencing)."""
    ctx = get_task_context(task)
    prompt = LOCAL_PATCH_PROMPT.format(
        instruction=ctx["instruction"][:500],
        answer_position=ctx["answer_position"],
        preview=ctx["preview"],
    )
    result = call_llm(prompt, model=LLM_MODEL, max_tokens=1500, temperature=0.3)
    return result.strip()


def build_round_robin_pairing(tasks):
    """Pair each task with the next task in the list (circular)."""
    pairs = {}
    for i, task in enumerate(tasks):
        source = tasks[(i + 1) % len(tasks)]
        pairs[str(task["id"])] = str(source["id"])
    return pairs


# ============================================================
# Main
# ============================================================

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--gen_skill_only", action="store_true")
    parser.add_argument("--local_patch_only", action="store_true")
    parser.add_argument("--workers", type=int, default=32)
    parser.add_argument("--dry_run", action="store_true")
    args = parser.parse_args()

    by_family = load_v3_tasks()
    print(f"Loaded tasks by family:")
    for fam, tasks in sorted(by_family.items()):
        print(f"  {fam}: {len(tasks)}")

    os.makedirs(ARTIFACT_DIR, exist_ok=True)

    # --- gen_skill ---
    if not args.local_patch_only:
        print(f"\n{'='*60}")
        print("GENERATING GEN_SKILL (one per family)")
        print(f"{'='*60}")
        for family, tasks in sorted(by_family.items()):
            if args.dry_run:
                print(f"  Would generate gen_skill for {family} ({len(tasks)} tasks)")
                continue
            print(f"  Generating gen_skill for {family}...", end=" ", flush=True)
            skill = generate_gen_skill(family, tasks)
            out_dir = os.path.join(ARTIFACT_DIR, "gen_skill")
            os.makedirs(out_dir, exist_ok=True)
            with open(os.path.join(out_dir, f"{family}.md"), "w") as f:
                f.write(skill)
            print(f"OK ({len(skill)} chars)")

    # --- local_patch ---
    if not args.gen_skill_only:
        print(f"\n{'='*60}")
        print("GENERATING LOCAL_PATCH (from each task's OWN original version)")
        print(f"{'='*60}")

        # No pairing needed — each task generates its own patch
        all_tasks_flat = []
        for family, tasks in by_family.items():
            all_tasks_flat.extend(tasks)

        # Save self-pairing info
        self_pairings = {str(t["id"]): str(t["id"]) for t in all_tasks_flat}
        with open(os.path.join(ARTIFACT_DIR, "local_patch_pairings.json"), "w") as f:
            json.dump(self_pairings, f, indent=2)
        print(f"  Tasks: {len(all_tasks_flat)} (self-pairing)")

        if args.dry_run:
            print(f"  Would generate {len(all_tasks_flat)} local patches")
            return

        out_dir = os.path.join(ARTIFACT_DIR, "local_patch")
        os.makedirs(out_dir, exist_ok=True)

        task_map = {str(t["id"]): t for t in all_tasks_flat}

        # Generate in parallel — each task generates from itself
        def gen_one(tid):
            task = task_map[tid]
            patch = generate_local_patch(task)
            with open(os.path.join(out_dir, f"{tid}.md"), "w") as f:
                f.write(patch)
            return tid, len(patch)

        done = 0
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            futures = {pool.submit(gen_one, tid): tid for tid in task_map}
            for future in as_completed(futures):
                tid = futures[future]
                try:
                    _, size = future.result()
                    done += 1
                    if done % 20 == 0:
                        print(f"  [{done}/{len(task_map)}] generated", flush=True)
                except Exception as e:
                    print(f"  ERROR {tid}: {e}")

        print(f"  Done: {done}/{len(task_map)} local patches generated")

    print(f"\nArtifacts saved to {ARTIFACT_DIR}/")


if __name__ == "__main__":
    main()
