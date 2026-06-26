#!/usr/bin/env python3
"""
Positive Control v2: Use actual successful code from original as hard_patch.
Only runs on the 31 tasks where no_patch_original passed.
"""

import json
import os
import sys
import shutil
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.path.insert(0, ".")
from scripts.run_pilot_experiment import (
    call_llm, extract_code, execute_code, evaluate_task, gen_spreadsheet_content,
    VERIFIED_DIR,
)

OUTPUT_DIR = "data/positive_control"
RESULTS_DIR = "scratch/experiment_results"
LLM_MODEL = "us.anthropic.claude-sonnet-4-5-20250929-v1:0"


def run_one(task_id, condition, patch_type, patch_content, manifest):
    """Run a single experiment."""
    if condition == "original":
        spread_dir = os.path.join(VERIFIED_DIR, "spreadsheet", str(task_id))
        with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
            dataset = json.load(f)
        meta = next(d for d in dataset if str(d["id"]) == str(task_id))
    else:
        spread_dir = os.path.join(OUTPUT_DIR, "perturbed", str(task_id))
        with open(os.path.join(spread_dir, "task_meta.json")) as f:
            meta = json.load(f)

    input_file = f"1_{task_id}_init.xlsx"
    golden_file = f"1_{task_id}_golden.xlsx"
    input_path = os.path.join(spread_dir, input_file)
    golden_path = os.path.join(spread_dir, golden_file)

    if not os.path.exists(input_path):
        return {"task_id": task_id, "condition": condition, "patch_type": patch_type,
                "pass": False, "error": "input not found", "code": ""}

    with tempfile.TemporaryDirectory() as work_dir:
        shutil.copy2(input_path, os.path.join(work_dir, input_file))
        output_path = os.path.join(work_dir, f"1_{task_id}_output.xlsx")
        preview = gen_spreadsheet_content(os.path.join(work_dir, input_file))

        artifact_section = ""
        if patch_content:
            artifact_section = f"{patch_content}\n\n"

        prompt = f"""You are a spreadsheet expert who can manipulate spreadsheets through Python code.

{artifact_section}You need to solve the given spreadsheet manipulation question:
### instruction
{meta['instruction']}

### spreadsheet_path
{os.path.join(work_dir, input_file)}

### spreadsheet_content
{preview}

### instruction_type
{meta.get('instruction_type', 'Cell-Level Manipulation')}

### answer_position
{meta['answer_position']}

### output_path
{output_path}

You should generate Python code for the final solution of the question."""

        try:
            response = call_llm(prompt, LLM_MODEL, None, None)
        except Exception as e:
            return {"task_id": task_id, "condition": condition, "patch_type": patch_type,
                    "pass": False, "error": str(e), "code": ""}

        code = extract_code(response)
        success, exec_output = execute_code(code, work_dir)
        passed = False
        if success and os.path.exists(output_path):
            passed = evaluate_task(golden_path, output_path, meta["answer_position"])

        # Diagnostic: old anchor usage
        old_anchor_used = []
        new_anchor_used = []
        if manifest:
            for anchor in manifest.get("old_anchors", []):
                if len(anchor) > 2 and anchor in code:
                    old_anchor_used.append(anchor)
            for anchor in manifest.get("new_anchors", []):
                if len(anchor) > 2 and anchor in code:
                    new_anchor_used.append(anchor)

        return {
            "task_id": task_id,
            "condition": condition,
            "patch_type": patch_type,
            "pass": passed,
            "code_executed": success,
            "old_anchor_used": old_anchor_used,
            "new_anchor_used": new_anchor_used,
            "old_anchor_count": len(old_anchor_used),
            "new_anchor_count": len(new_anchor_used),
            "code": code[:1500],
            "error": exec_output[:200] if not success else None,
        }


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--workers", type=int, default=128)
    parser.add_argument("--repeats", type=int, default=2)
    parser.add_argument("--dry_run", action="store_true")
    args = parser.parse_args()

    # Load successful codes
    with open(os.path.join(OUTPUT_DIR, "successful_codes.json")) as f:
        successful_codes = json.load(f)

    # Load manifests for anchor info
    manifests = {}
    for tid in successful_codes:
        manifest_path = os.path.join(OUTPUT_DIR, "perturbed", tid, "manifest.json")
        if os.path.exists(manifest_path):
            with open(manifest_path) as f:
                manifests[tid] = json.load(f)

    task_ids = [tid for tid in successful_codes if tid in manifests]
    print(f"Tasks with successful code + perturbed version: {len(task_ids)}")

    # Build patches
    patches = {}
    for tid in task_ids:
        code = successful_codes[tid]
        # Hard patch: the actual successful code wrapped as prior solution
        hard_patch = f"""## Prior Working Solution

Below is a prior working solution for the original version of this task. Reuse it as much as possible, adapting only when necessary.

```python
{code}
```"""
        # Soft patch: natural language description with old anchors
        old_anchors = manifests[tid].get("old_anchors", [])
        soft_patch = f"""## Prior Solution Note

I previously solved this task. Key details:
- Critical anchors used: {', '.join(f'"{a}"' for a in old_anchors)}
- I used openpyxl to load the workbook, found target columns by exact header name match
- Header row was row 1, data started from row 2
- I accessed columns by their exact header name string, not by position
"""
        patches[tid] = {
            "no_patch": "",
            "soft_patch": soft_patch,
            "hard_patch": hard_patch,
        }

    # Build jobs
    jobs = []
    for tid in task_ids:
        for condition in ["original", "perturbed"]:
            for patch_type in ["no_patch", "soft_patch", "hard_patch"]:
                for repeat in range(args.repeats):
                    jobs.append((tid, condition, patch_type, patches[tid][patch_type],
                                 manifests[tid], repeat))

    print(f"Total jobs: {len(jobs)}")
    if args.dry_run:
        return

    results = []
    lock = threading.Lock()
    done = [0]
    total = len(jobs)

    def run_job(job):
        tid, condition, patch_type, patch_content, manifest, repeat = job
        result = run_one(tid, condition, patch_type, patch_content, manifest)
        result["repeat"] = repeat
        with lock:
            done[0] += 1
            if done[0] % 20 == 0:
                s = "✅" if result["pass"] else "❌"
                print(f"[{done[0]}/{total}] {s} {tid}|{condition}|{patch_type}", flush=True)
        return result

    print(f"Running {total} jobs with {args.workers} workers...")
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(run_job, j): j for j in jobs}
        for future in as_completed(futures):
            try:
                results.append(future.result())
            except Exception as e:
                print(f"ERROR: {e}")

    # Save
    results_path = os.path.join(RESULTS_DIR, "positive_control_v2_results.json")
    with open(results_path, "w") as f:
        json.dump(results, f, indent=2)

    # Analysis
    print(f"\n{'='*60}")
    print("RESULTS")
    print(f"{'='*60}")
    for patch_type in ["no_patch", "soft_patch", "hard_patch"]:
        for condition in ["original", "perturbed"]:
            subset = [r for r in results if r["patch_type"] == patch_type and r["condition"] == condition]
            if subset:
                pr = sum(r["pass"] for r in subset) / len(subset)
                print(f"  {patch_type:12s} | {condition:10s} | {pr:.1%} ({sum(r['pass'] for r in subset)}/{len(subset)})")

    print(f"\n{'='*60}")
    print("PATCH GAP & TRANSFER LOSS")
    print(f"{'='*60}")
    no_orig = sum(r["pass"] for r in results if r["patch_type"] == "no_patch" and r["condition"] == "original")
    no_pert = sum(r["pass"] for r in results if r["patch_type"] == "no_patch" and r["condition"] == "perturbed")
    no_n = len([r for r in results if r["patch_type"] == "no_patch" and r["condition"] == "original"])

    for patch_type in ["no_patch", "soft_patch", "hard_patch"]:
        orig = [r for r in results if r["patch_type"] == patch_type and r["condition"] == "original"]
        pert = [r for r in results if r["patch_type"] == patch_type and r["condition"] == "perturbed"]
        if orig and pert:
            og = sum(r["pass"] for r in orig) / len(orig)
            pg = sum(r["pass"] for r in pert) / len(pert)
            no_og = no_orig / no_n if no_n else 0
            no_pg = no_pert / no_n if no_n else 0
            gain_orig = og - no_og
            gain_pert = pg - no_pg
            transfer_loss = gain_orig - gain_pert
            print(f"  {patch_type:12s} | Gap={og-pg:+.1%} | Gain_orig={gain_orig:+.1%} | Gain_pert={gain_pert:+.1%} | Transfer_loss={transfer_loss:+.1%}")

    print(f"\n{'='*60}")
    print("OLD ANCHOR USAGE")
    print(f"{'='*60}")
    for patch_type in ["no_patch", "soft_patch", "hard_patch"]:
        for condition in ["original", "perturbed"]:
            subset = [r for r in results if r["patch_type"] == patch_type and r["condition"] == condition]
            if subset:
                old_usage = sum(1 for r in subset if r.get("old_anchor_count", 0) > 0) / len(subset)
                new_usage = sum(1 for r in subset if r.get("new_anchor_count", 0) > 0) / len(subset)
                print(f"  {patch_type:12s} | {condition:10s} | old={old_usage:.0%} new={new_usage:.0%}")


if __name__ == "__main__":
    main()
