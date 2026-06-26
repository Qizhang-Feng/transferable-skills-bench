#!/usr/bin/env python3
"""
Experiment runner for v3 perturbed tasks.
Runs 285 tasks × 2 conditions (original/perturbed) × no_skill × N repeats.
Reuses LLM call, code execution, and evaluation from run_pilot_experiment.py.
"""

import json
import os
import sys
import argparse
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter, defaultdict

# Reuse from pilot experiment
sys.path.insert(0, ".")
from scripts.run_pilot_experiment import (
    run_single, evaluate_task, gen_spreadsheet_content,
    call_llm, extract_code, execute_code,
    VERIFIED_DIR,
)

V3_PERTURBED_DIR = "data/perturbed_v3"
ARTIFACT_DIR = "data/v3_artifacts"
RESULTS_DIR = "scratch/experiment_results"

FAMILY_DIR_MAP = {
    "formula_writeback": "family_a_formula_writeback",
    "lookup_match_merge": "family_b_lookup_match_merge",
    "conditional_partition": "family_c_conditional_partition",
    "grouped_aggregation": "family_d_grouped_aggregation",
    "structural_edit": "family_e_structural_edit",
}


def load_artifact(artifact_type, family, task_id=None):
    """Load artifact content."""
    if artifact_type == "no_skill":
        return ""
    if artifact_type == "gen_skill":
        path = os.path.join(ARTIFACT_DIR, "gen_skill", f"{family}.md")
        if os.path.exists(path):
            with open(path) as f:
                return f.read()
    if artifact_type == "local_patch" and task_id:
        path = os.path.join(ARTIFACT_DIR, "local_patch", f"{task_id}.md")
        if os.path.exists(path):
            with open(path) as f:
                return f.read()
    return ""


def load_v3_tasks():
    """Load all pass tasks from v3 manifests."""
    manifest_path = os.path.join(V3_PERTURBED_DIR, "all_manifests.json")
    with open(manifest_path) as f:
        manifests = json.load(f)

    with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
        dataset = json.load(f)
    meta_map = {str(d["id"]): d for d in dataset}

    tasks = []
    for tid, m in manifests.items():
        if m["status"] != "pass":
            continue
        if tid not in meta_map:
            continue
        tasks.append({
            "id": tid,
            "family": m.get("family", "unknown"),
            "anchor_types": m.get("anchor_types_used", []),
        })
    return tasks


def run_single_v3(task_id, condition, artifact_type, family, model, api_key, base_url):
    """Run a single experiment for v3 tasks."""
    import tempfile, shutil, openpyxl

    if condition == "original":
        spread_dir = os.path.join(VERIFIED_DIR, "spreadsheet", str(task_id))
        input_file = f"1_{task_id}_init.xlsx"
        golden_file = f"1_{task_id}_golden.xlsx"
        with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
            dataset = json.load(f)
        meta = next(d for d in dataset if str(d["id"]) == str(task_id))
    else:
        spread_dir = os.path.join(V3_PERTURBED_DIR, str(task_id))
        input_file = f"1_{task_id}_init.xlsx"
        golden_file = f"1_{task_id}_golden.xlsx"
        meta_path = os.path.join(spread_dir, "task_meta.json")
        if not os.path.exists(meta_path):
            return {"task_id": task_id, "condition": condition, "artifact_type": artifact_type,
                    "pass": False, "error": "task_meta.json not found"}
        with open(meta_path) as f:
            meta = json.load(f)

    input_path = os.path.join(spread_dir, input_file)
    golden_path = os.path.join(spread_dir, golden_file)

    if not os.path.exists(input_path):
        return {"task_id": task_id, "condition": condition, "artifact_type": artifact_type,
                "pass": False, "error": "input file not found"}

    artifact_content = load_artifact(artifact_type, family, task_id=task_id)
    artifact_section = ""
    if artifact_content:
        artifact_section = (
            "### knowledge_artifact\n"
            "The following knowledge artifact may help you solve this task. "
            "Use it as reference if relevant.\n\n"
            f"{artifact_content}\n\n"
        )

    with tempfile.TemporaryDirectory() as work_dir:
        shutil.copy2(input_path, os.path.join(work_dir, input_file))
        output_file = f"1_{task_id}_output.xlsx"
        output_path = os.path.join(work_dir, output_file)

        spreadsheet_content = gen_spreadsheet_content(os.path.join(work_dir, input_file))

        prompt = f"""You are a spreadsheet expert who can manipulate spreadsheets through Python code.

{artifact_section}You need to solve the given spreadsheet manipulation question, which contains six types of information:
- instruction: The question about spreadsheet manipulation.
- spreadsheet_path: The path of the spreadsheet file you need to manipulate.
- spreadsheet_content: The first few rows of the content of spreadsheet file.
- instruction_type: There are two values (Cell-Level Manipulation, Sheet-Level Manipulation).
- answer_position: The position need to be modified or filled.
- output_path: You need to generate the modified spreadsheet file in this new path.

Below is the spreadsheet manipulation question you need to solve:
### instruction
{meta['instruction']}

### spreadsheet_path
{os.path.join(work_dir, input_file)}

### spreadsheet_content
{spreadsheet_content}

### instruction_type
{meta.get('instruction_type', 'Cell-Level Manipulation')}

### answer_position
{meta['answer_position']}

### output_path
{output_path}

You should generate Python code for the final solution of the question."""

        try:
            response = call_llm(prompt, model, api_key, base_url)
        except Exception as e:
            return {"task_id": task_id, "condition": condition,
                    "pass": False, "error": f"LLM error: {e}"}

        code = extract_code(response)
        success, exec_output = execute_code(code, work_dir)

        passed = False
        if success and os.path.exists(output_path):
            passed = evaluate_task(golden_path, output_path, meta["answer_position"])

        return {
            "task_id": task_id,
            "condition": condition,
            "pass": passed,
            "code_executed": success,
            "model": model,
            "output_file_exists": os.path.exists(output_path),
            "error": exec_output[:200] if not success else None,
        }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default="us.anthropic.claude-sonnet-4-5-20250929-v1:0")
    parser.add_argument("--api_key", default=None)
    parser.add_argument("--base_url", default=None)
    parser.add_argument("--workers", type=int, default=64)
    parser.add_argument("--repeats", type=int, default=1)
    parser.add_argument("--max_tasks", type=int, default=None, help="Limit number of tasks")
    parser.add_argument("--dry_run", action="store_true")
    args = parser.parse_args()

    v3_tasks = load_v3_tasks()
    if args.max_tasks:
        v3_tasks = v3_tasks[:args.max_tasks]

    print(f"Loaded {len(v3_tasks)} v3 pass tasks")

    os.makedirs(RESULTS_DIR, exist_ok=True)

    # Build job list: each task × 2 conditions × artifact_types × repeats
    artifact_types = ["no_skill", "gen_skill", "local_patch"]
    jobs = []
    for task in v3_tasks:
        for condition in ["original", "perturbed"]:
            for artifact_type in artifact_types:
                for repeat in range(args.repeats):
                    jobs.append((task["id"], condition, artifact_type, task["family"], task["anchor_types"], repeat))

    results_path = os.path.join(RESULTS_DIR, f"results_v3_{args.model.split('/')[-1]}.json")

    # Resume support
    existing_results = []
    completed_keys = set()
    if os.path.exists(results_path):
        with open(results_path) as f:
            existing_results = json.load(f)
        for r in existing_results:
            completed_keys.add((str(r["task_id"]), r["condition"], r.get("artifact_type", "no_skill"), r.get("repeat", 0)))
        print(f"Resuming: {len(completed_keys)} done, {len(jobs) - len(completed_keys)} remaining")

    remaining = [j for j in jobs if (j[0], j[1], j[2], j[5]) not in completed_keys]

    if args.dry_run:
        print(f"Would run {len(remaining)} jobs")
        for cond in ["original", "perturbed"]:
            n = sum(1 for j in remaining if j[1] == cond)
            print(f"  {cond}: {n}")
        return

    results = list(existing_results)
    lock = threading.Lock()
    done_count = [len(completed_keys)]
    total = len(jobs)

    def run_job(job):
        task_id, condition, artifact_type, family, anchors, repeat = job
        result = run_single_v3(task_id, condition, artifact_type, family, args.model, args.api_key, args.base_url)
        result["family"] = family
        result["anchor_types"] = anchors
        result["artifact_type"] = artifact_type
        result["repeat"] = repeat
        with lock:
            done_count[0] += 1
            status = "✅" if result["pass"] else "❌"
            print(f"[{done_count[0]}/{total}] {status} {task_id} | {condition} | {artifact_type} | r{repeat}")
        return result

    print(f"Running {len(remaining)} jobs with {args.workers} workers...")

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(run_job, job): job for job in remaining}
        for future in as_completed(futures):
            try:
                result = future.result()
                with lock:
                    results.append(result)
                    if len(results) % 50 == 0:
                        with open(results_path, "w") as f:
                            json.dump(results, f, indent=2)
            except Exception as e:
                job = futures[future]
                print(f"ERROR on {job[0]}: {e}")

    # Final save
    with open(results_path, "w") as f:
        json.dump(results, f, indent=2)

    # Summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    for artifact_type in ["no_skill", "gen_skill", "local_patch"]:
        for condition in ["original", "perturbed"]:
            subset = [r for r in results if r.get("artifact_type", "no_skill") == artifact_type and r["condition"] == condition]
            if subset:
                pass_rate = sum(r["pass"] for r in subset) / len(subset)
                print(f"  {artifact_type:12s} | {condition:10s} | {pass_rate:.1%} ({sum(r['pass'] for r in subset)}/{len(subset)})")

    # Patch Gap
    print(f"\n{'='*60}")
    print("PATCH GAP (original_rate - perturbed_rate)")
    print(f"{'='*60}")
    for artifact_type in ["no_skill", "gen_skill", "local_patch"]:
        orig = [r for r in results if r.get("artifact_type", "no_skill") == artifact_type and r["condition"] == "original"]
        pert = [r for r in results if r.get("artifact_type", "no_skill") == artifact_type and r["condition"] == "perturbed"]
        if orig and pert:
            orig_rate = sum(r["pass"] for r in orig) / len(orig)
            pert_rate = sum(r["pass"] for r in pert) / len(pert)
            gap = orig_rate - pert_rate
            print(f"  {artifact_type:12s} | Patch Gap = {gap:+.1%}  ({orig_rate:.1%} -> {pert_rate:.1%})")

    # Per-family breakdown
    print(f"\n{'='*60}")
    print("PER FAMILY")
    print(f"{'='*60}")
    families = sorted(set(r.get("family", "?") for r in results))
    for fam in families:
        for artifact_type in ["no_skill", "gen_skill", "local_patch"]:
            for cond in ["original", "perturbed"]:
                subset = [r for r in results if r.get("family") == fam and r.get("artifact_type", "no_skill") == artifact_type and r["condition"] == cond]
                if subset:
                    pr = sum(r["pass"] for r in subset) / len(subset)
                    print(f"  {fam:25s} | {artifact_type:10s} | {cond:10s} | {pr:.1%} ({sum(r['pass'] for r in subset)}/{len(subset)})")


if __name__ == "__main__":
    main()
