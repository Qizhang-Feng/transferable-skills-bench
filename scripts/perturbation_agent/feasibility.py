"""
Stage 2: Feasibility check and safety matrix evaluation.
"""


def choose_formula_policy(schema):
    """Choose formula handling policy based on golden file analysis."""
    unbakeable = schema["golden_info"]["unbakeable_in_answer"]
    if unbakeable > 0:
        return "exclude"  # Can't safely perturb
    return "cached_only"


def evaluate_safety(schema):
    """Determine which anchor types are safe for this task."""
    safe_anchors = []
    refs = schema["instruction_refs"]
    ref_inv = schema["reference_inventory"]

    # Lexical rename: safe if there are renameable names
    has_col_names = bool(refs["column_names_mentioned"])
    has_sheet_names = bool(refs["sheet_names_mentioned"])
    if has_col_names or has_sheet_names:
        safe_anchors.append("lexical")

    # Positional shift
    has_cell_refs = bool(refs["cell_refs"])
    has_row_refs = bool(refs["row_refs"])
    has_abs_refs = bool(refs["absolute_refs"])

    if not has_cell_refs and not has_row_refs and not has_abs_refs:
        safe_anchors.append("positional_title_row")
    else:
        # Can still do blank row insert (LLM will update instruction refs)
        safe_anchors.append("positional_blank_row")

    # Distractor column: safe unless charts/pivots
    if ref_inv["charts"] == 0 and ref_inv["pivot_tables"] == 0:
        safe_anchors.append("distractor_column")

    # Instruction paraphrase: always safe
    safe_anchors.append("instruction_paraphrase")

    return safe_anchors


def check_feasibility(schema):
    """Full feasibility check. Returns dict with feasible flag and details."""
    formula_policy = choose_formula_policy(schema)

    if formula_policy == "exclude":
        return {
            "feasible": False,
            "reason": f"unbakeable formulas in golden answer ({schema['golden_info']['unbakeable_in_answer']})",
            "formula_policy": formula_policy,
            "safe_anchors": [],
        }

    safe_anchors = evaluate_safety(schema)

    if not safe_anchors:
        return {
            "feasible": False,
            "reason": "no safe perturbation possible",
            "formula_policy": formula_policy,
            "safe_anchors": [],
        }

    return {
        "feasible": True,
        "formula_policy": formula_policy,
        "safe_anchors": safe_anchors,
    }
