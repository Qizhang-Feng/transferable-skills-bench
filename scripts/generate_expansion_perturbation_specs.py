#!/usr/bin/env python3
"""
Generate draft perturbation specs for expansion tasks.
Reads task schemas and instructions, produces rename maps and layout specs.
Output needs human review before use.
"""

import json
import os
import re
import openpyxl

VERIFIED_DIR = "vendor/SpreadsheetBench/data/spreadsheetbench_verified_400"

# Semantic rename suggestions (old -> new)
RENAME_SUGGESTIONS = {
    # Common column names
    "Name": "Full_Name",
    "Date": "Record_Date",
    "Amount": "Total_Amount",
    "Type": "Category",
    "Rate": "Percentage",
    "Week": "Week_Number",
    "Group": "Division",
    "ID": "Identifier",
    "Performance": "Score",
    "serial": "serial_code",
    "price": "unit_price",
    "Result Should Be": "Expected_Output",
    "Entry Time": "Check_In_Time",
    "Max": "Maximum",
    "Commission Rate": "Comm_Pct",
    "Employee Number": "Staff_ID",
    "Unique ID": "Record_Key",
    "May": "Month_05",
    "Reference": "Ref_Code",
    "Narrative": "Description",
    "DATA": "RECORDS",
    "Segment": "Category",
    "Country": "Nation",
    "Product": "Item",
    "Discount Band": "Discount_Tier",
    "Units Sold": "Qty_Sold",
    "Manufacturing Price": "Production_Cost",
    "Sale Price": "Retail_Price",
    "Gross Sales": "Total_Revenue",
    "Datum verzending": "Shipping_Date",
    "short code": "abbrev_code",
    "pillar#": "division_num",
    "key#": "index_num",
    "SUB CHANGE ORDERS": "CHANGE_ORDER_SUBS",
    "Address2": "Secondary_Address",
    "Code": "Identifier",
    "End Date": "Expiry_Date",
    # Sheet names
    "Summary": "Overview",
    "Main": "Primary",
    "Output Required": "Expected_Output",
    "Main unique ID": "Primary_Keys",
    "Result what i am getting": "Current_Result",
    "Result what i am expecting": "Expected_Result",
    "final Recon Items": "Reconciliation_Items",
    "Imported Data": "Source_Data",
    "Desired Results": "Target_Output",
    "Blad1": "DataSheet",
    "Sub CO's": "Change_Orders",
    "Data": "Records",
    "Records": "DataSource",
    "RateHurdles": "Rate_Thresholds",
    "Deposits": "Transactions",
    "MyResult": "Output_Sheet",
}


def suggest_rename(name):
    """Suggest a semantic rename for a column/sheet name."""
    if name in RENAME_SUGGESTIONS:
        return RENAME_SUGGESTIONS[name]
    # Try case-insensitive
    for k, v in RENAME_SUGGESTIONS.items():
        if k.lower() == name.lower():
            return v
    return None


def find_instruction_references(instruction, names):
    """Find which names appear in the instruction text."""
    found = []
    for name in names:
        if len(name) <= 2:
            continue
        pattern = re.compile(re.escape(name), re.IGNORECASE)
        if pattern.search(instruction):
            found.append(name)
    return found


def generate_instruction_rewrite_rules(col_rename_map, sheet_rename_map, instruction):
    """Generate instruction rewrite rules for renamed columns and sheets."""
    rules = []
    
    for old, new in sheet_rename_map.items():
        # Try various forms the sheet name might appear
        if f"'{old}'" in instruction:
            rules.append({"find": f"'{old}'", "replace": f"'{new}'"})
        if old in instruction:
            rules.append({"find": old, "replace": new, "case_insensitive": True})
    
    for old, new in col_rename_map.items():
        if old in instruction:
            rules.append({"find": old, "replace": new})
        # Also try case-insensitive
        if old.lower() != old and old.lower() in instruction.lower():
            rules.append({"find": old, "replace": new, "case_insensitive": True})
    
    return rules


def update_answer_position(answer_position, sheet_rename_map):
    """Update answer_position with renamed sheet names."""
    result = answer_position
    for old, new in sheet_rename_map.items():
        result = result.replace(f"'{old}'", f"'{new}'")
        result = result.replace(f"{old}!", f"{new}!")
        result = result.replace(f"'{old}!", f"'{new}!")
    return result


def generate_spec(task_id, schema, meta):
    """Generate a perturbation spec for one task."""
    instruction = meta.get("instruction", "")
    answer_position = meta.get("answer_position", "")
    sheet_names = schema.get("sheet_names", [])
    
    # Collect all header values
    all_headers = {}
    for sh in schema.get("sheets", []):
        for h in sh["headers"]:
            all_headers[h["value"]] = sh["name"]
    
    # Find what's mentioned in instruction
    mentioned_in_inst = find_instruction_references(
        instruction, list(all_headers.keys()) + sheet_names
    )
    
    # Build rename maps
    col_rename_map = {}
    for header_val in all_headers:
        if header_val in mentioned_in_inst or len(header_val) > 3:
            suggestion = suggest_rename(header_val)
            if suggestion and suggestion != header_val:
                col_rename_map[header_val] = suggestion
    
    sheet_rename_map = {}
    for sn in sheet_names:
        if sn in mentioned_in_inst or (sn not in ("Sheet1", "Sheet2", "Sheet3") and len(sn) > 2):
            suggestion = suggest_rename(sn)
            if suggestion and suggestion != sn:
                sheet_rename_map[sn] = suggestion
    
    # Generate instruction rewrite rules
    rewrite_rules = generate_instruction_rewrite_rules(
        col_rename_map, sheet_rename_map, instruction
    )
    
    # Update answer position
    new_answer_position = update_answer_position(answer_position, sheet_rename_map)
    
    # Check formula count for layout feasibility
    total_formulas = sum(sh.get("formula_count", 0) for sh in schema.get("sheets", []))
    layout_feasible = total_formulas < 50
    
    spec = {
        "task_id": task_id,
        "family": schema.get("family", ""),
        "instruction_type": meta.get("instruction_type", ""),
        "rename": {
            "feasible": bool(col_rename_map or sheet_rename_map),
            "column_rename_map": col_rename_map,
            "sheet_rename_map": sheet_rename_map,
            "instruction_rewrite_needed": bool(rewrite_rules),
            "answer_position_update": new_answer_position if new_answer_position != answer_position else None,
            "unsafe_to_rename_fields": [],
            "semantic_equivalence_notes": "",
            "instruction_rewrite_rules": rewrite_rules,
            "answer_position_update_is_hint": True,
        },
        "layout": {
            "feasible": layout_feasible,
            "shift_rows": 1,
            "title_text": "Report",
            "answer_position_update": "NEEDS_MANUAL_REVIEW",
            "semantic_equivalence_notes": "Answer position row offset needs manual verification.",
            "notes": f"Total formulas: {total_formulas}. " + (
                "Formula shift required." if total_formulas > 0 else "No formulas."
            ),
            "formula_shift_required": total_formulas > 0,
            "answer_position_update_is_hint": True,
        },
        "overall_perturbability": schema.get("perturbability", "medium"),
        "exclusion_reason": None,
        "review_status": "draft",
        "execution_order": "rename_first_then_layout",
    }
    
    return spec


def main():
    # Load schemas
    with open("scratch/expansion-task-schemas.json") as f:
        schemas = json.load(f)
    schema_map = {s["task_id"]: s for s in schemas}
    
    # Load dataset metadata
    with open(os.path.join(VERIFIED_DIR, "dataset.json")) as f:
        dataset = json.load(f)
    meta_map = {str(d["id"]): d for d in dataset}
    
    # New tasks to process
    new_tasks = [
        "24-23", "341-40",  # conditional_partition
        "1818", "31915", "49801", "36191", "49857", "58723", "57558",  # formula_writeback
        "52575", "203-15", "395-36",  # grouped_aggregation
        "367-23", "493-5", "488-14", "374-31", "486-17", "577-40",  # structural_edit
    ]
    
    specs = {}
    for tid in new_tasks:
        schema = schema_map.get(tid)
        meta = meta_map.get(tid)
        if not schema or not meta:
            print(f"WARNING: Missing data for {tid}")
            continue
        
        spec = generate_spec(tid, schema, meta)
        specs[tid] = spec
        
        # Summary
        n_col = len(spec["rename"]["column_rename_map"])
        n_sheet = len(spec["rename"]["sheet_rename_map"])
        n_rules = len(spec["rename"]["instruction_rewrite_rules"])
        layout = "yes" if spec["layout"]["feasible"] else "no"
        print(f"  {tid:12s} | rename: {n_col} cols, {n_sheet} sheets, {n_rules} rules | layout: {layout}")
    
    # Save
    output_path = "scratch/expansion-perturbation-specs-draft.json"
    with open(output_path, "w") as f:
        json.dump(specs, f, indent=2, ensure_ascii=False)
    
    print(f"\nSaved {len(specs)} specs to {output_path}")
    print("⚠️  These are DRAFTS — need human review before generating perturbed files!")


if __name__ == "__main__":
    main()
