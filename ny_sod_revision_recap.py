"""
ny_sod_revision_recap.py

Diffs two NY SOD "GOAL SHEET" .xlsb workbooks (an earlier INPUT/draft version
and a later FINAL/revised version) and reports what changed, added, or
removed, so the team doesn't have to eyeball the green highlighting by hand.

NY SOD-specific: this is keyed to the exact GOAL SHEET column layout used by
the Monthly Quota Planner template. It will not work against the SPP Goal
Builder's Tracking Table or any other state's template.

Usage:
    python ny_sod_revision_recap.py <input.xlsb> <final.xlsb> [output_csv]

Rows are matched between the two files by (Division, Supplier, Goal Name,
Product) rather than the '#' column, since inserting/deleting a row shifts
every '#' below it and would otherwise look like a wall of false changes.
"""

import os
import sys

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required.  Install with: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    import pyxlsb  # noqa: F401
except ImportError:
    print("ERROR: pyxlsb is required to read .xlsb files.  Install with: pip install pyxlsb", file=sys.stderr)
    sys.exit(1)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_lpm_upload as spp

# Columns worth calling out — human-entered fields, not the calculated
# lookback/comparison columns (LY Case, R3, R12, LYDEC, etc.)
COMPARE_COLUMNS = [
    "Supplier", "Goal Name", "Category", "Product Filter", "Product", "Size",
    "Premise - On or Off Only", "Goal Type", "Goal Criteria", "Goal Calculation",
    "Min", "Division", "Must Make", "Measure Decimal/9L/STD",
    "Unsold Start", "Unsold End", "Tracking Start", "Tracking End",
    "On Program Y/N?", "Goal",
]

KEY_COLUMNS = ["Division", "Supplier", "Goal Name", "Product"]


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return str(round(v, 4))
    return str(v).strip()


def load_rows(xlsb_path):
    df = pd.read_excel(xlsb_path, sheet_name="GOAL SHEET", engine="pyxlsb", header=4)
    rows = []
    for i, row in df.iterrows():
        if pd.isna(row.get("Goal Name")) and pd.isna(row.get("Division")):
            continue
        rows.append({c: _clean(row.get(c)) for c in COMPARE_COLUMNS})
    return rows


def key_for(row):
    return tuple(row.get(c, "") for c in KEY_COLUMNS)


def index_rows(rows):
    """Map key -> list of rows (duplicates possible on '-1'/'-2' double buckets)."""
    idx = {}
    for row in rows:
        idx.setdefault(key_for(row), []).append(row)
    return idx


def diff_rows(input_rows, final_rows):
    input_idx = index_rows(input_rows)
    final_idx = index_rows(final_rows)

    changes = []  # list of dicts: Goal Name, Division, Change Type, Changes

    all_keys = list(dict.fromkeys(list(input_idx.keys()) + list(final_idx.keys())))

    for key in all_keys:
        in_list = input_idx.get(key, [])
        fin_list = final_idx.get(key, [])
        n = max(len(in_list), len(fin_list))
        for i in range(n):
            in_row = in_list[i] if i < len(in_list) else None
            fin_row = fin_list[i] if i < len(fin_list) else None
            division, supplier, goal_name, product = key

            if in_row is not None and fin_row is None:
                changes.append({
                    "Goal Name": goal_name, "Division": division, "Change Type": "Removed",
                    "Changes": f'Removed goal "{goal_name}" ({supplier})',
                })
            elif in_row is None and fin_row is not None:
                changes.append({
                    "Goal Name": goal_name, "Division": division, "Change Type": "Added",
                    "Changes": f'Added new goal "{goal_name}" ({supplier})',
                })
            else:
                field_diffs = []
                for c in COMPARE_COLUMNS:
                    old, new = in_row.get(c, ""), fin_row.get(c, "")
                    if old != new:
                        field_diffs.append(f"{c}: {old or '(blank)'}→{new or '(blank)'}")
                if field_diffs:
                    changes.append({
                        "Goal Name": goal_name, "Division": division, "Change Type": "Modified",
                        "Changes": "; ".join(field_diffs),
                    })

    return changes


def write_recap_csv(changes, output_path):
    import csv
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["Goal Name", "Division", "Change Type", "Changes"], quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for c in changes:
            writer.writerow(c)


def main():
    if len(sys.argv) < 3:
        print("Usage: python ny_sod_revision_recap.py <input.xlsb> <final.xlsb> [output.csv]", file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    final_path = sys.argv[2]
    for p in (input_path, final_path):
        if not os.path.isfile(p):
            print(f"ERROR: File not found: {p}", file=sys.stderr)
            sys.exit(1)

    if len(sys.argv) >= 4:
        output_path = sys.argv[3]
    else:
        base = os.path.splitext(os.path.abspath(final_path))[0]
        output_path = base + "_revision_recap.csv"

    print(f"Input : {input_path}")
    print(f"Final : {final_path}")

    input_rows = load_rows(input_path)
    final_rows = load_rows(final_path)
    changes = diff_rows(input_rows, final_rows)

    write_recap_csv(changes, output_path)

    added    = sum(1 for c in changes if c["Change Type"] == "Added")
    removed  = sum(1 for c in changes if c["Change Type"] == "Removed")
    modified = sum(1 for c in changes if c["Change Type"] == "Modified")

    print()
    print("=" * 65)
    print("NY SOD Revision Recap Summary")
    print("=" * 65)
    print(f"Output file : {output_path}")
    print(f"Modified    : {modified}")
    print(f"Added       : {added}")
    print(f"Removed     : {removed}")
    if not changes:
        print("\nNo differences found.")
    print("=" * 65)


if __name__ == "__main__":
    main()
