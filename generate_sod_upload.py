"""
generate_sod_upload.py

Reads the "GOAL SHEET" table from a NY SOD Monthly Quota Planner .xlsb file and
generates LPM-format CSV files ready for upload to LiquidPerform. Fork of
generate_lpm_upload.py (the SPP generator) — reuses its CSV schema, POD
attribute map, and collection-lookup helpers, with NY SOD-specific column
mapping and date resolution.

Usage:
    python generate_sod_upload.py <goal_sheet.xlsb> [output_csv] [collection_report.xlsx]

One Tracker is generated per unique (Division, goal_type, goal_uom,
Tracking Start/End, unsold-period, Premise, Must Make) combination; every
GOAL SHEET row under that combination becomes a PTG row. SOD goes into LPM
as quotas, so program_class is the literal string "QUO" for every tracker
(not an enum path like SPP's "ProgramClass.Site.Sppa") — confirmed, not an
open assumption.

Open assumptions (confirm/adjust before relying on this in production):
  - "Reverse MS" (Goal of Unsold) is treated like "Market Share" (Unit Goal)
    for distribution purposes (min_objective_target blank, distribution_target
    = the "Goal" column value) — only basis_flag differs (FALSE vs TRUE).
  - product_collection_id / achievement_min are left blank — no equivalent
    source column has been identified yet in the GOAL SHEET.
"""

import os
import re
import sys
from collections import defaultdict

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required.  Install with: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    import pyxlsb  # noqa: F401  (engine used by pandas.read_excel)
except ImportError:
    print("ERROR: pyxlsb is required to read .xlsb files.  Install with: pip install pyxlsb", file=sys.stderr)
    sys.exit(1)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_lpm_upload as spp  # reuse CSV schema + helpers from the SPP generator

CSV_COLUMNS = spp.CSV_COLUMNS

SOD_CONFIG = {
    "program_class":                  "QUO",
    "score_by_td_linx_customer_code": "",
    "attainment_org_level":           "Salesperson",
    "recalculate_until_date":         "365",
    "exclude_from_os_sales_reports":  "FALSE",
    "salesforce_collection_ids":      "",
    "send_to_proof":                  "SendStartingInActive",
    "send_to_proof_date":             "",
    "distribution_level_path":        "Salesperson",
}

# GOAL SHEET "Goal Type" -> LPM goal_type (before unsold-period coercion)
GOAL_TYPE_MAP = {
    "Volume":        "VolumeCases",
    "DREV":          "VolumeRevenueDREV",
    "ACS":           "DistributionACS",
    "PODs by SKU":   "DistributionPODs",
    "PODs by Type":  "DistributionPODs",
    "PODs by Size":  "DistributionPODs",
}

# "PODs by X" -> pod_attribute
POD_TYPE_ATTR_MAP = {
    "PODs by SKU":  "ProductId",
    "PODs by Type": "PhSubGroup",
    "PODs by Size": "ProductSize",
}

# Only these support New-period coercion when an unsold window is present
UNSOLD_ELIGIBLE = {"DistributionACS", "DistributionPODs"}
UNSOLD_COERCION = {
    "DistributionACS":  "DistributionNewACS",
    "DistributionPODs": "DistributionNewPODs",
}

# GOAL SHEET "Measure Decimal/9L/STD" -> LPM goal_uom
UOM_MAP = {
    "Decimal": "Cases",
    "9L":      "NineLiter",
    "STD":     "STD",
}

# GOAL SHEET "Goal Criteria" -> goal_distribution concept (mirrors SPP's
# FIXED_GOAL_DISTRIBUTION / EVEN_GOAL_DISTRIBUTION treatment)
FIXED_CRITERIA = {"Per Rep"}
EVEN_CRITERIA  = {"Market Share", "Reverse MS"}
NO_MIN_CRITERIA = FIXED_CRITERIA | EVEN_CRITERIA

# Goal Criteria that trigger basis_flag = TRUE
BASIS_TRUE_CRITERIA = {"Market Share", "% Increase/Decrease"}

MONTH_NUM = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "SEPT": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

GOAL_SHEET_COLUMNS = [
    "#", "LIQUID TRACKER NUMBERS", "Supplier", "Goal Name", "Category",
    "Product Filter", "Product", "Size", "Premise - On or Off Only",
    "Goal Type", "Goal Criteria", "Goal Calculation", "Min", "Division",
    "Must Make", "Measure Decimal/9L/STD", "Unsold Start", "Unsold End",
    "Tracking Start", "Tracking End", "Commnets", "SD Notes",
    "On Program Y/N?", "Goal",
]


# ---------------------------------------------------------------------------
# Year/month resolution
# ---------------------------------------------------------------------------

def derive_anchor_year(filepath, goal_month_cell):
    """
    Program year for the file. Prefers a 4-digit year in the filename
    (e.g. "...2026 MONTHLY QUOTA PLANNER...09.14.26.xlsb" -> 2026); falls back
    to the current year if none is found.
    """
    name = os.path.basename(filepath)
    m = re.search(r"\b(20\d{2})\b", name)
    if m:
        return int(m.group(1))
    from datetime import date
    return date.today().year


def month_pair_to_yyyymm(start_abbr, end_abbr, anchor_year, anchor_month=None):
    """
    Resolve a (Start, End) month-abbreviation pair to (start_yyyymm, end_yyyymm).

    Current year is assumed for End unless anchor_month is given and End is
    clearly after it (rare — Tracking/Unsold periods run up to the anchor
    month, not past it). Start takes the same year as End unless Start's
    month number is greater than End's (a wraparound, e.g. Unsold Sep->Feb),
    in which case Start is End's year minus one.
    """
    s = safe_month(start_abbr)
    e = safe_month(end_abbr)
    if not s or not e:
        return "", ""

    end_year = anchor_year
    if anchor_month is not None and e > anchor_month:
        end_year -= 1

    start_year = end_year if s <= e else end_year - 1

    return f"{start_year}{s:02d}", f"{end_year}{e:02d}"


def safe_month(abbr):
    if not abbr:
        return None
    return MONTH_NUM.get(str(abbr).strip().upper())


# ---------------------------------------------------------------------------
# Load GOAL SHEET
# ---------------------------------------------------------------------------

def load_goal_sheet(xlsb_path):
    """
    Read the GOAL SHEET table and return (records, skipped_rows, anchor_year,
    anchor_month).
    """
    import pandas as pd

    raw = pd.read_excel(xlsb_path, sheet_name="GOAL SHEET", engine="pyxlsb", header=None, nrows=1)
    goal_month_abbr = raw.iat[0, 0] if not raw.empty else None
    anchor_month = safe_month(goal_month_abbr)
    anchor_year = derive_anchor_year(xlsb_path, goal_month_abbr)

    df = pd.read_excel(xlsb_path, sheet_name="GOAL SHEET", engine="pyxlsb", header=4)

    records = []
    skipped = []

    for i, row in df.iterrows():
        excel_row = i + 6  # header on sheet row 5 (1-indexed); data starts row 6

        def gv(col):
            v = row.get(col)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return None
            return v

        goal_name = spp.safe_str(gv("Goal Name"))
        division  = spp.safe_str(gv("Division"))
        goal_type_raw = spp.safe_str(gv("Goal Type"))
        goal_criteria = spp.safe_str(gv("Goal Criteria"))
        premise = spp.safe_str(gv("Premise - On or Off Only")).upper()
        must_make = spp.safe_str(gv("Must Make"))

        raw_row = [gv(c) for c in GOAL_SHEET_COLUMNS]

        # Skip blank rows
        if not goal_name and not division:
            continue

        goal_type = GOAL_TYPE_MAP.get(goal_type_raw)
        if not goal_type:
            skipped.append({"row_num": excel_row, "reason": f"unknown Goal Type: {goal_type_raw!r}", "raw_row": raw_row})
            continue

        unsold_start_raw = gv("Unsold Start")
        unsold_end_raw   = gv("Unsold End")
        unsold_present = bool(safe_month(unsold_start_raw)) and bool(safe_month(unsold_end_raw))

        if unsold_present and goal_type in UNSOLD_ELIGIBLE:
            goal_type = UNSOLD_COERCION[goal_type]

        pod_attribute = POD_TYPE_ATTR_MAP.get(goal_type_raw, "")

        tracking_start_raw = gv("Tracking Start")
        tracking_end_raw   = gv("Tracking End")
        start_yyyymm, end_yyyymm = month_pair_to_yyyymm(
            tracking_start_raw, tracking_end_raw, anchor_year, anchor_month
        )
        if not start_yyyymm or not end_yyyymm:
            skipped.append({"row_num": excel_row, "reason": "missing/invalid Tracking Start or End", "raw_row": raw_row})
            continue

        unsold_start_yyyymm, unsold_end_yyyymm = "", ""
        if unsold_present:
            end_month, end_year = int(end_yyyymm[4:6]), int(end_yyyymm[:4])
            unsold_start_yyyymm, unsold_end_yyyymm = month_pair_to_yyyymm(
                unsold_start_raw, unsold_end_raw, end_year, end_month
            )

        measure = spp.safe_str(gv("Measure Decimal/9L/STD"))
        if goal_type in {"VolumeRevenueDREV"}:
            goal_uom = ""
        else:
            goal_uom = UOM_MAP.get(measure, "Cases" if goal_type in spp.DISTRIBUTION_TYPES else "")

        goal_value_raw = gv("Goal")
        if isinstance(goal_value_raw, float):
            goal_value_raw = round(goal_value_raw, 4)
        goal_value = spp._numeric_str(goal_value_raw)
        if not goal_value:
            skipped.append({"row_num": excel_row, "reason": "no Goal value", "raw_row": raw_row})
            continue

        records.append({
            "row_num":              excel_row,
            "division":             division,
            "goal_type":            goal_type,
            "goal_uom":             goal_uom,
            "start_yyyymm":         start_yyyymm,
            "end_yyyymm":           end_yyyymm,
            "unsold_prd":           "unsold" if unsold_present else "",
            "unsold_start_yyyymm":  unsold_start_yyyymm,
            "unsold_end_yyyymm":    unsold_end_yyyymm,
            "ptg_name":             goal_name,
            "goal_criteria":        goal_criteria,
            "goal_value":           goal_value,
            "pod_attribute":        pod_attribute,
            "premise":              premise,
            "must_make":            must_make,
        })

    return records, skipped, anchor_year, anchor_month


# ---------------------------------------------------------------------------
# Group + build rows
# ---------------------------------------------------------------------------

def group_key(rec):
    # Key on the actual resolved unsold window, not just presence — two rows
    # with the same goal_type/tracking period but different unsold periods
    # (e.g. JUL-AUG vs AUG-AUG) must land in separate Trackers. Premise (On
    # premise/Off premise) has no LPM CSV field of its own, but still gates
    # grouping -- an On-only row and an Off-only row must not share a
    # Tracker just because everything else matches. Must Make (and NYU's
    # Anchor/Flex, which share the same GOAL SHEET column) likewise gates
    # grouping -- a Must Make row and a regular row must not land in the
    # same Tracker just because every other field matches, mirroring how
    # SPP's spp_tier (Anchor/Flex) is part of its own group_key.
    return (
        rec["division"], rec["goal_type"], rec["goal_uom"],
        rec["start_yyyymm"], rec["end_yyyymm"],
        rec["unsold_start_yyyymm"], rec["unsold_end_yyyymm"],
        rec["premise"], rec["must_make"],
    )


def group_records(records):
    groups = defaultdict(list)
    order = []
    for rec in records:
        key = group_key(rec)
        if key not in groups:
            order.append(key)
        groups[key].append(rec)
    return order, groups


def build_tracker_row(key, recs):
    division, goal_type, goal_uom, start, end, unsold_start, unsold_end, _premise, _must_make = key
    return {
        "goal_category":                  "Tracker",
        "goal_name":                      division,
        "tpm_nav_ref":                    "",
        "goal_type":                      goal_type,
        "goal_description":               "",
        "goal_start_date":                start,
        "goal_end_date":                  end,
        "basis_flag":                     "",  # set per-PTG-group below; trackers hold the group's shared value
        "score_by_td_linx_customer_code": SOD_CONFIG["score_by_td_linx_customer_code"],
        "attainment_org_level":           SOD_CONFIG["attainment_org_level"],
        "goal_uom":                       goal_uom,
        "recalculate_until_date":         SOD_CONFIG["recalculate_until_date"],
        "exclude_from_os_sales_reports":  SOD_CONFIG["exclude_from_os_sales_reports"],
        "salesforce_collection_ids":      spp.COLLECTION_LOOKUP.get(division.lower(), SOD_CONFIG["salesforce_collection_ids"]),
        "program_class":                  SOD_CONFIG["program_class"],
        "send_to_proof":                  SOD_CONFIG["send_to_proof"],
        "send_to_proof_date":             SOD_CONFIG["send_to_proof_date"],
        "unsold_start_date":              unsold_start,
        "unsold_end_date":                unsold_end,
        "product_collection_id":          "",
        "customer_collection_id":         "",
        "distribution_target":            "",
        "min_objective_target":           "",
        "distribution_level_path":        "",
        "pod_attribute":                  "",
        "achievement_min":                "",
    }


def build_ptg_row(rec):
    is_dist = rec["goal_type"] in spp.POD_ATTR_TYPES
    no_min  = rec["goal_criteria"] in NO_MIN_CRITERIA
    return {
        "goal_category":                  "PTG",
        "goal_name":                      rec["ptg_name"],
        "tpm_nav_ref":                    "",
        "goal_type":                      "",
        "goal_description":               "",
        "goal_start_date":                "",
        "goal_end_date":                  "",
        "basis_flag":                     "TRUE" if rec["goal_criteria"] in BASIS_TRUE_CRITERIA else "FALSE",
        "score_by_td_linx_customer_code": "",
        "attainment_org_level":           "",
        "goal_uom":                       "",
        "recalculate_until_date":         "",
        "exclude_from_os_sales_reports":  "",
        "salesforce_collection_ids":      "",
        "program_class":                  "",
        "send_to_proof":                  "",
        "send_to_proof_date":             "",
        "unsold_start_date":              "",
        "unsold_end_date":                "",
        "product_collection_id":          "",
        "customer_collection_id":         "",
        "distribution_target":            rec["goal_value"],
        "min_objective_target":           "" if no_min else "1",
        "distribution_level_path":        SOD_CONFIG["distribution_level_path"],
        "pod_attribute":                  (rec["pod_attribute"] or "ProductId") if is_dist else "",
        "achievement_min":                "",
    }


def find_duplicate_ptg_names(order, groups):
    """
    Every GOAL SHEET row always becomes its own PTG row here — rows are never
    merged. But LPM keys a PTG by its goal_name within a Tracker, so two
    distinct rows sharing the exact same Goal Name under the same Tracker
    (division/goal_type/period) can collide on import even though this script
    emits them as separate lines. Flag those so they get caught before upload.
    """
    dupes = []
    for key in order:
        recs = groups[key]
        by_name = defaultdict(list)
        for r in recs:
            by_name[r["ptg_name"]].append(r["row_num"])
        division = key[0]
        for name, row_nums in by_name.items():
            if len(row_nums) > 1:
                dupes.append({"tracker": division, "ptg_name": name, "row_nums": row_nums})
    return dupes


def generate_output(order, groups, output_path):
    import csv
    total_trackers = 0
    total_ptgs = 0
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for key in order:
            recs = groups[key]
            tracker_row = build_tracker_row(key, recs)
            writer.writerow(tracker_row)
            for rec in recs:
                writer.writerow(build_ptg_row(rec))
            total_trackers += 1
            total_ptgs += len(recs)
    return total_trackers, total_ptgs


def write_skipped_csv(skipped_path, skipped):
    import csv
    with open(skipped_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(["skip_reason"] + GOAL_SHEET_COLUMNS)
        for entry in skipped:
            writer.writerow([entry["reason"]] + entry["raw_row"])


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_sod_upload.py <goal_sheet.xlsb> [output.csv] [collection_report.xlsx]", file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base = os.path.splitext(os.path.abspath(input_path))[0]
        output_path = base + "_sod_upload.csv"

    collection_path = sys.argv[3] if len(sys.argv) >= 4 else None
    if not collection_path:
        folder = os.path.dirname(os.path.abspath(input_path))
        import glob as _glob
        candidates = _glob.glob(os.path.join(folder, "*Collection*Report*.xlsx"))
        if candidates:
            collection_path = candidates[0]

    if collection_path and os.path.isfile(collection_path):
        spp.COLLECTION_LOOKUP = spp.load_collection_lookup(collection_path, "NY")
        print(f"Collection IDs loaded: {len(spp.COLLECTION_LOOKUP)}  ({os.path.basename(collection_path)})")
    else:
        print("No collection report found — salesforce_collection_ids will be blank.")

    print(f"Reading: {input_path}")
    records, skipped, anchor_year, anchor_month = load_goal_sheet(input_path)
    print(f"Anchor: year={anchor_year} month={anchor_month}")

    if not records:
        print("ERROR: No valid records found in GOAL SHEET after filtering.", file=sys.stderr)
        sys.exit(1)

    order, groups = group_records(records)
    base = os.path.splitext(output_path)[0]
    skipped_path = base + "_skipped.csv"
    total_trackers, total_ptgs = generate_output(order, groups, output_path)
    duplicate_names = find_duplicate_ptg_names(order, groups)

    if skipped:
        write_skipped_csv(skipped_path, skipped)

    print()
    print("=" * 65)
    print("SOD Upload Generation Summary")
    print("=" * 65)
    print(f"Output file : {output_path}")
    print(f"Trackers    : {total_trackers}")
    print(f"PTGs        : {total_ptgs}")
    if skipped:
        print(f"\n*** {len(skipped)} row(s) skipped — saved to {skipped_path} ***")
        for entry in skipped:
            print(f"  row {entry['row_num']}: {entry['reason']}")
    if duplicate_names:
        print(f"\n*** {len(duplicate_names)} Goal Name(s) repeat within the same Tracker — LPM may treat them as one on import ***")
        for d in duplicate_names:
            print(f"  Tracker '{d['tracker']}' | Goal Name '{d['ptg_name']}' | source rows: {d['row_nums']}")
    print("=" * 65)


if __name__ == "__main__":
    main()
