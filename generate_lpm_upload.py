"""
generate_lpm_upload.py

Reads the "Tracking Table" sheet from a Goal Builder .xlsm file and generates
LPM-format CSV files ready for upload to LiquidPerform.

Usage:
    python generate_lpm_upload.py <goal_builder.xlsm> [output_dir]

One CSV is generated per tracker group. Trackers are grouped by:
    Goal Group + SPP Tier + Objective Type + Program Period (Start/End) + Unsold Period
"""

import csv
import os
import sys
from collections import defaultdict
from datetime import date, datetime

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl is required.  Install with: pip install openpyxl", file=sys.stderr)
    sys.exit(1)

# ---------------------------------------------------------------------------
# Site-wide defaults
# ---------------------------------------------------------------------------
CONFIG = {
    "basis_flag":                     "TRUE",
    "score_by_td_linx_customer_code": "",
    "attainment_org_level":           "Salesperson",
    "recalculate_until_date":         "365",
    "exclude_from_os_sales_reports":  "FALSE",
    "salesforce_collection_ids":      "",
    "send_to_proof":                  "SendStartingInActive",
    "send_to_proof_date":             "",
    "distribution_level_path":        "Salesperson",
}

# Tracking Table "Objective Type" (col 4) -> LPM goal_type
OBJECTIVE_TYPE_MAP = {
    "Volume (Cases)": "VolumeCases",
    "New POD":        "DistributionNewPODs",
    "POD":            "DistributionACS",
    "New ACS":        "DistributionNewACS",
    "ACS":            "DistributionACS",
    "Revenue":        "VolumeRevenue",
    "DREV":           "VolumeRevenueDREV",
}

# Tracker type short labels used in output filenames
TRACKER_TYPE_LABELS = {
    "VolumeCases":         "Vol",
    "DistributionNewPODs": "NPOD",
    "DistributionACS":     "POD",
    "DistributionNewACS":  "NACS",
}

DISTRIBUTION_TYPES = {"DistributionNewPODs", "DistributionACS", "DistributionNewACS"}

# Tracking Table "Measure" (col 10) -> LPM goal_uom
UOM_MAP = {
    "9L":      "NineLiter",
    "Cases":   "Cases",
    "Dollars": "Dollars",
    "STD":     "STD",
}

# SPP Tier (col 2) -> program_class
PROGRAM_CLASS_MAP = {
    "Anchor": "ProgramClass.Site.Sppa",
    "Flex":   "ProgramClass.Site.Sppf",
}

# Objective types and measures to skip entirely
SKIP_OBJECTIVE_TYPES = {"SPP My Sales"}
SKIP_MEASURES        = {"DigComRev"}

# LPM CSV column order (must match upload spec)
CSV_COLUMNS = [
    "goal_category", "goal_name", "tpm_nav_ref", "goal_type", "goal_description",
    "goal_start_date", "goal_end_date", "basis_flag", "score_by_td_linx_customer_code",
    "attainment_org_level", "goal_uom", "recalculate_until_date",
    "exclude_from_os_sales_reports", "salesforce_collection_ids", "program_class",
    "send_to_proof", "send_to_proof_date", "unsold_start_date", "unsold_end_date",
    "product_collection_id", "customer_collection_id", "distribution_target",
    "min_objective_target", "distribution_level_path", "pod_attribute", "achievement_min",
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def to_yyyymm(value):
    """Convert a cell value (datetime, int YYYYMM, or string) to a YYYYMM string."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y%m")
    s = str(value).strip()
    if not s:
        return ""
    if len(s) == 6 and s.isdigit():
        return s
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y%m")
        except ValueError:
            pass
    return s


def cytd_unsold_dates(goal_start_yyyymm):
    """
    CYTD: January of the goal_start year through the current calendar month.
    Returns (unsold_start_yyyymm, unsold_end_yyyymm).
    """
    if not goal_start_yyyymm or len(goal_start_yyyymm) < 4:
        return "", ""
    year = goal_start_yyyymm[:4]
    today = date.today()
    return f"{year}01", today.strftime("%Y%m")


def safe_str(val):
    if val is None:
        return ""
    return str(val).strip()


def _numeric_str(val):
    """Return the numeric portion of a value. Extracts leading number from strings like '5 per SC'."""
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    try:
        float(s)
        return s
    except ValueError:
        import re
        m = re.match(r"(\d+\.?\d*)", s)
        return m.group(1) if m else ""


def normalize_pod_attribute(val):
    s = safe_str(val)
    return "" if s == "-" else s


def ptg_name_from_row(ptg_name_col, selection_col, supplier_col):
    """
    PTG name: col 12 preferred; fall back to col 15 (Selection), then col 14 (Supplier).
    Truncate to 256 chars (LPM limit).
    """
    for v in (ptg_name_col, selection_col, supplier_col):
        s = safe_str(v)
        if s:
            return s[:256]
    return ""


def safe_filename(s):
    for ch in r'\/:*?"<>| ':
        s = s.replace(ch, "_")
    return s


# ---------------------------------------------------------------------------
# Load Tracking Table
# ---------------------------------------------------------------------------

def load_tracking_table(wb_path):
    """
    Read the Tracking Table sheet and return (records, skipped_rows).
    Each record is a dict of parsed/mapped fields ready for CSV output.
    """
    wb = openpyxl.load_workbook(wb_path, keep_vba=True, data_only=True)

    if "Tracking Table" not in wb.sheetnames:
        raise ValueError(
            f"No 'Tracking Table' sheet found in {wb_path}.\n"
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb["Tracking Table"]
    # Capture header row for the skipped-rows export
    header_row = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    records = []
    skipped = []  # each entry: {"row_num", "reason", "raw_row"}

    def make_skip(row_num, reason, raw):
        return {"row_num": row_num, "reason": reason, "raw_row": raw}

    for r in range(2, ws.max_row + 1):
        def cv(c):
            return ws.cell(r, c).value

        goal_group  = safe_str(cv(1))
        spp_tier    = safe_str(cv(2))
        goal_bucket = safe_str(cv(3))
        obj_type    = safe_str(cv(4))
        start_raw   = cv(5)
        end_raw     = cv(6)
        # cols 7-8 = Basis Period (not needed; basis_flag=TRUE handles it)
        unsold_prd  = safe_str(cv(9))
        measure     = safe_str(cv(10))
        mkt_seg     = cv(11)
        ptg_name_v  = cv(12)
        level_detail = safe_str(cv(13))
        supplier    = cv(14)
        selection   = cv(15)
        # cols 16-22 not used
        min_cases   = cv(23)
        # col 24 not used
        pod_attr    = cv(25)
        # col 26 not used

        raw_row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]

        # Skip completely blank rows (no logging)
        if not goal_group and ptg_name_v is None and selection is None:
            continue

        # Skip Digital goals (no logging)
        if measure in SKIP_MEASURES:
            continue

        # Skip "Select:" bucket rows with no objective type (no logging)
        if goal_bucket == "Select:" and not obj_type:
            continue

        # Skip SPP My Sales rows (no logging)
        if obj_type in SKIP_OBJECTIVE_TYPES:
            continue

        # Map objective type
        goal_type = OBJECTIVE_TYPE_MAP.get(obj_type)
        if not goal_type:
            skipped.append(make_skip(r, f"unknown objective type: {obj_type!r}", raw_row))
            continue

        # Resolve PTG name — explicit name always wins; fall back to "Total <type>" only if blank
        name = ptg_name_from_row(ptg_name_v, selection, supplier)
        if not name and level_detail == "Total":
            name = f"Total {obj_type}"
        if not name:
            skipped.append(make_skip(r, "no PTG name, selection, or supplier", raw_row))
            continue

        # Map UOM; distribution types default to Cases if measure is blank
        goal_uom = UOM_MAP.get(measure, "")
        if not goal_uom and goal_type in DISTRIBUTION_TYPES:
            goal_uom = "Cases"

        # Dates
        start_yyyymm = to_yyyymm(start_raw)
        end_yyyymm   = to_yyyymm(end_raw) or start_yyyymm

        records.append({
            "row_num":       r,
            "goal_group":    goal_group,
            "spp_tier":      spp_tier,
            "goal_type":     goal_type,
            "goal_uom":      goal_uom,
            "program_class": PROGRAM_CLASS_MAP.get(spp_tier, ""),
            "start_yyyymm":  start_yyyymm,
            "end_yyyymm":    end_yyyymm,
            "unsold_prd":    unsold_prd,
            "ptg_name":      name,
            "mkt_seg_goal":  _numeric_str(mkt_seg),
            "pod_attribute": normalize_pod_attribute(pod_attr),
            "min_cases":     safe_str(min_cases),
        })

    return records, skipped, header_row


# ---------------------------------------------------------------------------
# Group records
# ---------------------------------------------------------------------------

def group_key(rec):
    return (
        rec["goal_group"],
        rec["spp_tier"],
        rec["goal_type"],
        rec["start_yyyymm"],
        rec["end_yyyymm"],
        rec["unsold_prd"],
    )


def group_records(records):
    groups = defaultdict(list)
    order  = []
    for rec in records:
        key = group_key(rec)
        if key not in groups:
            order.append(key)
        groups[key].append(rec)
    return order, groups


# ---------------------------------------------------------------------------
# Build CSV rows
# ---------------------------------------------------------------------------

def build_tracker_row(key, recs):
    goal_group, spp_tier, goal_type, start, end, unsold_prd = key

    # UOM: use most common value across the group's records
    uom_counts = defaultdict(int)
    for r in recs:
        if r["goal_uom"]:
            uom_counts[r["goal_uom"]] += 1
    goal_uom = max(uom_counts, key=uom_counts.get) if uom_counts else ""

    program_class = recs[0]["program_class"]

    # Unsold dates
    unsold_start, unsold_end = "", ""
    if unsold_prd.upper() == "CYTD":
        unsold_start, unsold_end = cytd_unsold_dates(start)

    return {
        "goal_category":                  "Tracker",
        "goal_name":                      goal_group,
        "tpm_nav_ref":                    "",
        "goal_type":                      goal_type,
        "goal_description":               "",
        "goal_start_date":                start,
        "goal_end_date":                  end,
        "basis_flag":                     CONFIG["basis_flag"],
        "score_by_td_linx_customer_code": CONFIG["score_by_td_linx_customer_code"],
        "attainment_org_level":           CONFIG["attainment_org_level"],
        "goal_uom":                       goal_uom,
        "recalculate_until_date":         CONFIG["recalculate_until_date"],
        "exclude_from_os_sales_reports":  CONFIG["exclude_from_os_sales_reports"],
        "salesforce_collection_ids":      CONFIG["salesforce_collection_ids"],
        "program_class":                  program_class,
        "send_to_proof":                  CONFIG["send_to_proof"],
        "send_to_proof_date":             CONFIG["send_to_proof_date"],
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
    is_dist = rec["goal_type"] in DISTRIBUTION_TYPES
    return {
        "goal_category":                  "PTG",
        "goal_name":                      rec["ptg_name"],
        "tpm_nav_ref":                    "",
        "goal_type":                      "",
        "goal_description":               "",
        "goal_start_date":                "",
        "goal_end_date":                  "",
        "basis_flag":                     "",
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
        "distribution_target":            rec["mkt_seg_goal"],
        "min_objective_target":           rec["min_cases"],
        "distribution_level_path":        CONFIG["distribution_level_path"],
        "pod_attribute":                  rec["pod_attribute"],
        "achievement_min":                "1",
    }


# ---------------------------------------------------------------------------
# Write output
# ---------------------------------------------------------------------------

def generate_output(order, groups, output_path):
    total_trackers = 0
    total_ptgs = 0

    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()

        for key in order:
            recs = groups[key]
            tracker_row = build_tracker_row(key, recs)
            ptg_rows    = [build_ptg_row(r) for r in recs]
            writer.writerow(tracker_row)
            for ptg in ptg_rows:
                writer.writerow(ptg)
            total_trackers += 1
            total_ptgs += len(ptg_rows)

    return total_trackers, total_ptgs


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def write_skipped_csv(skipped_path, skipped, header_row):
    """Write skipped rows to a CSV with an extra 'skip_reason' column."""
    with open(skipped_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(["skip_reason"] + header_row)
        for entry in skipped:
            writer.writerow([entry["reason"]] + entry["raw_row"])


def print_summary(output_path, skipped_path, total_trackers, total_ptgs, skipped):
    print()
    print("=" * 65)
    print("LPM Upload Generation Summary")
    print("=" * 65)
    print(f"Output file : {output_path}")
    print(f"Trackers    : {total_trackers}")
    print(f"PTGs        : {total_ptgs}")
    if skipped:
        print(f"\n*** {len(skipped)} row(s) were skipped and need review ***")
        print(f"Skipped rows saved to: {skipped_path}")
        print()
        print(f"  {'Row':<5}  {'Goal Group':<20}  Reason")
        print(f"  {'-'*5}  {'-'*20}  {'-'*40}")
        for entry in skipped:
            print(f"  {entry['row_num']:<5}  {entry['raw_row'][0] or '':<20}  {entry['reason']}")
    else:
        print("\nAll rows processed successfully — no skipped rows.")
    print("=" * 65)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print(
            "Usage: python generate_lpm_upload.py <goal_builder.xlsm> [output_dir]",
            file=sys.stderr,
        )
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Output path: explicit arg, or same folder as input with .csv extension
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base = os.path.splitext(os.path.abspath(input_path))[0]
        output_path = base + "_lpm_upload.csv"

    base = os.path.splitext(output_path)[0]
    skipped_path = base + "_skipped.csv"

    print(f"Reading: {input_path}")
    records, skipped, header_row = load_tracking_table(input_path)

    if not records:
        print("ERROR: No valid records found in Tracking Table after filtering.", file=sys.stderr)
        sys.exit(1)

    order, groups = group_records(records)
    total_trackers, total_ptgs = generate_output(order, groups, output_path)

    if skipped:
        write_skipped_csv(skipped_path, skipped, header_row)

    print_summary(output_path, skipped_path, total_trackers, total_ptgs, skipped)


if __name__ == "__main__":
    main()
