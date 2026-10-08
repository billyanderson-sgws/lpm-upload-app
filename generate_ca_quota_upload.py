"""
generate_ca_quota_upload.py

Reads California Union Quota Input Form .xlsb files and generates LPM-format
CSV files ready for upload to LiquidPerform. Fork of generate_lpm_upload.py
(the SPP generator) -- reuses its CSV schema, month/unsold-date helpers, and
collection-lookup helpers.

The form has no fixed sheet name -- real filled-out forms rename the data tab
per team/region/month ("Jun Off", "SCA Off", "Full Book BDM ", "Quota", ...).
Every sheet in the workbook is scanned; a sheet is treated as a quota tab iff
it has a "Quota Type" / "Tracked Period/s" label pair followed by the known
"#, Supplier, Product Description, ... Goal Type, Goal, ..." header row.

PTG grouping follows the same ruleset as SPP/NY SOD: one Tracker per
(org-unit, goal_type, goal_uom, start, end, unsold window); every row that
matches on all of those becomes a PTG under that Tracker. "Sheet/tab name"
stands in for NY SOD's Division column (CA has no per-row Division). SPP
also splits on SPP Tier (Anchor/Flex) -- that dimension just drops out here
(same as NY SOD) since every CA Union Quota tracker is the same "class"
(QUO), there's no Tier to split on. basis_flag lives on the PTG (derived per
row from the Goal Type criteria column), not the Tracker, exactly like NY
SOD -- so, also like NY SOD, two rows with different Goal Type criteria can
still land in the same Tracker.

Usage:
    python generate_ca_quota_upload.py <quota_form.xlsb> [output_csv] [collection_report.xlsx]

Open assumptions (confirm/adjust before relying on this in production):
  - program_class is the literal string "QUO" for every tracker -- same
    unconfirmed assumption generate_sod_upload.py makes for NY SOD.
  - product_collection_id is left blank. "Product Codes" is free text
    ("...- Item", "...- Sub Group(s)", "...- Super Group", or just bare
    codes) and isn't parsed into a Product Collection ID.
  - pod_attribute defaults to "ProductId" for every POD/NewPOD row -- same
    unresolved gap as NY SOD; the "Product Codes" free text occasionally
    hints at Item vs Sub Group vs Super Group but isn't reliable enough.
  - salesforce_collection_ids is looked up by *sheet/tab name* (the closest
    thing this form has to NY's per-row Division column). Confirm/override
    per tab in the Streamlit UI before generating -- same human-in-the-loop
    step the SPP/NY SOD tabs already use, since tab names (e.g. "Full Book
    BDM") rarely match a Salesforce collection name exactly.
  - A blank "Tracked Period/s" cell on a tab skips every row on that tab
    (not guessed from another tab in the same file).
  - "Goal minimum" maps straight to min_objective_target and "POD/ACS Min
    Quantity" straight to achievement_min -- unlike NY SOD, this form gives
    real numeric values for both instead of just a per-criteria flag.
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
STATE = "CA"

CA_CONFIG = {
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

# "Quota Type" column (abbreviated) -> LPM goal_type
QUOTA_TYPE_MAP = {
    "VOL":    "VolumeCases",
    "POD":    "DistributionPODs",
    "ACS":    "DistributionACS",
    "NEWPOD": "DistributionNewPODs",
    "NEWACS": "DistributionNewACS",
}

# "Unit of Measure" column -> LPM goal_uom, for VOL rows
VOL_UOM_MAP = {
    "9L":      "NineLiter",
    "DECIMAL": "Cases",
    "STD":     "STD",
}

# "Goal Type" column values that trigger basis_flag = TRUE
BASIS_TRUE_CRITERIA = {"% INCREASE/DECREASE"}

# Required header tokens (normalized) that must appear in a sheet's header
# row for it to be treated as a quota data tab.
REQUIRED_HEADER_TOKENS = {"supplier", "goal type", "goal", "premise"}


# ---------------------------------------------------------------------------
# Header / sheet detection
# ---------------------------------------------------------------------------

def normalize_header(val):
    s = spp.safe_str(val)
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def find_quota_sheets(xlsb_path):
    """
    Return a list of (sheet_name, header_row_idx, quota_type_label,
    tracked_period_raw) for every sheet that looks like a quota data tab.
    """
    import pandas as pd
    xl = pd.ExcelFile(xlsb_path, engine="pyxlsb")
    found = []
    for sheet_name in xl.sheet_names:
        raw = pd.read_excel(xlsb_path, sheet_name=sheet_name, engine="pyxlsb", header=None, nrows=8)
        header_row_idx = None
        for i in range(len(raw)):
            tokens = {normalize_header(v) for v in raw.iloc[i].tolist()}
            if REQUIRED_HEADER_TOKENS <= tokens:
                header_row_idx = i
                break
        if header_row_idx is None or header_row_idx < 1:
            continue

        label_row = raw.iloc[header_row_idx - 1]
        marker_row = raw.iloc[header_row_idx - 2] if header_row_idx >= 2 else None
        marker_tokens = set(normalize_header(marker_row.iloc[1]).split()) if marker_row is not None else set()
        if not {"quota", "type"} <= marker_tokens:
            continue

        quota_type_label = spp.safe_str(label_row.iloc[1])
        tracked_period_raw = spp.safe_str(label_row.iloc[2]) if len(label_row) > 2 else ""
        found.append((sheet_name, header_row_idx, quota_type_label, tracked_period_raw))
    return found


def parse_tracked_period(text):
    """'OCTOBER 2026' -> '202610'. Returns '' if unparseable."""
    m = re.search(r"([A-Za-z]+)\D+(\d{4})", text or "")
    if not m:
        return ""
    month = spp.MONTH_NAME_TO_NUM.get(m.group(1).strip().lower())
    if not month:
        return ""
    return f"{m.group(2)}{month:02d}"


# ---------------------------------------------------------------------------
# Load a quota tab
# ---------------------------------------------------------------------------

def load_quota_tab(xlsb_path, sheet_name, header_row_idx, tracked_period_raw):
    """
    Read one quota data tab and return (records, skipped) for that tab.
    """
    import pandas as pd

    goal_yyyymm = parse_tracked_period(tracked_period_raw)
    if not goal_yyyymm:
        return [], [{
            "row_num": header_row_idx + 1,
            "reason": f"tab '{sheet_name}': blank/unparseable Tracked Period/s ({tracked_period_raw!r}) -- all rows skipped",
            "raw_row": [],
        }]

    df = pd.read_excel(xlsb_path, sheet_name=sheet_name, engine="pyxlsb", header=header_row_idx)
    df.columns = [normalize_header(c) for c in df.columns]

    records = []
    skipped = []

    for i, row in df.iterrows():
        excel_row = i + header_row_idx + 2  # 1-indexed; data starts right after the header row

        def gv(col):
            v = row.get(col)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return None
            return v

        supplier = spp.safe_str(gv("supplier"))
        product_desc = spp.safe_str(gv("product description"))
        quota_description = spp.safe_str(gv("quota description"))
        quota_type_raw = spp.safe_str(gv("quota type")).upper().replace(" ", "")

        if not supplier and not product_desc and not quota_type_raw:
            continue  # blank row

        raw_row = [supplier, product_desc, quota_type_raw, spp.safe_str(gv("goal type")), spp.safe_str(gv("goal"))]

        goal_type = QUOTA_TYPE_MAP.get(quota_type_raw)
        if not goal_type:
            skipped.append({"row_num": excel_row, "reason": f"unknown Quota Type: {quota_type_raw!r}", "raw_row": raw_row})
            continue

        goal_value = spp._numeric_str(gv("goal"))
        if not goal_value:
            skipped.append({"row_num": excel_row, "reason": "no Goal value", "raw_row": raw_row})
            continue

        # goal_uom always comes from "Unit of Measure" (9L/Decimal/STD), for
        # every goal_type. "Cases/ Bottles" is a separate column describing
        # the unit of "POD/ACS Min Quantity" (-> achievement_min) only -- it
        # is not the tracker's overall goal_uom, even on distribution rows.
        uom_raw = spp.safe_str(gv("unit of measure")).upper()
        goal_uom = VOL_UOM_MAP.get(uom_raw, "")

        goal_criteria = spp.safe_str(gv("goal type")).upper()
        basis_flag = "TRUE" if goal_criteria in BASIS_TRUE_CRITERIA else "FALSE"

        min_objective_target = spp._numeric_str(gv("goal minimum"))

        achievement_min = ""
        if goal_type in spp.DISTRIBUTION_TYPES:
            achievement_min = spp._numeric_str(gv("pod/acs min quantity"))

        pod_attribute = "ProductId" if goal_type in spp.POD_ATTR_TYPES else ""

        unsold_start_yyyymm, unsold_end_yyyymm = "", ""
        if goal_type in spp.UNSOLD_TYPES:
            unsold_period_raw = spp.safe_str(gv("unsold period"))
            if unsold_period_raw:
                unsold_start_yyyymm, unsold_end_yyyymm = spp.compute_unsold_dates(unsold_period_raw, goal_yyyymm)

        premise = spp.safe_str(gv("premise")).upper()

        name = quota_description or product_desc or supplier
        if not name:
            skipped.append({"row_num": excel_row, "reason": "no Product Description/Quota Description/Supplier to name the goal", "raw_row": raw_row})
            continue

        records.append({
            "row_num":              excel_row,
            "premise":              premise,
            "sheet_name":           sheet_name,
            "name":                 name[:256],
            "goal_type":            goal_type,
            "goal_uom":             goal_uom,
            "goal_start_yyyymm":    goal_yyyymm,
            "goal_end_yyyymm":      goal_yyyymm,
            "basis_flag":           basis_flag,
            "goal_value":           goal_value,
            "min_objective_target": min_objective_target,
            "achievement_min":      achievement_min,
            "pod_attribute":        pod_attribute,
            "unsold_start_yyyymm":  unsold_start_yyyymm,
            "unsold_end_yyyymm":    unsold_end_yyyymm,
        })

    return records, skipped


def load_workbook(xlsb_path):
    """Scan every sheet and return (records, skipped, tab_names)."""
    quota_sheets = find_quota_sheets(xlsb_path)
    all_records, all_skipped = [], []
    tab_names = []
    for sheet_name, header_row_idx, _quota_type_label, tracked_period_raw in quota_sheets:
        tab_names.append(sheet_name)
        records, skipped = load_quota_tab(xlsb_path, sheet_name, header_row_idx, tracked_period_raw)
        all_records.extend(records)
        all_skipped.extend(skipped)
    return all_records, all_skipped, tab_names


def extract_goal_groups(xlsb_path):
    """Tab names found in the workbook that look like quota data tabs -- used
    by the Streamlit UI to drive the Collection ID mapping, one entry per tab."""
    return [s[0] for s in find_quota_sheets(xlsb_path)]


# ---------------------------------------------------------------------------
# Group + build rows
#
# Same ruleset as the SPP/NY SOD generators: a Tracker is the (org-unit,
# goal_type, goal_uom, start, end, unsold window) combination; every row that
# matches on all of those becomes a PTG under that one Tracker. SPP also
# splits on SPP Tier (Anchor/Flex) -- that dimension drops out here (and in
# NY SOD) because every CA Union Quota tracker is the same "class" (QUO).
# "Sheet/tab name" stands in for NY SOD's Division column -- CA has no
# per-row Division, so the tab is the closest equivalent org-unit.
# basis_flag lives on the PTG (per-row, from the Goal Type criteria column),
# not the Tracker, exactly like NY SOD -- so it correctly does NOT gate
# grouping, since two rows can share a Tracker with different Goal Type
# criteria (Fixed vs % Increase/Decrease) the same way NY SOD allows it.
# Premise (ALL/ON/OFF) has no LPM CSV field of its own, but still gates
# grouping -- an ON-only row and an OFF-only row must not share a Tracker
# just because everything else matches (same fix applied to SPP's
# "Applicable Premise" and NY SOD's "Premise - On or Off Only" columns).
# ---------------------------------------------------------------------------

def group_key(rec):
    return (
        rec["sheet_name"], rec["goal_type"], rec["goal_uom"],
        rec["goal_start_yyyymm"], rec["goal_end_yyyymm"],
        rec["unsold_start_yyyymm"], rec["unsold_end_yyyymm"],
        rec["premise"],
    )


def group_records(records):
    order = []
    groups = defaultdict(list)
    for rec in records:
        key = group_key(rec)
        if key not in groups:
            order.append(key)
        groups[key].append(rec)
    return order, groups


def build_tracker_row(key, recs):
    sheet_name, goal_type, goal_uom, start, end, unsold_start, unsold_end, _premise = key
    return {
        "goal_category":                  "Tracker",
        "goal_name":                      sheet_name,
        "tpm_nav_ref":                    "",
        "goal_type":                      goal_type,
        "goal_description":               "",
        "goal_start_date":                start,
        "goal_end_date":                  end,
        "basis_flag":                     "",  # set per-PTG below, same as NY SOD
        "score_by_td_linx_customer_code": CA_CONFIG["score_by_td_linx_customer_code"],
        "attainment_org_level":           CA_CONFIG["attainment_org_level"],
        "goal_uom":                       goal_uom,
        "recalculate_until_date":         CA_CONFIG["recalculate_until_date"],
        "exclude_from_os_sales_reports":  CA_CONFIG["exclude_from_os_sales_reports"],
        "salesforce_collection_ids":      spp.COLLECTION_LOOKUP.get(sheet_name.lower(), CA_CONFIG["salesforce_collection_ids"]),
        "program_class":                  CA_CONFIG["program_class"],
        "send_to_proof":                  CA_CONFIG["send_to_proof"],
        "send_to_proof_date":             CA_CONFIG["send_to_proof_date"],
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
    return {
        "goal_category":                  "PTG",
        "goal_name":                      rec["name"],
        "tpm_nav_ref":                    "",
        "goal_type":                      "",
        "goal_description":               "",
        "goal_start_date":                "",
        "goal_end_date":                  "",
        "basis_flag":                     rec["basis_flag"],
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
        "min_objective_target":           rec["min_objective_target"],
        "distribution_level_path":        CA_CONFIG["distribution_level_path"],
        "pod_attribute":                  rec["pod_attribute"],
        "achievement_min":                rec["achievement_min"],
    }


def find_duplicate_ptg_names(order, groups):
    """Flag goal names that repeat within the same Tracker -- LPM keys a PTG
    by goal_name within its Tracker, so two rows sharing both a group key and
    a name would collide on import even though this script emits them as
    separate lines. Same check as NY SOD's."""
    dupes = []
    for key in order:
        recs = groups[key]
        by_name = defaultdict(list)
        for r in recs:
            by_name[r["name"]].append(r["row_num"])
        sheet_name = key[0]
        for name, row_nums in by_name.items():
            if len(row_nums) > 1:
                dupes.append({"tracker": sheet_name, "ptg_name": name, "row_nums": row_nums})
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
            writer.writerow(build_tracker_row(key, recs))
            for rec in recs:
                writer.writerow(build_ptg_row(rec))
            total_trackers += 1
            total_ptgs += len(recs)
    return total_trackers, total_ptgs


RAW_COLUMNS = ["supplier", "product_description", "quota_type", "goal_type", "goal"]


def write_skipped_csv(skipped_path, skipped):
    import csv
    with open(skipped_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(["skip_reason"] + RAW_COLUMNS)
        for entry in skipped:
            writer.writerow([entry["reason"]] + entry["raw_row"])


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_ca_quota_upload.py <quota_form.xlsb> [output.csv] [collection_report.xlsx]", file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base = os.path.splitext(os.path.abspath(input_path))[0]
        output_path = base + "_ca_quota_upload.csv"

    collection_path = sys.argv[3] if len(sys.argv) >= 4 else None
    if not collection_path:
        folder = os.path.dirname(os.path.abspath(input_path))
        import glob as _glob
        candidates = _glob.glob(os.path.join(folder, "*Collection*Report*.xlsx"))
        if candidates:
            collection_path = candidates[0]

    if collection_path and os.path.isfile(collection_path):
        spp.COLLECTION_LOOKUP = spp.load_collection_lookup(collection_path, STATE)
        print(f"Collection IDs loaded: {len(spp.COLLECTION_LOOKUP)}  ({os.path.basename(collection_path)})")
    else:
        print("No collection report found -- salesforce_collection_ids will be blank.")

    print(f"Reading: {input_path}")
    records, skipped, tab_names = load_workbook(input_path)
    print(f"Quota tabs found: {tab_names}")

    if not records:
        print("ERROR: No valid records found in any quota tab after filtering.", file=sys.stderr)
        if skipped:
            for entry in skipped:
                print(f"  row {entry['row_num']}: {entry['reason']}")
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
    print("CA Union Quota Upload Generation Summary")
    print("=" * 65)
    print(f"Output file : {output_path}")
    print(f"Trackers    : {total_trackers}")
    print(f"PTGs        : {total_ptgs}")
    if skipped:
        print(f"\n*** {len(skipped)} row(s)/tab(s) skipped -- saved to {skipped_path} ***")
        for entry in skipped:
            print(f"  row {entry['row_num']}: {entry['reason']}")
    if duplicate_names:
        print(f"\n*** {len(duplicate_names)} goal name(s) repeat within the same tab ***")
        for d in duplicate_names:
            print(f"  Tab '{d['tracker']}' | Goal Name '{d['ptg_name']}' | source rows: {d['row_nums']}")
    print("=" * 65)


if __name__ == "__main__":
    main()
