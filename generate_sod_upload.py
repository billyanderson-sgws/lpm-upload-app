"""
generate_sod_upload.py

Reads the "GOAL SHEET" table from a NY Monthly Quota Planner .xlsb file and
generates LPM-format CSV files ready for upload to LiquidPerform ("NY
Uploader" -- handles both NY Metro Quota and NY Upstate SPP divisions from
one file). Fork of generate_lpm_upload.py (the SPP generator) — reuses its
CSV schema, POD attribute map, and collection-lookup helpers, with NY
Uploader-specific column mapping and date resolution.

Usage:
    python generate_sod_upload.py <goal_sheet.xlsb> [output_csv] [collection_report.xlsx]

One Tracker is generated per unique (Division, goal_type, goal_uom,
Tracking Start/End, unsold-period, Premise, Must Make) combination; every
GOAL SHEET row under that combination becomes a PTG row.

Every Division must start with "NYM" or "NYU" (division_group()) — rows
under anything else are skipped and flagged. The two groups are handled
differently end to end:
  - NYM (NY Metro) rows go into LPM as Quotas: program_class is always the
    flat enum path "ProgramClass.Site.Quo".
  - NYU (NY Upstate) rows go into LPM as SPP: program_class is derived from
    the Must Make column via spp.PROGRAM_CLASS_MAP (Anchor -> ...Sppa, Flex
    -> ...Sppf), the same convention SPP itself uses. An NYU row whose Must
    Make isn't "Anchor" or "Flex" is skipped and flagged rather than
    guessing. Must Make already gates Tracker grouping (see group_key), so
    Anchor/Flex trackers are kept separate automatically.
  - Exception: NYU divisions starting with "NYU SIG" (e.g. "NYU SIG
    On/Off") are themselves plain Quotas, not SPP -- program_class is the
    same "ProgramClass.Site.Quo" literal as NYM, and Must Make on these
    rows is just a mandatory-goal flag (blank or the literal "Must Make"),
    never an Anchor/Flex selector. It still gates Tracker grouping like any
    other NYU division. See is_nyu_quota_division().
  - Both groups now set send_to_proof to "SendStartingOnCalendarDate" (not
    SPP's "SendStartingInActive") and resolve send_to_proof_date from
    SEND_TO_PROOF_SCHEDULE, keyed by (anchor_year, anchor_month) -> group.
    A month with no schedule entry leaves send_to_proof_date blank.
  - "GSC" (Go Spot Check) is NY's SET alternative that lives outside LPM —
    reps manually update tracking, so NY builds a "dummy" ACS tracker in LPM
    just to hold the goal. A GSC row maps to the same DistributionACS
    goal_type as a normal ACS row, but: the tracking period is forced to a
    single month matching the file's close month (the Tracking Start/End
    columns are ignored), the Tracker's goal_name gets a " (GSC)" suffix,
    and the PTG's achievement_min is forced to "1" (a 1-bottle qualifier)
    regardless of goal_type. Everything else — PTG naming, distribution_target
    from Goal Calculation, etc. — follows normal ACS rules. GSC rows never
    group into the same Tracker as a real ACS row (see group_key's is_gsc).

Open assumptions (confirm/adjust before relying on this in production):
  - "Reverse MS" (Goal of Unsold) is treated like "Market Share" (Unit Goal)
    for distribution purposes.
  - basis_flag is always "TRUE" on the Tracker row (PTG rows leave it blank),
    matching how the SPP generator sets it — a constant, not derived from
    Goal Criteria.
  - distribution_target comes from the "Goal Calculation" column (the
    per-rep/per-unit figure), not "Goal" (the aggregate total). "Goal
    Calculation" is also the required completeness check — a row with no
    Goal Calculation value is skipped. The "Goal" column is not read.
  - min_objective_target is always "1" on the PTG row. achievement_min is
    only populated for distribution objectives (goal_type in
    spp.POD_ATTR_TYPES) and carries just the numeric value from the "Min"
    column with its unit suffix/whitespace stripped (e.g. "3 btl" -> "3")
    — blank for non-distro objectives.
  - product_collection_id is left blank — no equivalent source column has
    been identified yet in the GOAL SHEET.
  - Unsold End is always force-corrected to the month immediately before
    Tracking Start — the unsold/basis period must never overlap the
    tracking period, and submissions have been seen with Unsold End
    mistakenly typed as the Tracking Start month. If Unsold Start itself is
    on/after Tracking Start (the whole window overlaps), Unsold Start is
    also force-corrected to that same month-before value, collapsing it to
    a single-month window alongside End — not skipped, but called out in
    unsold_warnings since it's a more significant source error worth the
    state double-checking.
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
    "program_class":                  "ProgramClass.Site.Quo",  # NYM (Quota) only -- NYU (SPP) uses spp.PROGRAM_CLASS_MAP via Must Make
    "score_by_td_linx_customer_code": "",
    "attainment_org_level":           "Salesperson",
    "recalculate_until_date":         "0",
    "exclude_from_os_sales_reports":  "FALSE",
    "salesforce_collection_ids":      "",
    "send_to_proof":                  "SendStartingOnCalendarDate",
    "send_to_proof_date":             "",
    "distribution_level_path":        "Salesperson",
}

# Division prefix -> which half of the NY Uploader a row belongs to:
# "NYM" (NY Metro) rows go in as Quotas, "NYU" (NY Upstate) rows go in as
# SPP, following SPP's own Anchor/Flex convention via the Must Make column.
DIVISION_GROUPS = ("NYM", "NYU")

# send_to_proof_date by (anchor_year, anchor_month) -> {division group: MM/DD/YYYY}.
# NYM and NYU follow different send-to-proof schedules some months and the
# same date other months -- add each month here as it's confirmed; a month
# with no entry here leaves send_to_proof_date blank rather than guessing.
SEND_TO_PROOF_SCHEDULE = {
    (2026, 11): {"NYM": "11/5/2026", "NYU": "11/4/2026"},
    (2026, 12): {"NYM": "12/3/2026", "NYU": "12/3/2026"},
}


def division_group(division):
    """Return "NYM" or "NYU" based on the Division prefix, or None if neither."""
    d = (division or "").strip().upper()
    for group in DIVISION_GROUPS:
        if d.startswith(group):
            return group
    return None


# NYU divisions that are plain Quotas despite being NYU -- their Must Make
# column is just a mandatory-goal flag, not an Anchor/Flex selector.
NYU_QUOTA_PREFIXES = ("NYU SIG",)


def is_nyu_quota_division(division):
    d = (division or "").strip().upper()
    return any(d.startswith(p) for p in NYU_QUOTA_PREFIXES)


def send_to_proof_date_for(anchor_year, anchor_month, group):
    sched = SEND_TO_PROOF_SCHEDULE.get((anchor_year, anchor_month))
    return sched.get(group, "") if sched else ""

# GOAL SHEET "Goal Type" -> LPM goal_type (before unsold-period coercion)
# "GSC" (Go Spot Check) is NY's SET alternative that lives outside LPM --
# these are "dummy" ACS goals tracked manually, so they map onto the same
# DistributionACS path as a normal ACS row. GSC_GOAL_TYPES below flags which
# raw values additionally get the GSC-only treatment (single-month tracking
# period forced to the close month, " (GSC)" tracker-name suffix, forced
# 1-bottle achievement_min) -- see is_gsc handling in load_goal_sheet.
GOAL_TYPE_MAP = {
    "Volume":        "VolumeCases",
    "DREV":          "VolumeRevenueDREV",
    "ACS":           "DistributionACS",
    "GSC":           "DistributionACS",
    "PODs by SKU":   "DistributionPODs",
    "PODs by Type":  "DistributionPODs",
    "PODs by Size":  "DistributionPODs",
}

GSC_GOAL_TYPES = {"gsc"}  # normalized (lowercase/stripped) raw Goal Type values treated as GSC

# "PODs by X" -> pod_attribute
POD_TYPE_ATTR_MAP = {
    "PODs by SKU":  "ProductId",
    "PODs by Type": "PhSubGroup",
    "PODs by Size": "ProductSize",
}

def _norm_goal_type(s):
    return s.strip().lower()

GOAL_TYPE_MAP_NORM = {_norm_goal_type(k): v for k, v in GOAL_TYPE_MAP.items()}
POD_TYPE_ATTR_MAP_NORM = {_norm_goal_type(k): v for k, v in POD_TYPE_ATTR_MAP.items()}

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


def strip_unit(value):
    """
    Pull just the leading numeric value out of a "Min" cell like "3 btl" or
    "1 cs" -- unit suffix and surrounding whitespace dropped, e.g. "3 btl" -> "3".
    """
    if not value:
        return ""
    m = re.match(r"\s*(-?\d+(?:\.\d+)?)", str(value))
    return m.group(1) if m else str(value).strip()


def month_before(yyyymm):
    """Return the YYYYMM immediately preceding the given YYYYMM."""
    year, month = int(yyyymm[:4]), int(yyyymm[4:6])
    if month == 1:
        return f"{year - 1}12"
    return f"{year}{month - 1:02d}"


# ---------------------------------------------------------------------------
# Load GOAL SHEET
# ---------------------------------------------------------------------------

def load_goal_sheet(xlsb_path):
    """
    Read the GOAL SHEET table and return (records, skipped_rows,
    unsold_warnings, anchor_year, anchor_month).
    """
    import pandas as pd

    raw = pd.read_excel(xlsb_path, sheet_name="GOAL SHEET", engine="pyxlsb", header=None, nrows=1)
    goal_month_abbr = raw.iat[0, 0] if not raw.empty else None
    anchor_month = safe_month(goal_month_abbr)
    anchor_year = derive_anchor_year(xlsb_path, goal_month_abbr)

    df = pd.read_excel(xlsb_path, sheet_name="GOAL SHEET", engine="pyxlsb", header=4)

    records = []
    skipped = []
    unsold_warnings = []

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
        min_value = strip_unit(gv("Min"))

        raw_row = [gv(c) for c in GOAL_SHEET_COLUMNS]

        # Skip blank rows
        if not goal_name and not division:
            continue

        group = division_group(division)
        if not group:
            skipped.append({
                "row_num": excel_row,
                "reason": f"Division {division!r} doesn't start with NYM or NYU — unrecognized group, needs review",
                "raw_row": raw_row,
            })
            continue

        if group == "NYU" and not is_nyu_quota_division(division) and must_make not in spp.PROGRAM_CLASS_MAP:
            skipped.append({
                "row_num": excel_row,
                "reason": f"NYU (SPP) row has no valid Anchor/Flex in Must Make: {must_make!r}",
                "raw_row": raw_row,
            })
            continue

        goal_type = GOAL_TYPE_MAP_NORM.get(_norm_goal_type(goal_type_raw))
        if not goal_type:
            skipped.append({"row_num": excel_row, "reason": f"unknown Goal Type: {goal_type_raw!r}", "raw_row": raw_row})
            continue

        unsold_start_raw = gv("Unsold Start")
        unsold_end_raw   = gv("Unsold End")
        unsold_present = bool(safe_month(unsold_start_raw)) and bool(safe_month(unsold_end_raw))

        if unsold_present and goal_type in UNSOLD_ELIGIBLE:
            goal_type = UNSOLD_COERCION[goal_type]

        pod_attribute = POD_TYPE_ATTR_MAP_NORM.get(_norm_goal_type(goal_type_raw), "")

        is_gsc = _norm_goal_type(goal_type_raw) in GSC_GOAL_TYPES

        if is_gsc:
            # GSC is tracked manually outside LPM -- the dummy ACS tracker is
            # always a single month matching the file's close month, not
            # whatever (if anything) is in the Tracking Start/End columns.
            if not anchor_month:
                skipped.append({
                    "row_num": excel_row,
                    "reason": "GSC row requires a resolvable close month (Goal Month cell) but none was found",
                    "raw_row": raw_row,
                })
                continue
            close_yyyymm = f"{anchor_year}{anchor_month:02d}"
            start_yyyymm, end_yyyymm = close_yyyymm, close_yyyymm
        else:
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
            # The unsold/basis period is the comparison window used to judge
            # "new" distribution and must never overlap the tracking period --
            # it has to end the month immediately before Tracking Start (e.g.
            # a Sep-Oct tracker's unsold period must end in August). The GOAL
            # SHEET sometimes has Unsold End mistakenly typed as the Tracking
            # Start month instead, so it's always force-corrected here rather
            # than trusting what was typed. If Unsold Start itself is on/after
            # Tracking Start, the whole unsold window was entered as (or
            # inside) the tracking period -- Start is force-corrected to the
            # same month-before-Tracking-Start value as End (collapsing to a
            # single-month window) rather than left overlapping, and it's
            # called out as a warning since it's a more significant source
            # error worth the state double-checking.
            expected_unsold_end = month_before(start_yyyymm)
            if unsold_start_yyyymm > expected_unsold_end:
                unsold_warnings.append({
                    "row_num": excel_row,
                    "reason": (
                        f"Unsold period matches/overlaps the Tracking period "
                        f"entirely: Unsold Start resolved to "
                        f"{unsold_start_yyyymm}, which is on or after Tracking "
                        f"Start ({start_yyyymm}) -- Unsold Start and End both "
                        f"forced to {expected_unsold_end}, verify at the source"
                    ),
                    "raw_row": raw_row,
                })
                unsold_start_yyyymm = expected_unsold_end
            unsold_end_yyyymm = expected_unsold_end

        measure = spp.safe_str(gv("Measure Decimal/9L/STD"))
        if goal_type in {"VolumeRevenueDREV"}:
            goal_uom = ""
        else:
            goal_uom = UOM_MAP.get(measure, "Cases" if goal_type in spp.DISTRIBUTION_TYPES else "")

        goal_calc_raw = gv("Goal Calculation")
        if isinstance(goal_calc_raw, float):
            goal_calc_raw = round(goal_calc_raw, 4)
        goal_calculation = spp._numeric_str(goal_calc_raw)
        if not goal_calculation:
            skipped.append({"row_num": excel_row, "reason": "no Goal Calculation value", "raw_row": raw_row})
            continue

        goal_value = goal_calculation

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
            "goal_calculation":     goal_calculation,
            "pod_attribute":        pod_attribute,
            "premise":              premise,
            "must_make":            must_make,
            "min_value":            min_value,
            "is_gsc":               is_gsc,
        })

    return records, skipped, unsold_warnings, anchor_year, anchor_month


# ---------------------------------------------------------------------------
# Group + build rows
# ---------------------------------------------------------------------------

def group_key(rec):
    # Key on the actual resolved unsold window, not just presence — two rows
    # with the same goal_type/tracking period but different unsold periods
    # (e.g. JUL-AUG vs AUG-AUG) must land in separate Trackers. Premise (On
    # premise/Off premise) has no LPM CSV field of its own and does not gate
    # grouping -- an On-only row and an Off-only row can share a Tracker.
    # Must Make (and NYU's Anchor/Flex, which share the same GOAL SHEET
    # column) still gates grouping -- a Must Make row and a regular row must
    # not land in the same Tracker just because every other field matches,
    # mirroring how SPP's spp_tier (Anchor/Flex) is part of its own group_key.
    # is_gsc is part of the key too -- a GSC "dummy" tracker must never merge
    # with a real ACS tracker just because division/period/etc happen to
    # match, since it gets its own tracker-name suffix and achievement_min.
    return (
        rec["division"], rec["goal_type"], rec["goal_uom"],
        rec["start_yyyymm"], rec["end_yyyymm"],
        rec["unsold_start_yyyymm"], rec["unsold_end_yyyymm"],
        rec["must_make"], rec["is_gsc"],
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


def build_tracker_row(key, recs, anchor_year=None, anchor_month=None):
    division, goal_type, goal_uom, start, end, unsold_start, unsold_end, must_make, is_gsc = key
    group = division_group(division)
    if group == "NYU" and not is_nyu_quota_division(division):
        program_class = spp.PROGRAM_CLASS_MAP.get(must_make, "")
    else:
        program_class = SOD_CONFIG["program_class"]
    send_to_proof_date = send_to_proof_date_for(anchor_year, anchor_month, group)
    return {
        "goal_category":                  "Tracker",
        "goal_name":                      division + (" (GSC)" if is_gsc else ""),
        "tpm_nav_ref":                    "",
        "goal_type":                      goal_type,
        "goal_description":               "",
        "goal_start_date":                start,
        "goal_end_date":                  end,
        "basis_flag":                     "TRUE",
        "score_by_td_linx_customer_code": SOD_CONFIG["score_by_td_linx_customer_code"],
        "attainment_org_level":           SOD_CONFIG["attainment_org_level"],
        "goal_uom":                       goal_uom,
        "recalculate_until_date":         SOD_CONFIG["recalculate_until_date"],
        "exclude_from_os_sales_reports":  SOD_CONFIG["exclude_from_os_sales_reports"],
        "salesforce_collection_ids":      spp.COLLECTION_LOOKUP.get(division.lower(), SOD_CONFIG["salesforce_collection_ids"]),
        "program_class":                  program_class,
        "send_to_proof":                  SOD_CONFIG["send_to_proof"],
        "send_to_proof_date":             spp.to_yyyymmdd(send_to_proof_date),
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
    # GSC's dummy ACS tracker always gets a 1-bottle achievement_min
    # qualifier, regardless of whether the goal_type is a POD-attribute type.
    if rec["is_gsc"]:
        achievement_min = "1"
    else:
        achievement_min = rec["min_value"] if is_dist else ""
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
        "distribution_target":            rec["goal_calculation"],
        "min_objective_target":           "1",
        "distribution_level_path":        SOD_CONFIG["distribution_level_path"],
        "pod_attribute":                  (rec["pod_attribute"] or "ProductId") if is_dist else "",
        "achievement_min":                achievement_min,
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


def generate_output(order, groups, output_path, anchor_year=None, anchor_month=None):
    import csv
    total_trackers = 0
    total_ptgs = 0
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        for key in order:
            recs = groups[key]
            tracker_row = build_tracker_row(key, recs, anchor_year, anchor_month)
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
    records, skipped, unsold_warnings, anchor_year, anchor_month = load_goal_sheet(input_path)
    print(f"Anchor: year={anchor_year} month={anchor_month}")

    if not records:
        print("ERROR: No valid records found in GOAL SHEET after filtering.", file=sys.stderr)
        sys.exit(1)

    order, groups = group_records(records)
    base = os.path.splitext(output_path)[0]
    skipped_path = base + "_skipped.csv"
    total_trackers, total_ptgs = generate_output(order, groups, output_path, anchor_year, anchor_month)
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
    if unsold_warnings:
        print(f"\n*** {len(unsold_warnings)} row(s) had an Unsold period overlapping Tracking Start — Unsold End was auto-corrected, verify at the source ***")
        for w in unsold_warnings:
            print(f"  row {w['row_num']}: {w['reason']}")
    print("=" * 65)


if __name__ == "__main__":
    main()
