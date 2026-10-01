"""
Streamlit web app for the LPM Upload Generator (SPP + NY SOD).

Run with:
    streamlit run app_lpm_upload.py
"""

import io
import os
import sys
import tempfile
from pathlib import Path

import openpyxl
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_lpm_upload as gen
import generate_sod_upload as sodgen
import ny_sod_revision_recap as recap
import check_spp_template as checker

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
_APP_DIR = Path(__file__).resolve().parent
BUNDLED_COLLECTION = str(_APP_DIR / "LPM Salesforce and Overlay Collection Report.xlsx")
CI_LOGO = _APP_DIR / "assets" / "ci_logo.png"

st.set_page_config(
    page_title="LPM Upload Generator",
    page_icon=str(CI_LOGO) if CI_LOGO.is_file() else "📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def extract_goal_groups(xlsm_bytes):
    """Return unique non-blank goal groups from Tracking Table, in order."""
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as f:
        f.write(xlsm_bytes)
        tmp = f.name
    try:
        wb = openpyxl.load_workbook(tmp, keep_vba=True, data_only=True, read_only=True)
        if "Tracking Table" not in wb.sheetnames:
            return []
        ws = wb["Tracking Table"]
        groups, seen = [], set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            g = str(row[0]).strip() if row[0] else ""
            if g and g not in seen:
                seen.add(g)
                groups.append(g)
        wb.close()
        return groups
    finally:
        os.unlink(tmp)


def get_state_collections(state, source):
    """
    Return {display_name: collection_id} for the given state.
    Delegates to gen._parse_collection_sheet_raw which reads raw XML to avoid
    openpyxl float64 precision loss. Returns original-cased display names.
    source: file path string or bytes.
    """
    try:
        import zipfile, xml.etree.ElementTree as ET
        NS     = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

        if isinstance(source, (bytes, bytearray)):
            zf = zipfile.ZipFile(io.BytesIO(source))
        else:
            zf = zipfile.ZipFile(str(source))

        # Map rId -> sheet name
        with zf.open("xl/workbook.xml") as f:
            wb_tree = ET.parse(f)
        rid_to_name = {}
        for sh in wb_tree.findall(f".//{{{NS}}}sheet"):
            rid = sh.get(f"{{{REL_NS}}}id")
            rid_to_name[rid] = sh.get("name")

        # Map rId -> file path
        rid_to_path = {}
        if "xl/_rels/workbook.xml.rels" in zf.namelist():
            with zf.open("xl/_rels/workbook.xml.rels") as f:
                rels_tree = ET.parse(f)
            for rel in rels_tree.findall(f"{{{PKG_NS}}}Relationship"):
                target = rel.get("Target")
                path = target.lstrip("/") if target.startswith("/") else "xl/" + target
                rid_to_path[rel.get("Id")] = path

        target = "Collection ID List"
        candidate = None
        for rid, name in rid_to_name.items():
            if name == target:
                candidate = rid_to_path.get(rid)
                break

        if not candidate or candidate not in zf.namelist():
            zf.close()
            return {}

        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            with zf.open("xl/sharedStrings.xml") as f:
                ss_tree = ET.parse(f)
            for si in ss_tree.findall(f"{{{NS}}}si"):
                parts = si.findall(f".//{{{NS}}}t")
                shared_strings.append("".join(p.text or "" for p in parts))

        def cell_val(c_elem):
            t = c_elem.get("t", "n")
            if t == "inlineStr":
                is_elem = c_elem.find(f"{{{NS}}}is")
                if is_elem is None:
                    return ""
                parts = is_elem.findall(f".//{{{NS}}}t")
                return "".join(p.text or "" for p in parts)
            v = c_elem.find(f"{{{NS}}}v")
            if v is None or v.text is None:
                return ""
            if t == "s":
                idx = int(v.text)
                return shared_strings[idx] if idx < len(shared_strings) else ""
            val = v.text.strip()
            if val.endswith(".0"):
                val = val[:-2]
            return val

        result = {}
        with zf.open(candidate) as f:
            ws_tree = ET.parse(f)
        for row_elem in ws_tree.findall(f".//{{{NS}}}row"):
            if int(row_elem.get("r", 0)) < 2:
                continue
            col_vals = {}
            for c in row_elem.findall(f"{{{NS}}}c"):
                ref = c.get("r", "")
                col = "".join(ch for ch in ref if ch.isalpha()).upper()
                col_vals[col] = cell_val(c)
            row_state = col_vals.get("A", "").strip().upper()
            name      = col_vals.get("B", "").strip()
            cid       = col_vals.get("C", "").strip()
            if row_state != state.upper():
                continue
            if not name or not cid:
                continue
            if "(do not use)" in name.lower():
                continue
            result[name] = cid
        zf.close()
        return result
    except Exception:
        return {}


_CI_PREFIX_TEMPLATES = [
    "CI - {state} SPP - ",
    "CI - {state} - SPP - ",
    "CI - SPP - {state} - ",
]


def _bare_group_label(group_name, state):
    """Strip a known CI prefix from group_name, if present, to get the bare
    short label (e.g. 'CI - SPP - OK - ATLANTIC PACKAGE' -> 'ATLANTIC PACKAGE').
    Returns group_name unchanged if no known prefix matches (already-bare
    labels like 'Combo')."""
    if not group_name:
        return group_name
    lower = group_name.lower()
    for template in _CI_PREFIX_TEMPLATES:
        prefix = template.format(state=state.upper())
        if lower.startswith(prefix.lower()):
            return group_name[len(prefix):]
    return group_name


def auto_match(group_name, state, collections):
    """
    Pre-select a collection for a goal group.
    Some Goal Builders already store the full Salesforce collection name in
    the group field (e.g. 'CI - SPP - MN - ALD On Premise'); others store
    just the short label (e.g. 'Combo'). Different states also use different
    prefix conventions for their actual Salesforce collections -- state
    before SPP ('CI - OK - SPP - X') or after ('CI - SPP - OK - X') -- and
    a Goal Builder's embedded prefix doesn't always match its own state's
    collection convention (e.g. OK Goal Builders embed 'CI - SPP - OK - X'
    but OK's collections are actually named 'CI - OK - SPP - X').
    So: strip whichever prefix (if any) is already in group_name to get the
    bare label, then try the bare label as-is plus all three prefix
    conventions rebuilt from it, plus the original group_name unchanged.
    Falls back to None if no match found.
    """
    bare = _bare_group_label(group_name, state)
    candidates = [
        group_name,
        bare,
        f"CI - {state.upper()} SPP - {bare}",
        f"CI - {state.upper()} - SPP - {bare}",
        f"CI - SPP - {state.upper()} - {bare}",
    ]
    for candidate in candidates:
        for cname in collections:
            if cname.lower() == candidate.lower():
                return cname
    return None


# ---------------------------------------------------------------------------
# Tab 1: SPP Generator
# ---------------------------------------------------------------------------

def render_spp_tab():
    st.subheader("SPP LPM Upload Generator")
    st.info("📌 Designed for use with the **Standardized SPP template**.")
    st.caption(
        "Upload a Goal Builder `.xlsm` file, confirm the Collection ID mapping, "
        "and generate the LPM upload CSV. State is read from the sheet's Goal "
        "Group text (falls back to the filename if that's not found)."
    )

    goal_builder_file = st.file_uploader(
        "Goal Builder (.xlsm)",
        type=["xlsm"],
        key="spp_goal_builder_file",
    )

    collection_file = st.file_uploader(
        "Override Collection Report (.xlsx) — optional",
        type=["xlsx"],
        help="Leave blank to use the bundled collection report.",
        key="spp_collection_file",
    )

    if "spp_last_gb_name" not in st.session_state:
        st.session_state.spp_last_gb_name = None

    if goal_builder_file and goal_builder_file.name != st.session_state.spp_last_gb_name:
        st.session_state.spp_last_gb_name   = goal_builder_file.name
        st.session_state.spp_goal_groups    = None
        st.session_state.spp_manual_mapping = {}
        st.session_state.spp_result         = None

    if goal_builder_file:
        state = gen.derive_state_from_filename(goal_builder_file.name)

        if collection_file:
            coll_source = collection_file.getvalue()
            coll_source_label = f"Uploaded: {collection_file.name}"
        elif Path(BUNDLED_COLLECTION).is_file():
            coll_source = BUNDLED_COLLECTION
            coll_source_label = f"Bundled: {Path(BUNDLED_COLLECTION).name}"
        else:
            coll_source = None
            coll_source_label = "No collection report found"

        collections = get_state_collections(state, coll_source) if coll_source else {}
        st.caption(f"Collection report: {coll_source_label}")

        if st.session_state.get("spp_goal_groups") is None:
            with st.spinner("Reading Goal Builder…"):
                st.session_state.spp_goal_groups = extract_goal_groups(goal_builder_file.getvalue())

        goal_groups = st.session_state.spp_goal_groups

        if goal_groups and collections:
            with st.expander("Collection ID Mapping", expanded=True):
                st.caption(
                    f"State: **{state}** — {len(collections)} collection(s) available. "
                    "Match each Goal Group to its Salesforce Collection ID."
                )
                st.caption(
                    "ℹ️ Collection IDs below are auto-matched by Goal Group name — "
                    "verify each one before generating. A blank **(none)** means no "
                    "match was found, not that no collection exists for that group."
                )
                options = ["(none)"] + list(collections.keys())
                manual_mapping = {}

                with st.container(height=400):
                    for group in goal_groups:
                        best = auto_match(group, state, collections)
                        default_idx = options.index(best) if best and best in options else 0
                        selected = st.selectbox(
                            group,
                            options,
                            index=default_idx,
                            key=f"spp_cmap_{group}",
                        )
                        if selected != "(none)":
                            manual_mapping[group.lower()] = collections[selected]

                st.session_state.spp_manual_mapping = manual_mapping

        elif goal_groups and not collections:
            st.info(
                f"State: **{state}** — no collections found in the collection report. "
                "`salesforce_collection_ids` will be blank."
            )

    generate_clicked = st.button(
        "Generate CSV",
        type="primary",
        disabled=(goal_builder_file is None),
        use_container_width=True,
        key="spp_generate_button",
    )

    if generate_clicked and goal_builder_file is not None:
        with st.spinner("Processing Tracking Table…"):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    gb_path = os.path.join(tmpdir, goal_builder_file.name)
                    with open(gb_path, "wb") as f:
                        f.write(goal_builder_file.getvalue())

                    output_path  = os.path.join(tmpdir, "lpm_upload.csv")
                    skipped_path = os.path.join(tmpdir, "lpm_skipped.csv")

                    wb_for_fiscal = openpyxl.load_workbook(gb_path, keep_vba=True, data_only=True)
                    gen.INTERNAL_FISCAL_LOOKUP = gen.load_internal_fiscal_lookup(wb_for_fiscal)
                    state_from_sheet = ""
                    if "Tracking Table" in wb_for_fiscal.sheetnames:
                        tt = wb_for_fiscal["Tracking Table"]
                        for r in range(2, tt.max_row + 1):
                            ggv = tt.cell(r, 1).value
                            if ggv:
                                state_from_sheet = gen.derive_state_from_goal_group(str(ggv))
                                if state_from_sheet:
                                    break
                    wb_for_fiscal.close()

                    state = state_from_sheet or gen.derive_state_from_filename(gb_path)
                    gen.COLLECTION_LOOKUP = st.session_state.get("spp_manual_mapping", {})
                    gen.CURRENT_STATE = state

                    fiscal_path = _APP_DIR / "Supplier Fiscal Start Month.xlsx"
                    gen.FISCAL_LOOKUP = gen.load_fiscal_lookup(str(fiscal_path)) if fiscal_path.is_file() else {}

                    records, skipped, header_row = gen.load_tracking_table(gb_path)
                    if not records:
                        raise ValueError("No valid records found in the Tracking Table after filtering.")

                    order, groups = gen.group_records(records)
                    total_trackers, total_ptgs = gen.generate_output(order, groups, output_path)
                    duplicate_names = gen.find_duplicate_ptg_names(order, groups)

                    if skipped:
                        gen.write_skipped_csv(skipped_path, skipped, header_row)

                    with open(output_path, "rb") as f:
                        output_bytes = f.read()

                    skipped_bytes = None
                    if skipped and os.path.exists(skipped_path):
                        with open(skipped_path, "rb") as f:
                            skipped_bytes = f.read()

                    collection_info = [
                        f"{group}  →  {cid}"
                        for group, cid in sorted(gen.COLLECTION_LOOKUP.items())
                    ]

                st.session_state.spp_result = {
                    "error":            None,
                    "state":            state,
                    "total_trackers":   total_trackers,
                    "total_ptgs":       total_ptgs,
                    "skipped":          skipped,
                    "duplicate_names":  duplicate_names,
                    "output_bytes":     output_bytes,
                    "skipped_bytes":    skipped_bytes,
                    "base_name":        os.path.splitext(goal_builder_file.name)[0],
                    "collection_info":  collection_info,
                }

            except Exception as exc:
                st.session_state.spp_result = {"error": str(exc)}

    result = st.session_state.get("spp_result")

    if result:
        st.divider()

        if result.get("error"):
            st.error(f"**Error:** {result['error']}")
        else:
            if result["collection_info"]:
                with st.expander(
                    f"State: **{result['state']}** — {len(result['collection_info'])} collection ID(s) applied",
                    expanded=False,
                ):
                    for line in result["collection_info"]:
                        st.text(line)
            else:
                st.info("`salesforce_collection_ids` will be blank — no collections were mapped.")

            c1, c2, c3 = st.columns(3)
            c1.metric("Trackers", result["total_trackers"])
            c2.metric("PTGs", result["total_ptgs"])
            c3.metric("Skipped rows", len(result["skipped"]))

            if result["skipped"]:
                st.warning(f"⚠️ {len(result['skipped'])} row(s) were skipped — review before uploading.")
                with st.expander("View skipped rows"):
                    for entry in result["skipped"]:
                        goal_group = (entry["raw_row"][0] or "") if entry["raw_row"] else ""
                        st.markdown(
                            f"**Row {entry['row_num']}** &nbsp;·&nbsp; "
                            f"`{goal_group}` &nbsp;·&nbsp; {entry['reason']}"
                        )
            else:
                st.success("✅ All rows processed — no skipped rows.")

            if result.get("duplicate_names"):
                st.warning(
                    f"⚠️ {len(result['duplicate_names'])} PTG name(s) repeat within the same Tracker — "
                    "LPM may treat them as one on import."
                )
                with st.expander("View repeated PTG names"):
                    for d in result["duplicate_names"]:
                        st.markdown(
                            f"Tracker `{d['tracker']}` · PTG `{d['ptg_name']}` · source rows: {d['row_nums']}"
                        )

            st.divider()

            base = result["base_name"]
            st.download_button(
                label="⬇️  Download LPM Upload CSV",
                data=result["output_bytes"],
                file_name=f"{base}_lpm_upload.csv",
                mime="text/csv",
                use_container_width=True,
                key="spp_download_csv",
            )
            if result["skipped_bytes"]:
                st.download_button(
                    label="⬇️  Download Skipped Rows CSV",
                    data=result["skipped_bytes"],
                    file_name=f"{base}_skipped.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="spp_download_skipped",
                )


# ---------------------------------------------------------------------------
# Tab 2: NY SOD Generator
# ---------------------------------------------------------------------------

def render_sod_tab():
    st.subheader("NY SOD LPM Upload Generator")
    st.caption(
        "Upload a NY Monthly Quota Planner `.xlsb` file (the GOAL SHEET tab) "
        "and generate the LPM upload CSV."
    )

    sod_file = st.file_uploader(
        "NY SOD Goal Sheet (.xlsb)",
        type=["xlsb"],
        key="sod_goal_sheet_file",
    )

    generate_clicked = st.button(
        "Generate CSV",
        type="primary",
        disabled=(sod_file is None),
        use_container_width=True,
        key="sod_generate_button",
    )

    if generate_clicked and sod_file is not None:
        with st.spinner("Processing GOAL SHEET…"):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    sod_path = os.path.join(tmpdir, sod_file.name)
                    with open(sod_path, "wb") as f:
                        f.write(sod_file.getvalue())

                    output_path  = os.path.join(tmpdir, "sod_upload.csv")
                    skipped_path = os.path.join(tmpdir, "sod_skipped.csv")

                    if Path(BUNDLED_COLLECTION).is_file():
                        sodgen.spp.COLLECTION_LOOKUP = sodgen.spp.load_collection_lookup(BUNDLED_COLLECTION, "NY")
                    else:
                        sodgen.spp.COLLECTION_LOOKUP = {}

                    records, skipped, anchor_year, anchor_month = sodgen.load_goal_sheet(sod_path)
                    if not records:
                        raise ValueError("No valid records found in the GOAL SHEET after filtering.")

                    order, groups = sodgen.group_records(records)
                    total_trackers, total_ptgs = sodgen.generate_output(order, groups, output_path)
                    duplicate_names = sodgen.find_duplicate_ptg_names(order, groups)

                    if skipped:
                        sodgen.write_skipped_csv(skipped_path, skipped)

                    with open(output_path, "rb") as f:
                        output_bytes = f.read()

                    skipped_bytes = None
                    if skipped and os.path.exists(skipped_path):
                        with open(skipped_path, "rb") as f:
                            skipped_bytes = f.read()

                st.session_state.sod_result = {
                    "error":           None,
                    "anchor_year":     anchor_year,
                    "anchor_month":    anchor_month,
                    "total_trackers":  total_trackers,
                    "total_ptgs":      total_ptgs,
                    "skipped":         skipped,
                    "duplicate_names": duplicate_names,
                    "output_bytes":    output_bytes,
                    "skipped_bytes":   skipped_bytes,
                    "base_name":       os.path.splitext(sod_file.name)[0],
                }

            except Exception as exc:
                st.session_state.sod_result = {"error": str(exc)}

    result = st.session_state.get("sod_result")

    if result:
        st.divider()

        if result.get("error"):
            st.error(f"**Error:** {result['error']}")
        else:
            st.caption(f"Anchor: year={result['anchor_year']} month={result['anchor_month']}")

            c1, c2, c3 = st.columns(3)
            c1.metric("Trackers", result["total_trackers"])
            c2.metric("PTGs", result["total_ptgs"])
            c3.metric("Skipped rows", len(result["skipped"]))

            if result["skipped"]:
                st.warning(f"⚠️ {len(result['skipped'])} row(s) were skipped — review before uploading.")
                with st.expander("View skipped rows"):
                    for entry in result["skipped"]:
                        st.markdown(f"**Row {entry['row_num']}** &nbsp;·&nbsp; {entry['reason']}")
            else:
                st.success("✅ All rows processed — no skipped rows.")

            if result.get("duplicate_names"):
                st.warning(
                    f"⚠️ {len(result['duplicate_names'])} Goal Name(s) repeat within the same Tracker — "
                    "LPM may treat them as one on import."
                )
                with st.expander("View repeated Goal Names"):
                    for d in result["duplicate_names"]:
                        st.markdown(
                            f"Tracker `{d['tracker']}` · Goal Name `{d['ptg_name']}` · source rows: {d['row_nums']}"
                        )

            st.divider()

            base = result["base_name"]
            st.download_button(
                label="⬇️  Download SOD Upload CSV",
                data=result["output_bytes"],
                file_name=f"{base}_sod_upload.csv",
                mime="text/csv",
                use_container_width=True,
                key="sod_download_csv",
            )
            if result["skipped_bytes"]:
                st.download_button(
                    label="⬇️  Download Skipped Rows CSV",
                    data=result["skipped_bytes"],
                    file_name=f"{base}_skipped.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="sod_download_skipped",
                )


# ---------------------------------------------------------------------------
# Tab 3: NY SOD Revision Recap
# ---------------------------------------------------------------------------

def render_recap_tab():
    st.subheader("NY SOD Revision Recap")
    st.caption(
        "Upload the earlier INPUT/draft version and the later FINAL/revised version "
        "of a NY Monthly Quota Planner `.xlsb` file to see exactly what changed, "
        "was added, or was removed — no need to eyeball the green highlighting by hand. "
        "NY SOD-specific: this is keyed to the GOAL SHEET column layout and won't work "
        "against the SPP Goal Builder's Tracking Table."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        input_file = st.file_uploader("INPUT / draft (.xlsb)", type=["xlsb"], key="recap_input_file")
    with col_b:
        final_file = st.file_uploader("FINAL / revised (.xlsb)", type=["xlsb"], key="recap_final_file")

    compare_clicked = st.button(
        "Compare",
        type="primary",
        disabled=(input_file is None or final_file is None),
        use_container_width=True,
        key="recap_compare_button",
    )

    if compare_clicked and input_file is not None and final_file is not None:
        with st.spinner("Comparing GOAL SHEET tabs…"):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    input_path = os.path.join(tmpdir, input_file.name)
                    final_path = os.path.join(tmpdir, final_file.name)
                    with open(input_path, "wb") as f:
                        f.write(input_file.getvalue())
                    with open(final_path, "wb") as f:
                        f.write(final_file.getvalue())

                    input_rows = recap.load_rows(input_path)
                    final_rows = recap.load_rows(final_path)
                    changes = recap.diff_rows(input_rows, final_rows)

                    output_path = os.path.join(tmpdir, "revision_recap.csv")
                    recap.write_recap_csv(changes, output_path)
                    with open(output_path, "rb") as f:
                        output_bytes = f.read()

                st.session_state.recap_result = {
                    "error":        None,
                    "changes":      changes,
                    "output_bytes": output_bytes,
                    "base_name":    os.path.splitext(final_file.name)[0],
                }

            except Exception as exc:
                st.session_state.recap_result = {"error": str(exc)}

    result = st.session_state.get("recap_result")

    if result:
        st.divider()

        if result.get("error"):
            st.error(f"**Error:** {result['error']}")
        else:
            changes = result["changes"]
            added    = [c for c in changes if c["Change Type"] == "Added"]
            removed  = [c for c in changes if c["Change Type"] == "Removed"]
            modified = [c for c in changes if c["Change Type"] == "Modified"]

            c1, c2, c3 = st.columns(3)
            c1.metric("Modified", len(modified))
            c2.metric("Added", len(added))
            c3.metric("Removed", len(removed))

            if changes:
                st.dataframe(changes, use_container_width=True, hide_index=True)
            else:
                st.success("✅ No differences found.")

            st.divider()

            base = result["base_name"]
            st.download_button(
                label="⬇️  Download Revision Recap CSV",
                data=result["output_bytes"],
                file_name=f"{base}_revision_recap.csv",
                mime="text/csv",
                use_container_width=True,
                key="recap_download_csv",
            )


# ---------------------------------------------------------------------------
# Tab 4: SPP Template Checker
# ---------------------------------------------------------------------------

_SEVERITY_ICON = {"warning": "⚠️", "info": "ℹ️"}


def render_checker_tab():
    st.subheader("SPP Template Checker")
    st.caption(
        "Upload a Goal Builder `.xlsm` and scan its Tracking Table for data-entry "
        "problems — missing/blank end dates, duplicated or concatenated text, "
        "Supplier/Selection mismatches, non-numeric goals, header/column changes, "
        "silent skips, and Anchor/Flex ratio imbalance. This never generates a CSV "
        "and never blocks anything downstream — it's a heads-up before you run the "
        "actual generator."
    )

    check_file = st.file_uploader(
        "Goal Builder (.xlsm)",
        type=["xlsm"],
        key="checker_goal_builder_file",
    )

    check_clicked = st.button(
        "Check Template",
        type="primary",
        disabled=(check_file is None),
        use_container_width=True,
        key="checker_run_button",
    )

    if check_clicked and check_file is not None:
        with st.spinner("Scanning Tracking Table…"):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    gb_path = os.path.join(tmpdir, check_file.name)
                    with open(gb_path, "wb") as f:
                        f.write(check_file.getvalue())
                    flags, header_issues, row_count = checker.check_tracking_table(gb_path)

                st.session_state.checker_result = {
                    "error": None,
                    "flags": flags,
                    "header_issues": header_issues,
                    "row_count": row_count,
                }
            except Exception as exc:
                st.session_state.checker_result = {"error": str(exc)}

    result = st.session_state.get("checker_result")

    if result:
        st.divider()

        if result.get("error"):
            st.error(f"**Error:** {result['error']}")
        else:
            st.caption(f"Checked {result['row_count']} row(s).")

            if result["header_issues"]:
                st.error(f"🛑 {len(result['header_issues'])} header/column issue(s) — the template structure itself has changed.")
                with st.expander("View header issues", expanded=True):
                    for issue in result["header_issues"]:
                        st.markdown(f"- {issue}")

            flags = result["flags"]
            warnings = [f for f in flags if f["severity"] == "warning"]
            infos = [f for f in flags if f["severity"] == "info"]

            c1, c2 = st.columns(2)
            c1.metric("Warnings", len(warnings))
            c2.metric("Informational", len(infos))

            if not flags and not result["header_issues"]:
                st.success("✅ No issues found.")
            else:
                for f in warnings + infos:
                    where = f"Row {f['row_num']}" if f["row_num"] else "File-level"
                    icon = _SEVERITY_ICON.get(f["severity"], "")
                    st.markdown(f"{icon} **{where}** · `{f['category']}` — {f['message']}")


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    if CI_LOGO.is_file():
        st.image(str(CI_LOGO), use_container_width=True)
    st.header("LPM Upload Generator")
    st.markdown("""
Four tools, one app:
- **SPP Generator** — SPP Goal Builder `.xlsm` → LPM upload CSV
- **NY SOD Generator** — NY Monthly Quota Planner `.xlsb` → LPM upload CSV
- **NY SOD Revision Recap** — diff an INPUT vs FINAL NY SOD file
- **SPP Template Checker** — scan a Goal Builder for data-entry problems before generating

The **LPM Collection Report** is bundled automatically for both generators.
""")
    st.divider()
    st.caption("LPM Upload Generator · Southern Glazer's")

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
st.title("LPM Upload Generator")

tab_spp, tab_sod, tab_recap, tab_checker = st.tabs(
    ["SPP Generator", "NY SOD Generator", "NY SOD Revision Recap", "SPP Template Checker"]
)

with tab_spp:
    render_spp_tab()

with tab_sod:
    render_sod_tab()

with tab_checker:
    render_checker_tab()

with tab_recap:
    render_recap_tab()
