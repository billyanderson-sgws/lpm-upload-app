"""
Streamlit web app for the SPP LPM Upload Generator.

Run with:
    streamlit run app_lpm_upload.py
"""

import os
import sys
import tempfile

import streamlit as st

# Make sure generate_lpm_upload is importable from the same folder
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_lpm_upload as gen

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SPP LPM Upload Generator",
    page_icon="📊",
    layout="centered",
)

# ---------------------------------------------------------------------------
# Bundled collection report path (committed alongside the app)
# ---------------------------------------------------------------------------
BUNDLED_COLLECTION = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "LPM Salesforce and Overlay Collection Report.xlsx",
)

# ---------------------------------------------------------------------------
# Sidebar — instructions
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("How to use")
    st.markdown("""
1. Upload your **Goal Builder** `.xlsm` file
2. Click **Generate CSV**
3. Download the output and review any skipped rows

---
**State** is read from the first word of the Goal Builder filename
(e.g. `SD SPP Goal Builder.xlsm` → **SD**)

---
The **LPM Collection Report** is bundled automatically.
Upload a replacement below only if you have a newer version.
""")
    st.markdown("---")
    st.caption("LPM Upload Generator · Southern Glazer's")

# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------
st.title("SPP LPM Upload Generator")

goal_builder_file = st.file_uploader(
    "Goal Builder (.xlsm)",
    type=["xlsm"],
    help="The state SPP Goal Builder file. State abbreviation is derived from the filename.",
)

collection_file = st.file_uploader(
    "Override Collection Report (.xlsx) — optional",
    type=["xlsx"],
    help="Leave blank to use the bundled collection report. Upload only if you have a newer version.",
)

# Clear prior result when new files are uploaded
if "last_gb_name" not in st.session_state:
    st.session_state.last_gb_name = None

if goal_builder_file and goal_builder_file.name != st.session_state.last_gb_name:
    st.session_state.last_gb_name = goal_builder_file.name
    st.session_state.result = None

generate_clicked = st.button(
    "Generate CSV",
    type="primary",
    disabled=(goal_builder_file is None),
    use_container_width=True,
)

# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------
if generate_clicked and goal_builder_file is not None:
    with st.spinner("Processing Tracking Table…"):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Write uploaded files to temp paths
                gb_path = os.path.join(tmpdir, goal_builder_file.name)
                with open(gb_path, "wb") as f:
                    f.write(goal_builder_file.getvalue())

                # Collection report: use uploaded override, else fall back to bundled file
                coll_path = None
                if collection_file:
                    coll_path = os.path.join(tmpdir, collection_file.name)
                    with open(coll_path, "wb") as f:
                        f.write(collection_file.getvalue())
                elif os.path.isfile(BUNDLED_COLLECTION):
                    coll_path = BUNDLED_COLLECTION

                output_path  = os.path.join(tmpdir, "lpm_upload.csv")
                skipped_path = os.path.join(tmpdir, "lpm_skipped.csv")

                # Load collection IDs
                state = gen.derive_state_from_filename(gb_path)
                if coll_path:
                    gen.COLLECTION_LOOKUP = gen.load_collection_lookup(coll_path, state)
                else:
                    gen.COLLECTION_LOOKUP = {}

                # Parse and generate
                records, skipped, header_row = gen.load_tracking_table(gb_path)

                if not records:
                    raise ValueError("No valid records found in the Tracking Table after filtering.")

                order, groups = gen.group_records(records)
                total_trackers, total_ptgs = gen.generate_output(order, groups, output_path)

                if skipped:
                    gen.write_skipped_csv(skipped_path, skipped, header_row)

                # Read output bytes while temp dir is still alive
                with open(output_path, "rb") as f:
                    output_bytes = f.read()

                skipped_bytes = None
                if skipped and os.path.exists(skipped_path):
                    with open(skipped_path, "rb") as f:
                        skipped_bytes = f.read()

                collection_info = [
                    f"{name}  →  {cid}"
                    for name, cid in sorted(gen.COLLECTION_LOOKUP.items())
                ]

            st.session_state.result = {
                "error":           None,
                "state":           state,
                "total_trackers":  total_trackers,
                "total_ptgs":      total_ptgs,
                "skipped":         skipped,
                "output_bytes":    output_bytes,
                "skipped_bytes":   skipped_bytes,
                "base_name":       os.path.splitext(goal_builder_file.name)[0],
                "collection_info": collection_info,
            }

        except Exception as exc:
            st.session_state.result = {"error": str(exc)}

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
result = st.session_state.get("result")

if result:
    st.markdown("---")

    if result.get("error"):
        st.error(f"**Error:** {result['error']}")

    else:
        # Collection ID summary
        if result["collection_info"]:
            with st.expander(
                f"State: **{result['state']}** — {len(result['collection_info'])} collection ID(s) loaded",
                expanded=False,
            ):
                for line in result["collection_info"]:
                    st.text(line)
        else:
            st.info(
                f"State: **{result['state']}** — no collection report provided; "
                "`salesforce_collection_ids` will be blank."
            )

        # Metrics
        c1, c2, c3 = st.columns(3)
        c1.metric("Trackers", result["total_trackers"])
        c2.metric("PTGs", result["total_ptgs"])
        c3.metric("Skipped rows", len(result["skipped"]))

        # Skipped rows detail
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

        st.markdown("")

        # Download buttons
        base = result["base_name"]
        st.download_button(
            label="⬇️  Download LPM Upload CSV",
            data=result["output_bytes"],
            file_name=f"{base}_lpm_upload.csv",
            mime="text/csv",
            use_container_width=True,
        )

        if result["skipped_bytes"]:
            st.download_button(
                label="⬇️  Download Skipped Rows CSV",
                data=result["skipped_bytes"],
                file_name=f"{base}_skipped.csv",
                mime="text/csv",
                use_container_width=True,
            )
