"""
Delta Analysis Tool — Main Streamlit Application

Designed for government analysts performing reconciliation, audit
support, case tracking, DFAS-style package reviews, and operational
reporting. Upload two datasets, configure your match keys, and
generate a full delta comparison with visualizations and a
leadership-ready Excel report.
"""
import io
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.delta_engine import DeltaResult, run_delta
from src.io_utils import (
    check_file_size,
    get_column_preview,
    get_display_frame,
    get_excel_sheet_names,
    read_uploaded_file,
)
from src.reporting import build_change_frequency, build_summary_df, export_to_excel

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Delta Analysis Tool",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------

st.markdown(
    """
    <style>
        .main-header {
            font-size: 2rem;
            font-weight: 700;
            color: #1F4E79;
            margin-bottom: 0.2rem;
        }
        .sub-header {
            font-size: 0.93rem;
            color: #555;
            margin-bottom: 1.2rem;
        }
        .section-title {
            font-size: 1.05rem;
            font-weight: 600;
            color: #1F4E79;
            border-bottom: 2px solid #1F4E79;
            padding-bottom: 4px;
            margin-top: 0.5rem;
            margin-bottom: 10px;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #dce4ef;
            border-radius: 6px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def _csv_bytes(df: pd.DataFrame) -> bytes:
    """Return a DataFrame serialised as UTF-8 CSV bytes."""
    return df.to_csv(index=False).encode("utf-8")


def _show_df(df: pd.DataFrame, key: str = "") -> None:
    """Display a DataFrame, truncating to PREVIEW_MAX_ROWS and noting if truncated."""
    from src.io_utils import PREVIEW_MAX_ROWS
    display_df, truncated = get_display_frame(df, max_rows=PREVIEW_MAX_ROWS)
    if truncated:
        st.caption(
            f"Showing first {PREVIEW_MAX_ROWS:,} of {len(df):,} rows. "
            "Download the CSV for the full dataset."
        )
    st.dataframe(display_df, use_container_width=True, hide_index=True, key=key or None)


def _csv_download(label: str, df: pd.DataFrame, filename: str) -> None:
    """Render a CSV download button if the DataFrame is non-empty."""
    if df is not None and not df.empty:
        st.download_button(
            label=f"⬇ Download {label} (CSV)",
            data=_csv_bytes(df),
            file_name=filename,
            mime="text/csv",
            key=f"csv_{filename}",
        )
    else:
        st.caption("No records to download for this category.")


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown('<div class="main-header">🔍 Delta Analysis Tool</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">'
    "Upload two datasets, select your match keys, and generate a full bidirectional comparison. "
    "Supports government reconciliation, receipt tracking, case audits, and duplicate submission review."
    "</div>",
    unsafe_allow_html=True,
)
st.divider()

# ---------------------------------------------------------------------------
# Step 1 — File Upload
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">Step 1 — Upload Files</div>', unsafe_allow_html=True)

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**File A** — Baseline / Prior Period / Source of Record")
    file_a = st.file_uploader(
        "Upload File A",
        type=["csv", "xlsx", "xls"],
        key="upload_a",
        label_visibility="collapsed",
    )

with col_b:
    st.markdown("**File B** — Comparison / Current Period / Received File")
    file_b = st.file_uploader(
        "Upload File B",
        type=["csv", "xlsx", "xls"],
        key="upload_b",
        label_visibility="collapsed",
    )

df_a: pd.DataFrame | None = None
df_b: pd.DataFrame | None = None
sheet_a_selected: str | None = None
sheet_b_selected: str | None = None
_large_file_blocked = False

if file_a:
    with col_a:
        sheets_a = get_excel_sheet_names(file_a)
        if len(sheets_a) > 1:
            sheet_a_selected = st.selectbox(
                "Sheet (File A)",
                options=sheets_a,
                key="sheet_a",
            )
        elif sheets_a:
            sheet_a_selected = sheets_a[0]
    try:
        df_a = read_uploaded_file(file_a, sheet_name=sheet_a_selected)
        with col_a:
            st.success(f"✔ {len(df_a):,} rows × {len(df_a.columns)} columns loaded")
            size_status, size_msg = check_file_size(len(df_a))
            if size_status == "warn":
                st.warning(f"File A: {size_msg}")
            elif size_status == "hard":
                st.error(f"File A: {size_msg}")
                if not st.checkbox("I understand the risk — proceed anyway (File A)", key="hard_a"):
                    _large_file_blocked = True
    except ValueError as exc:
        with col_a:
            st.error(f"File A error: {exc}")

if file_b:
    with col_b:
        sheets_b = get_excel_sheet_names(file_b)
        if len(sheets_b) > 1:
            sheet_b_selected = st.selectbox(
                "Sheet (File B)",
                options=sheets_b,
                key="sheet_b",
            )
        elif sheets_b:
            sheet_b_selected = sheets_b[0]
    try:
        df_b = read_uploaded_file(file_b, sheet_name=sheet_b_selected)
        with col_b:
            st.success(f"✔ {len(df_b):,} rows × {len(df_b.columns)} columns loaded")
            size_status, size_msg = check_file_size(len(df_b))
            if size_status == "warn":
                st.warning(f"File B: {size_msg}")
            elif size_status == "hard":
                st.error(f"File B: {size_msg}")
                if not st.checkbox("I understand the risk — proceed anyway (File B)", key="hard_b"):
                    _large_file_blocked = True
    except ValueError as exc:
        with col_b:
            st.error(f"File B error: {exc}")

if df_a is not None or df_b is not None:
    with st.expander("Preview uploaded data (first 5 rows)", expanded=False):
        p1, p2 = st.columns(2)
        if df_a is not None:
            with p1:
                st.markdown("**File A**")
                st.dataframe(get_column_preview(df_a), use_container_width=True, hide_index=True)
        if df_b is not None:
            with p2:
                st.markdown("**File B**")
                st.dataframe(get_column_preview(df_b), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Step 2 — Key column configuration
# ---------------------------------------------------------------------------

if df_a is not None and df_b is not None:
    st.divider()
    st.markdown('<div class="section-title">Step 2 — Configure Match Keys</div>', unsafe_allow_html=True)
    st.info(
        "**Key columns** uniquely identify each record — e.g., Case ID, Contract Number. "
        "Select the matching column(s) from each file. Multiple columns are combined into a "
        "composite key and matched positionally."
    )

    k1, k2 = st.columns(2)
    with k1:
        key_cols_a = st.multiselect(
            "Key column(s) from File A",
            options=df_a.columns.tolist(),
            help="Column(s) that uniquely identify each record in File A.",
        )
    with k2:
        key_cols_b = st.multiselect(
            "Key column(s) from File B",
            options=df_b.columns.tolist(),
            help="Column(s) that uniquely identify each record in File B.",
        )

    if key_cols_a and key_cols_b:
        if len(key_cols_a) != len(key_cols_b):
            st.warning(
                f"Key column count mismatch — File A: {len(key_cols_a)}, "
                f"File B: {len(key_cols_b)}. Counts must match."
            )
        else:
            st.markdown("**Key mapping (positional):**")
            st.dataframe(
                pd.DataFrame({"File A Key": key_cols_a, "File B Key": key_cols_b}),
                use_container_width=False,
                hide_index=True,
            )

    # -----------------------------------------------------------------------
    # Step 3 — Comparison columns
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown('<div class="section-title">Step 3 — Select Fields to Compare for Changes</div>', unsafe_allow_html=True)
    st.info(
        "**Comparison columns** are non-key fields to diff between files. "
        "Only matched records are checked. Leave blank to skip field-change analysis."
    )

    common_non_key = [
        c for c in df_a.columns
        if c in df_b.columns and c not in key_cols_a and c not in key_cols_b
    ]

    c1, c2 = st.columns(2)
    with c1:
        compare_cols_a = st.multiselect(
            "Comparison columns from File A",
            options=[c for c in df_a.columns if c not in key_cols_a],
            default=common_non_key,
        )
    with c2:
        compare_cols_b = st.multiselect(
            "Corresponding columns from File B",
            options=[c for c in df_b.columns if c not in key_cols_b],
            default=[c for c in common_non_key if c in df_b.columns],
        )

    if compare_cols_a and compare_cols_b and len(compare_cols_a) != len(compare_cols_b):
        st.warning(
            f"Comparison column count mismatch — File A: {len(compare_cols_a)}, "
            f"File B: {len(compare_cols_b)}. Counts must match."
        )
    elif compare_cols_a and compare_cols_b:
        st.markdown("**Comparison mapping (positional):**")
        st.dataframe(
            pd.DataFrame({"File A Column": compare_cols_a, "File B Column": compare_cols_b}),
            use_container_width=False,
            hide_index=True,
        )

    # Advanced comparison settings ----------------------------------------
    # Uses a type-grouped design: one multiselect per type + shared settings.
    # This scales to 20+ comparison columns without becoming unusable.
    comparison_rules: list[dict] | None = None
    if compare_cols_a and compare_cols_b and len(compare_cols_a) == len(compare_cols_b):
        col_pairs = list(zip(compare_cols_a, compare_cols_b))
        col_labels = [ca if ca == cb else f"{ca} / {cb}" for ca, cb in col_pairs]

        with st.expander("Advanced Comparison Settings (optional)", expanded=False):
            st.info(
                "By default all fields are compared as **text**. "
                "Use the selectors below to assign numeric tolerance or date-aware "
                "comparison to specific fields. Fields not assigned remain text."
            )

            adv1, adv2 = st.columns(2)

            with adv1:
                st.markdown("**Numeric fields**")
                numeric_labels = st.multiselect(
                    "Fields to compare as numeric",
                    options=col_labels,
                    key="adv_numeric",
                    help="Currency symbols and commas are stripped automatically.",
                )
                numeric_tolerance = st.number_input(
                    "Tolerance (applies to all numeric fields)",
                    min_value=0.0,
                    value=0.0,
                    step=0.01,
                    format="%.4f",
                    key="adv_tol",
                    help="Maximum allowed difference — 0.00 means exact match.",
                )

            with adv2:
                st.markdown("**Date fields**")
                date_labels = st.multiselect(
                    "Fields to compare as dates",
                    options=[lbl for lbl in col_labels if lbl not in numeric_labels],
                    key="adv_date",
                    help="Compares calendar values regardless of input format.",
                )
                date_mode = st.selectbox(
                    "Date precision (applies to all date fields)",
                    options=["date_only", "datetime_precision"],
                    key="adv_date_mode",
                    help=(
                        "date_only: ignore time — 08:30 and 14:00 on the same day are equal. "
                        "datetime_precision: include time in the comparison."
                    ),
                )

            if numeric_labels or date_labels:
                built_rules: list[dict] = []
                for label, (ca, cb) in zip(col_labels, col_pairs):
                    if label in numeric_labels:
                        built_rules.append({
                            "column_a":  ca, "column_b": cb,
                            "type": "numeric",
                            "tolerance": numeric_tolerance,
                            "date_mode": None,
                        })
                    elif label in date_labels:
                        built_rules.append({
                            "column_a":  ca, "column_b": cb,
                            "type": "date",
                            "tolerance": None,
                            "date_mode": date_mode,
                        })
                    else:
                        built_rules.append({
                            "column_a":  ca, "column_b": cb,
                            "type": "text",
                            "tolerance": None,
                            "date_mode": None,
                        })
                comparison_rules = built_rules

    # -----------------------------------------------------------------------
    # Step 4 — Run
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown('<div class="section-title">Step 4 — Run Analysis</div>', unsafe_allow_html=True)

    run_ready = bool(
        key_cols_a and key_cols_b
        and len(key_cols_a) == len(key_cols_b)
        and not _large_file_blocked
    )
    if not run_ready and not _large_file_blocked:
        st.warning("Select at least one key column from each file (with matching counts) to proceed.")
    if _large_file_blocked:
        st.warning("Confirm the large-file warning above before running analysis.")

    if st.button("▶  Run Delta Analysis", type="primary", disabled=not run_ready):
        with st.spinner("Analysing…"):
            try:
                result: DeltaResult = run_delta(
                    df_a=df_a,
                    df_b=df_b,
                    key_cols_a=key_cols_a,
                    key_cols_b=key_cols_b,
                    compare_cols_a=compare_cols_a or None,
                    compare_cols_b=compare_cols_b or None,
                    comparison_rules=comparison_rules,
                    sheet_a=sheet_a_selected,
                    sheet_b=sheet_b_selected,
                )
                st.session_state["result"]      = result
                st.session_state["file_a_name"] = file_a.name
                st.session_state["file_b_name"] = file_b.name
                st.success("Analysis complete — review the results below.")
            except ValueError as exc:
                st.error(f"Analysis failed: {exc}")
            except Exception as exc:
                st.error(f"Unexpected error: {exc}")

# ===========================================================================
# RESULTS SECTION
# ===========================================================================

if "result" in st.session_state:
    result: DeltaResult = st.session_state["result"]
    file_a_name: str    = st.session_state.get("file_a_name", "File A")
    file_b_name: str    = st.session_state.get("file_b_name", "File B")

    n_only_a  = len(result.only_in_a)
    n_only_b  = len(result.only_in_b)
    n_matched = len(result.matched)
    n_changed = len(result.changed)
    n_dup_a   = len(result.duplicates_a)
    n_dup_b   = len(result.duplicates_b)
    n_blank_a = len(result.blank_keys_a)
    n_blank_b = len(result.blank_keys_b)

    # -----------------------------------------------------------------------
    # Metric cards
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown('<div class="section-title">Results Dashboard</div>', unsafe_allow_html=True)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("File A — Total",       f"{result.total_a:,}", help="Total records loaded from File A")
    k2.metric("File B — Total",       f"{result.total_b:,}", help="Total records loaded from File B")
    k3.metric("Only in File A",       f"{n_only_a:,}",       help="Key present in A but not B")
    k4.metric("Only in File B",       f"{n_only_b:,}",       help="Key present in B but not A")
    k5.metric("Matched Records",      f"{n_matched:,}",      help="Common key in both files")

    k6, k7, k8, k9, k10 = st.columns(5)
    k6.metric("Records with Changes", f"{n_changed:,}",      help="Matched records where a compared field differs")
    k7.metric("Duplicate Keys — A",   f"{n_dup_a:,}",        help="Rows sharing a key within File A")
    k8.metric("Duplicate Keys — B",   f"{n_dup_b:,}",        help="Rows sharing a key within File B")
    k9.metric("Blank Keys — A",       f"{n_blank_a:,}",      help="Rows with null/blank key in File A")
    k10.metric("Blank Keys — B",      f"{n_blank_b:,}",      help="Rows with null/blank key in File B")

    # -----------------------------------------------------------------------
    # Visualizations
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown('<div class="section-title">Visual Summary</div>', unsafe_allow_html=True)

    viz1, viz2, viz3 = st.columns([2, 1.4, 1.6])

    # Chart 1 — Delta category bar chart
    with viz1:
        st.markdown("**Delta Categories**")
        cat_labels = [
            "Only in A", "Only in B", "Matched",
            "Changed", "Dupes A", "Dupes B",
            "Blank Keys A", "Blank Keys B",
        ]
        cat_values = [
            n_only_a, n_only_b, n_matched,
            n_changed, n_dup_a, n_dup_b,
            n_blank_a, n_blank_b,
        ]
        cat_colors = [
            "#C00000", "#375623", "#1F4E79",
            "#7B3F00", "#843C0C", "#843C0C",
            "#403152", "#403152",
        ]
        fig_bar = go.Figure(go.Bar(
            x=cat_labels,
            y=cat_values,
            marker_color=cat_colors,
            text=cat_values,
            textposition="outside",
            hovertemplate="%{x}: %{y:,}<extra></extra>",
        ))
        fig_bar.update_layout(
            margin=dict(l=20, r=20, t=10, b=40),
            height=320,
            yaxis_title="Records",
            plot_bgcolor="white",
            paper_bgcolor="white",
            yaxis=dict(gridcolor="#EEEEEE"),
            font=dict(size=11),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # Chart 2 — Match coverage donut
    with viz2:
        st.markdown("**Match Coverage**")
        unmatched_a = n_only_a
        unmatched_b = n_only_b
        pie_labels = ["Matched", f"Unmatched (A only)", f"Unmatched (B only)"]
        pie_values = [n_matched, unmatched_a, unmatched_b]
        pie_colors = ["#1F4E79", "#C00000", "#375623"]

        fig_pie = go.Figure(go.Pie(
            labels=pie_labels,
            values=pie_values,
            hole=0.48,
            marker_colors=pie_colors,
            textinfo="percent",
            hovertemplate="%{label}: %{value:,}<extra></extra>",
        ))
        fig_pie.update_layout(
            margin=dict(l=10, r=10, t=10, b=10),
            height=320,
            showlegend=True,
            legend=dict(orientation="v", font=dict(size=10)),
            paper_bgcolor="white",
            font=dict(size=11),
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # Chart 3 — Field-change frequency
    with viz3:
        st.markdown("**Field Change Frequency**")
        freq_df = build_change_frequency(result)
        if freq_df.empty:
            st.info("Run with comparison columns selected to see field-change frequency.")
        else:
            fig_freq = go.Figure(go.Bar(
                x=freq_df["Changes"],
                y=freq_df["Field"],
                orientation="h",
                marker_color="#7B3F00",
                text=freq_df["Changes"],
                textposition="outside",
                hovertemplate="%{y}: %{x:,} changes<extra></extra>",
            ))
            fig_freq.update_layout(
                margin=dict(l=20, r=40, t=10, b=30),
                height=320,
                xaxis_title="# of Changes",
                plot_bgcolor="white",
                paper_bgcolor="white",
                xaxis=dict(gridcolor="#EEEEEE"),
                yaxis=dict(autorange="reversed"),
                font=dict(size=11),
            )
            st.plotly_chart(fig_freq, use_container_width=True)

    # -----------------------------------------------------------------------
    # Detailed result tabs
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown('<div class="section-title">Detailed Results</div>', unsafe_allow_html=True)

    tabs = st.tabs([
        "📋 Summary",
        f"⬅ Only in A ({n_only_a:,})",
        f"➡ Only in B ({n_only_b:,})",
        f"✅ Matched ({n_matched:,})",
        f"⚠ Changed ({n_changed:,})",
        f"⚡ Dupes A ({n_dup_a:,})",
        f"⚡ Dupes B ({n_dup_b:,})",
        "🔴 Data Quality",
    ])

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Tab 0 — Summary
    with tabs[0]:
        st.dataframe(build_summary_df(result), use_container_width=True, hide_index=True)

    # Tab 1 — Only in A
    with tabs[1]:
        st.markdown(
            f"**{n_only_a:,} record(s)** whose key appears in **{file_a_name}** "
            f"but not in **{file_b_name}**."
        )
        if not result.only_in_a.empty:
            _show_df(result.only_in_a, key="only_a")
        else:
            st.info("No records in this category.")
        _csv_download("Only in File A", result.only_in_a, f"only_in_a_{ts}.csv")

    # Tab 2 — Only in B
    with tabs[2]:
        st.markdown(
            f"**{n_only_b:,} record(s)** whose key appears in **{file_b_name}** "
            f"but not in **{file_a_name}**."
        )
        if not result.only_in_b.empty:
            _show_df(result.only_in_b, key="only_b")
        else:
            st.info("No records in this category.")
        _csv_download("Only in File B", result.only_in_b, f"only_in_b_{ts}.csv")

    # Tab 3 — Matched
    with tabs[3]:
        st.markdown(
            f"**{n_matched:,} record(s)** with a matching key in both files. "
            "Columns prefixed **A:** (File A) and **B:** (File B)."
        )
        if not result.matched.empty:
            _show_df(result.matched, key="matched")
        else:
            st.info("No matched records found.")
        _csv_download("Matched Records", result.matched, f"matched_{ts}.csv")

    # Tab 4 — Changed
    with tabs[4]:
        if not result.compare_cols_a:
            st.info(
                "No comparison columns were selected. "
                "Re-run with comparison columns to see field-level changes."
            )
        elif result.changed.empty:
            st.success("No field-level changes detected among matched records.")
        else:
            st.markdown(
                f"**{n_changed:,} matched record(s)** where at least one compared field "
                "differs. Each changed field shows its **File A** and **File B** values."
            )
            _show_df(result.changed, key="changed")
        _csv_download("Changed Records", result.changed, f"changed_{ts}.csv")
        if result.compare_parse_issues is not None and not result.compare_parse_issues.empty:
            with st.expander(
                f"⚠ {len(result.compare_parse_issues)} parse issue(s) detected during comparison",
                expanded=False,
            ):
                st.dataframe(result.compare_parse_issues, use_container_width=True, hide_index=True)

    # Tab 5 — Duplicates A
    with tabs[5]:
        st.markdown(
            f"**{n_dup_a:,} row(s)** sharing a key with at least one other row "
            f"in **{file_a_name}**."
        )
        if not result.duplicates_a.empty:
            _show_df(result.duplicates_a, key="dup_a")
        else:
            st.success("No duplicate keys in File A.")
        _csv_download("Duplicate Keys File A", result.duplicates_a, f"dupes_a_{ts}.csv")

    # Tab 6 — Duplicates B
    with tabs[6]:
        st.markdown(
            f"**{n_dup_b:,} row(s)** sharing a key with at least one other row "
            f"in **{file_b_name}**."
        )
        if not result.duplicates_b.empty:
            _show_df(result.duplicates_b, key="dup_b")
        else:
            st.success("No duplicate keys in File B.")
        _csv_download("Duplicate Keys File B", result.duplicates_b, f"dupes_b_{ts}.csv")

    # Tab 7 — Data Quality
    with tabs[7]:
        st.markdown("**Rows excluded from matching due to blank or null key values.**")
        dq_rows: list[dict] = []
        for _, row in result.blank_keys_a.iterrows():
            dq_rows.append({"Source File": file_a_name, "Issue": "Blank / Null Key", **row.to_dict()})
        for _, row in result.blank_keys_b.iterrows():
            dq_rows.append({"Source File": file_b_name, "Issue": "Blank / Null Key", **row.to_dict()})
        dq_df = pd.DataFrame(dq_rows) if dq_rows else pd.DataFrame(columns=["Source File", "Issue"])
        if not dq_df.empty:
            _show_df(dq_df, key="dq")
        else:
            st.success("No data quality issues detected.")
        _csv_download("Data Quality Issues", dq_df, f"data_quality_{ts}.csv")

    # -----------------------------------------------------------------------
    # Export
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown('<div class="section-title">Export Results</div>', unsafe_allow_html=True)

    exp_col, info_col = st.columns([1, 3])

    with exp_col:
        with st.spinner("Building Excel workbook…"):
            excel_bytes = export_to_excel(result, file_a_name, file_b_name)

        st.download_button(
            label="⬇ Download Excel Report",
            data=excel_bytes,
            file_name=f"delta_analysis_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    with info_col:
        st.markdown(
            """
            The Excel workbook contains **11 tabs**:
            - **Executive Summary** — auto-generated plain-English briefing text
            - **Analysis Metadata** — file names, sheets, row counts, key columns, timestamp
            - **Comparison Rules** — per-column comparison type, tolerance, date mode
            - **Delta Counts** — flat count table for pivot/charting
            - **Only in File A / B** — unmatched records
            - **Matched Records** — side-by-side view
            - **Changed Records** — before/after for each changed field
            - **Duplicate Keys A / B** — non-unique key rows
            - **Data Quality Issues** — blank/null key rows
            """
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.divider()
st.markdown(
    "<div style='text-align:center; color:#888; font-size:0.8rem;'>"
    "Delta Analysis Tool &nbsp;|&nbsp; Streamlit + Pandas + Plotly &nbsp;|&nbsp; "
    f"{datetime.now().strftime('%Y-%m-%d')}"
    "</div>",
    unsafe_allow_html=True,
)
