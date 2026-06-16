"""
Delta Analysis Tool - Main Streamlit Application

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
    RAW_PREVIEW_ROWS,
    check_file_size,
    get_column_preview,
    get_display_frame,
    get_excel_sheet_names,
    get_raw_preview,
    prepare_dataframe_from_raw,
    read_uploaded_file_raw,
)
from src.reporting import build_change_frequency, build_summary_df, export_to_excel

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Delta Analysis Tool",
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
        .footer-primary {
            text-align: center;
            font-size: 0.92rem;
            font-weight: 600;
            color: #1F4E79;
            margin-bottom: 0.2rem;
        }
        .footer-secondary {
            text-align: center;
            font-size: 0.82rem;
            font-weight: 500;
            color: #44546A;
            margin-bottom: 0.2rem;
        }
        .footer-tertiary {
            text-align: center;
            font-size: 0.74rem;
            font-weight: 400;
            color: #888;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def _csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def _show_df(df: pd.DataFrame, key: str = "") -> None:
    from src.io_utils import PREVIEW_MAX_ROWS
    display_df, truncated = get_display_frame(df, max_rows=PREVIEW_MAX_ROWS)
    if truncated:
        st.caption(
            f"Showing first {PREVIEW_MAX_ROWS:,} of {len(df):,} rows. "
            "Download the CSV for the complete dataset."
        )
    st.dataframe(display_df, use_container_width=True, hide_index=True, key=key or None)


def _csv_download(label: str, df: pd.DataFrame, filename: str) -> None:
    if df is not None and not df.empty:
        st.download_button(
            label=f"Download {label} (CSV)",
            data=_csv_bytes(df),
            file_name=filename,
            mime="text/csv",
            key=f"csv_{filename}",
        )
    else:
        st.caption("No records to download for this category.")


def _prep_controls(
    raw_df: pd.DataFrame,
    label: str,
    key_prefix: str,
):
    """
    Render data preparation controls for one dataset.
    Returns (prepared_df, metadata) or (None, None) on error.
    Row numbers shown to the user are 1-based; internally 0-based.
    """
    n_rows = len(raw_df)
    st.markdown(f"**{label}**")

    st.caption(
        f"Raw file preview (first {min(RAW_PREVIEW_ROWS, n_rows)} of {n_rows} rows). "
        "The 'Source Row' column shows 1-based row numbers as they appear in the file."
    )
    st.dataframe(
        get_raw_preview(raw_df, max_rows=RAW_PREVIEW_ROWS),
        use_container_width=True,
        hide_index=True,
    )

    # 1-based display; convert to 0-based for internal use
    header_row_display = int(st.number_input(
        "Header row",
        min_value=1,
        max_value=max(1, n_rows),
        value=1,
        step=1,
        key=f"{key_prefix}_header_row",
        help=(
            "Row number (starting from 1) that contains the column headers. "
            "Row 1 is the first row in the file."
        ),
    ))
    header_row = header_row_display - 1  # 0-based

    if 0 <= header_row < n_rows:
        header_vals = raw_df.iloc[header_row].tolist()
        preview_str = " | ".join(str(v)[:25] for v in header_vals[:12])
        if len(header_vals) > 12:
            preview_str += " ..."
        st.caption(f"Row {header_row_display} header preview: {preview_str}")

    drop_above = st.checkbox(
        "Drop rows above header",
        value=True,
        key=f"{key_prefix}_drop_above",
        help="Remove all rows that appear before the selected header row.",
    )
    drop_blank = st.checkbox(
        "Drop fully blank rows",
        value=True,
        key=f"{key_prefix}_drop_blank",
        help="Remove rows where every cell is empty or contains only whitespace.",
    )

    use_end_row = st.checkbox(
        "Apply end row limit",
        value=False,
        key=f"{key_prefix}_use_end",
        help=(
            "Stop reading data after a specific row. "
            "Use this to exclude footer rows, totals, or trailing notes."
        ),
    )
    end_row = None
    if use_end_row:
        end_row_display = int(st.number_input(
            "End row (inclusive)",
            min_value=header_row_display + 1,
            max_value=max(header_row_display + 1, n_rows),
            value=max(header_row_display + 1, n_rows),
            step=1,
            key=f"{key_prefix}_end_row",
            help="Last row number (1-based) to include in the data body.",
        ))
        end_row = end_row_display - 1  # 0-based

    try:
        prep_df, meta, warns = prepare_dataframe_from_raw(
            raw_df,
            header_row_index=header_row,
            drop_rows_above=drop_above,
            drop_blank_rows=drop_blank,
            end_row_index=end_row,
        )
        for w in warns:
            st.warning(w)
        st.caption(
            f"Prepared dataset: {meta['rows_in_prepared']:,} data row(s), "
            f"{meta['columns_in_prepared']} column(s)."
        )
        if not prep_df.empty:
            st.dataframe(
                get_column_preview(prep_df),
                use_container_width=True,
                hide_index=True,
            )
        return prep_df, meta
    except ValueError as exc:
        st.error(f"Preparation error: {exc}")
        return None, None


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown('<div class="main-header">Delta Analysis Tool</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">'
    "Upload a Source Dataset and a Comparison Dataset, define your match keys, and generate "
    "a full bidirectional reconciliation. Supports government financial reconciliation, "
    "receipt tracking, contract audits, and duplicate submission review."
    "</div>",
    unsafe_allow_html=True,
)
st.divider()

# ---------------------------------------------------------------------------
# Step 1: File Upload
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">Step 1: Upload Datasets</div>', unsafe_allow_html=True)

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**Source Dataset:** Prior period, source of record, or authoritative extract")
    file_a = st.file_uploader(
        "Upload Source Dataset",
        type=["csv", "xlsx", "xls"],
        key="upload_a",
        label_visibility="collapsed",
    )

with col_b:
    st.markdown("**Comparison Dataset:** Current period, received file, or updated extract")
    file_b = st.file_uploader(
        "Upload Comparison Dataset",
        type=["csv", "xlsx", "xls"],
        key="upload_b",
        label_visibility="collapsed",
    )

# Detect sheets for any uploaded Excel files
sheets_a: list[str] = get_excel_sheet_names(file_a) if file_a else []
sheets_b: list[str] = get_excel_sheet_names(file_b) if file_b else []

sheet_a_selected: str | None = sheets_a[0] if len(sheets_a) == 1 else None
sheet_b_selected: str | None = sheets_b[0] if len(sheets_b) == 1 else None

# ---------------------------------------------------------------------------
# Step 2: Select Sheet (only shown for Excel workbooks with multiple tabs)
# ---------------------------------------------------------------------------

if (file_a is not None or file_b is not None) and (len(sheets_a) > 1 or len(sheets_b) > 1):
    st.divider()
    st.markdown(
        '<div class="section-title">Step 2: Select Sheet</div>',
        unsafe_allow_html=True,
    )
    st.info(
        "Select the worksheet tab to use from each Excel workbook. "
        "Only tabs that contain the data applicable to this analysis should be selected."
    )
    sh_col_a, sh_col_b = st.columns(2)
    with sh_col_a:
        if len(sheets_a) > 1:
            sheet_a_selected = st.selectbox(
                "Worksheet: Source Dataset",
                options=sheets_a,
                key="sheet_a",
            )
        elif file_a:
            st.caption("Source Dataset: single sheet or CSV, no selection required.")
    with sh_col_b:
        if len(sheets_b) > 1:
            sheet_b_selected = st.selectbox(
                "Worksheet: Comparison Dataset",
                options=sheets_b,
                key="sheet_b",
            )
        elif file_b:
            st.caption("Comparison Dataset: single sheet or CSV, no selection required.")

# Read raw files (no UI step; happens automatically after sheet selection)
raw_a: pd.DataFrame | None = None
raw_b: pd.DataFrame | None = None
_large_file_blocked = False

if file_a:
    try:
        raw_a = read_uploaded_file_raw(file_a, sheet_name=sheet_a_selected)
        col_a.success(f"{len(raw_a):,} rows detected (header assignment in Step 3)")
        size_status, size_msg = check_file_size(len(raw_a))
        if size_status == "warn":
            col_a.warning(f"Source Dataset: {size_msg}")
        elif size_status == "hard":
            col_a.error(f"Source Dataset: {size_msg}")
            if not col_a.checkbox(
                "I understand the risk; proceed anyway (Source Dataset)", key="hard_a"
            ):
                _large_file_blocked = True
    except ValueError as exc:
        col_a.error(f"Source Dataset error: {exc}")

if file_b:
    try:
        raw_b = read_uploaded_file_raw(file_b, sheet_name=sheet_b_selected)
        col_b.success(f"{len(raw_b):,} rows detected (header assignment in Step 3)")
        size_status, size_msg = check_file_size(len(raw_b))
        if size_status == "warn":
            col_b.warning(f"Comparison Dataset: {size_msg}")
        elif size_status == "hard":
            col_b.error(f"Comparison Dataset: {size_msg}")
            if not col_b.checkbox(
                "I understand the risk; proceed anyway (Comparison Dataset)", key="hard_b"
            ):
                _large_file_blocked = True
    except ValueError as exc:
        col_b.error(f"Comparison Dataset error: {exc}")

# ---------------------------------------------------------------------------
# Step 2: Prepare Data
# ---------------------------------------------------------------------------

df_a: pd.DataFrame | None = None
df_b: pd.DataFrame | None = None
source_prep_metadata: dict | None = None
comparison_prep_metadata: dict | None = None

if raw_a is not None and raw_b is not None:
    st.divider()
    st.markdown(
        '<div class="section-title">Step 3: Prepare Headers and Rows</div>',
        unsafe_allow_html=True,
    )
    st.info(
        "Review the raw file contents below and assign the correct header row for each dataset. "
        "Rows above the header (report titles, agency metadata, blank separators) can be "
        "excluded automatically. Fully blank rows are removed by default. "
        "Use the end row setting to exclude footer rows, totals, or trailing notes."
    )

    prep_col_a, prep_col_b = st.columns(2)

    with prep_col_a:
        df_a, source_prep_metadata = _prep_controls(raw_a, "Source Dataset", "prep_a")

    with prep_col_b:
        df_b, comparison_prep_metadata = _prep_controls(raw_b, "Comparison Dataset", "prep_b")

# ---------------------------------------------------------------------------
# Step 3: Key column configuration
# ---------------------------------------------------------------------------

if df_a is not None and df_b is not None:
    st.divider()
    st.markdown(
        '<div class="section-title">Step 4: Configure Match Keys</div>',
        unsafe_allow_html=True,
    )
    st.info(
        "**Match key columns** uniquely identify each record (e.g., Contract Number, Case ID, "
        "Award ID). Select the corresponding key column(s) from each dataset. "
        "When using multiple columns, they are combined into a composite key and matched positionally."
    )

    k1, k2 = st.columns(2)
    with k1:
        key_cols_a = st.multiselect(
            "Match key column(s): Source Dataset",
            options=df_a.columns.tolist(),
            help="Column(s) that uniquely identify each record in the Source Dataset.",
        )
    with k2:
        key_cols_b = st.multiselect(
            "Match key column(s): Comparison Dataset",
            options=df_b.columns.tolist(),
            help="Column(s) that uniquely identify each record in the Comparison Dataset.",
        )

    if key_cols_a and key_cols_b:
        if len(key_cols_a) != len(key_cols_b):
            st.warning(
                f"Match key column count mismatch. Source: {len(key_cols_a)} column(s), "
                f"Comparison: {len(key_cols_b)} column(s). Counts must match."
            )
        else:
            st.markdown("**Key mapping (matched positionally):**")
            st.dataframe(
                pd.DataFrame({"Source Key": key_cols_a, "Comparison Key": key_cols_b}),
                use_container_width=False,
                hide_index=True,
            )

    # -----------------------------------------------------------------------
    # Step 4: Comparison columns
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown(
        '<div class="section-title">Step 5: Select Fields to Compare</div>',
        unsafe_allow_html=True,
    )
    st.info(
        "**Comparison fields** are non-key columns to diff across matched records (e.g., "
        "obligation amounts, status codes, dates). Leave blank to skip field-level change detection."
    )

    common_non_key = [
        c for c in df_a.columns
        if c in df_b.columns and c not in key_cols_a and c not in key_cols_b
    ]

    c1, c2 = st.columns(2)
    with c1:
        compare_cols_a = st.multiselect(
            "Comparison fields: Source Dataset",
            options=[c for c in df_a.columns if c not in key_cols_a],
            default=common_non_key,
        )
    with c2:
        compare_cols_b = st.multiselect(
            "Corresponding fields: Comparison Dataset",
            options=[c for c in df_b.columns if c not in key_cols_b],
            default=[c for c in common_non_key if c in df_b.columns],
        )

    if compare_cols_a and compare_cols_b and len(compare_cols_a) != len(compare_cols_b):
        st.warning(
            f"Comparison field count mismatch. Source: {len(compare_cols_a)}, "
            f"Comparison: {len(compare_cols_b)}. Counts must match."
        )
    elif compare_cols_a and compare_cols_b:
        st.markdown("**Field mapping (matched positionally):**")
        st.dataframe(
            pd.DataFrame({"Source Field": compare_cols_a, "Comparison Field": compare_cols_b}),
            use_container_width=False,
            hide_index=True,
        )

    # Advanced comparison settings
    comparison_rules: list[dict] | None = None
    if compare_cols_a and compare_cols_b and len(compare_cols_a) == len(compare_cols_b):
        col_pairs = list(zip(compare_cols_a, compare_cols_b))
        col_labels = [ca if ca == cb else f"{ca} / {cb}" for ca, cb in col_pairs]

        with st.expander("Advanced Comparison Settings (optional)", expanded=False):
            st.info(
                "By default, all fields are compared as **text**. "
                "Assign numeric tolerance or date-aware comparison to specific fields below. "
                "Fields not assigned here will continue to use exact text matching."
            )

            adv1, adv2 = st.columns(2)

            with adv1:
                st.markdown("**Numeric fields**")
                numeric_labels = st.multiselect(
                    "Fields to compare as numeric values",
                    options=col_labels,
                    key="adv_numeric",
                    help=(
                        "Currency symbols ($, £, €), commas, and parenthesised negatives "
                        "are stripped automatically before comparison."
                    ),
                )
                numeric_tolerance = st.number_input(
                    "Tolerance (applies to all numeric fields)",
                    min_value=0.0,
                    value=0.0,
                    step=0.01,
                    format="%.4f",
                    key="adv_tol",
                    help=(
                        "Maximum allowed absolute difference before flagging a record as changed. "
                        "Set to 0.00 for exact match. Example: 0.01 ignores rounding differences."
                    ),
                )

            with adv2:
                st.markdown("**Date fields**")
                date_labels = st.multiselect(
                    "Fields to compare as dates",
                    options=[lbl for lbl in col_labels if lbl not in numeric_labels],
                    key="adv_date",
                    help=(
                        "Parses both ISO (YYYY-MM-DD) and US (MM/DD/YYYY) date formats "
                        "before comparing. Mixed formats in the same column are handled automatically."
                    ),
                )
                date_mode = st.selectbox(
                    "Date precision (applies to all date fields)",
                    options=["date_only", "datetime_precision"],
                    key="adv_date_mode",
                    help=(
                        "date_only: compare calendar date only; time component is ignored "
                        "(e.g., 08:30 and 14:00 on the same date are treated as equal). "
                        "datetime_precision: time is included in the comparison."
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
    # Step 5: Run
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown(
        '<div class="section-title">Step 6: Run Analysis</div>',
        unsafe_allow_html=True,
    )

    run_ready = bool(
        key_cols_a and key_cols_b
        and len(key_cols_a) == len(key_cols_b)
        and not _large_file_blocked
    )
    if not run_ready and not _large_file_blocked:
        st.warning(
            "Select at least one match key column from each dataset (with matching counts) to proceed."
        )
    if _large_file_blocked:
        st.warning("Confirm the large-file warning above before running the analysis.")

    if st.button("Run Delta Analysis", type="primary", disabled=not run_ready):
        with st.spinner("Running reconciliation..."):
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
                st.session_state["result"]                   = result
                st.session_state["file_a_name"]              = file_a.name
                st.session_state["file_b_name"]              = file_b.name
                st.session_state["source_prep_metadata"]     = source_prep_metadata
                st.session_state["comparison_prep_metadata"] = comparison_prep_metadata
                st.success("Analysis complete. Review the results below.")
            except ValueError as exc:
                st.error(f"Analysis error: {exc}")
            except Exception as exc:
                st.error(f"Unexpected error: {exc}")

# ===========================================================================
# RESULTS SECTION
# ===========================================================================

if "result" in st.session_state:
    result: DeltaResult = st.session_state["result"]
    file_a_name: str    = st.session_state.get("file_a_name", "Source Dataset")
    file_b_name: str    = st.session_state.get("file_b_name", "Comparison Dataset")

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
    st.markdown(
        '<div class="section-title">Results Dashboard</div>',
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric(
        "Source Records", f"{result.total_a:,}",
        help=f"Total records loaded from {file_a_name}",
    )
    k2.metric(
        "Comparison Records", f"{result.total_b:,}",
        help=f"Total records loaded from {file_b_name}",
    )
    k3.metric(
        "Source Only", f"{n_only_a:,}",
        help=f"Records whose key appears in {file_a_name} but not in {file_b_name}",
    )
    k4.metric(
        "Comparison Only", f"{n_only_b:,}",
        help=f"Records whose key appears in {file_b_name} but not in {file_a_name}",
    )
    k5.metric(
        "Matched Records", f"{n_matched:,}",
        help="Records with a common key present in both datasets",
    )

    k6, k7, k8, k9, k10 = st.columns(5)
    k6.metric(
        "Changed Records", f"{n_changed:,}",
        help="Matched records where at least one compared field differs between datasets",
    )
    k7.metric(
        "Source Duplicate Keys", f"{n_dup_a:,}",
        help=f"Rows sharing a match key with at least one other row in {file_a_name}",
    )
    k8.metric(
        "Comparison Duplicate Keys", f"{n_dup_b:,}",
        help=f"Rows sharing a match key with at least one other row in {file_b_name}",
    )
    k9.metric(
        "Source Missing Keys", f"{n_blank_a:,}",
        help=f"Rows with a blank or null match key in {file_a_name} (excluded from reconciliation)",
    )
    k10.metric(
        "Comparison Missing Keys", f"{n_blank_b:,}",
        help=f"Rows with a blank or null match key in {file_b_name} (excluded from reconciliation)",
    )

    # -----------------------------------------------------------------------
    # Visualizations
    # Blue palette: #1F4E79 navy, #2F75B5 medium blue, #9DC3E6 light blue,
    #               #44546A slate, #F2F2F2 light gray
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown(
        '<div class="section-title">Visual Summary</div>',
        unsafe_allow_html=True,
    )

    viz1, viz2, viz3 = st.columns([2, 1.4, 1.6])

    # Chart 1: Reconciliation Summary bar chart
    with viz1:
        st.markdown("**Reconciliation Summary**")
        cat_labels = [
            "Source Only",
            "Comparison Only",
            "Matched",
            "Changed",
            "Source Duplicates",
            "Comparison Duplicates",
            "Missing (Source)",
            "Missing (Comparison)",
        ]
        cat_values = [
            n_only_a, n_only_b, n_matched,
            n_changed, n_dup_a, n_dup_b,
            n_blank_a, n_blank_b,
        ]
        cat_colors = [
            "#1F4E79",
            "#2F75B5",
            "#9DC3E6",
            "#1F4E79",
            "#44546A",
            "#44546A",
            "#9DC3E6",
            "#9DC3E6",
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
            yaxis_title="Record Count",
            plot_bgcolor="white",
            paper_bgcolor="white",
            yaxis=dict(gridcolor="#EEEEEE"),
            font=dict(size=11),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # Chart 2: Match Coverage donut
    with viz2:
        st.markdown("**Match Coverage**")
        pie_labels = ["Matched", "Source Only", "Comparison Only"]
        pie_values = [n_matched, n_only_a, n_only_b]
        pie_colors = ["#1F4E79", "#2F75B5", "#9DC3E6"]

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

    # Chart 3: Field-Level Differences frequency
    with viz3:
        st.markdown("**Field-Level Differences**")
        freq_df = build_change_frequency(result)
        if freq_df.empty:
            st.info("Select comparison fields and re-run to see field-level difference frequency.")
        else:
            fig_freq = go.Figure(go.Bar(
                x=freq_df["Changes"],
                y=freq_df["Field"],
                orientation="h",
                marker_color="#2F75B5",
                text=freq_df["Changes"],
                textposition="outside",
                hovertemplate="%{y}: %{x:,} difference(s)<extra></extra>",
            ))
            fig_freq.update_layout(
                margin=dict(l=20, r=40, t=10, b=30),
                height=320,
                xaxis_title="Number of Differences",
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
    st.markdown(
        '<div class="section-title">Detailed Results</div>',
        unsafe_allow_html=True,
    )

    tabs = st.tabs([
        "Summary",
        f"Source Only ({n_only_a:,})",
        f"Comparison Only ({n_only_b:,})",
        f"Matched ({n_matched:,})",
        f"Changed Records ({n_changed:,})",
        f"Source Duplicates ({n_dup_a:,})",
        f"Comparison Duplicates ({n_dup_b:,})",
        "Data Quality Flags",
    ])

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Tab 0: Summary
    with tabs[0]:
        st.dataframe(build_summary_df(result), use_container_width=True, hide_index=True)

    # Tab 1: Source Only
    with tabs[1]:
        st.markdown(
            f"**{n_only_a:,} record(s)** whose match key appears in **{file_a_name}** "
            f"but has no corresponding entry in **{file_b_name}**. "
            "These may represent withdrawn entries, pending submissions, or records not yet "
            "reflected in the comparison dataset."
        )
        if not result.only_in_a.empty:
            _show_df(result.only_in_a, key="only_a")
        else:
            st.info("No source-only records found.")
        _csv_download("Source Only Records", result.only_in_a, f"source_only_{ts}.csv")

    # Tab 2: Comparison Only
    with tabs[2]:
        st.markdown(
            f"**{n_only_b:,} record(s)** whose match key appears in **{file_b_name}** "
            f"but has no corresponding entry in **{file_a_name}**. "
            "These may represent new submissions, late arrivals, or records not yet "
            "posted to the source dataset."
        )
        if not result.only_in_b.empty:
            _show_df(result.only_in_b, key="only_b")
        else:
            st.info("No comparison-only records found.")
        _csv_download("Comparison Only Records", result.only_in_b, f"comparison_only_{ts}.csv")

    # Tab 3: Matched
    with tabs[3]:
        st.markdown(
            f"**{n_matched:,} record(s)** with a matching key present in both datasets. "
            "Columns prefixed **Source:** show values from the source dataset; "
            "columns prefixed **Comparison:** show values from the comparison dataset."
        )
        if not result.matched.empty:
            _show_df(result.matched, key="matched")
        else:
            st.info("No matched records found.")
        _csv_download("Matched Records", result.matched, f"matched_{ts}.csv")

    # Tab 4: Changed Records
    with tabs[4]:
        if not result.compare_cols_a:
            st.info(
                "No comparison fields were selected. "
                "Re-run the analysis with comparison fields configured to see field-level differences."
            )
        elif result.changed.empty:
            st.success(
                "No field-level differences detected among matched records. "
                "All compared fields are identical between datasets."
            )
        else:
            st.markdown(
                f"**{n_changed:,} matched record(s)** where at least one compared field differs. "
                "Each row shows the **Source** and **Comparison** values for every changed field."
            )
            _show_df(result.changed, key="changed")
        _csv_download("Changed Records", result.changed, f"changed_records_{ts}.csv")
        if result.compare_parse_issues is not None and not result.compare_parse_issues.empty:
            with st.expander(
                f"{len(result.compare_parse_issues)} parse warning(s): values that could not "
                "be interpreted as the configured type",
                expanded=False,
            ):
                st.dataframe(result.compare_parse_issues, use_container_width=True, hide_index=True)

    # Tab 5: Source Duplicates
    with tabs[5]:
        st.markdown(
            f"**{n_dup_a:,} row(s)** in **{file_a_name}** that share a match key with at least "
            "one other row in the same dataset. Only the first occurrence of each duplicate key "
            "was used in the matching process. Duplicate submissions should be resolved at the "
            "source system before official reporting."
        )
        if not result.duplicates_a.empty:
            _show_df(result.duplicates_a, key="dup_a")
        else:
            st.success("No duplicate identifiers detected in the Source Dataset.")
        _csv_download(
            "Source Duplicate Keys", result.duplicates_a, f"source_duplicates_{ts}.csv"
        )

    # Tab 6: Comparison Duplicates
    with tabs[6]:
        st.markdown(
            f"**{n_dup_b:,} row(s)** in **{file_b_name}** that share a match key with at least "
            "one other row in the same dataset. Only the first occurrence of each duplicate key "
            "was used in the matching process. Duplicate submissions should be resolved at the "
            "source system before official reporting."
        )
        if not result.duplicates_b.empty:
            _show_df(result.duplicates_b, key="dup_b")
        else:
            st.success("No duplicate identifiers detected in the Comparison Dataset.")
        _csv_download(
            "Comparison Duplicate Keys", result.duplicates_b, f"comparison_duplicates_{ts}.csv"
        )

    # Tab 7: Data Quality Flags
    with tabs[7]:
        st.markdown(
            "**Rows excluded from reconciliation** due to a blank or null match key value. "
            "These records cannot be matched and are flagged here for source-system correction."
        )
        dq_rows: list[dict] = []
        for _, row in result.blank_keys_a.iterrows():
            dq_rows.append({"Dataset": "Source", "Filename": file_a_name,
                             "Flag": "Missing Identifier", **row.to_dict()})
        for _, row in result.blank_keys_b.iterrows():
            dq_rows.append({"Dataset": "Comparison", "Filename": file_b_name,
                             "Flag": "Missing Identifier", **row.to_dict()})
        dq_df = pd.DataFrame(dq_rows) if dq_rows else pd.DataFrame(
            columns=["Dataset", "Filename", "Flag"]
        )
        if not dq_df.empty:
            _show_df(dq_df, key="dq")
        else:
            st.success("No data quality flags detected. All records carry valid match key values.")
        _csv_download("Data Quality Flags", dq_df, f"data_quality_flags_{ts}.csv")

    # -----------------------------------------------------------------------
    # Export
    # -----------------------------------------------------------------------

    st.divider()
    st.markdown(
        '<div class="section-title">Export Results</div>',
        unsafe_allow_html=True,
    )

    exp_col, info_col = st.columns([1, 3])

    with exp_col:
        with st.spinner("Building Excel workbook..."):
            excel_bytes = export_to_excel(
                result,
                file_a_name,
                file_b_name,
                source_prep_metadata=st.session_state.get("source_prep_metadata"),
                comparison_prep_metadata=st.session_state.get("comparison_prep_metadata"),
            )

        st.download_button(
            label="Download Excel Report",
            data=excel_bytes,
            file_name=f"delta_analysis_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    with info_col:
        st.markdown(
            """
            The Excel workbook contains **11 tabs**:
            - **Executive Summary:** auto-generated plain-English briefing narrative
            - **Analysis Metadata:** dataset names, sheets, record counts, key columns, timestamp, and header/row preparation settings
            - **Comparison Rules:** per-field comparison type, tolerance, and date precision settings
            - **Delta Counts:** flat count table suitable for pivot tables and downstream reporting
            - **Source Only Records:** records present in the source but absent from comparison
            - **Comparison Only Records:** records present in the comparison but absent from source
            - **Matched Records:** side-by-side view of all matched records
            - **Changed Records:** before/after values for every field-level difference
            - **Source Duplicate Keys / Comparison Duplicate Keys:** rows with non-unique match keys
            - **Data Quality Flags:** rows excluded due to blank or null match key values
            """
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.divider()
st.markdown(
    "<div class='footer-primary'>"
    "Delta Analysis Tool &nbsp;|&nbsp; Operational Analytics and Reconciliation Platform"
    " &nbsp;|&nbsp; Version 1.2.0"
    "</div>"
    "<div class='footer-secondary'>Designed and Developed by Richie Garafola</div>"
    "<div class='footer-tertiary'>"
    "Supports CSV and Excel comparison, multi-sheet workbooks, numeric tolerance matching,"
    " date-aware comparisons, and audit-ready reporting."
    "</div>",
    unsafe_allow_html=True,
)
