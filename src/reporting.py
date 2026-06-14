"""
Excel reporting module.

Builds an audit-grade, multi-tab Excel workbook from a DeltaResult.
Includes an Executive Narrative tab with auto-generated plain-English
summary language suitable for leadership briefings.
"""
from __future__ import annotations

import io
from datetime import datetime
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from src.delta_engine import DeltaResult


# ---------------------------------------------------------------------------
# Color palette
# ---------------------------------------------------------------------------

COLORS = {
    "navy":        "1F4E79",
    "dark_red":    "C00000",
    "dark_green":  "375623",
    "amber":       "7B3F00",
    "orange":      "843C0C",
    "purple":      "403152",
    "slate":       "44546A",
    "light_blue":  "BDD7EE",
    "light_red":   "FCE4D6",
    "light_green": "E2EFDA",
    "light_amber": "FFF2CC",
    "white":       "FFFFFF",
    "light_gray":  "F2F2F2",
}

_HEADER_COLORS = {
    "Summary":               COLORS["navy"],
    "Executive Summary":     COLORS["slate"],
    "Analysis Metadata":     COLORS["slate"],
    "Comparison Rules":      COLORS["slate"],
    "Delta Counts":          COLORS["navy"],
    "Only in File A":        COLORS["dark_red"],
    "Only in File B":        COLORS["dark_green"],
    "Matched Records":       COLORS["navy"],
    "Changed Records":       COLORS["amber"],
    "Duplicate Keys File A": COLORS["orange"],
    "Duplicate Keys File B": COLORS["orange"],
    "Data Quality Issues":   COLORS["purple"],
}

_ROW_FILL_COLORS = {
    "Only in File A":        COLORS["light_red"],
    "Only in File B":        COLORS["light_green"],
    "Changed Records":       COLORS["light_amber"],
    "Duplicate Keys File A": COLORS["light_amber"],
    "Duplicate Keys File B": COLORS["light_amber"],
}


# ---------------------------------------------------------------------------
# Public: summary DataFrame
# ---------------------------------------------------------------------------

def build_summary_df(result: DeltaResult) -> pd.DataFrame:
    """Produce a human-readable summary statistics table."""
    total_a = result.total_a
    total_b = result.total_b
    n_matched = len(result.matched)

    def pct(n: int, denom: int) -> str:
        if denom == 0:
            return "N/A"
        return f"{n / denom * 100:.1f}%"

    rows = [
        ("File A — Total Records",          total_a,                   ""),
        ("File B — Total Records",          total_b,                   ""),
        ("Records Only in File A",          len(result.only_in_a),     pct(len(result.only_in_a), total_a)),
        ("Records Only in File B",          len(result.only_in_b),     pct(len(result.only_in_b), total_b)),
        ("Records Matched (common keys)",   n_matched,                 pct(n_matched, total_a)),
        ("Records with Field Changes",      len(result.changed),       pct(len(result.changed), n_matched) if n_matched else "N/A"),
        ("Duplicate Key Rows — File A",     len(result.duplicates_a),  pct(len(result.duplicates_a), total_a)),
        ("Duplicate Key Rows — File B",     len(result.duplicates_b),  pct(len(result.duplicates_b), total_b)),
        ("Blank / Null Key Rows — File A",  len(result.blank_keys_a),  pct(len(result.blank_keys_a), total_a)),
        ("Blank / Null Key Rows — File B",  len(result.blank_keys_b),  pct(len(result.blank_keys_b), total_b)),
    ]

    return pd.DataFrame(rows, columns=["Metric", "Count", "% of File Total"])


# ---------------------------------------------------------------------------
# Public: field-change frequency
# ---------------------------------------------------------------------------

def build_change_frequency(result: DeltaResult) -> pd.DataFrame:
    """Count how many times each compared field changed."""
    if result.changed.empty or not result.compare_cols_a:
        return pd.DataFrame(columns=["Field", "Changes"])

    rows = []
    for col in result.compare_cols_a:
        a_col = f"{col} — File A"
        b_col = f"{col} — File B"
        if a_col in result.changed.columns and b_col in result.changed.columns:
            n_changed = (result.changed[a_col] != result.changed[b_col]).sum()
            rows.append({"Field": col, "Changes": int(n_changed)})

    df = pd.DataFrame(rows).sort_values("Changes", ascending=False).reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# Public: Excel export
# ---------------------------------------------------------------------------

def export_to_excel(
    result: DeltaResult,
    file_a_name: str = "File A",
    file_b_name: str = "File B",
) -> bytes:
    """
    Build a polished, multi-tab Excel workbook from a DeltaResult.
    Returns raw bytes for Streamlit download.
    """
    run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    summary_df   = build_summary_df(result)
    narrative    = _build_narrative(result, file_a_name, file_b_name, run_timestamp)
    delta_counts = _build_delta_counts_df(result)

    # Data Quality Issues — blank keys only (duplicates have their own tabs)
    dq_rows: list[dict] = []
    for _, row in result.blank_keys_a.iterrows():
        dq_rows.append({"Source File": file_a_name, "Issue Type": "Blank / Null Key", **row.to_dict()})
    for _, row in result.blank_keys_b.iterrows():
        dq_rows.append({"Source File": file_b_name, "Issue Type": "Blank / Null Key", **row.to_dict()})
    dq_df = pd.DataFrame(dq_rows) if dq_rows else pd.DataFrame(columns=["Source File", "Issue Type"])

    metadata_df = _build_metadata_df(result, file_a_name, file_b_name, run_timestamp)
    rules_df    = _build_rules_df(result)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary_df.to_excel(writer,             sheet_name="Summary",               index=False)
        _write_narrative(narrative, writer)
        metadata_df.to_excel(writer,            sheet_name="Analysis Metadata",     index=False)
        rules_df.to_excel(writer,               sheet_name="Comparison Rules",      index=False)
        delta_counts.to_excel(writer,           sheet_name="Delta Counts",          index=False)
        _write_sheet(result.only_in_a,    writer, "Only in File A")
        _write_sheet(result.only_in_b,    writer, "Only in File B")
        _write_sheet(result.matched,      writer, "Matched Records")
        _write_sheet(result.changed,      writer, "Changed Records")
        _write_sheet(result.duplicates_a, writer, "Duplicate Keys File A")
        _write_sheet(result.duplicates_b, writer, "Duplicate Keys File B")
        _write_sheet(dq_df,               writer, "Data Quality Issues")

    # Post-process: styling
    buffer.seek(0)
    wb = load_workbook(buffer)
    _apply_styles(wb)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------------------------------------------------------------------------
# Narrative builder
# ---------------------------------------------------------------------------

def _build_narrative(
    result: DeltaResult,
    file_a_name: str,
    file_b_name: str,
    run_timestamp: str,
) -> list[tuple[str, str]]:
    """
    Return a list of (label, text) pairs for the Executive Narrative tab.
    All numbers are filled in from the actual DeltaResult.
    """
    total_a   = result.total_a
    total_b   = result.total_b
    n_only_a  = len(result.only_in_a)
    n_only_b  = len(result.only_in_b)
    n_matched = len(result.matched)
    n_changed = len(result.changed)
    n_dup_a   = len(result.duplicates_a)
    n_dup_b   = len(result.duplicates_b)
    n_blank_a = len(result.blank_keys_a)
    n_blank_b = len(result.blank_keys_b)

    def pct(n, d):
        return f"{n / d * 100:.1f}%" if d else "N/A"

    key_label = " + ".join(result.key_cols_a) if result.key_cols_a else "configured key(s)"

    overview = (
        f"This analysis compared {total_a:,} records from {file_a_name} against "
        f"{total_b:,} records from {file_b_name} using the key field(s): {key_label}. "
        f"The analysis was run on {run_timestamp}."
    )

    matching = (
        f"{n_matched:,} records ({pct(n_matched, total_a)} of File A) were successfully matched "
        f"between both files. "
        f"{n_only_a:,} records ({pct(n_only_a, total_a)} of File A) were present only in {file_a_name} "
        f"and may represent deleted, withdrawn, or unsubmitted entries. "
        f"{n_only_b:,} records ({pct(n_only_b, total_b)} of File B) were present only in {file_b_name} "
        f"and may represent new submissions or entries not yet in the baseline."
    )

    if result.compare_cols_a:
        compare_label = ", ".join(result.compare_cols_a)
        changes = (
            f"Of the {n_matched:,} matched records, {n_changed:,} ({pct(n_changed, n_matched)}) "
            f"had at least one field-level difference in the compared columns: {compare_label}. "
            f"These records are listed in the 'Changed Records' tab with before/after values for each changed field."
        )
    else:
        changes = (
            "No comparison columns were configured for this run. "
            "Field-level change detection was not performed. "
            "Re-run the analysis with comparison columns selected to identify value changes."
        )

    if n_dup_a > 0 or n_dup_b > 0:
        duplicates = (
            f"{n_dup_a:,} rows in {file_a_name} and {n_dup_b:,} rows in {file_b_name} "
            f"share a key with at least one other row in the same file. "
            f"Duplicate submissions or data entry errors should be reviewed before this data "
            f"is used for official reporting. Only the first occurrence of each duplicate key "
            f"was used in the matching process."
        )
    else:
        duplicates = (
            f"No duplicate keys were detected in either file. "
            f"Both datasets appear to have unique records per key."
        )

    if n_blank_a > 0 or n_blank_b > 0:
        blanks = (
            f"{n_blank_a:,} rows in {file_a_name} and {n_blank_b:,} rows in {file_b_name} "
            f"had blank or null values in the key field(s) and could not be matched. "
            f"These rows are listed in the 'Data Quality Issues' tab and should be corrected "
            f"in the source system before the next analysis cycle."
        )
    else:
        blanks = (
            f"No blank or null key values were detected. "
            f"All rows in both files had valid key values."
        )

    recommendation = (
        f"Recommended next steps: (1) Resolve the {n_only_a:,} records only in {file_a_name} "
        f"by confirming whether they were intentionally removed or require resubmission. "
        f"(2) Validate the {n_only_b:,} records only in {file_b_name} against the authoritative "
        f"source of record. "
        f"(3) Review the {n_changed:,} changed records for unauthorized modifications or "
        f"legitimate updates. "
        f"(4) Correct all data quality issues (duplicates and blank keys) in the source system."
    )

    return [
        ("Analysis Overview",    overview),
        ("Matching Results",     matching),
        ("Field-Level Changes",  changes),
        ("Duplicate Key Review", duplicates),
        ("Blank Key Review",     blanks),
        ("Recommended Actions",  recommendation),
    ]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _build_delta_counts_df(result: DeltaResult) -> pd.DataFrame:
    """Flat table of every delta count — useful for quick reference."""
    return pd.DataFrame([
        {"Category": "File A — Total Records",          "Count": result.total_a,              "File": "A"},
        {"Category": "File B — Total Records",          "Count": result.total_b,              "File": "B"},
        {"Category": "Only in File A",                  "Count": len(result.only_in_a),       "File": "A"},
        {"Category": "Only in File B",                  "Count": len(result.only_in_b),       "File": "B"},
        {"Category": "Matched Records",                 "Count": len(result.matched),         "File": "Both"},
        {"Category": "Changed Records",                 "Count": len(result.changed),         "File": "Both"},
        {"Category": "Duplicate Key Rows — File A",     "Count": len(result.duplicates_a),    "File": "A"},
        {"Category": "Duplicate Key Rows — File B",     "Count": len(result.duplicates_b),    "File": "B"},
        {"Category": "Blank / Null Key Rows — File A",  "Count": len(result.blank_keys_a),    "File": "A"},
        {"Category": "Blank / Null Key Rows — File B",  "Count": len(result.blank_keys_b),    "File": "B"},
    ])


def _write_sheet(df: pd.DataFrame, writer: pd.ExcelWriter, sheet_name: str) -> None:
    if df is not None and not df.empty:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        pd.DataFrame([["No records found for this category."]]).to_excel(
            writer, sheet_name=sheet_name, index=False, header=False
        )


def _build_metadata_df(
    result: DeltaResult,
    file_a_name: str,
    file_b_name: str,
    run_timestamp: str,
) -> pd.DataFrame:
    """Tabular run metadata for the Analysis Metadata tab."""
    rows = [
        ("Run Timestamp",            run_timestamp),
        ("File A Name",              file_a_name),
        ("File A Sheet",             result.sheet_a or "N/A (first sheet / CSV)"),
        ("File A Row Count",         str(result.total_a)),
        ("File B Name",              file_b_name),
        ("File B Sheet",             result.sheet_b or "N/A (first sheet / CSV)"),
        ("File B Row Count",         str(result.total_b)),
        ("Key Columns (File A)",     ", ".join(result.key_cols_a) if result.key_cols_a else "None"),
        ("Key Columns (File B)",     ", ".join(result.key_cols_b) if result.key_cols_b else "None"),
        ("Comparison Columns (A)",   ", ".join(result.compare_cols_a) if result.compare_cols_a else "None"),
        ("Comparison Columns (B)",   ", ".join(result.compare_cols_b) if result.compare_cols_b else "None"),
        ("Comparison Rule Count",    str(len(result.comparison_rules))),
        ("Parse Issues Detected",    str(len(result.compare_parse_issues)) if result.compare_parse_issues is not None else "0"),
    ]
    return pd.DataFrame(rows, columns=["Parameter", "Value"])


def _build_rules_df(result: DeltaResult) -> pd.DataFrame:
    """Tabular view of every comparison rule."""
    if not result.comparison_rules:
        return pd.DataFrame(
            [["No comparison rules configured."]],
            columns=["Note"],
        )
    rows = []
    for rule in result.comparison_rules:
        rows.append({
            "Column (File A)":  rule.get("column_a", ""),
            "Column (File B)":  rule.get("column_b", ""),
            "Type":             rule.get("type", "text"),
            "Tolerance":        rule.get("tolerance", "") if rule.get("tolerance") is not None else "",
            "Date Mode":        rule.get("date_mode", "") if rule.get("date_mode") is not None else "",
        })
    return pd.DataFrame(rows)


def _write_narrative(
    narrative: list[tuple[str, str]],
    writer: pd.ExcelWriter,
) -> None:
    """Write the Executive Summary as a two-column label/text sheet."""
    rows = [{"Section": label, "Narrative": text} for label, text in narrative]
    pd.DataFrame(rows).to_excel(writer, sheet_name="Executive Summary", index=False)


def _apply_styles(wb) -> None:
    """Apply header colors, row alternation, column widths, and frozen panes."""
    white_bold = Font(color="FFFFFF", bold=True, size=11)
    thin_border = Border(
        bottom=Side(style="thin", color="CCCCCC")
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        fill_hex = _HEADER_COLORS.get(sheet_name, COLORS["navy"])
        header_fill = PatternFill(fill_type="solid", fgColor=fill_hex)
        alt_fill = PatternFill(fill_type="solid", fgColor=COLORS["light_gray"])
        row_highlight = _ROW_FILL_COLORS.get(sheet_name)
        highlight_fill = PatternFill(fill_type="solid", fgColor=row_highlight) if row_highlight else None

        # Header row
        for cell in ws[1]:
            cell.font = white_bold
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 22

        # Data rows
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            use_fill = highlight_fill if highlight_fill else (alt_fill if row_idx % 2 == 0 else None)
            for cell in row:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                cell.border = thin_border
                if use_fill:
                    cell.fill = use_fill

        # Narrative tab: wider text column, taller rows
        if sheet_name == "Executive Summary":
            ws.column_dimensions["A"].width = 28
            ws.column_dimensions["B"].width = 100
            for row_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = 80
                for cell in ws[row_idx]:
                    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        else:
            # Auto-fit all other columns (capped at 60 chars)
            for col_idx, col_cells in enumerate(ws.columns, start=1):
                max_len = 0
                for cell in col_cells:
                    try:
                        cell_len = len(str(cell.value)) if cell.value is not None else 0
                        max_len = max(max_len, cell_len)
                    except Exception:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 60)

        ws.freeze_panes = ws["A2"]
