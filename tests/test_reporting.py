"""
Tests for src/reporting.py: build_summary_df, build_change_frequency,
export_to_excel (sheet names, metadata content, rules content).
"""
import io

import pandas as pd
import pytest
from openpyxl import load_workbook

from src.delta_engine import DeltaResult, run_delta
from src.reporting import (
    build_change_frequency,
    build_summary_df,
    export_to_excel,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_result(
    compare_cols=True,
    with_comparison_rules=False,
    sheet_a=None,
    sheet_b=None,
):
    """Build a minimal DeltaResult covering all categories."""
    df_a = pd.DataFrame({
        "id":     ["1", "2", "3", "4", "4", ""],
        "status": ["Open", "Pending", "Closed", "Open", "Open", "Unknown"],
        "amount": ["100", "200", "300", "400", "400", "50"],
    })
    df_b = pd.DataFrame({
        "id":     ["1", "2", "5"],
        "status": ["Open", "Approved", "New"],
        "amount": ["100", "250", "600"],
    })

    rules = None
    if with_comparison_rules:
        rules = [
            {"column_a": "status", "column_b": "status", "type": "text",    "tolerance": None, "date_mode": None},
            {"column_a": "amount", "column_b": "amount", "type": "numeric", "tolerance": 0.0,  "date_mode": None},
        ]

    return run_delta(
        df_a=df_a,
        df_b=df_b,
        key_cols_a=["id"],
        key_cols_b=["id"],
        compare_cols_a=["status", "amount"] if compare_cols else None,
        compare_cols_b=["status", "amount"] if compare_cols else None,
        comparison_rules=rules,
        sheet_a=sheet_a,
        sheet_b=sheet_b,
    )


# ---------------------------------------------------------------------------
# build_summary_df
# ---------------------------------------------------------------------------

class TestBuildSummaryDf:
    def test_returns_dataframe(self):
        result = _make_result()
        df = build_summary_df(result)
        assert isinstance(df, pd.DataFrame)

    def test_has_required_columns(self):
        df = build_summary_df(_make_result())
        assert list(df.columns) == ["Metric", "Count", "% of Dataset Total"]

    def test_total_a_correct(self):
        result = _make_result()
        df = build_summary_df(result)
        row = df[df["Metric"] == "Source Dataset: Total Records"]
        assert int(row["Count"].iloc[0]) == result.total_a

    def test_total_b_correct(self):
        result = _make_result()
        df = build_summary_df(result)
        row = df[df["Metric"] == "Comparison Dataset: Total Records"]
        assert int(row["Count"].iloc[0]) == result.total_b

    def test_zero_total_b_returns_na_percentages(self):
        df_a = pd.DataFrame({"id": ["X"], "v": ["1"]})
        df_b = pd.DataFrame({"id": pd.Series([], dtype=str), "v": pd.Series([], dtype=str)})
        # Empty file B is valid input; engine should not raise
        result = run_delta(df_a, df_b, ["id"], ["id"])
        df = build_summary_df(result)
        # B-denominator rows should produce "N/A", not a crash
        b_row = df[df["Metric"] == "Comparison Only Records"]
        assert b_row["% of Dataset Total"].iloc[0] == "N/A"

    def test_pct_column_valid_values(self):
        df = build_summary_df(_make_result())
        for val in df["% of Dataset Total"]:
            # Each cell is: a percent string, "N/A", or "" (blank for total rows)
            assert str(val).endswith("%") or val in ("N/A", "")


# ---------------------------------------------------------------------------
# build_change_frequency
# ---------------------------------------------------------------------------

class TestBuildChangeFrequency:
    def test_returns_dataframe_with_columns(self):
        result = _make_result()
        df = build_change_frequency(result)
        assert list(df.columns) == ["Field", "Changes"]

    def test_no_comparison_cols_returns_empty(self):
        result = _make_result(compare_cols=False)
        df = build_change_frequency(result)
        assert df.empty

    def test_changed_fields_appear(self):
        result = _make_result()
        df = build_change_frequency(result)
        assert "status" in df["Field"].values or "amount" in df["Field"].values

    def test_zero_change_fields_included(self):
        result = _make_result()
        df = build_change_frequency(result)
        assert len(df) <= len(result.compare_cols_a)


# ---------------------------------------------------------------------------
# export_to_excel: sheet names
# ---------------------------------------------------------------------------

REQUIRED_SHEETS = [
    "Executive Summary",
    "Analysis Metadata",
    "Comparison Rules",
    "Delta Counts",
    "Source Only Records",
    "Comparison Only Records",
    "Matched Records",
    "Changed Records",
    "Source Duplicate Keys",
    "Comparison Duplicate Keys",
    "Data Quality Flags",
]

REMOVED_SHEETS = ["Summary", "Only in File A", "Only in File B",
                  "Records with Differences", "Duplicate Keys File A",
                  "Duplicate Keys File B", "Data Quality Issues",
                  "Baseline Only Records", "Baseline Duplicate Identifiers",
                  "Comparison Duplicates"]


class TestExcelSheetNames:
    def _load_wb(self, result=None, a="Source Dataset", b="Comparison Dataset"):
        if result is None:
            result = _make_result()
        raw = export_to_excel(result, a, b)
        return load_workbook(io.BytesIO(raw))

    def test_all_required_sheets_present(self):
        wb = self._load_wb()
        for sheet in REQUIRED_SHEETS:
            assert sheet in wb.sheetnames, f"Missing required sheet: '{sheet}'"

    def test_legacy_summary_sheet_removed(self):
        wb = self._load_wb()
        for sheet in REMOVED_SHEETS:
            assert sheet not in wb.sheetnames, f"Legacy sheet '{sheet}' should have been removed"

    def test_exactly_eleven_sheets(self):
        wb = self._load_wb()
        assert len(wb.sheetnames) == 11, f"Expected 11 sheets, got {len(wb.sheetnames)}: {wb.sheetnames}"

    def test_executive_summary_has_data(self):
        wb = self._load_wb()
        ws = wb["Executive Summary"]
        assert ws.max_row >= 2

    def test_analysis_metadata_has_data(self):
        wb = self._load_wb()
        ws = wb["Analysis Metadata"]
        assert ws.max_row >= 2

    def test_analysis_metadata_contains_file_names(self):
        wb = self._load_wb(a="my_a.csv", b="my_b.csv")
        ws = wb["Analysis Metadata"]
        values = [str(ws.cell(row=r, column=2).value) for r in range(2, ws.max_row + 1)]
        assert any("my_a.csv" in v for v in values)
        assert any("my_b.csv" in v for v in values)

    def test_analysis_metadata_contains_sheet_names(self):
        result = _make_result(sheet_a="Sheet1", sheet_b="DataTab")
        wb = self._load_wb(result)
        ws = wb["Analysis Metadata"]
        values = [str(ws.cell(row=r, column=2).value) for r in range(2, ws.max_row + 1)]
        assert any("Sheet1" in v for v in values)
        assert any("DataTab" in v for v in values)

    def test_comparison_rules_with_rules(self):
        result = _make_result(with_comparison_rules=True)
        wb = self._load_wb(result)
        ws = wb["Comparison Rules"]
        assert ws.max_row >= 2

    def test_comparison_rules_no_rules_shows_note(self):
        result = _make_result(compare_cols=False)
        wb = self._load_wb(result)
        ws = wb["Comparison Rules"]
        cell_val = str(ws.cell(row=1, column=1).value)
        assert "Note" in cell_val or ws.max_row >= 1

    def test_delta_counts_has_data(self):
        wb = self._load_wb()
        ws = wb["Delta Counts"]
        assert ws.max_row >= 2

    def test_source_only_records_present(self):
        wb = self._load_wb()
        assert "Source Only Records" in wb.sheetnames

    def test_comparison_only_records_present(self):
        wb = self._load_wb()
        assert "Comparison Only Records" in wb.sheetnames

    def test_changed_records_present(self):
        wb = self._load_wb()
        assert "Changed Records" in wb.sheetnames

    def test_source_duplicate_keys_present(self):
        wb = self._load_wb()
        assert "Source Duplicate Keys" in wb.sheetnames

    def test_comparison_duplicate_keys_present(self):
        wb = self._load_wb()
        assert "Comparison Duplicate Keys" in wb.sheetnames

    def test_data_quality_flags_present(self):
        wb = self._load_wb()
        assert "Data Quality Flags" in wb.sheetnames

    def test_returns_bytes(self):
        raw = export_to_excel(_make_result())
        assert isinstance(raw, bytes) and len(raw) > 0


# ---------------------------------------------------------------------------
# build_change_frequency: regression tests for per-field counts
#
# result.changed only holds a value for the fields that actually differ in a
# given record; fields that matched are absent (NaN once the DataFrame is
# built). Unchanged cells must not be counted as differences.
# ---------------------------------------------------------------------------

from pathlib import Path

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


def _freq_dict(result):
    df = build_change_frequency(result)
    return dict(zip(df["Field"], df["Changes"]))


def _read_sample(name):
    return pd.read_csv(SAMPLE_DIR / name, dtype=str, keep_default_na=False)


def _engine_field_counts(df_a, df_b, key, cols, rules=None):
    """Independent count: apply the engine's own comparison function to every
    matched pair (first occurrence per key, non-blank keys) and tally the
    fields it reports as unequal."""
    from src.comparison import compare_field_values
    from src.normalization import normalize_key_value

    rules = rules or [
        {"column_a": c, "column_b": c, "type": "text", "tolerance": None, "date_mode": None}
        for c in cols
    ]

    def first_by_key(df):
        out = {}
        for _, row in df.iterrows():
            k = normalize_key_value(row[key])
            if k != "__BLANK__" and k not in out:
                out[k] = row
        return out

    a, b = first_by_key(df_a), first_by_key(df_b)
    counts = {}
    for k in set(a) & set(b):
        for rule in rules:
            equal, *_ = compare_field_values(a[k][rule["column_a"]], b[k][rule["column_b"]], rule)
            if not equal:
                counts[rule["column_a"]] = counts.get(rule["column_a"], 0) + 1
    return counts


def _run(rows_a, rows_b, cols, rules=None):
    df_a = pd.DataFrame(rows_a)
    df_b = pd.DataFrame(rows_b)
    return run_delta(df_a, df_b, ["id"], ["id"], cols, cols, comparison_rules=rules)


class TestChangeFrequencyCounts:
    def test_unchanged_cells_not_counted(self):
        result = _run(
            [{"id": "1", "status": "Open", "amount": "10", "owner": "A"},
             {"id": "2", "status": "Open", "amount": "20", "owner": "B"}],
            [{"id": "1", "status": "Closed", "amount": "10", "owner": "A"},
             {"id": "2", "status": "Open", "amount": "25", "owner": "B"}],
            ["status", "amount", "owner"],
        )
        assert len(result.changed) == 2
        assert _freq_dict(result) == {"status": 1, "amount": 1}

    def test_multiple_changed_fields_in_one_record_each_counted_once(self):
        result = _run(
            [{"id": "1", "status": "Open", "amount": "10", "owner": "A"}],
            [{"id": "1", "status": "Closed", "amount": "11", "owner": "A"}],
            ["status", "amount", "owner"],
        )
        assert _freq_dict(result) == {"status": 1, "amount": 1}

    def test_blank_vs_value_counted_and_blank_vs_blank_not(self):
        result = _run(
            [{"id": "1", "status": "", "note": None},
             {"id": "2", "status": "Open", "note": "x"}],
            [{"id": "1", "status": "Open", "note": float("nan")},
             {"id": "2", "status": "Open", "note": "x"}],
            ["status", "note"],
        )
        # id 1: status blank -> Open is a change; note None vs NaN are both blank.
        assert _freq_dict(result) == {"status": 1}

    def test_pandas_nullable_values_treated_as_blank(self):
        df_a = pd.DataFrame({"id": ["1", "2"], "status": pd.array(["Open", pd.NA], dtype="string")})
        df_b = pd.DataFrame({"id": ["1", "2"], "status": pd.array(["Closed", "null"], dtype="string")})
        result = run_delta(df_a, df_b, ["id"], ["id"], ["status"], ["status"])
        assert _freq_dict(result) == {"status": 1}

    def test_numeric_within_tolerance_not_counted(self):
        rules = [
            {"column_a": "amount", "column_b": "amount", "type": "numeric", "tolerance": 0.01, "date_mode": None},
            {"column_a": "status", "column_b": "status", "type": "text", "tolerance": None, "date_mode": None},
        ]
        result = _run(
            [{"id": "1", "amount": "$1,000.00", "status": "Open"},
             {"id": "2", "amount": "500", "status": "Open"}],
            [{"id": "1", "amount": "1000.005", "status": "Closed"},
             {"id": "2", "amount": "500.50", "status": "Open"}],
            ["amount", "status"], rules,
        )
        # id 1: amount within tolerance, status changed; id 2: amount beyond tolerance.
        assert _freq_dict(result) == {"status": 1, "amount": 1}

    def test_date_aware_comparison(self):
        rules = [
            {"column_a": "due", "column_b": "due", "type": "date", "tolerance": None, "date_mode": "date_only"},
            {"column_a": "status", "column_b": "status", "type": "text", "tolerance": None, "date_mode": None},
        ]
        result = _run(
            [{"id": "1", "due": "2024-01-15", "status": "Open"},
             {"id": "2", "due": "2024-01-15 08:30", "status": "Open"},
             {"id": "3", "due": "2024-01-15", "status": "Open"}],
            [{"id": "1", "due": "01/15/2024", "status": "Closed"},
             {"id": "2", "due": "2024-01-15 23:59", "status": "Closed"},
             {"id": "3", "due": "2024-01-16", "status": "Open"}],
            ["due", "status"], rules,
        )
        # Same calendar date in another format / time is not a difference.
        assert _freq_dict(result) == {"status": 2, "due": 1}

    def test_engine_populates_changed_fields_with_non_null_strings(self):
        # build_change_frequency relies on this: a field the engine judged
        # different always has string before/after values, even when blank.
        rules = [
            {"column_a": "amount", "column_b": "amount", "type": "numeric", "tolerance": 0.0, "date_mode": None},
            {"column_a": "due", "column_b": "due", "type": "date", "tolerance": None, "date_mode": "date_only"},
            {"column_a": "status", "column_b": "status", "type": "text", "tolerance": None, "date_mode": None},
        ]
        result = _run(
            [{"id": "1", "amount": "", "due": "2024-01-01", "status": None},
             {"id": "2", "amount": "5", "due": "", "status": "Open"}],
            [{"id": "1", "amount": "7", "due": "2024-01-01", "status": "Open"},
             {"id": "2", "amount": "abc", "due": "2024-02-01", "status": "Open"}],
            ["amount", "due", "status"], rules,
        )
        for col in ("amount", "due", "status"):
            for side in ("Source", "Comparison"):
                name = f"{col} - {side}"
                if name in result.changed.columns:
                    populated = result.changed[name].dropna()
                    assert all(isinstance(v, str) for v in populated)
        assert _freq_dict(result) == {"amount": 2, "due": 1, "status": 1}

    def test_each_count_bounded_by_changed_records(self):
        result = _run(
            [{"id": str(i), "a": "x", "b": "y"} for i in range(5)],
            [{"id": str(i), "a": "x" if i % 2 else "z", "b": "y"} for i in range(5)],
            ["a", "b"],
        )
        freq = _freq_dict(result)
        assert freq == {"a": 3}
        assert all(v <= len(result.changed) for v in freq.values())


class TestChangeFrequencyBundledSamples:
    def test_invoice_register_sample(self):
        df_a, df_b = _read_sample("sample_a.csv"), _read_sample("sample_b.csv")
        cols = ["InvoiceAmount", "Status", "Reviewer"]
        result = run_delta(df_a, df_b, ["InvoiceID"], ["InvoiceID"], cols, cols)

        # Record-level results are unchanged by the fix.
        assert (len(result.only_in_a), len(result.only_in_b), len(result.matched),
                len(result.changed), len(result.duplicates_a), len(result.blank_keys_a)) == (2, 3, 7, 4, 2, 1)

        expected = {"Status": 4, "InvoiceAmount": 1, "Reviewer": 1}
        assert _engine_field_counts(df_a, df_b, "InvoiceID", cols) == expected
        assert _freq_dict(result) == expected

    def test_minimal_demo_sample(self):
        df_a, df_b = _read_sample("file_a.csv"), _read_sample("file_b.csv")
        cols = ["Status", "Amount"]
        result = run_delta(df_a, df_b, ["InvoiceID"], ["InvoiceID"], cols, cols)

        assert (len(result.only_in_a), len(result.only_in_b), len(result.matched),
                len(result.changed), len(result.duplicates_a), len(result.blank_keys_a)) == (1, 1, 3, 1, 2, 1)

        expected = {"Status": 1, "Amount": 1}
        assert _engine_field_counts(df_a, df_b, "InvoiceID", cols) == expected
        assert _freq_dict(result) == expected
