"""Unit tests for src/delta_engine.py"""
import pandas as pd
import pytest

from src.delta_engine import run_delta, DeltaResult


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_df(rows: list[dict]) -> pd.DataFrame:
    """Build a DataFrame from a list of dicts (all values as str)."""
    df = pd.DataFrame(rows)
    return df.astype(str).replace("nan", "")


@pytest.fixture
def base_a():
    return _make_df([
        {"id": "C001", "name": "Alpha",   "status": "Approved", "amount": "1000"},
        {"id": "C002", "name": "Bravo",   "status": "Pending",  "amount": "2000"},
        {"id": "C003", "name": "Charlie", "status": "Denied",   "amount": "3000"},
        {"id": "C004", "name": "Delta",   "status": "Approved", "amount": "4000"},
    ])


@pytest.fixture
def base_b():
    return _make_df([
        {"id": "C001", "name": "Alpha",   "status": "Approved", "amount": "1000"},  # unchanged
        {"id": "C002", "name": "Bravo",   "status": "Approved", "amount": "2500"},  # changed: status + amount
        {"id": "C003", "name": "Charlie", "status": "Denied",   "amount": "3000"},  # unchanged
        {"id": "C005", "name": "Echo",    "status": "Pending",  "amount": "5000"},  # only in B
    ])


# ---------------------------------------------------------------------------
# Basic delta categories
# ---------------------------------------------------------------------------

class TestOnlyInAOnlyInB:
    def test_only_in_a(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        assert set(result.only_in_a["id"]) == {"C004"}

    def test_only_in_b(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        assert set(result.only_in_b["id"]) == {"C005"}

    def test_only_in_a_is_empty_when_fully_matched(self):
        df = _make_df([{"id": "X1", "v": "a"}, {"id": "X2", "v": "b"}])
        result = run_delta(df, df.copy(), ["id"], ["id"])
        assert result.only_in_a.empty
        assert result.only_in_b.empty

    def test_all_only_in_a(self, base_a):
        empty_b = _make_df([{"id": "Z99", "name": "Ghost", "status": "N/A", "amount": "0"}])
        result = run_delta(base_a, empty_b, ["id"], ["id"])
        assert len(result.only_in_a) == len(base_a)
        assert len(result.only_in_b) == 1


class TestMatchedRecords:
    def test_matched_count(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        # C001, C002, C003 match
        assert len(result.matched) == 3

    def test_matched_has_prefixed_columns(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        for col in result.matched.columns:
            assert col.startswith("A: ") or col.startswith("B: ")

    def test_matched_includes_both_sides(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        assert "A: id" in result.matched.columns
        assert "B: id" in result.matched.columns


class TestChangedRecords:
    def test_changed_detected(self, base_a, base_b):
        result = run_delta(
            base_a, base_b,
            key_cols_a=["id"], key_cols_b=["id"],
            compare_cols_a=["status", "amount"],
            compare_cols_b=["status", "amount"],
        )
        # Only C002 has changes
        assert len(result.changed) == 1
        row = result.changed.iloc[0]
        assert row["status — File A"] == "Pending"
        assert row["status — File B"] == "Approved"

    def test_no_changes_when_identical(self):
        df = _make_df([{"id": "X1", "v": "same"}, {"id": "X2", "v": "also_same"}])
        result = run_delta(df, df.copy(), ["id"], ["id"],
                           compare_cols_a=["v"], compare_cols_b=["v"])
        assert result.changed.empty

    def test_no_comparison_cols_skips_change_detection(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        assert result.changed.empty

    def test_changed_key_cols_present(self, base_a, base_b):
        result = run_delta(
            base_a, base_b, ["id"], ["id"],
            compare_cols_a=["status"], compare_cols_b=["status"],
        )
        assert "Key: id" in result.changed.columns

    def test_multiple_changed_fields(self, base_a, base_b):
        result = run_delta(
            base_a, base_b, ["id"], ["id"],
            compare_cols_a=["status", "amount"],
            compare_cols_b=["status", "amount"],
        )
        row = result.changed.iloc[0]
        # Both status and amount changed for C002
        assert "status — File A" in result.changed.columns
        assert "amount — File A" in result.changed.columns


# ---------------------------------------------------------------------------
# Duplicate detection
# ---------------------------------------------------------------------------

class TestDuplicates:
    def test_duplicates_detected_in_a(self):
        df_a = _make_df([
            {"id": "C001", "v": "first"},
            {"id": "C001", "v": "second"},
            {"id": "C002", "v": "unique"},
        ])
        df_b = _make_df([{"id": "C001", "v": "x"}])
        result = run_delta(df_a, df_b, ["id"], ["id"])
        assert len(result.duplicates_a) == 2
        assert all(result.duplicates_a["id"] == "C001")

    def test_duplicates_detected_in_b(self):
        df_a = _make_df([{"id": "C001", "v": "x"}])
        df_b = _make_df([
            {"id": "C001", "v": "first"},
            {"id": "C001", "v": "second"},
        ])
        result = run_delta(df_a, df_b, ["id"], ["id"])
        assert len(result.duplicates_b) == 2

    def test_no_duplicates_returns_empty(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        assert result.duplicates_a.empty
        assert result.duplicates_b.empty

    def test_duplicate_matching_uses_first_occurrence(self):
        """Duplicates in A should not cause extra rows in 'only_in_a'."""
        df_a = _make_df([
            {"id": "C001", "v": "alpha"},
            {"id": "C001", "v": "beta"},
        ])
        df_b = _make_df([{"id": "C001", "v": "alpha"}])
        result = run_delta(df_a, df_b, ["id"], ["id"])
        # C001 matched, so nothing should appear in only_in_a
        assert result.only_in_a.empty


# ---------------------------------------------------------------------------
# Blank key handling
# ---------------------------------------------------------------------------

class TestBlankKeys:
    def test_blank_keys_separated(self):
        df_a = _make_df([
            {"id": "C001", "v": "good"},
            {"id": "",     "v": "bad"},
            {"id": "nan",  "v": "also_bad"},
        ])
        df_b = _make_df([{"id": "C001", "v": "good"}])
        result = run_delta(df_a, df_b, ["id"], ["id"])
        assert len(result.blank_keys_a) == 2
        assert result.only_in_a.empty  # C001 matched, blanks excluded

    def test_blank_keys_not_matched(self):
        """Two rows with blank keys should NOT be matched to each other."""
        df_a = _make_df([{"id": "", "v": "x"}])
        df_b = _make_df([{"id": "", "v": "y"}])
        result = run_delta(df_a, df_b, ["id"], ["id"])
        assert result.matched.empty
        assert len(result.blank_keys_a) == 1
        assert len(result.blank_keys_b) == 1


# ---------------------------------------------------------------------------
# Multi-column keys
# ---------------------------------------------------------------------------

class TestCompositeKeys:
    def test_composite_key_match(self):
        df_a = _make_df([
            {"fiscal_year": "2024", "case_id": "001", "status": "Open"},
            {"fiscal_year": "2024", "case_id": "002", "status": "Closed"},
            {"fiscal_year": "2025", "case_id": "001", "status": "Pending"},
        ])
        df_b = _make_df([
            {"fiscal_year": "2024", "case_id": "001", "status": "Open"},
            {"fiscal_year": "2025", "case_id": "001", "status": "Approved"},  # changed
        ])
        result = run_delta(
            df_a, df_b,
            key_cols_a=["fiscal_year", "case_id"],
            key_cols_b=["fiscal_year", "case_id"],
            compare_cols_a=["status"],
            compare_cols_b=["status"],
        )
        assert len(result.matched) == 2
        assert len(result.only_in_a) == 1  # FY2024-002
        assert len(result.changed) == 1    # FY2025-001

    def test_different_key_column_names(self):
        """Key columns can have different names in each file."""
        df_a = _make_df([{"CaseID": "001", "val": "A"}, {"CaseID": "002", "val": "B"}])
        df_b = _make_df([{"CaseNumber": "001", "val": "A"}, {"CaseNumber": "003", "val": "C"}])
        result = run_delta(df_a, df_b, ["CaseID"], ["CaseNumber"])
        assert len(result.matched) == 1
        assert len(result.only_in_a) == 1
        assert len(result.only_in_b) == 1


# ---------------------------------------------------------------------------
# Validation and error handling
# ---------------------------------------------------------------------------

class TestValidation:
    def test_missing_key_column_raises(self, base_a, base_b):
        with pytest.raises(ValueError, match="not found in File A"):
            run_delta(base_a, base_b, ["nonexistent"], ["id"])

    def test_mismatched_key_column_counts_raises(self, base_a, base_b):
        with pytest.raises(ValueError, match="Key column counts must match"):
            run_delta(base_a, base_b, ["id", "name"], ["id"])

    def test_mismatched_compare_column_counts_raises(self, base_a, base_b):
        with pytest.raises(ValueError, match="Comparison column counts must match"):
            run_delta(base_a, base_b, ["id"], ["id"],
                      compare_cols_a=["status", "amount"],
                      compare_cols_b=["status"])

    def test_empty_dataframe_a(self, base_b):
        df_a = pd.DataFrame(columns=["id", "name", "status", "amount"])
        result = run_delta(df_a, base_b, ["id"], ["id"])
        assert result.only_in_a.empty
        assert len(result.only_in_b) == len(base_b)
        assert result.matched.empty

    def test_empty_dataframe_b(self, base_a):
        df_b = pd.DataFrame(columns=["id", "name", "status", "amount"])
        result = run_delta(base_a, df_b, ["id"], ["id"])
        assert result.only_in_b.empty
        assert len(result.only_in_a) == len(base_a)
        assert result.matched.empty


# ---------------------------------------------------------------------------
# Totals preserved
# ---------------------------------------------------------------------------

class TestTotals:
    def test_total_counts_match_input(self, base_a, base_b):
        result = run_delta(base_a, base_b, ["id"], ["id"])
        assert result.total_a == len(base_a)
        assert result.total_b == len(base_b)
