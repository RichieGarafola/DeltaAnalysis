"""
Tests for src/io_utils.py: sheet names, file reading, size checks, display frames.
"""
import io

import pandas as pd
import pytest

from src.io_utils import (
    LARGE_FILE_HARD_ROWS,
    LARGE_FILE_WARN_ROWS,
    PREVIEW_MAX_ROWS,
    check_file_size,
    get_display_frame,
    get_excel_sheet_names,
    read_uploaded_file,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

class MockUploadedFile:
    """Minimal Streamlit UploadedFile substitute for unit tests."""

    def __init__(self, name: str, content: bytes):
        self.name = name
        self._buf = io.BytesIO(content)

    def read(self, n: int = -1) -> bytes:
        return self._buf.read(n)

    def seek(self, pos: int, whence: int = 0) -> int:
        return self._buf.seek(pos, whence)

    def tell(self) -> int:
        return self._buf.tell()

    def seekable(self) -> bool:
        return True


def _make_excel_bytes(sheets: dict) -> bytes:
    """Build an Excel file in memory. sheets = {sheet_name: pd.DataFrame}."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()


def _make_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# ---------------------------------------------------------------------------
# TestGetSheetNames
# ---------------------------------------------------------------------------

class TestGetSheetNames:
    def test_single_sheet_excel(self):
        df = pd.DataFrame({"A": [1, 2]})
        xlsx = _make_excel_bytes({"Sheet1": df})
        f = MockUploadedFile("data.xlsx", xlsx)
        sheets = get_excel_sheet_names(f)
        assert sheets == ["Sheet1"]

    def test_multiple_sheets_excel(self):
        df = pd.DataFrame({"A": [1]})
        xlsx = _make_excel_bytes({"Alpha": df, "Beta": df, "Gamma": df})
        f = MockUploadedFile("data.xlsx", xlsx)
        sheets = get_excel_sheet_names(f)
        assert sheets == ["Alpha", "Beta", "Gamma"]

    def test_seekable_after_call(self):
        df = pd.DataFrame({"A": [1]})
        xlsx = _make_excel_bytes({"Sheet1": df})
        f = MockUploadedFile("data.xlsx", xlsx)
        get_excel_sheet_names(f)
        assert f.tell() == 0

    def test_csv_returns_empty_list(self):
        df = pd.DataFrame({"A": [1]})
        f = MockUploadedFile("data.csv", _make_csv_bytes(df))
        assert get_excel_sheet_names(f) == []


# ---------------------------------------------------------------------------
# TestReadUploadedFileWithSheet
# ---------------------------------------------------------------------------

class TestReadUploadedFileWithSheet:
    def test_default_reads_first_sheet(self):
        df1 = pd.DataFrame({"id": ["1"], "val": ["x"]})
        df2 = pd.DataFrame({"id": ["2"], "val": ["y"]})
        xlsx = _make_excel_bytes({"First": df1, "Second": df2})
        f = MockUploadedFile("data.xlsx", xlsx)
        result = read_uploaded_file(f)
        assert list(result["id"]) == ["1"]

    def test_named_sheet(self):
        df1 = pd.DataFrame({"id": ["1"]})
        df2 = pd.DataFrame({"id": ["99"]})
        xlsx = _make_excel_bytes({"A": df1, "B": df2})
        f = MockUploadedFile("data.xlsx", xlsx)
        result = read_uploaded_file(f, sheet_name="B")
        assert list(result["id"]) == ["99"]

    def test_invalid_sheet_raises(self):
        df = pd.DataFrame({"id": ["1"]})
        xlsx = _make_excel_bytes({"Sheet1": df})
        f = MockUploadedFile("data.xlsx", xlsx)
        with pytest.raises(ValueError, match="Could not parse"):
            read_uploaded_file(f, sheet_name="DoesNotExist")

    def test_csv_ignores_sheet_param(self):
        df = pd.DataFrame({"id": ["A"], "v": ["1"]})
        f = MockUploadedFile("data.csv", _make_csv_bytes(df))
        result = read_uploaded_file(f, sheet_name="IrrelevantSheet")
        assert list(result["id"]) == ["A"]

    def test_empty_sheet_raises(self):
        xlsx = _make_excel_bytes({"Empty": pd.DataFrame()})
        f = MockUploadedFile("data.xlsx", xlsx)
        with pytest.raises(ValueError, match="no data rows"):
            read_uploaded_file(f)


# ---------------------------------------------------------------------------
# TestFileSizeChecks
# ---------------------------------------------------------------------------

class TestFileSizeChecks:
    def test_ok_below_warn(self):
        status, msg = check_file_size(0)
        assert status == "ok" and msg == ""

    def test_ok_just_below_warn(self):
        status, msg = check_file_size(LARGE_FILE_WARN_ROWS - 1)
        assert status == "ok"

    def test_warn_at_threshold(self):
        status, msg = check_file_size(LARGE_FILE_WARN_ROWS)
        assert status == "warn" and msg != ""

    def test_warn_between_thresholds(self):
        mid = (LARGE_FILE_WARN_ROWS + LARGE_FILE_HARD_ROWS) // 2
        status, _ = check_file_size(mid)
        assert status == "warn"

    def test_hard_at_threshold(self):
        status, msg = check_file_size(LARGE_FILE_HARD_ROWS)
        assert status == "hard" and msg != ""

    def test_hard_above_threshold(self):
        status, _ = check_file_size(LARGE_FILE_HARD_ROWS + 1)
        assert status == "hard"


# ---------------------------------------------------------------------------
# TestGetDisplayFrame
# ---------------------------------------------------------------------------

class TestGetDisplayFrame:
    def test_small_df_unchanged(self):
        df = pd.DataFrame({"a": range(10)})
        result, truncated = get_display_frame(df)
        assert len(result) == 10 and not truncated

    def test_large_df_truncated(self):
        df = pd.DataFrame({"a": range(PREVIEW_MAX_ROWS + 50)})
        result, truncated = get_display_frame(df)
        assert len(result) == PREVIEW_MAX_ROWS and truncated

    def test_exact_limit_not_truncated(self):
        df = pd.DataFrame({"a": range(PREVIEW_MAX_ROWS)})
        result, truncated = get_display_frame(df)
        assert len(result) == PREVIEW_MAX_ROWS and not truncated

    def test_custom_max_rows(self):
        df = pd.DataFrame({"a": range(200)})
        result, truncated = get_display_frame(df, max_rows=50)
        assert len(result) == 50 and truncated

    def test_empty_df_not_truncated(self):
        df = pd.DataFrame({"a": []})
        result, truncated = get_display_frame(df)
        assert len(result) == 0 and not truncated
