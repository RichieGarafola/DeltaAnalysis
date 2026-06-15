"""
Tests for src/io_utils.py: read_raw_file and prepare_dataframe_from_raw,
plus verification that prep metadata flows through to the Excel report.
"""
import io

import pandas as pd
import pytest
from openpyxl import load_workbook

from src.io_utils import prepare_dataframe_from_raw, read_raw_file
from src.delta_engine import run_delta
from src.reporting import export_to_excel


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

class _FakeFile:
    """Minimal stand-in for a Streamlit UploadedFile backed by bytes."""

    def __init__(self, content: bytes, name: str):
        self._buf = io.BytesIO(content)
        self.name = name

    def seek(self, pos: int) -> None:
        self._buf.seek(pos)

    def read(self, n: int = -1) -> bytes:
        return self._buf.read(n)

    # pandas calls tell() on some paths
    def tell(self) -> int:
        return self._buf.tell()


def _csv_file(rows: list[list], name: str = "test.csv") -> _FakeFile:
    lines = [",".join(str(c) for c in row) for row in rows]
    content = "\n".join(lines).encode("utf-8")
    return _FakeFile(content, name)


def _make_raw_df(*rows) -> pd.DataFrame:
    """Build a raw (integer-column) DataFrame the same way read_raw_file would."""
    data = [list(r) for r in rows]
    df = pd.DataFrame(data, dtype=str)
    df.columns = range(len(df.columns))
    return df


# ---------------------------------------------------------------------------
# TestReadRawFile
# ---------------------------------------------------------------------------

class TestReadRawFile:
    def test_csv_returns_integer_columns(self):
        f = _csv_file([["id", "name"], ["1", "Alice"], ["2", "Bob"]])
        df = read_raw_file(f)
        assert list(df.columns) == [0, 1]

    def test_csv_first_row_is_data_not_header(self):
        f = _csv_file([["id", "name"], ["1", "Alice"]])
        df = read_raw_file(f)
        assert df.iloc[0, 0] == "id"

    def test_csv_all_values_are_strings(self):
        f = _csv_file([["col"], ["42"], ["3.14"]])
        df = read_raw_file(f)
        assert all(isinstance(v, str) for v in df.iloc[:, 0])

    def test_csv_blank_cells_become_empty_string(self):
        content = b"a,,c\n1,,3\n"
        f = _FakeFile(content, "test.csv")
        df = read_raw_file(f)
        assert df.iloc[1, 1] == ""

    def test_seek_restored_after_read(self):
        f = _csv_file([["x"], ["1"]])
        read_raw_file(f)
        assert f._buf.tell() == 0

    def test_none_file_raises_value_error(self):
        with pytest.raises(ValueError, match="No file"):

            class _Null:
                name = "x.csv"
                def __call__(self): pass

            # simulate None
            read_raw_file(None)

    def test_unsupported_extension_raises_value_error(self):
        f = _FakeFile(b"data", "data.txt")
        with pytest.raises(ValueError, match="Unsupported file format"):
            read_raw_file(f)

    def test_row_count_matches_content(self):
        f = _csv_file([["h1", "h2"], ["a", "b"], ["c", "d"], ["e", "f"]])
        df = read_raw_file(f)
        assert len(df) == 4

    def test_column_count_matches_content(self):
        f = _csv_file([["a", "b", "c", "d"]])
        df = read_raw_file(f)
        assert len(df.columns) == 4


# ---------------------------------------------------------------------------
# TestPrepareDataframeFromRaw
# ---------------------------------------------------------------------------

class TestPrepareDataframeFromRaw:
    def _simple_raw(self):
        return _make_raw_df(
            ["id", "name", "amount"],
            ["1",  "Alice", "100"],
            ["2",  "Bob",   "200"],
            ["3",  "Carol", "300"],
        )

    def test_header_row_0_default(self):
        df, meta, warns = prepare_dataframe_from_raw(self._simple_raw())
        assert list(df.columns) == ["id", "name", "amount"]
        assert len(df) == 3
        assert warns == []

    def test_header_row_index_2(self):
        raw = _make_raw_df(
            ["Report Title", "", ""],
            ["Agency: DoD",  "", ""],
            ["id", "name",   "amount"],
            ["1",  "Alice",  "100"],
        )
        df, meta, _ = prepare_dataframe_from_raw(raw, header_row_index=2)
        assert list(df.columns) == ["id", "name", "amount"]
        assert len(df) == 1
        assert df.iloc[0]["id"] == "1"

    def test_rows_above_dropped_by_default(self):
        raw = _make_raw_df(
            ["metadata", ""],
            ["id", "name"],
            ["1", "Alice"],
        )
        df, meta, _ = prepare_dataframe_from_raw(raw, header_row_index=1, drop_rows_above=True)
        assert len(df) == 1
        assert meta["rows_dropped_above_header"] == 1

    def test_rows_above_kept_when_disabled(self):
        raw = _make_raw_df(
            ["metadata", "notes"],
            ["id", "name"],
            ["1", "Alice"],
        )
        df, meta, _ = prepare_dataframe_from_raw(raw, header_row_index=1, drop_rows_above=False)
        assert len(df) == 2
        assert meta["rows_dropped_above_header"] == 0

    def test_blank_rows_dropped_by_default(self):
        raw = _make_raw_df(
            ["id", "name"],
            ["1", "Alice"],
            ["",  ""],
            ["2", "Bob"],
        )
        df, meta, _ = prepare_dataframe_from_raw(raw)
        assert len(df) == 2
        assert meta["rows_dropped_blank"] == 1

    def test_blank_rows_kept_when_disabled(self):
        raw = _make_raw_df(
            ["id", "name"],
            ["1", "Alice"],
            ["",  ""],
        )
        df, meta, _ = prepare_dataframe_from_raw(raw, drop_blank_rows=False)
        assert len(df) == 2
        assert meta["rows_dropped_blank"] == 0

    def test_end_row_filter(self):
        raw = _make_raw_df(
            ["id"],
            ["1"],
            ["2"],
            ["3"],
            ["Total", "10"],
        )
        # end_row_index=3 means include raw rows 1, 2, 3 (data rows 0-2)
        df, meta, _ = prepare_dataframe_from_raw(raw, end_row_index=3)
        assert len(df) == 3
        assert meta["end_row_applied"] == 3

    def test_end_row_none_includes_all(self):
        raw = self._simple_raw()
        df, meta, _ = prepare_dataframe_from_raw(raw, end_row_index=None)
        assert len(df) == 3
        assert meta["end_row_applied"] is None

    def test_out_of_bounds_header_raises(self):
        raw = _make_raw_df(["id"], ["1"])
        with pytest.raises(ValueError, match="out of range"):
            prepare_dataframe_from_raw(raw, header_row_index=99)

    def test_empty_dataframe_raises(self):
        empty = pd.DataFrame()
        with pytest.raises(ValueError, match="empty"):
            prepare_dataframe_from_raw(empty)

    def test_blank_header_renamed_to_unnamed(self):
        raw = _make_raw_df(
            ["id", "", "amount"],
            ["1",  "X", "100"],
        )
        df, _, warns = prepare_dataframe_from_raw(raw)
        assert "Unnamed_1" in df.columns
        assert any("blank header" in w for w in warns)

    def test_duplicate_headers_renamed(self):
        raw = _make_raw_df(
            ["id", "val", "val"],
            ["1",  "a",   "b"],
        )
        df, _, warns = prepare_dataframe_from_raw(raw)
        assert "val" in df.columns
        assert "val_2" in df.columns
        assert any("Duplicate" in w for w in warns)

    def test_whitespace_stripped_from_headers(self):
        raw = _make_raw_df(
            ["  id  ", " name ", "amount"],
            ["1", "Alice", "100"],
        )
        df, _, _ = prepare_dataframe_from_raw(raw)
        assert "id" in df.columns
        assert "name" in df.columns

    def test_metadata_structure(self):
        df, meta, _ = prepare_dataframe_from_raw(self._simple_raw())
        required_keys = {
            "header_row_selected",
            "rows_dropped_above_header",
            "rows_dropped_blank",
            "end_row_applied",
            "final_row_count",
            "final_column_count",
        }
        assert required_keys == set(meta.keys())

    def test_metadata_values_correct(self):
        raw = _make_raw_df(
            ["ignore", ""],
            ["id", "name"],
            ["1", "Alice"],
            ["",  ""],
        )
        df, meta, _ = prepare_dataframe_from_raw(raw, header_row_index=1)
        assert meta["header_row_selected"] == 1
        assert meta["rows_dropped_above_header"] == 1
        assert meta["rows_dropped_blank"] == 1
        assert meta["final_row_count"] == 1
        assert meta["final_column_count"] == 2

    def test_returns_tuple_of_three(self):
        result = prepare_dataframe_from_raw(self._simple_raw())
        assert len(result) == 3
        df, meta, warns = result
        assert isinstance(df, pd.DataFrame)
        assert isinstance(meta, dict)
        assert isinstance(warns, list)


# ---------------------------------------------------------------------------
# TestPreparationMetadataInReport
# ---------------------------------------------------------------------------

class TestPreparationMetadataInReport:
    def _make_result(self):
        df_a = pd.DataFrame({"id": ["1", "2"], "val": ["a", "b"]})
        df_b = pd.DataFrame({"id": ["1", "2"], "val": ["a", "c"]})
        return run_delta(df_a, df_b, ["id"], ["id"], ["val"], ["val"])

    def _load_metadata_values(self, wb) -> list[str]:
        ws = wb["Analysis Metadata"]
        return [str(ws.cell(row=r, column=2).value) for r in range(2, ws.max_row + 1)]

    def _load_metadata_params(self, wb) -> list[str]:
        ws = wb["Analysis Metadata"]
        return [str(ws.cell(row=r, column=1).value) for r in range(2, ws.max_row + 1)]

    def test_metadata_without_prep_meta_has_no_prep_rows(self):
        result = self._make_result()
        raw = export_to_excel(result)
        wb = load_workbook(io.BytesIO(raw))
        params = self._load_metadata_params(wb)
        assert not any("Header Row" in p for p in params)

    def test_metadata_with_prep_meta_a_includes_source_rows(self):
        result = self._make_result()
        prep_a = {
            "header_row_selected": 2,
            "rows_dropped_above_header": 2,
            "rows_dropped_blank": 1,
            "end_row_applied": None,
            "final_row_count": 5,
            "final_column_count": 3,
        }
        raw = export_to_excel(result, prep_meta_a=prep_a)
        wb = load_workbook(io.BytesIO(raw))
        params = self._load_metadata_params(wb)
        assert any("Source Header Row Selected" in p for p in params)
        assert any("Source Rows Dropped Above Header" in p for p in params)
        assert any("Source Blank Rows Dropped" in p for p in params)

    def test_metadata_with_prep_meta_b_includes_comparison_rows(self):
        result = self._make_result()
        prep_b = {
            "header_row_selected": 0,
            "rows_dropped_above_header": 0,
            "rows_dropped_blank": 3,
            "end_row_applied": None,
            "final_row_count": 10,
            "final_column_count": 4,
        }
        raw = export_to_excel(result, prep_meta_b=prep_b)
        wb = load_workbook(io.BytesIO(raw))
        params = self._load_metadata_params(wb)
        assert any("Comparison Header Row Selected" in p for p in params)
        assert any("Comparison Blank Rows Dropped" in p for p in params)

    def test_end_row_row_present_only_when_set(self):
        result = self._make_result()
        prep_a_no_end = {
            "header_row_selected": 0,
            "rows_dropped_above_header": 0,
            "rows_dropped_blank": 0,
            "end_row_applied": None,
            "final_row_count": 2,
            "final_column_count": 2,
        }
        prep_a_with_end = dict(prep_a_no_end, end_row_applied=50)

        raw_no_end = export_to_excel(result, prep_meta_a=prep_a_no_end)
        raw_with_end = export_to_excel(result, prep_meta_a=prep_a_with_end)

        wb_no = load_workbook(io.BytesIO(raw_no_end))
        wb_yes = load_workbook(io.BytesIO(raw_with_end))

        params_no = self._load_metadata_params(wb_no)
        params_yes = self._load_metadata_params(wb_yes)

        assert not any("Source End Row Applied" in p for p in params_no)
        assert any("Source End Row Applied" in p for p in params_yes)

    def test_prep_meta_values_written_to_excel(self):
        result = self._make_result()
        prep_a = {
            "header_row_selected": 3,
            "rows_dropped_above_header": 3,
            "rows_dropped_blank": 7,
            "end_row_applied": None,
            "final_row_count": 100,
            "final_column_count": 5,
        }
        raw = export_to_excel(result, prep_meta_a=prep_a)
        wb = load_workbook(io.BytesIO(raw))
        values = self._load_metadata_values(wb)
        assert "3" in values
        assert "7" in values
