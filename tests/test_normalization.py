"""Unit tests for src/normalization.py"""
import math
import pandas as pd
import pytest

from src.normalization import normalize_key_value, normalize_key_series, build_composite_key


class TestNormalizeKeyValue:
    def test_none_returns_blank(self):
        assert normalize_key_value(None) == "__BLANK__"

    def test_float_nan_returns_blank(self):
        assert normalize_key_value(float("nan")) == "__BLANK__"

    def test_empty_string_returns_blank(self):
        assert normalize_key_value("") == "__BLANK__"

    def test_whitespace_only_returns_blank(self):
        assert normalize_key_value("   ") == "__BLANK__"

    def test_string_nan_returns_blank(self):
        assert normalize_key_value("nan") == "__BLANK__"
        assert normalize_key_value("NaN") == "__BLANK__"
        assert normalize_key_value("NAN") == "__BLANK__"

    def test_string_none_returns_blank(self):
        assert normalize_key_value("none") == "__BLANK__"
        assert normalize_key_value("None") == "__BLANK__"
        assert normalize_key_value("NONE") == "__BLANK__"

    def test_string_null_returns_blank(self):
        assert normalize_key_value("null") == "__BLANK__"
        assert normalize_key_value("NULL") == "__BLANK__"

    def test_na_returns_blank(self):
        assert normalize_key_value("n/a") == "__BLANK__"
        assert normalize_key_value("N/A") == "__BLANK__"
        assert normalize_key_value("na") == "__BLANK__"

    def test_decimal_artifact_stripped(self):
        assert normalize_key_value("1234.0") == "1234"
        assert normalize_key_value("0.0") == "0"
        assert normalize_key_value("99999.0") == "99999"

    def test_non_decimal_float_preserved(self):
        # "12.5" is not an artifact — it's a real decimal
        assert normalize_key_value("12.5") == "12.5"

    def test_whitespace_stripped(self):
        assert normalize_key_value("  C001  ") == "C001"
        assert normalize_key_value("\tABC\n") == "ABC"

    def test_normal_string_unchanged(self):
        assert normalize_key_value("C001") == "C001"
        assert normalize_key_value("ACME-001") == "ACME-001"

    def test_integer_converted_to_string(self):
        assert normalize_key_value(42) == "42"

    def test_integer_float_artifact(self):
        # When Excel reads an integer ID it often stores as float
        assert normalize_key_value(1001.0) == "1001"

    def test_boolean_preserved_as_string(self):
        result = normalize_key_value(True)
        assert result in ("True", "true", "1")  # str(True) == "True"


class TestNormalizeKeySeries:
    def test_series_normalization(self):
        s = pd.Series(["  C001  ", "1002.0", None, "", "C003"])
        result = normalize_key_series(s)
        assert result.tolist() == ["C001", "1002", "__BLANK__", "__BLANK__", "C003"]

    def test_all_blank(self):
        s = pd.Series([None, float("nan"), "", "null"])
        result = normalize_key_series(s)
        assert all(v == "__BLANK__" for v in result)


class TestBuildCompositeKey:
    def test_single_column(self):
        df = pd.DataFrame({"id": ["C001", "C002", "C003"]})
        keys = build_composite_key(df, ["id"])
        assert keys.tolist() == ["C001", "C002", "C003"]

    def test_multi_column(self):
        df = pd.DataFrame({"a": ["X", "Y"], "b": ["1", "2"]})
        keys = build_composite_key(df, ["a", "b"])
        assert keys.tolist() == ["X||1", "Y||2"]

    def test_normalization_applied(self):
        df = pd.DataFrame({"id": ["  C001  ", "1002.0", None]})
        keys = build_composite_key(df, ["id"])
        assert keys.tolist() == ["C001", "1002", "__BLANK__"]

    def test_empty_cols_list_raises(self):
        df = pd.DataFrame({"id": ["C001"]})
        with pytest.raises(ValueError, match="At least one key column"):
            build_composite_key(df, [])
