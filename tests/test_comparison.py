"""
Tests for src/comparison.py — parse_numeric, parse_date_value, parse_datetime_value,
compare_field_values (including date_only and datetime_precision modes).
"""
import pytest
from datetime import date, datetime

from src.comparison import (
    compare_field_values,
    parse_date_value,
    parse_datetime_value,
    parse_numeric,
)


# ---------------------------------------------------------------------------
# parse_numeric
# ---------------------------------------------------------------------------

class TestParseNumeric:
    def test_plain_integer(self):
        val, ok = parse_numeric("1000")
        assert ok and val == 1000.0

    def test_comma_formatted(self):
        val, ok = parse_numeric("1,234,567.89")
        assert ok and val == pytest.approx(1234567.89)

    def test_dollar_sign(self):
        val, ok = parse_numeric("$1,500.00")
        assert ok and val == pytest.approx(1500.0)

    def test_euro_sign(self):
        val, ok = parse_numeric("€2500")
        assert ok and val == pytest.approx(2500.0)

    def test_parentheses_negative(self):
        val, ok = parse_numeric("(1,000.00)")
        assert ok and val == pytest.approx(-1000.0)

    def test_blank_string(self):
        val, ok = parse_numeric("")
        assert ok and val is None

    def test_nan_string(self):
        val, ok = parse_numeric("nan")
        assert ok and val is None

    def test_none_input(self):
        val, ok = parse_numeric(None)
        assert ok and val is None

    def test_null_string(self):
        val, ok = parse_numeric("null")
        assert ok and val is None

    def test_invalid_text(self):
        val, ok = parse_numeric("not-a-number")
        assert not ok and val is None

    def test_float_string(self):
        val, ok = parse_numeric("3.14")
        assert ok and val == pytest.approx(3.14)

    def test_negative_plain(self):
        val, ok = parse_numeric("-500")
        assert ok and val == pytest.approx(-500.0)


# ---------------------------------------------------------------------------
# parse_date_value
# ---------------------------------------------------------------------------

class TestParseDateValue:
    def test_iso_format(self):
        d, ok = parse_date_value("2024-01-15")
        assert ok and d == date(2024, 1, 15)

    def test_us_format(self):
        d, ok = parse_date_value("01/15/2024")
        assert ok and d == date(2024, 1, 15)

    def test_short_year(self):
        d, ok = parse_date_value("1/5/24")
        assert ok and d is not None

    def test_datetime_string(self):
        d, ok = parse_date_value("2024-01-15 08:30:00")
        assert ok and d == date(2024, 1, 15)

    def test_equivalent_formats_match(self):
        d1, ok1 = parse_date_value("2024-06-01")
        d2, ok2 = parse_date_value("06/01/2024")
        assert ok1 and ok2 and d1 == d2

    def test_blank_string(self):
        d, ok = parse_date_value("")
        assert ok and d is None

    def test_none_input(self):
        d, ok = parse_date_value(None)
        assert ok and d is None

    def test_nan_string(self):
        d, ok = parse_date_value("nan")
        assert ok and d is None

    def test_invalid_date(self):
        d, ok = parse_date_value("not-a-date")
        assert not ok and d is None


# ---------------------------------------------------------------------------
# parse_datetime_value
# ---------------------------------------------------------------------------

class TestParseDatetimeValue:
    def test_datetime_string_preserves_time(self):
        dt, ok = parse_datetime_value("2024-01-15 08:30:00")
        assert ok and dt == datetime(2024, 1, 15, 8, 30, 0)

    def test_date_only_string_midnight(self):
        dt, ok = parse_datetime_value("2024-01-15")
        assert ok and dt == datetime(2024, 1, 15, 0, 0, 0)

    def test_blank_returns_none(self):
        dt, ok = parse_datetime_value("")
        assert ok and dt is None

    def test_none_input(self):
        dt, ok = parse_datetime_value(None)
        assert ok and dt is None

    def test_invalid_returns_false(self):
        dt, ok = parse_datetime_value("not-a-date")
        assert not ok and dt is None


# ---------------------------------------------------------------------------
# compare_field_values
# ---------------------------------------------------------------------------

class TestCompareFieldValues:
    # --- text ---
    def test_text_equal(self):
        is_eq, _, _, issue = compare_field_values("Approved", "Approved", {"type": "text"})
        assert is_eq and issue is None

    def test_text_different(self):
        is_eq, _, _, issue = compare_field_values("Approved", "Pending", {"type": "text"})
        assert not is_eq and issue is None

    def test_text_normalizes_whitespace(self):
        is_eq, _, _, _ = compare_field_values("  Alpha  ", "Alpha", {"type": "text"})
        assert is_eq

    def test_text_default_type(self):
        is_eq, _, _, _ = compare_field_values("X", "X", {})
        assert is_eq

    # --- numeric ---
    def test_numeric_exact_match(self):
        rule = {"type": "numeric", "tolerance": 0.0}
        is_eq, _, _, issue = compare_field_values("1000", "1000", rule)
        assert is_eq and issue is None

    def test_numeric_within_tolerance(self):
        rule = {"type": "numeric", "tolerance": 0.01}
        is_eq, _, _, issue = compare_field_values("1000.005", "1000.00", rule)
        assert is_eq and issue is None

    def test_numeric_outside_tolerance(self):
        rule = {"type": "numeric", "tolerance": 0.01}
        is_eq, _, _, issue = compare_field_values("1000.10", "1000.00", rule)
        assert not is_eq and issue is None

    def test_numeric_currency_symbols(self):
        rule = {"type": "numeric", "tolerance": 0.0}
        is_eq, _, _, _ = compare_field_values("$1,500.00", "1500", rule)
        assert is_eq

    def test_numeric_both_blank(self):
        rule = {"type": "numeric", "tolerance": 0.0}
        is_eq, _, _, issue = compare_field_values("", "", rule)
        assert is_eq and issue is None

    def test_numeric_one_blank(self):
        rule = {"type": "numeric", "tolerance": 0.0}
        is_eq, _, _, _ = compare_field_values("100", "", rule)
        assert not is_eq

    def test_numeric_invalid_returns_issue(self):
        rule = {"type": "numeric", "tolerance": 0.0}
        _, _, _, issue = compare_field_values("abc", "100", rule)
        assert issue is not None

    # --- date: date_only mode (default) ---
    def test_date_equivalent_formats(self):
        rule = {"type": "date", "date_mode": "date_only"}
        is_eq, _, _, issue = compare_field_values("2024-06-01", "06/01/2024", rule)
        assert is_eq and issue is None

    def test_date_different(self):
        rule = {"type": "date", "date_mode": "date_only"}
        is_eq, _, _, _ = compare_field_values("2024-06-01", "2024-06-02", rule)
        assert not is_eq

    def test_date_only_ignores_time(self):
        rule = {"type": "date", "date_mode": "date_only"}
        is_eq, _, _, _ = compare_field_values("2024-01-15 08:30:00", "2024-01-15 14:00:00", rule)
        assert is_eq

    def test_date_only_default_when_mode_absent(self):
        rule = {"type": "date"}
        is_eq, _, _, _ = compare_field_values("2024-01-15 08:30:00", "2024-01-15 23:59:00", rule)
        assert is_eq

    def test_date_both_blank(self):
        rule = {"type": "date", "date_mode": "date_only"}
        is_eq, _, _, issue = compare_field_values("", "", rule)
        assert is_eq and issue is None

    def test_date_invalid_returns_issue(self):
        rule = {"type": "date", "date_mode": "date_only"}
        _, _, _, issue = compare_field_values("not-a-date", "2024-01-01", rule)
        assert issue is not None

    # --- date: datetime_precision mode ---
    def test_datetime_precision_same_time_equal(self):
        rule = {"type": "date", "date_mode": "datetime_precision"}
        is_eq, _, _, _ = compare_field_values("2024-01-15 08:30:00", "2024-01-15 08:30:00", rule)
        assert is_eq

    def test_datetime_precision_different_time_not_equal(self):
        rule = {"type": "date", "date_mode": "datetime_precision"}
        is_eq, _, _, _ = compare_field_values("2024-01-15 08:30:00", "2024-01-15 14:00:00", rule)
        assert not is_eq

    def test_datetime_precision_date_only_strings_equal(self):
        rule = {"type": "date", "date_mode": "datetime_precision"}
        is_eq, _, _, _ = compare_field_values("2024-01-15", "2024-01-15 00:00:00", rule)
        assert is_eq

    def test_datetime_precision_different_dates_not_equal(self):
        rule = {"type": "date", "date_mode": "datetime_precision"}
        is_eq, _, _, _ = compare_field_values("2024-01-15", "2024-01-16", rule)
        assert not is_eq
