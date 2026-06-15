"""
Field-level comparison utilities.

Provides type-aware value comparison supporting:
- Numeric: optional tolerance for rounding/currency differences
- Date: comparison precision control (date_only vs datetime_precision)
- Text: normalized string comparison (default)

Date modes
----------
date_only           Compare only the calendar date; time components are
                    discarded. "2024-01-15 08:30" == "2024-01-15 14:00".
                    This is the default for the 'date' comparison type.
datetime_precision  Compare the full datetime including time. Strings
                    without an explicit time part are treated as midnight
                    (00:00:00), so "2024-01-15" == "2024-01-15 00:00:00"
                    but != "2024-01-15 06:00:00".
"""
from __future__ import annotations

import re
from datetime import date, datetime
from typing import Optional, Tuple

import pandas as pd


# ---------------------------------------------------------------------------
# Numeric parsing
# ---------------------------------------------------------------------------

_CURRENCY_STRIP = re.compile(r"[\$£€¥,\s]")
_PAREN_NEGATIVE = re.compile(r"^\(([0-9,.\s]+)\)$")


def parse_numeric(value) -> Tuple[Optional[float], bool]:
    """
    Parse a value as a float, stripping common currency/formatting characters.

    Handles: plain integers, comma-formatted numbers, currency symbols ($£€¥),
    parenthesised negatives like (1,000.00), and leading/trailing whitespace.

    Returns (float_value, True) on success, (None, False) on failure.
    Blank / NaN strings return (None, True) - treated as missing, not invalid.
    """
    if value is None:
        return None, True

    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "null", "n/a", "na"):
        return None, True

    # Parenthesised negative: (1,234.56) -> -1234.56
    paren_match = _PAREN_NEGATIVE.match(s)
    if paren_match:
        cleaned = _CURRENCY_STRIP.sub("", paren_match.group(1))
        try:
            return -float(cleaned), True
        except ValueError:
            return None, False

    cleaned = _CURRENCY_STRIP.sub("", s)
    try:
        return float(cleaned), True
    except ValueError:
        return None, False


# ---------------------------------------------------------------------------
# Date / datetime parsing
# ---------------------------------------------------------------------------

_BLANK_STRINGS = {"", "nan", "none", "null", "n/a", "na"}

_DATE_FORMATS = [
    "%Y-%m-%d",
    "%m/%d/%Y",
    "%m/%d/%y",
    "%d/%m/%Y",
    "%Y%m%d",
    "%m-%d-%Y",
    "%d-%m-%Y",
    "%B %d, %Y",
    "%b %d, %Y",
]


def parse_date_value(value) -> Tuple[Optional[date], bool]:
    """
    Parse a value and return only its calendar date (time discarded).

    Used by date_only comparison mode.

    Returns (date_object, True) on success, (None, False) on failure.
    Blank / NaN strings return (None, True) - missing, not invalid.
    """
    if value is None:
        return None, True

    s = str(value).strip()
    if s.lower() in _BLANK_STRINGS:
        return None, True

    try:
        return pd.to_datetime(s, dayfirst=False).date(), True
    except (ValueError, TypeError):
        pass

    try:
        return pd.to_datetime(s, dayfirst=True).date(), True
    except (ValueError, TypeError):
        pass

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date(), True
        except ValueError:
            continue

    return None, False


def parse_datetime_value(value) -> Tuple[Optional[datetime], bool]:
    """
    Parse a value as a full datetime, preserving time components.

    Used by datetime_precision comparison mode. Values without an explicit
    time part are treated as midnight (00:00:00).

    Returns (datetime_object, True) on success, (None, False) on failure.
    Blank / NaN strings return (None, True) - missing, not invalid.
    """
    if value is None:
        return None, True

    s = str(value).strip()
    if s.lower() in _BLANK_STRINGS:
        return None, True

    try:
        return pd.to_datetime(s, dayfirst=False).to_pydatetime(), True
    except (ValueError, TypeError):
        pass

    try:
        return pd.to_datetime(s, dayfirst=True).to_pydatetime(), True
    except (ValueError, TypeError):
        pass

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt), True
        except ValueError:
            continue

    return None, False


# ---------------------------------------------------------------------------
# Field comparison
# ---------------------------------------------------------------------------

def compare_field_values(
    val_a,
    val_b,
    rule: dict,
) -> Tuple[bool, str, str, Optional[str]]:
    """
    Compare two field values according to a comparison rule.

    Parameters
    ----------
    val_a, val_b : Raw values from the Baseline and Comparison datasets.
    rule         : Dict with keys:
                   - type       : 'text' | 'numeric' | 'date'  (default 'text')
                   - tolerance  : float, used when type == 'numeric'
                   - date_mode  : 'date_only' | 'datetime_precision'
                                  (default 'date_only' when type == 'date')

    Returns
    -------
    (is_equal, str_a, str_b, issue_note)

    is_equal    : True if the values are considered equal under the rule
    str_a       : Normalised display string for the Baseline value
    str_b       : Normalised display string for the Comparison value
    issue_note  : Non-None string describing a parse warning, if any
    """
    ctype = rule.get("type", "text")

    if ctype == "numeric":
        return _compare_numeric(val_a, val_b, rule)
    if ctype == "date":
        return _compare_date(val_a, val_b, rule)
    return _compare_text(val_a, val_b)


# ---------------------------------------------------------------------------
# Type-specific helpers
# ---------------------------------------------------------------------------

def _compare_text(val_a, val_b) -> Tuple[bool, str, str, Optional[str]]:
    from src.normalization import normalize_key_value
    s_a = normalize_key_value(val_a)
    s_b = normalize_key_value(val_b)
    return s_a == s_b, s_a, s_b, None


def _compare_numeric(val_a, val_b, rule: dict) -> Tuple[bool, str, str, Optional[str]]:
    tolerance = rule.get("tolerance", 0.0) or 0.0
    n_a, ok_a = parse_numeric(val_a)
    n_b, ok_b = parse_numeric(val_b)

    issue = None
    if not ok_a and not ok_b:
        issue = f"Could not parse either value as numeric: '{val_a}' / '{val_b}'"
    elif not ok_a:
        issue = f"Could not parse Baseline value as numeric: '{val_a}'"
    elif not ok_b:
        issue = f"Could not parse Comparison value as numeric: '{val_b}'"

    if issue:
        from src.normalization import normalize_key_value
        s_a = normalize_key_value(val_a)
        s_b = normalize_key_value(val_b)
        return s_a == s_b, s_a, s_b, issue

    if n_a is None and n_b is None:
        return True, "", "", None

    if n_a is None or n_b is None:
        s_a = "" if n_a is None else str(n_a)
        s_b = "" if n_b is None else str(n_b)
        return False, s_a, s_b, None

    diff = abs(n_a - n_b)
    is_equal = diff <= tolerance
    return is_equal, str(n_a), str(n_b), None


def _compare_date(val_a, val_b, rule: dict) -> Tuple[bool, str, str, Optional[str]]:
    """
    Compare two date/datetime values.

    date_mode='date_only' (default): strip time, compare calendar dates only.
    date_mode='datetime_precision':  compare full datetime including time.
    """
    mode = rule.get("date_mode") or "date_only"

    if mode == "datetime_precision":
        return _compare_datetime_precision(val_a, val_b)
    return _compare_date_only(val_a, val_b)


def _compare_date_only(val_a, val_b) -> Tuple[bool, str, str, Optional[str]]:
    d_a, ok_a = parse_date_value(val_a)
    d_b, ok_b = parse_date_value(val_b)
    return _finish_date_compare(val_a, val_b, d_a, ok_a, d_b, ok_b)


def _compare_datetime_precision(val_a, val_b) -> Tuple[bool, str, str, Optional[str]]:
    d_a, ok_a = parse_datetime_value(val_a)
    d_b, ok_b = parse_datetime_value(val_b)
    return _finish_date_compare(val_a, val_b, d_a, ok_a, d_b, ok_b)


def _finish_date_compare(
    raw_a, raw_b, d_a, ok_a, d_b, ok_b
) -> Tuple[bool, str, str, Optional[str]]:
    issue = None
    if not ok_a and not ok_b:
        issue = f"Could not parse either value as date: '{raw_a}' / '{raw_b}'"
    elif not ok_a:
        issue = f"Could not parse Baseline value as date: '{raw_a}'"
    elif not ok_b:
        issue = f"Could not parse Comparison value as date: '{raw_b}'"

    if issue:
        from src.normalization import normalize_key_value
        return normalize_key_value(raw_a) == normalize_key_value(raw_b), \
               normalize_key_value(raw_a), normalize_key_value(raw_b), issue

    if d_a is None and d_b is None:
        return True, "", "", None

    if d_a is None or d_b is None:
        s_a = "" if d_a is None else str(d_a)
        s_b = "" if d_b is None else str(d_b)
        return False, s_a, s_b, None

    is_equal = d_a == d_b
    return is_equal, str(d_a), str(d_b), None
