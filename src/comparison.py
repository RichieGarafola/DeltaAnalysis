"""
Field-level comparison utilities.

Provides type-aware value comparison supporting:
- Numeric: optional tolerance for rounding/currency differences
- Date: format-agnostic comparison (MM/DD/YYYY vs YYYY-MM-DD, etc.)
- Text: normalized string comparison (default)
"""
from __future__ import annotations

import re
from datetime import date
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
    parenthesised negatives like (1000.00), and leading/trailing whitespace.

    Returns (float_value, True) on success, (None, False) on failure.
    Blank / NaN strings return (None, True) — treated as missing, not invalid.
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
# Date parsing
# ---------------------------------------------------------------------------

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
    Parse a value as a date, trying multiple common formats.

    Returns (date_object, True) on success, (None, False) on failure.
    Blank / NaN strings return (None, True) — treated as missing, not invalid.
    """
    if value is None:
        return None, True

    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "null", "n/a", "na"):
        return None, True

    # Try pandas first — it handles ISO 8601 and datetime strings well
    try:
        result = pd.to_datetime(s, dayfirst=False)
        return result.date(), True
    except (ValueError, TypeError):
        pass

    # Try dayfirst fallback for ambiguous formats like "01/02/2024"
    try:
        result = pd.to_datetime(s, dayfirst=True)
        return result.date(), True
    except (ValueError, TypeError):
        pass

    # Explicit format loop for edge cases
    for fmt in _DATE_FORMATS:
        try:
            from datetime import datetime
            return datetime.strptime(s, fmt).date(), True
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
    val_a, val_b : Raw values from each file.
    rule         : Dict with keys:
                   - type       : 'text' | 'numeric' | 'date'  (default 'text')
                   - tolerance  : float, used when type == 'numeric'
                   - date_mode  : 'us' | 'iso' | 'auto'        (default 'auto')

    Returns
    -------
    (is_equal, str_a, str_b, issue_note)

    is_equal    : True if values should be considered the same
    str_a       : Normalised display string for value A
    str_b       : Normalised display string for value B
    issue_note  : Non-None string describing a parse problem, if any
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
        issue = f"Could not parse File A value as numeric: '{val_a}'"
    elif not ok_b:
        issue = f"Could not parse File B value as numeric: '{val_b}'"

    if issue:
        # Fall back to text comparison when parse fails
        from src.normalization import normalize_key_value
        s_a = normalize_key_value(val_a)
        s_b = normalize_key_value(val_b)
        return s_a == s_b, s_a, s_b, issue

    # Both are None (both blank) -> equal
    if n_a is None and n_b is None:
        return True, "", "", None

    # One is None -> unequal
    if n_a is None or n_b is None:
        s_a = "" if n_a is None else str(n_a)
        s_b = "" if n_b is None else str(n_b)
        return False, s_a, s_b, None

    diff = abs(n_a - n_b)
    is_equal = diff <= tolerance
    return is_equal, str(n_a), str(n_b), None


def _compare_date(val_a, val_b, rule: dict) -> Tuple[bool, str, str, Optional[str]]:
    d_a, ok_a = parse_date_value(val_a)
    d_b, ok_b = parse_date_value(val_b)

    issue = None
    if not ok_a and not ok_b:
        issue = f"Could not parse either value as date: '{val_a}' / '{val_b}'"
    elif not ok_a:
        issue = f"Could not parse File A value as date: '{val_a}'"
    elif not ok_b:
        issue = f"Could not parse File B value as date: '{val_b}'"

    if issue:
        from src.normalization import normalize_key_value
        s_a = normalize_key_value(val_a)
        s_b = normalize_key_value(val_b)
        return s_a == s_b, s_a, s_b, issue

    if d_a is None and d_b is None:
        return True, "", "", None

    if d_a is None or d_b is None:
        s_a = "" if d_a is None else d_a.isoformat()
        s_b = "" if d_b is None else d_b.isoformat()
        return False, s_a, s_b, None

    is_equal = d_a == d_b
    return is_equal, d_a.isoformat(), d_b.isoformat(), None
