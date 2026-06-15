"""
Key normalization utilities.

Ensures consistent comparison across datasets that may have
inconsistent formatting, a common issue in government data exports.
"""
import math
import pandas as pd


def normalize_key_value(value) -> str:
    """
    Normalize a single key value to a clean, comparable string.

    Handles:
    - True None / float NaN
    - Whitespace padding
    - Float-formatted integers (e.g. "1234.0" → "1234")
    - String representations of null ("nan", "null", "none", "n/a")
    """
    # Catch None and float NaN
    if value is None:
        return "__BLANK__"
    if isinstance(value, float) and math.isnan(value):
        return "__BLANK__"

    s = str(value).strip()

    # Normalize string representations of null values
    if s == "" or s.lower() in ("nan", "none", "null", "na", "n/a", "<na>"):
        return "__BLANK__"

    # Strip accidental decimal artifacts from IDs stored as floats
    # e.g. Excel reads integer 1001 as 1001.0
    if s.endswith(".0"):
        try:
            s = str(int(float(s)))
        except (ValueError, OverflowError):
            pass

    return s


def normalize_key_series(series: pd.Series) -> pd.Series:
    """Apply normalize_key_value to every element of a Series."""
    return series.apply(normalize_key_value)


def build_composite_key(df: pd.DataFrame, key_cols: list) -> pd.Series:
    """
    Build a single composite key string from one or more columns.

    Columns are normalized and joined with '||' so that
    ('A', 'B') and ('A||B',) cannot collide.
    """
    if not key_cols:
        raise ValueError("At least one key column must be provided.")

    parts = [normalize_key_series(df[col]) for col in key_cols]
    combined = pd.Series(
        ["||".join(row) for row in zip(*[p.values for p in parts])],
        index=df.index,
        name="__composite_key__",
    )
    return combined
