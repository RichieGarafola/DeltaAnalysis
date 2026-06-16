"""
Core delta analysis engine.

Performs a bidirectional comparison between two DataFrames and
produces a structured DeltaResult containing every category of
difference needed for audit-grade reporting.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional

import pandas as pd

from src.normalization import build_composite_key, normalize_key_value


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------

@dataclass
class DeltaResult:
    """Holds the full output of a delta comparison run."""

    only_in_a: pd.DataFrame
    only_in_b: pd.DataFrame
    matched: pd.DataFrame          # All common-key records, side-by-side
    changed: pd.DataFrame          # Matched records where compared fields differ
    duplicates_a: pd.DataFrame     # Rows sharing a key within the Source Dataset
    duplicates_b: pd.DataFrame     # Rows sharing a key within the Comparison Dataset
    blank_keys_a: pd.DataFrame     # Rows with blank/null keys in the Source Dataset
    blank_keys_b: pd.DataFrame     # Rows with blank/null keys in the Comparison Dataset

    key_cols_a: List[str]
    key_cols_b: List[str]
    compare_cols_a: List[str]      # Comparison field names from the Source Dataset
    compare_cols_b: List[str]      # Corresponding field names from the Comparison Dataset

    total_a: int                   # Original row count before any filtering
    total_b: int

    # v1.1 additions: optional, with defaults for backward compatibility
    comparison_rules: List[dict] = field(default_factory=list)
    sheet_a: Optional[str] = None
    sheet_b: Optional[str] = None
    compare_parse_issues: Optional[pd.DataFrame] = None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def run_delta(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_cols_a: List[str],
    key_cols_b: List[str],
    compare_cols_a: Optional[List[str]] = None,
    compare_cols_b: Optional[List[str]] = None,
    comparison_rules: Optional[List[dict]] = None,
    sheet_a: Optional[str] = None,
    sheet_b: Optional[str] = None,
) -> DeltaResult:
    """
    Run a full bidirectional delta analysis.

    Parameters
    ----------
    df_a, df_b           : Source DataFrames (read as strings recommended)
    key_cols_a/b         : Columns that uniquely identify a record in each file.
                           Matched positionally; first col in A pairs with
                           first col in B, etc.
    compare_cols_a/b     : Non-key columns to diff for changes.
                           Must be the same length; compared positionally.
    comparison_rules     : List of rule dicts controlling per-column comparison
                           type, tolerance, and date handling.  When None,
                           defaults to text comparison for all columns (same
                           behaviour as v1.0).
    sheet_a, sheet_b     : Excel sheet names used when reading each file;
                           stored on the result for reporting purposes only.

    Returns
    -------
    DeltaResult dataclass with all comparison categories populated.
    """
    compare_cols_a = compare_cols_a or []
    compare_cols_b = compare_cols_b or []

    # --- Validate inputs ---------------------------------------------------
    _validate_columns(df_a, key_cols_a, "Source Dataset")
    _validate_columns(df_b, key_cols_b, "Comparison Dataset")

    if len(key_cols_a) != len(key_cols_b):
        raise ValueError(
            f"Key column counts must match: "
            f"Source has {len(key_cols_a)}, Comparison has {len(key_cols_b)}."
        )

    if compare_cols_a or compare_cols_b:
        _validate_columns(df_a, compare_cols_a, "Source Dataset comparison fields")
        _validate_columns(df_b, compare_cols_b, "Comparison Dataset comparison fields")
        if len(compare_cols_a) != len(compare_cols_b):
            raise ValueError(
                f"Comparison column counts must match: "
                f"Source has {len(compare_cols_a)}, Comparison has {len(compare_cols_b)}."
            )

    total_a = len(df_a)
    total_b = len(df_b)

    # Build or normalise comparison rules
    rules = _normalise_rules(comparison_rules, compare_cols_a, compare_cols_b)

    # --- Build working copies with composite keys --------------------------
    a = df_a.copy().reset_index(drop=True)
    b = df_b.copy().reset_index(drop=True)

    a["__key__"] = build_composite_key(a, key_cols_a)
    b["__key__"] = build_composite_key(b, key_cols_b)

    # --- Separate blank-key rows -------------------------------------------
    blank_a = a[a["__key__"] == "__BLANK__"].drop(columns=["__key__"]).reset_index(drop=True)
    blank_b = b[b["__key__"] == "__BLANK__"].drop(columns=["__key__"]).reset_index(drop=True)

    a = a[a["__key__"] != "__BLANK__"].reset_index(drop=True)
    b = b[b["__key__"] != "__BLANK__"].reset_index(drop=True)

    # --- Identify duplicates (full duplicate group, including first row) ----
    dup_mask_a = a.duplicated(subset=["__key__"], keep=False)
    dup_mask_b = b.duplicated(subset=["__key__"], keep=False)

    dupes_a = (
        a[dup_mask_a]
        .sort_values("__key__")
        .drop(columns=["__key__"])
        .reset_index(drop=True)
    )
    dupes_b = (
        b[dup_mask_b]
        .sort_values("__key__")
        .drop(columns=["__key__"])
        .reset_index(drop=True)
    )

    # Keep first occurrence of each key for matching
    a_dedup = a.drop_duplicates(subset=["__key__"], keep="first")
    b_dedup = b.drop_duplicates(subset=["__key__"], keep="first")

    # --- Set membership ----------------------------------------------------
    keys_a = set(a_dedup["__key__"])
    keys_b = set(b_dedup["__key__"])

    only_a_keys = keys_a - keys_b
    only_b_keys = keys_b - keys_a
    common_keys = keys_a & keys_b

    only_a = (
        a_dedup[a_dedup["__key__"].isin(only_a_keys)]
        .drop(columns=["__key__"])
        .reset_index(drop=True)
    )
    only_b = (
        b_dedup[b_dedup["__key__"].isin(only_b_keys)]
        .drop(columns=["__key__"])
        .reset_index(drop=True)
    )

    # --- Matched records (side-by-side) ------------------------------------
    matched_a = a_dedup[a_dedup["__key__"].isin(common_keys)].set_index("__key__")
    matched_b = b_dedup[b_dedup["__key__"].isin(common_keys)].set_index("__key__")

    a_prefixed = matched_a.rename(columns=lambda c: f"Source: {c}")
    b_prefixed = matched_b.rename(columns=lambda c: f"Comparison: {c}")
    matched_combined = (
        a_prefixed.join(b_prefixed, how="inner")
        .reset_index(drop=True)
    )

    # --- Changed records ---------------------------------------------------
    changed, parse_issues = _find_changed(
        matched_a, matched_b, common_keys,
        key_cols_a, rules,
    )

    parse_issues_df = pd.DataFrame(parse_issues) if parse_issues else None

    return DeltaResult(
        only_in_a=only_a,
        only_in_b=only_b,
        matched=matched_combined,
        changed=changed,
        duplicates_a=dupes_a,
        duplicates_b=dupes_b,
        blank_keys_a=blank_a,
        blank_keys_b=blank_b,
        key_cols_a=key_cols_a,
        key_cols_b=key_cols_b,
        compare_cols_a=compare_cols_a,
        compare_cols_b=compare_cols_b,
        total_a=total_a,
        total_b=total_b,
        comparison_rules=rules,
        sheet_a=sheet_a,
        sheet_b=sheet_b,
        compare_parse_issues=parse_issues_df,
    )


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _normalise_rules(
    comparison_rules: Optional[List[dict]],
    compare_cols_a: List[str],
    compare_cols_b: List[str],
) -> List[dict]:
    """
    Return a validated rule list.

    If comparison_rules is None or empty, build a default text rule for each
    column pair, preserving v1.0 behaviour unchanged.
    """
    if comparison_rules:
        return comparison_rules
    return _build_default_rules(compare_cols_a, compare_cols_b)


def _build_default_rules(
    cols_a: List[str],
    cols_b: List[str],
) -> List[dict]:
    """Build text-type comparison rules for every column pair."""
    return [
        {"column_a": ca, "column_b": cb, "type": "text", "tolerance": None, "date_mode": None}
        for ca, cb in zip(cols_a, cols_b)
    ]


def _find_changed(
    matched_a: pd.DataFrame,
    matched_b: pd.DataFrame,
    common_keys: set,
    key_cols_a: List[str],
    rules: List[dict],
):
    """
    Return (changed_df, parse_issue_list) where any compared field differs.

    Uses comparison_rules for type-aware comparison; falls back to text when
    rules is empty (no comparison columns selected).
    """
    from src.comparison import compare_field_values

    if not rules:
        return pd.DataFrame(), []

    rows = []
    parse_issues = []

    for key in sorted(common_keys):
        row_a = matched_a.loc[key]
        row_b = matched_b.loc[key]

        diffs = {}
        for rule in rules:
            col_a = rule["column_a"]
            col_b = rule["column_b"]

            val_a = row_a.get(col_a, "")
            val_b = row_b.get(col_b, "")

            is_equal, str_a, str_b, issue = compare_field_values(val_a, val_b, rule)

            if issue:
                parse_issues.append({
                    "key": key,
                    "column_a": col_a,
                    "column_b": col_b,
                    "value_a": val_a,
                    "value_b": val_b,
                    "issue": issue,
                })

            if not is_equal:
                diffs[col_a] = (str_a, str_b)

        if diffs:
            record: dict = {}
            for kc in key_cols_a:
                record[f"Key: {kc}"] = row_a.get(kc, "")
            for col, (before, after) in diffs.items():
                record[f"{col} - Source"] = before
                record[f"{col} - Comparison"] = after
            rows.append(record)

    changed_df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return changed_df, parse_issues


def _validate_columns(df: pd.DataFrame, cols: List[str], label: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Columns not found in {label}: {missing}")
