"""
Core delta analysis engine.

Performs a bidirectional comparison between two DataFrames and
produces a structured DeltaResult containing every category of
difference needed for audit-grade reporting.
"""
from __future__ import annotations

from dataclasses import dataclass
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
    duplicates_a: pd.DataFrame     # All rows that share a key in File A
    duplicates_b: pd.DataFrame     # All rows that share a key in File B
    blank_keys_a: pd.DataFrame     # Rows with blank/null keys in File A
    blank_keys_b: pd.DataFrame     # Rows with blank/null keys in File B

    key_cols_a: List[str]
    key_cols_b: List[str]
    compare_cols_a: List[str]      # Comparison column names from File A
    compare_cols_b: List[str]      # Corresponding column names from File B

    total_a: int                   # Original row count before any filtering
    total_b: int


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
) -> DeltaResult:
    """
    Run a full bidirectional delta analysis.

    Parameters
    ----------
    df_a, df_b         : Source DataFrames (read as strings recommended)
    key_cols_a/b       : Columns that uniquely identify a record in each file.
                         Matched positionally — first col in A pairs with
                         first col in B, etc.
    compare_cols_a/b   : Non-key columns to diff for changes.
                         Must be the same length; compared positionally.

    Returns
    -------
    DeltaResult dataclass with all comparison categories populated.
    """
    compare_cols_a = compare_cols_a or []
    compare_cols_b = compare_cols_b or []

    # --- Validate inputs ---------------------------------------------------
    _validate_columns(df_a, key_cols_a, "File A key columns")
    _validate_columns(df_b, key_cols_b, "File B key columns")

    if len(key_cols_a) != len(key_cols_b):
        raise ValueError(
            f"Key column counts must match: "
            f"File A has {len(key_cols_a)}, File B has {len(key_cols_b)}."
        )

    if compare_cols_a or compare_cols_b:
        _validate_columns(df_a, compare_cols_a, "File A comparison columns")
        _validate_columns(df_b, compare_cols_b, "File B comparison columns")
        if len(compare_cols_a) != len(compare_cols_b):
            raise ValueError(
                f"Comparison column counts must match: "
                f"File A has {len(compare_cols_a)}, File B has {len(compare_cols_b)}."
            )

    total_a = len(df_a)
    total_b = len(df_b)

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

    a_prefixed = matched_a.rename(columns=lambda c: f"A: {c}")
    b_prefixed = matched_b.rename(columns=lambda c: f"B: {c}")
    matched_combined = (
        a_prefixed.join(b_prefixed, how="inner")
        .reset_index(drop=True)
    )

    # --- Changed records ---------------------------------------------------
    changed = _find_changed(
        matched_a, matched_b, common_keys,
        key_cols_a, compare_cols_a, compare_cols_b,
    )

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
    )


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _find_changed(
    matched_a: pd.DataFrame,
    matched_b: pd.DataFrame,
    common_keys: set,
    key_cols_a: List[str],
    compare_cols_a: List[str],
    compare_cols_b: List[str],
) -> pd.DataFrame:
    """Return rows where any compared field differs between A and B."""
    if not compare_cols_a:
        return pd.DataFrame()

    rows = []
    for key in sorted(common_keys):
        row_a = matched_a.loc[key]
        row_b = matched_b.loc[key]

        diffs = {}
        for col_a, col_b in zip(compare_cols_a, compare_cols_b):
            val_a = normalize_key_value(row_a.get(col_a, ""))
            val_b = normalize_key_value(row_b.get(col_b, ""))
            if val_a != val_b:
                diffs[col_a] = (val_a, val_b)

        if diffs:
            record: dict = {}
            # Show the key field(s) so the user knows which record changed
            for kc in key_cols_a:
                record[f"Key: {kc}"] = row_a.get(kc, "")
            # Show before/after for each changed field
            for col, (before, after) in diffs.items():
                record[f"{col} — File A"] = before
                record[f"{col} — File B"] = after
            rows.append(record)

    return pd.DataFrame(rows) if rows else pd.DataFrame()


def _validate_columns(df: pd.DataFrame, cols: List[str], label: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"Columns not found in {label}: {missing}")
