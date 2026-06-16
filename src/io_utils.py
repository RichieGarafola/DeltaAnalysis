"""
File I/O utilities.

Reads CSV and Excel uploads into DataFrames with consistent error
handling that surfaces meaningful messages to end users.
"""
import pandas as pd
from typing import List, Optional, Tuple

SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xls")

LARGE_FILE_WARN_ROWS = 100_000
LARGE_FILE_HARD_ROWS = 500_000
PREVIEW_MAX_ROWS = 1_000
RAW_PREVIEW_ROWS = 25


def get_excel_sheet_names(uploaded_file) -> List[str]:
    """
    Return the sheet names from an Excel upload without consuming the file.

    Seeks back to position 0 after reading so the caller can subsequently
    call read_uploaded_file on the same object.

    Returns an empty list for non-Excel files (CSVs, etc.).
    """
    name = uploaded_file.name.lower()
    if not name.endswith((".xlsx", ".xls")):
        return []

    try:
        uploaded_file.seek(0)
        xf = pd.ExcelFile(uploaded_file)
        sheets = xf.sheet_names
    except Exception:
        sheets = []
    finally:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass

    return sheets


def read_uploaded_file(uploaded_file, sheet_name=None) -> pd.DataFrame:
    """
    Parse a Streamlit UploadedFile into a DataFrame.

    All values are read as strings to prevent silent type coercion.
    Column names are whitespace-stripped on load.

    Parameters
    ----------
    uploaded_file : Streamlit UploadedFile object
    sheet_name    : Excel sheet name to read; None reads the first sheet.
                    Ignored for CSV files.

    Raises ValueError with a user-friendly message on any failure.
    """
    if uploaded_file is None:
        raise ValueError("No file was provided.")

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
        except Exception as exc:
            raise ValueError(f"Could not parse CSV file '{uploaded_file.name}': {exc}") from exc

    elif name.endswith((".xlsx", ".xls")):
        try:
            kwargs = {"dtype": str, "keep_default_na": False}
            if sheet_name is not None:
                kwargs["sheet_name"] = sheet_name
            df = pd.read_excel(uploaded_file, **kwargs)
        except Exception as exc:
            raise ValueError(
                f"Could not parse Excel file '{uploaded_file.name}': {exc}"
            ) from exc

    else:
        ext = name.rsplit(".", 1)[-1] if "." in name else "unknown"
        raise ValueError(
            f"Unsupported file format '.{ext}'. "
            f"Please upload a CSV (.csv) or Excel (.xlsx / .xls) file."
        )

    if df.empty:
        raise ValueError(f"The uploaded file '{uploaded_file.name}' contains no data rows.")

    # Strip whitespace from column headers
    df.columns = [str(c).strip() for c in df.columns]

    # Reject duplicate column names (they cause ambiguous merges)
    dupes = df.columns[df.columns.duplicated()].tolist()
    if dupes:
        raise ValueError(
            f"Duplicate column names detected in '{uploaded_file.name}': {dupes}. "
            "Please rename or remove duplicate columns before uploading."
        )

    return df


def check_file_size(n_rows: int) -> Tuple[str, str]:
    """
    Classify a row count and return a status/message pair.

    Returns
    -------
    ('ok',   '')                           - under warn threshold
    ('warn', human-readable message)       - between warn and hard threshold
    ('hard', human-readable message)       - at or above hard threshold
    """
    if n_rows >= LARGE_FILE_HARD_ROWS:
        return (
            "hard",
            f"This file has {n_rows:,} rows, which exceeds the {LARGE_FILE_HARD_ROWS:,}-row "
            "limit. Processing may exhaust available memory. Confirm below to proceed anyway.",
        )
    if n_rows >= LARGE_FILE_WARN_ROWS:
        return (
            "warn",
            f"This file has {n_rows:,} rows. Processing large files may be slow. "
            "Consider pre-filtering the data before uploading.",
        )
    return "ok", ""


def get_display_frame(
    df: pd.DataFrame, max_rows: int = PREVIEW_MAX_ROWS
) -> Tuple[pd.DataFrame, bool]:
    """
    Return a display-safe slice of a DataFrame.

    Returns (df_slice, was_truncated). When df has more than max_rows rows,
    returns the first max_rows rows and sets was_truncated=True.
    """
    if len(df) > max_rows:
        return df.iloc[:max_rows].copy(), True
    return df, False


def get_column_preview(df: pd.DataFrame, n_rows: int = 5) -> pd.DataFrame:
    """Return the first n rows for display in the UI."""
    return df.head(n_rows)


def read_uploaded_file_raw(uploaded_file, sheet_name=None) -> pd.DataFrame:
    """
    Read a CSV or Excel file without assuming a header row.

    Returns a DataFrame with sequential integer column indices (0, 1, 2, ...).
    All cell values are returned as strings. Use prepare_dataframe_from_raw to
    assign column headers and clean the data before analysis.

    Seeks the file back to position 0 after reading.

    Raises ValueError with a user-readable message on any parse failure.
    """
    if uploaded_file is None:
        raise ValueError("No file was provided.")

    name = uploaded_file.name.lower()

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    if name.endswith(".csv"):
        try:
            df = pd.read_csv(
                uploaded_file,
                header=None,
                dtype=str,
                keep_default_na=False,
                na_values=[],
                skip_blank_lines=False,
            )
        except Exception as exc:
            raise ValueError(
                f"Could not parse CSV file '{uploaded_file.name}': {exc}"
            ) from exc

    elif name.endswith((".xlsx", ".xls")):
        try:
            kwargs: dict = {"header": None, "dtype": str, "keep_default_na": False}
            if sheet_name is not None:
                kwargs["sheet_name"] = sheet_name
            df = pd.read_excel(uploaded_file, **kwargs)
        except Exception as exc:
            raise ValueError(
                f"Could not parse Excel file '{uploaded_file.name}': {exc}"
            ) from exc

    else:
        ext = name.rsplit(".", 1)[-1] if "." in name else "unknown"
        raise ValueError(
            f"Unsupported file format '.{ext}'. "
            "Please upload a CSV (.csv) or Excel (.xlsx / .xls) file."
        )

    if df.empty:
        raise ValueError(
            f"The uploaded file '{uploaded_file.name}' contains no rows."
        )

    # Normalize all values to plain strings; replace any residual NaN with "".
    df = df.fillna("").astype(str)
    # Re-index columns as 0, 1, 2, ... regardless of what pandas produced.
    df.columns = range(len(df.columns))

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    return df


# Backward-compatible alias kept for existing call sites and tests.
read_raw_file = read_uploaded_file_raw


def get_raw_preview(
    raw_df: pd.DataFrame,
    max_rows: int = 25,
    max_cols: int = 15,
) -> pd.DataFrame:
    """
    Return a display-safe preview of a raw DataFrame.

    Prepends a 'Source Row' column with 1-based row numbers matching what the
    user sees in the header-row selector. Does not mutate the original DataFrame.

    Parameters
    ----------
    raw_df   : DataFrame returned by read_uploaded_file_raw
    max_rows : Maximum rows to include (default 25)
    max_cols : Maximum data columns to include (default 15)
    """
    sliced = raw_df.head(max_rows)
    if max_cols < len(sliced.columns):
        sliced = sliced.iloc[:, :max_cols]
    preview = sliced.copy()
    preview.insert(0, "Source Row", range(1, len(preview) + 1))
    return preview


def prepare_dataframe_from_raw(
    raw_df: pd.DataFrame,
    header_row_index: int = 0,
    drop_rows_above: bool = True,
    drop_blank_rows: bool = True,
    end_row_index: Optional[int] = None,
) -> Tuple[pd.DataFrame, dict, list]:
    """
    Convert a raw (integer-columned) DataFrame into an analysis-ready DataFrame.

    Parameters
    ----------
    raw_df           : DataFrame returned by read_uploaded_file_raw
    header_row_index : 0-based row index containing column headers (default 0)
    drop_rows_above  : Exclude rows above the header when True (default True);
                       when False those rows are prepended to the data body
    drop_blank_rows  : Remove rows where every cell is empty (default True)
    end_row_index    : 0-based raw row index (inclusive) beyond which rows are
                       discarded; None means include all rows

    Returns
    -------
    (prepared_df, metadata, warnings)

    prepared_df : DataFrame with proper column names and cleaned rows
    metadata    : Dict — see keys below
    warnings    : List of professional warning strings

    Metadata keys
    -------------
    header_row_selected, drop_rows_above, drop_blank_rows, end_row_selected,
    rows_in_raw, rows_dropped_above_header, rows_dropped_blank,
    rows_dropped_after_end_row, rows_in_prepared, columns_in_prepared,
    blank_headers_renamed, duplicate_headers_renamed
    """
    warnings_out: list = []
    n_raw = len(raw_df)

    if n_raw == 0:
        raise ValueError("The raw DataFrame is empty; cannot assign a header row.")

    if not (0 <= header_row_index < n_raw):
        raise ValueError(
            f"Header row index {header_row_index} is out of range. "
            f"The file has {n_raw} row(s) (valid indices 0 to {n_raw - 1})."
        )

    # ------------------------------------------------------------------
    # Validate end_row_index; keep original for metadata, use effective for logic
    # ------------------------------------------------------------------
    end_row_effective = end_row_index
    if end_row_index is not None:
        if end_row_index <= header_row_index:
            warnings_out.append(
                f"The end row (row {end_row_index + 1}) is at or before the "
                f"header row (row {header_row_index + 1}). "
                "The end row setting has been ignored and all available rows will be included."
            )
            end_row_effective = None
        elif end_row_index >= n_raw:
            warnings_out.append(
                f"The end row (row {end_row_index + 1}) exceeds the available "
                f"row count ({n_raw}). All available rows will be included."
            )
            end_row_effective = None

    # ------------------------------------------------------------------
    # Build column names from the selected header row
    # ------------------------------------------------------------------
    header_values = raw_df.iloc[header_row_index]
    raw_names: list = [str(v).strip() for v in header_values]

    blank_count = 0
    col_names: list = []
    unnamed_counter = 1
    for i, raw_name in enumerate(raw_names):
        if raw_name in ("", "nan", "None", "<NA>", "NaN"):
            auto = f"Unnamed_{unnamed_counter}"
            unnamed_counter += 1
            blank_count += 1
            warnings_out.append(
                f"Column {i} has a blank header and has been renamed to '{auto}'."
            )
            col_names.append(auto)
        else:
            col_names.append(raw_name)

    dup_count = 0
    seen: dict = {}
    deduped: list = []
    for col_name in col_names:
        if col_name in seen:
            seen[col_name] += 1
            new_name = f"{col_name}_{seen[col_name]}"
            dup_count += 1
            warnings_out.append(
                f"Duplicate column header '{col_name}' has been renamed to '{new_name}'."
            )
            deduped.append(new_name)
        else:
            seen[col_name] = 1
            deduped.append(col_name)
    col_names = deduped

    # ------------------------------------------------------------------
    # Build full data body (without end-row filtering)
    # ------------------------------------------------------------------
    if drop_rows_above:
        full_data = raw_df.iloc[header_row_index + 1 :].copy()
        rows_dropped_above = header_row_index
    else:
        above = raw_df.iloc[:header_row_index]
        below = raw_df.iloc[header_row_index + 1 :]
        full_data = pd.concat([above, below], ignore_index=True)
        rows_dropped_above = 0

    n_full = len(full_data)

    # ------------------------------------------------------------------
    # Apply end-row filter
    # When drop_rows_above=True: n_keep = end_row_effective - header_row_index
    # When drop_rows_above=False: n_keep = end_row_effective
    # Both formulas give the correct count of rows in full_data to keep.
    # ------------------------------------------------------------------
    if end_row_effective is not None:
        if drop_rows_above:
            n_keep = max(0, end_row_effective - header_row_index)
        else:
            n_keep = max(0, end_row_effective)
        data_slice = full_data.iloc[:n_keep]
    else:
        data_slice = full_data
    rows_dropped_after_end = n_full - len(data_slice)

    df = data_slice.reset_index(drop=True)

    # ------------------------------------------------------------------
    # Assign column names (guard against column-count mismatch)
    # ------------------------------------------------------------------
    n_cols_df = len(df.columns) if not df.empty else len(col_names)
    if len(col_names) == n_cols_df:
        df.columns = col_names
    elif len(col_names) < n_cols_df:
        extras = [f"Extra_{i + 1}" for i in range(n_cols_df - len(col_names))]
        df.columns = col_names + extras
    else:
        df.columns = col_names[:n_cols_df]

    # ------------------------------------------------------------------
    # Drop blank rows
    # ------------------------------------------------------------------
    n_before_blank = len(df)
    if drop_blank_rows and not df.empty:
        blank_mask = df.apply(
            lambda row: all(str(v).strip() == "" for v in row),
            axis=1,
        )
        df = df[~blank_mask].reset_index(drop=True)
    n_blank_dropped = n_before_blank - len(df)

    # ------------------------------------------------------------------
    # Post-preparation warnings
    # ------------------------------------------------------------------
    if df.empty:
        warnings_out.append(
            "The prepared dataset has no rows after applying the selected settings. "
            "Review the header row, end row, and blank-row settings."
        )
    if not df.empty and len(df.columns) == 0:
        warnings_out.append("The prepared dataset has no columns.")

    # ------------------------------------------------------------------
    # Metadata
    # ------------------------------------------------------------------
    metadata: dict = {
        "header_row_selected":        header_row_index,
        "drop_rows_above":            drop_rows_above,
        "drop_blank_rows":            drop_blank_rows,
        "end_row_selected":           end_row_index,
        "rows_in_raw":                n_raw,
        "rows_dropped_above_header":  rows_dropped_above,
        "rows_dropped_blank":         n_blank_dropped,
        "rows_dropped_after_end_row": rows_dropped_after_end,
        "rows_in_prepared":           len(df),
        "columns_in_prepared":        len(df.columns),
        "blank_headers_renamed":      blank_count,
        "duplicate_headers_renamed":  dup_count,
    }

    return df, metadata, warnings_out
