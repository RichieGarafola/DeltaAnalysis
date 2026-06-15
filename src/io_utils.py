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


def read_raw_file(uploaded_file, sheet_name=None) -> pd.DataFrame:
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
    raw_df           : DataFrame returned by read_raw_file
    header_row_index : 0-based index of the row that contains column names
    drop_rows_above  : When True (default), rows before the header are excluded
                       from the data body; when False they are prepended to data
    drop_blank_rows  : When True (default), rows where every cell is empty are removed
    end_row_index    : Optional 0-based raw row index (inclusive) at which data ends;
                       rows beyond this index are discarded

    Returns
    -------
    (prepared_df, metadata, warnings)

    prepared_df : DataFrame with proper column names and cleaned rows
    metadata    : Dict with keys: header_row_selected, rows_dropped_above_header,
                  rows_dropped_blank, end_row_applied, final_row_count,
                  final_column_count
    warnings    : List of warning strings (blank headers renamed, duplicates renamed)
    """
    warnings_out: list = []
    n_raw = len(raw_df)

    if n_raw == 0:
        raise ValueError("The raw DataFrame is empty; cannot assign a header row.")

    if not (0 <= header_row_index < n_raw):
        raise ValueError(
            f"Header row index {header_row_index} is out of range. "
            f"The file has {n_raw} row(s) (valid indices: 0 to {n_raw - 1})."
        )

    # ------------------------------------------------------------------
    # Build column names from the selected header row
    # ------------------------------------------------------------------
    header_values = raw_df.iloc[header_row_index]
    raw_names: list = [str(v).strip() for v in header_values]

    # Replace blank/null headers with Unnamed_N
    col_names: list = []
    unnamed_counter = 1
    for i, name in enumerate(raw_names):
        if name in ("", "nan", "None", "<NA>", "NaN"):
            auto = f"Unnamed_{unnamed_counter}"
            unnamed_counter += 1
            warnings_out.append(
                f"Column {i} has a blank header; renamed to '{auto}'."
            )
            col_names.append(auto)
        else:
            col_names.append(name)

    # Deduplicate: suffix repeats with _2, _3, ...
    seen: dict = {}
    deduped: list = []
    for name in col_names:
        if name in seen:
            seen[name] += 1
            new_name = f"{name}_{seen[name]}"
            warnings_out.append(
                f"Duplicate column header '{name}' renamed to '{new_name}'."
            )
            deduped.append(new_name)
        else:
            seen[name] = 1
            deduped.append(name)
    col_names = deduped

    # ------------------------------------------------------------------
    # Build data body
    # ------------------------------------------------------------------
    if drop_rows_above:
        if end_row_index is not None:
            data_slice = raw_df.iloc[header_row_index + 1 : end_row_index + 1]
        else:
            data_slice = raw_df.iloc[header_row_index + 1 :]
        rows_dropped_above = header_row_index
    else:
        above = raw_df.iloc[:header_row_index]
        if end_row_index is not None:
            below = raw_df.iloc[header_row_index + 1 : end_row_index + 1]
        else:
            below = raw_df.iloc[header_row_index + 1 :]
        data_slice = pd.concat([above, below], ignore_index=True)
        rows_dropped_above = 0

    df = data_slice.copy().reset_index(drop=True)

    # Assign column names (guard against column-count mismatch)
    n_cols_raw = len(df.columns) if not df.empty else len(col_names)
    if len(col_names) == n_cols_raw:
        df.columns = col_names
    elif len(col_names) < n_cols_raw:
        extras = [f"Extra_{i + 1}" for i in range(n_cols_raw - len(col_names))]
        df.columns = col_names + extras
    else:
        df.columns = col_names[:n_cols_raw]

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
    # Metadata
    # ------------------------------------------------------------------
    metadata: dict = {
        "header_row_selected":      header_row_index,
        "rows_dropped_above_header": rows_dropped_above,
        "rows_dropped_blank":        n_blank_dropped,
        "end_row_applied":           end_row_index,
        "final_row_count":           len(df),
        "final_column_count":        len(df.columns),
    }

    return df, metadata, warnings_out
