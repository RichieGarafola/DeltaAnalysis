"""
File I/O utilities.

Reads CSV and Excel uploads into DataFrames with consistent error
handling that surfaces meaningful messages to end users.
"""
import pandas as pd
from typing import List, Tuple

SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xls")

LARGE_FILE_WARN_ROWS = 100_000
LARGE_FILE_HARD_ROWS = 500_000
PREVIEW_MAX_ROWS = 1_000


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
