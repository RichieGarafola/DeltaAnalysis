"""
File I/O utilities.

Reads CSV and Excel uploads into DataFrames with consistent error
handling that surfaces meaningful messages to end users.
"""
import pandas as pd

SUPPORTED_EXTENSIONS = (".csv", ".xlsx", ".xls")


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """
    Parse a Streamlit UploadedFile into a DataFrame.

    All values are read as strings to prevent silent type coercion.
    Column names are whitespace-stripped on load.

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
            df = pd.read_excel(uploaded_file, dtype=str, keep_default_na=False)
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

    # Reject duplicate column names — they cause ambiguous merges
    dupes = df.columns[df.columns.duplicated()].tolist()
    if dupes:
        raise ValueError(
            f"Duplicate column names detected in '{uploaded_file.name}': {dupes}. "
            "Please rename or remove duplicate columns before uploading."
        )

    return df


def get_column_preview(df: pd.DataFrame, n_rows: int = 5) -> pd.DataFrame:
    """Return the first n rows for display in the UI."""
    return df.head(n_rows)
