# DeltaAnalysis

> In-depth Delta Analysis platform. Performs bidirectional delta analysis and provides a detailed, client-facing report.

[![Tests](https://github.com/RichieGarafola/DeltaAnalysis/actions/workflows/tests.yml/badge.svg)](https://github.com/RichieGarafola/DeltaAnalysis/actions/workflows/tests.yml)

A production-quality Python + Streamlit application for comparing two datasets side-by-side. Built for government analysts performing reconciliation, receipt tracking, case audits, DFAS-style package reviews, and duplicate submission checks.

---

## Why this tool exists

Government data workflows constantly produce two versions of the same dataset — last week's extract vs this week's, contractor-submitted vs agency-received, a source system export vs a downstream database snapshot. Identifying *what changed*, *what's missing*, and *what's new* is a recurring, error-prone manual process. This tool automates it with a reproducible, auditable, briefing-ready output.

---

## What it does

Upload **File A** (baseline) and **File B** (comparison), configure the columns that uniquely identify each record, and click Run.

| Category | Description |
|---|---|
| **Only in File A** | Records present in A but not B — potential deletions or missing submissions |
| **Only in File B** | Records present in B but not A — new arrivals or unmatched entries |
| **Matched Records** | All records with a common key, shown side-by-side |
| **Changed Records** | Matched records where a compared field has a different value |
| **Duplicate Keys** | Rows that share a key within the same file — data quality flag |
| **Blank / Null Keys** | Rows excluded because the key field is empty — flagged for correction |

Results are shown in an interactive dashboard with:
- **10 KPI cards** for at-a-glance counts (totals, matched, changed, duplicates, blanks)
- **3 Plotly charts**: delta category bar chart, match coverage donut, field-change frequency
- **8 tabbed result tables** with per-category CSV downloads
- **11-tab Excel workbook** with an auto-generated Executive Summary and metadata tabs

---

## Installation

### 1 — Clone the repo

```bash
git clone https://github.com/RichieGarafola/DeltaAnalysis.git
cd DeltaAnalysis
```

### 2 — Create a virtual environment

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 3 — Install dependencies

```bash
pip install -r requirements.txt
```

**Dependencies:** `streamlit`, `pandas`, `openpyxl`, `xlrd`, `plotly`, `pytest`, `pytest-cov`

---

## Running the application

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`. No configuration required.

---

## Running the tests

```bash
pytest tests/ -v
```

Expected output: **103 tests passing** across four test modules.

With coverage:

```bash
pytest tests/ --cov=src --cov-report=term-missing
```

---

## Generating reports

1. Launch the app (`streamlit run app.py`)
2. Upload File A and File B (CSV or Excel)
3. Select key column(s) from each file
4. Select comparison column(s) to diff
5. Click **Run Delta Analysis**
6. (Optional) Expand **Advanced Comparison Settings** to configure numeric tolerance or date-aware comparison per column
7. Use the **Download Excel Report** button for the full 11-tab workbook
8. Use the per-tab **Download CSV** buttons for individual category exports

The Excel workbook tabs:
- **Executive Summary** — auto-generated plain-English briefing text
- **Analysis Metadata** — file names, sheets used, row counts, key columns, timestamp
- **Comparison Rules** — per-column comparison type, tolerance, and date mode
- **Delta Counts** — flat table suitable for pivot tables
- **Only in File A / B** — unmatched records
- **Matched Records** — side-by-side A + B view
- **Changed Records** — before/after values per changed field
- **Duplicate Keys File A / B** — non-unique key rows
- **Data Quality Issues** — blank/null key rows

---

## Using the sample data

Two sets of sample data are included:

### Minimal demo (`file_a.csv` / `file_b.csv`)

Designed to exercise every delta category in a single small dataset.

```
sample_data/file_a.csv  — 6 rows: 1 matched-unchanged, 1 matched-changed,
                           1 only-in-A, 1 duplicate key (x2), 1 blank key
sample_data/file_b.csv  — 4 rows: corresponding comparison data
```

Recommended settings:
- Key column: `RecordID` (both files)
- Comparison columns: `Status`, `Amount` (both files)

Expected results:

| Category | Count |
|---|---|
| Only in File A | 1 (R003) |
| Only in File B | 1 (R005) |
| Matched | 3 (R001 unchanged, R002 changed, R004 from first occurrence) |
| Changed | 1 (R002: Status Pending→Approved, Amount 8500→9000) |
| Duplicate Keys A | 2 (both R004 rows flagged) |
| Blank Keys A | 1 |

### Full contracting dataset (`sample_a.csv` / `sample_b.csv`)

```
sample_data/sample_a.csv  — 11-row baseline (includes 1 duplicate, 1 blank key)
sample_data/sample_b.csv  — 10-row comparison (changed fields, new entries)
```

Recommended settings:
- Key column: `CaseID` (both files)
- Comparison columns: `ContractAmount`, `Status`, `ReviewedBy` (both files)

---

## Project structure

```
DeltaAnalysis/
├── app.py                      # Streamlit UI — main entry point
├── requirements.txt
├── README.md
├── .gitignore
├── .github/
│   └── workflows/
│       └── tests.yml           # CI: runs pytest on push and pull_request
├── src/
│   ├── __init__.py
│   ├── normalization.py        # Key cleaning: trim, fix "1234.0", handle nulls
│   ├── io_utils.py             # File upload parsing, sheet names, size guards
│   ├── comparison.py           # Type-aware field comparison (numeric, date, text)
│   ├── delta_engine.py         # Core comparison → DeltaResult dataclass
│   └── reporting.py            # 12-tab Excel export with Executive Summary
├── tests/
│   ├── __init__.py
│   ├── test_normalization.py   # 21 unit tests
│   ├── test_delta_engine.py    # 26 unit tests
│   ├── test_comparison.py      # 32 unit tests (numeric, date, field comparison)
│   └── test_io_utils.py        # 24 unit tests (sheets, size checks, display)
└── sample_data/
    ├── file_a.csv              # Minimal demo — one of each delta category
    ├── file_b.csv              # Minimal demo — comparison side
    ├── sample_a.csv            # Full contracting dataset (11 rows)
    └── sample_b.csv            # Full contracting dataset (10 rows)
```

---

## CI/CD

A GitHub Actions workflow (`.github/workflows/tests.yml`) runs the full test suite automatically on every push and pull request. The build fails if any test fails.

---

## Key design decisions

**All data read as strings.** Prevents silent type coercion — a common cause of false mismatches when Excel stores IDs as numbers (e.g., `1001` read as `1001.0`).

**Composite key support.** Multiple columns combine into one match key (e.g., Fiscal Year + Case ID), joined with `||` to prevent collisions between single-column values.

**Blank keys are quarantined, not silently dropped.** Rows with null keys are reported separately so analysts see what was excluded and why — important for audit trails.

**Duplicates use first-occurrence for matching.** All duplicate rows are surfaced in the Duplicates tab; matching uses only the first occurrence so totals remain predictable.

**Executive Narrative is fully data-driven.** Every number in the narrative text is derived from the actual DeltaResult — no manual editing required before briefing leadership.

---

## Supported file formats

| Format | Extension |
|---|---|
| CSV | `.csv` |
| Excel (modern) | `.xlsx` |
| Excel (legacy) | `.xls` |

---

## Screenshots

*Screenshots placeholder — add after first deployment.*

---

## Government use cases

- **DFAS receipt reconciliation** — match invoices by contract number; flag amount or status changes
- **M&RA package tracking** — compare weekly extracts; surface new submissions and status deltas
- **Duplicate submission review** — identify cases where the same ID appears multiple times
- **Operational reporting** — produce a briefing-ready delta between two reporting periods
- **Audit support** — every result category is traceable to its source rows; Excel export is sharable and self-contained

---

## Current Limitations

- **Memory-bound processing.** Very large files, especially 500k+ rows, may be slow or exhaust browser/session memory. The app warns at 100k rows and requires confirmation at 500k rows.
- **No authentication.** The Streamlit app has no login wall. Do not deploy to a public URL with sensitive government data unless appropriate security controls are in place.
- **No persistent saved configurations.** Key selections and comparison mappings must be selected each session.
- **No row-level analyst annotations yet.**
- **No scheduled/automated runs yet.**

### Implemented in v1.1

- ✅ Multi-sheet Excel support with sheet selector (shown only when file has multiple sheets)
- ✅ Numeric tolerance comparison — strip `$`, `£`, `€`, commas; handle `(1,000)` negatives; configurable tolerance per column
- ✅ Date-aware comparison — `01/15/2024` and `2024-01-15` are recognised as the same date
- ✅ Large-file safety — warning banner at 100k rows, confirmation gate at 500k rows, table previews capped at 1,000 rows
- ✅ Analysis Metadata and Comparison Rules tabs in the Excel export
- ✅ Parse-issue log surfaced in the Changed Records tab when numeric/date parsing encounters unexpected values

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `Duplicate column names detected` | Two columns share the same name | Rename the duplicate column in the source file |
| `Columns not found in File A key columns` | Selected column no longer matches after re-upload | Re-select key columns after uploading |
| `Key column counts must match` | Different number of keys selected for File A vs File B | Select the same number of key columns in each file |
| `Could not parse Excel file` | File may be password-protected, corrupted, or unsupported | Save as a new `.xlsx` from Excel and re-upload |
| `The uploaded file contains no data rows` | File has headers only | Confirm the file has data rows below the header |
| No comparison columns selected | User selected keys only | Select non-key fields if field-level change detection is needed |
| Report export failed | Workbook generation issue or invalid sheet data | Re-run analysis and check for unsupported values |

---

## Roadmap

### Near-Term
- [ ] Saved configuration profiles
- [ ] Column rename/mapping UI
- [ ] More robust data quality checks
- [ ] PDF report export for leadership briefings

### Medium-Term
- [ ] Scheduled or automated comparison runs via CLI
- [ ] Row-level audit log with analyst annotations
- [ ] Reconciliation modes for DFAS, M&RA Package Tracking, Veteran Case Tracking, and Generic Delta Analysis
- [ ] Optional deployment guide for secure internal environments

### Longer-Term
- [ ] Authentication and role-based access
- [ ] Database-backed processing for large files
- [ ] SharePoint/OneDrive file integration
- [ ] ADVANA/export pipeline integration
- [ ] Power BI-ready output tables
