# Delta Analysis Tool

> Bidirectional dataset reconciliation for government analysts. Generates auditable, briefing-ready comparison reports from two CSV or Excel files.

[![Tests](https://github.com/RichieGarafola/DeltaAnalysis/actions/workflows/tests.yml/badge.svg)](https://github.com/RichieGarafola/DeltaAnalysis/actions/workflows/tests.yml)

A Python + Streamlit application built for analysts performing financial reconciliation, receipt tracking, contract audits, DFAS-style package reviews, and duplicate submission checks. Upload a Baseline Dataset and a Comparison Dataset, define your match keys, and produce a complete delta analysis in seconds.

---

## Business Use Cases

Government data workflows routinely generate two versions of the same dataset: last week's extract vs. this week's, contractor-submitted vs. agency-received, a source system pull vs. a downstream database snapshot. Identifying *what changed*, *what's missing*, and *what's new* is a recurring, error-prone manual process. This tool automates that reconciliation with a reproducible, audit-traceable, briefing-ready output.

Typical scenarios:

- **DFAS receipt reconciliation:** match invoices by contract number; flag obligation amount or status changes
- **M&RA package tracking:** compare weekly extracts; surface new submissions and status deltas between periods
- **Duplicate submission review:** identify records where the same identifier appears multiple times in the same file
- **Operational reporting:** produce a leadership-ready delta between two reporting periods
- **Audit support:** every result category is traceable to its source rows; the Excel export is self-contained and shareable with oversight stakeholders

---

## Key Features

| Feature | Detail |
|---|---|
| Bidirectional comparison | Six result categories: Source Only, Comparison Only, Matched, Records with Differences, Duplicate Identifiers, Data Quality Flags |
| Data preparation workflow | User-controlled header row selection, row trimming above/below, and blank row removal before reconciliation |
| Multi-sheet Excel support | Sheet selector appears automatically when an uploaded file contains multiple worksheets |
| Numeric tolerance | Strips currency symbols ($, £, €), commas, and parenthesised negatives; configurable tolerance per field |
| Date-aware comparison | Parses ISO and US date formats; `date_only` mode ignores time; `datetime_precision` mode preserves time |
| Large-file safeguards | Warning banner at 100,000 rows; confirmation gate at 500,000 rows; table previews capped at 1,000 rows |
| Executive Summary | Auto-generated plain-English briefing narrative (every number is derived from the actual results) |
| 11-tab Excel workbook | Audit-ready export with metadata, comparison rules, delta counts, and per-category data tabs |
| 253-test suite | Covers normalization, comparison logic, I/O utilities, data preparation, reporting, and end-to-end scenarios |

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/RichieGarafola/DeltaAnalysis.git
cd DeltaAnalysis
```

### 2. Create a virtual environment

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

**Core dependencies:** `streamlit`, `pandas`, `openpyxl`, `xlrd`, `plotly`, `pytest`, `pytest-cov`

---

## Usage

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`. No configuration file required.

**Workflow:**

1. Upload the **Source Dataset** and **Comparison Dataset** (CSV or Excel)
2. Select the worksheet (if the file contains multiple sheets)
3. Prepare headers and rows (select the header row, trim title/footer rows, drop blank rows)
4. Select the **match key column(s)** from each dataset (these uniquely identify each record, e.g., Contract Number, Case ID)
5. Select **comparison fields** to diff across matched records (e.g., obligation amount, status, date)
6. Click **Run Delta Analysis**
7. Review results in the interactive dashboard; download per-category CSVs or the full Excel report

---

## Comparison Modes

### Text (default)
All fields default to exact text comparison after whitespace normalization. Leading/trailing spaces and case are preserved as-is.

### Numeric
Strips currency symbols (`$`, `£`, `€`, `¥`), commas, and parenthesised negatives `(1,500.00)` before comparing. A configurable tolerance allows small rounding differences to be treated as equal.

### Date: `date_only`
Parses both ISO (`YYYY-MM-DD`) and US (`MM/DD/YYYY`) formats, then compares calendar date only. The time component is ignored; `2024-01-15 08:30` and `01/15/2024 23:59` are treated as equal.

### Date: `datetime_precision`
Same format parsing as `date_only`, but time is included in the comparison. `2024-01-15 08:30` and `2024-01-15 14:00` are treated as different.

---

## Data Preparation Workflow

Many government extracts include title rows, agency metadata, or footer totals above or below the actual data. The Data Preparation step (Step 3 in the 6-step workflow) lets you identify the correct header row and discard non-data rows before reconciliation begins.

### Example

A file downloaded from a case management system might look like this:

| Row | Content |
|---|---|
| 1 | Monthly Reconciliation Report |
| 2 | Exported: 2025-06-01 |
| 3 | (blank) |
| 4 | CaseID, Status, Amount |
| 5 | 1001, Open, 500 |
| 6 | 1002, Closed, 750 |
| 7 | Total, , 1250 |

The correct preparation settings for this file:

- **Header Row:** 4 (the row containing column names, displayed 1-based in the UI)
- **Drop rows above header:** Yes (removes rows 1-3 before processing)
- **Drop fully blank rows:** Yes (removes the blank row if any exist after the header)
- **End Row:** 6 (stops before the "Total" footer row; displayed 1-based)

After preparation, the dataset contains two records with columns CaseID, Status, and Amount. The preparation summary is written to the Analysis Metadata tab of the exported Excel workbook so the selection is auditable.

### Preparation Controls

| Control | Default | Description |
|---|---|---|
| Header Row | 1 | The 1-based row number containing column names |
| Drop rows above header | Yes | Removes all rows above the selected header row |
| Drop fully blank rows | Yes | Removes rows where every cell is empty |
| End Row | (none) | The last data row to include; rows after this row are excluded |

### Blank and Duplicate Header Handling

If a column header is blank, it is renamed `Unnamed_1`, `Unnamed_2`, etc. If the same header appears more than once, duplicates are renamed `ColumnName_2`, `ColumnName_3`, etc. Both are reported as warnings and counted in the preparation summary.

---

## Report Outputs

### Interactive Dashboard
- **10 KPI cards:** totals, baseline-only, comparison-only, matched, records with differences, duplicates, missing identifiers
- **3 Plotly charts:** Reconciliation Summary bar chart, Match Coverage donut, Field-Level Differences frequency
- **8 tabbed result tables:** each with a per-category CSV download

### Excel Workbook (11 tabs)

| Tab | Contents |
|---|---|
| Executive Summary | Auto-generated plain-English narrative with six labelled sections |
| Analysis Metadata | Dataset names, sheet tabs, record counts, match key fields, timestamp |
| Comparison Rules | Per-field comparison type, tolerance value, and date precision setting |
| Delta Counts | Flat count table suitable for pivot tables and downstream reporting |
| Baseline Only Records | Records present in the Baseline Dataset but absent from the Comparison Dataset |
| Comparison Only Records | Records present in the Comparison Dataset but absent from the Baseline Dataset |
| Matched Records | Side-by-side view of all matched records (prefixed Baseline:/Comparison:) |
| Records with Differences | Before/after values for every field-level difference |
| Baseline Duplicate Identifiers | Rows sharing a match key with another row in the Baseline Dataset |
| Comparison Duplicates | Rows sharing a match key with another row in the Comparison Dataset |
| Data Quality Flags | Rows excluded from reconciliation due to blank or null match key values |

---

## Sample Data

Two scenarios are included to demonstrate the tool without uploading real data.

### Minimal demo (`file_a.csv` / `file_b.csv`)

Exercises every delta category in a small, purpose-built dataset.

```
sample_data/file_a.csv  - 6 rows (1 matched-unchanged, 1 matched-changed,
                           1 baseline-only, 1 duplicate key pair, 1 blank key)
sample_data/file_b.csv  - 4 rows (corresponding comparison data)
```

Recommended settings:
- Match key: `RecordID` (both datasets)
- Comparison fields: `Status`, `Amount` (both datasets)

Expected results:

| Category | Count |
|---|---|
| Baseline Only Records | 1 (R003) |
| Comparison Only Records | 1 (R005) |
| Matched Records | 3 (R001, R002, R004) |
| Records with Differences | 1 (R002, Status and Amount changed) |
| Baseline Duplicate Identifiers | 2 (both R004 rows flagged) |
| Data Quality Flags | 1 (blank key row) |

### Full contracting dataset (`sample_a.csv` / `sample_b.csv`)

```
sample_data/sample_a.csv  - 11-row baseline (includes 1 duplicate, 1 blank key)
sample_data/sample_b.csv  - 10-row comparison (changed fields, new entries)
```

Recommended settings:
- Match key: `CaseID` (both datasets)
- Comparison fields: `ContractAmount`, `Status`, `ReviewedBy` (both datasets)

---

## Running the Tests

```bash
pytest tests/ -v
```

Expected: **253 tests passing** across seven test modules.

With coverage:

```bash
pytest tests/ --cov=src --cov-report=term-missing
```

---

## Project Structure

```
DeltaAnalysis/
├── app.py                          # Streamlit UI - main entry point
├── requirements.txt
├── README.md
├── .gitignore
├── .github/
│   └── workflows/
│       └── tests.yml               # CI: runs pytest on push and pull_request
├── src/
│   ├── __init__.py
│   ├── normalization.py            # Key cleaning: trim, fix "1234.0" artifacts, handle nulls
│   ├── io_utils.py                 # File I/O, raw read, data preparation, sheet names
│   ├── comparison.py               # Type-aware field comparison (numeric, date, text)
│   ├── delta_engine.py             # Core comparison engine → DeltaResult dataclass
│   └── reporting.py                # 11-tab Excel export with Executive Summary narrative
├── tests/
│   ├── __init__.py
│   ├── test_normalization.py       # 21 unit tests - key normalization
│   ├── test_delta_engine.py        # 26 unit tests - comparison categories and validation
│   ├── test_comparison.py          # 47 unit tests - numeric, date, and field-level comparison
│   ├── test_io_utils.py            # 24 unit tests - file parsing, sheet selection, size checks
│   ├── test_prepare_dataframe.py   # 35 unit tests - raw read, prepare_dataframe, report metadata
│   ├── test_data_preparation.py    # 41 unit tests - v1.2 data preparation workflow
│   └── test_e2e_validation.py      # 59 integration tests - full workbook and scenario coverage
└── sample_data/
    ├── file_a.csv                  # Minimal demo - one of each delta category
    ├── file_b.csv                  # Minimal demo - comparison side
    ├── sample_a.csv                # Full contracting dataset (11-row source)
    └── sample_b.csv                # Full contracting dataset (10-row comparison)
```

---

## Architecture Overview

**`DeltaResult` dataclass** is the central output container. `run_delta()` accepts two DataFrames and produces a fully populated `DeltaResult` in a single call (no intermediate state, no side effects). Every downstream consumer (the UI, the Excel exporter, the test suite) works exclusively from this object.

**All data is read as strings.** This prevents silent type coercion, a common source of false mismatches when Excel stores numeric IDs (e.g., `1001` read as `1001.0`). Numeric and date parsing is performed explicitly only when the user selects a typed comparison field.

**Composite key support.** Multiple match key columns are concatenated with a `||` separator into a single match key before comparison, preventing collisions between single-column values that happen to share parts.

**Blank keys are quarantined, not silently dropped.** Rows with null or empty key values are captured in the Data Quality Flags category so analysts see exactly what was excluded and why, a requirement for audit-traceable outputs.

**Duplicates use first-occurrence for matching.** All duplicate rows are surfaced in the Duplicate Identifiers tabs; only the first occurrence participates in matching, keeping totals predictable.

---

## Design Decisions

**Type-grouped Advanced Comparison Settings.** Rather than creating per-field widgets (which becomes unusable at 20+ fields), the UI groups fields by type: one multiselect for numeric fields, one for date fields, shared tolerance and date-mode inputs. This scales to arbitrarily wide datasets.

**`comparison_rules` backward compatibility.** When no rules are provided, the engine defaults to text comparison for all fields, identical behavior to v1.0. Existing integrations and tests continue to work without modification.

**Executive Summary is fully data-driven.** Every statistic in the narrative is interpolated directly from `DeltaResult` at export time. Analysts do not need to edit the narrative before sharing it with leadership.

---

## CI/CD

A GitHub Actions workflow (`.github/workflows/tests.yml`) runs the full test suite on every push and pull request. The build fails if any test fails. All seven test modules are included in the CI run.

---

## Current Limitations

- **Memory-bound processing.** Files larger than 500,000 rows may be slow or exhaust session memory. The application warns at 100,000 rows and requires explicit acknowledgment at 500,000 rows.
- **No authentication.** The Streamlit application has no login gate. Do not deploy to a public URL with sensitive government data unless appropriate security controls are in place.
- **No persistent configuration.** Match key selections and comparison field mappings must be re-entered each session.
- **Single-run comparison only.** The tool performs one comparison per session; scheduled or batch comparisons are not currently supported.

---

## Troubleshooting

| Symptom | Likely Cause | Resolution |
|---|---|---|
| `Columns not found in Baseline Dataset` | Selected column no longer matches after re-upload | Re-select match key columns after uploading a new file |
| `Key column counts must match` | Different number of key columns selected for each dataset | Select the same number of match key columns in both datasets |
| `Could not parse Excel file` | File is password-protected, corrupted, or unsupported format | Open the file in Excel, save as `.xlsx`, and re-upload |
| `The uploaded file contains no data rows` | File has a header row only | Confirm the file contains data rows below the header |
| No field-level differences shown | No comparison fields selected | Select non-key fields in Step 3 to enable field-level change detection |
| Workbook tab name truncated warning | Sheet name exceeds Excel's 31-character limit | This is an openpyxl warning only; the workbook opens and functions correctly |

---

## Roadmap

### Near-Term
- [ ] Saved configuration profiles (persist key/comparison column selections across sessions)
- [ ] Column rename/alias mapping (compare columns with different names without renaming source data)
- [ ] PDF export option for leadership briefings

### Medium-Term
- [ ] Row-level analyst annotations in the Excel export
- [ ] CLI entry point for batch or scheduled comparisons
- [ ] Configurable reconciliation modes for common government workflows (DFAS, M&RA, Veteran Case Tracking)

### Longer-Term
- [ ] Optional deployment guide for secure internal environments
- [ ] Database-backed processing for very large file comparisons
