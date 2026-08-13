# KPI Validation Automation Framework

## Overview

The KPI Validation Automation Framework is a generic, configuration-driven Python framework that validates KPI data (Excel or CSV) against a JSON-defined rule set and produces a consolidated Excel validation report.

The framework is built around a clear separation of concerns:

- **Python** defines *how* a validation works (the validation engine and logic).
- **JSON** defines *what* needs to be validated (sheets, fields, rules, thresholds).

This separation allows new validation rules to be configured without modifying the underlying Python code, making the framework reusable across different KPI datasets and reporting structures.

## Key Capabilities

- Validates Excel (`.xlsx`, `.xls`) and CSV input files against a JSON rule configuration.
- Supports single files, folders of files, and split-workbook folders (where each Excel file represents one worksheet of a logical workbook).
- Performs ten categories of validation, including data quality, business rules, reconciliation, cross-sheet, and trend checks.
- Applies Excel-aware numeric precision validation using cell number formatting, avoiding false failures caused by floating-point representation.
- Generates a structured Excel validation report summarizing pass/fail status and failure counts per check.
- Handles missing/invalid configuration, missing sheets/columns, and corrupt files gracefully without aborting the full run.
- Processes folders in batch mode, where a failure in one file does not stop processing of the others.

## Architecture / Flow

1. Parse command-line arguments (`--config`, `--input`, `--output`).
2. Load and validate the JSON configuration file.
3. Resolve the input path:
   - Single Excel/CSV file
   - Folder of supported files
   - Split-workbook folder (multiple Excel files representing individual sheets)
4. Normalize the input into a workbook the validation engine can operate on (CSV and split-workbook inputs are converted into a temporary Excel workbook internally).
5. Run each configured validation category against the workbook using `validations.py`.
6. Collect results (status, failed record counts) for every configured check.
7. Generate `Validation_Report.xlsx` via `report_generator.py`.
8. Clean up any temporary workbooks created during the run.
9. Print a run summary to the terminal.

## Project Structure

```
KPI_Automation_Framework/
├── config/
│   └── Kpi_automation_phase1.json
├── input/
├── output/
├── validations/
│   └── validations.py
├── utils/
│   ├── excel_reader.py
│   ├── report_generator.py
│   └── helper.py
├── main.py
└── README.md
```

## Installation

**Dependencies:**

```bash
pip install pandas openpyxl
```

## Usage

**Single Excel file:**

```bash
python main.py --config config/Kpi_automation_phase1.json --input input/KPI_Sample.xlsx
```

**CSV file:**

```bash
python main.py --config config/Kpi_automation_phase1.json --input input/data.csv
```

**Folder of files:**

```bash
python main.py --config config/Kpi_automation_phase1.json --input input/
```

**Split-workbook folder:**

```bash
python main.py --config config/Kpi_automation_phase1.json --input "input/KPI_Sample (2)"
```

**Custom output path:**

```bash
python main.py --config config/Kpi_automation_phase1.json --input input/KPI_Sample.xlsx --output output/MyReport.xlsx
```

**Help:**

```bash
python main.py --help
```

## Supported Input Formats

- `.xlsx`
- `.xls`
- `.csv`
- A folder containing any combination of the above supported files
- A split-workbook folder, where individual Excel files each represent one worksheet of a logical workbook

### CSV Behavior

CSV input is converted internally into a temporary single-sheet Excel workbook named `Sheet1` so it can be processed by the same validation engine used for Excel files. Since CSV files have no concept of multiple worksheets, any configured rule that expects a different sheet name will report that sheet as missing. This is expected behavior for CSV input, not a defect.

## Split Workbook Processing

The framework can detect when the input path is a folder containing multiple Excel files, where each file represents one logical worksheet rather than a standalone workbook.

Example folder layout:

```
input/KPI_Sample (2)/
├── Aging Recon.xlsx
├── Data Currency.xlsx
├── Data Pipeline Status.xlsx
├── FIN Recon.xlsx
├── KPI.xlsx
├── Missing Practices Summary.xlsx
├── Practice Detail.xlsx
└── Read Me.xlsx
```

Processing steps:

1. Detect the split-workbook folder.
2. Locate the individual Excel files.
3. Read the first (or only) sheet from each file.
4. Combine the sheets into a temporary multi-sheet Excel workbook, using each file's name as the corresponding worksheet name.
5. Run the standard validation engine against the temporary combined workbook.
6. Generate `Validation_Report.xlsx`.
7. Close the temporary workbook.
8. Remove the temporary workbook from disk.

This workflow was tested successfully against the `KPI_Sample (2)` dataset, with all eight files correctly combined into worksheets: Aging Recon, Data Currency, Data Pipeline Status, FIN Recon, KPI, Missing Practices Summary, Practice Detail, and Read Me.

The temporary combined workbook is created under `output/` and is only used for the duration of the run. The workbook is explicitly closed before the cleanup step, which avoids file-lock issues (such as `WinError 32`) on Windows.

## Validation Types

The framework currently supports the following validation categories:

1. Sheet existence validation
2. Field/column existence validation
3. Duplicate check
4. Null check
5. Formula validation
6. Numeric precision validation
7. Business rule validation
8. Reconciliation validation
9. Cross-sheet validation
10. Trend validation

### Business Rule Types

- `comparison`
- `staledata`
- `lookupmatch`

### Cross-Sheet Validation Types

- `valuecomparison`
- `existencecheck`

## Configuration Architecture

Validation behavior is driven entirely by the JSON configuration file. The configuration is organized into the following sections:

- `sheetlist`
- `sheetvalidation`
- `dataqualityvalidation.duplicatecheck`
- `dataqualityvalidation.nullcheck`
- `dataqualityvalidation.numericprecision`
- `formulavalidation`
- `businessrulevalidation`
- `reconciliationvalidation`
- `crosssheetvalidation`
- `trendvalidation`

Each section maps to a corresponding validator in `validations.py`. Adding or adjusting validation coverage typically only requires updating the JSON configuration, not the codebase.

## Numeric Precision Handling

The original numeric precision check compared raw floating-point values, which produced false failures when Excel stored more decimal digits internally than it displayed.

Example:

| Raw stored value | Excel displayed value |
|---|---|
| 133062.020162 | 133062.02 |

To resolve this, the framework now inspects the Excel cell's `number_format` using `openpyxl`:

- If the cell has an explicit numeric format (e.g. `#,##0.00`), the displayed decimal precision is used for comparison.
- If the cell has `General` or no explicit formatting, the framework falls back to raw-value validation with a tolerance to absorb normal floating-point noise.

This behavior can be overridden per configuration using:

```json
"usecellformat": false
```

which forces raw-value/tolerance-based validation regardless of cell formatting.

## Notable Fixes

- **Lookup-based business rule validation:** The lookup-side field was previously validated against the source sheet instead of the lookup sheet (e.g. `KPI.PracticeStatus` should be checked against `Practice Detail.ProblemReason`, not against the KPI sheet itself). This has been corrected so lookup-side fields are validated against the correct lookup dataframe.
- **Cross-sheet value comparison over-counting:** Because the KPI sheet can contain multiple rows per `PracticeCode` across fiscal periods, the same logical mismatch could previously be counted more than once. Results are now deduplicated on the configured match key, with appropriate handling of null keys.
- **Trend validation operator handling:** Per-row trend validation now consistently applies the operator defined in the configuration, rather than defaulting to a fixed comparison.
- **Reconciliation floating-point comparison:** Exact float equality has been replaced with a configurable numeric tolerance (default: `0.01`).

## Error Handling

The framework is designed to handle the following conditions without crashing:

- Missing configuration file
- Missing input path
- Invalid configuration path
- Unsupported file extension
- Invalid or malformed JSON
- Empty JSON configuration
- Missing configuration sections
- Missing worksheets
- Missing columns
- Validator-level errors
- Corrupt or unreadable input files

If a configuration section required by a particular validator is missing, that validator is skipped and a Configuration Validation failure is recorded rather than aborting the entire run. In folder/batch mode, an issue with one file does not stop the remaining files from being processed.

## Output / Report

By default, the framework writes its report to:

```
output/Validation_Report.xlsx
```

The report includes, for each validation check:

- Sheet Name
- Field
- Test Type
- Test Name
- Status (P/F)
- Failed Count

The terminal output also prints validation results and a run summary as the framework executes.

## Example Execution

**Command:**

```bash
python main.py --config config/Kpi_automation_phase1.json --input "input/KPI_Sample (2)"
```

**Result:** The split-workbook folder was detected and the following sheets were combined successfully: Aging Recon, Data Currency, Data Pipeline Status, FIN Recon, KPI, Missing Practices Summary, Practice Detail, Read Me.

**Output file:**

```
output\Validation_Report.xlsx
```

**Run summary:**

```
Total files: 1
Succeeded: 1
Failed: 0
```

**Sample validation results from this run:**

Sheet and field existence checks passed for all configured sheets/fields.

*Numeric precision:*

| Field | Status | Failed Count |
|---|---|---|
| AvgDailyChargesByPractice | FAIL | 37 |
| AvgDailyChargesByProvider | FAIL | 31 |
| AvgDailyChargesByLocation | FAIL | 34 |

*Business rules:*

| Rule | Status | Failed Count |
|---|---|---|
| Latest Transaction Date should be less than Latest Loaded Date | FAIL | 13 |
| Latest Transaction Date older than 3 months should display No LatestTransactionDate or NULL | PASS | 0 |
| Practice Status of KPI sheet should match with Problem Reason of Practice Detail | FAIL | 2 |

*Reconciliation:*

| Field | Status | Failed Count |
|---|---|---|
| Current Charges | FAIL | 3 |
| Current Payments | FAIL | 2 |
| Current Adjustments | FAIL | 3 |
| Gross AR | FAIL | 4 |

*Cross-sheet:*

| Rule | Status | Failed Count |
|---|---|---|
| Practice Status Validation | FAIL | 2 |
| Practice Availability Validation | FAIL | 2 |

*Trend:*

| Field | Status | Failed Count |
|---|---|---|
| Current Charges | FAIL | 18 |
| Current Payments | FAIL | 19 |
| Current Adjustments | FAIL | 22 |
| Current Encounters | FAIL | 12 |
| Gross AR | FAIL | 4 |

**Important:** A successful framework execution means the validation pipeline ran to completion and a report was generated. The individual P/F results shown above reflect whether the *input data* passed the configured validation rules — they are data quality findings in the sample dataset, not defects in the framework itself. The framework is expected to surface these kinds of discrepancies when they exist in the source data.

## Testing / Verification

The framework has been verified end-to-end using the `KPI_Sample (2)` split-workbook dataset described above, covering:

- Split-workbook detection, combination, and cleanup
- All ten validation categories executing against real KPI data
- Correct temporary workbook closure and removal (no residual file locks)
- Batch/folder execution reporting a run summary (Total / Succeeded / Failed)

## Known Limitations

- CSV input has no native worksheet concept; rules referencing sheet names other than `Sheet1` will report those sheets as missing when validating a CSV file.
- Numeric precision validation using cell formatting depends on the source Excel file having explicit number formats applied; unformatted (`General`) cells fall back to tolerance-based comparison.
- The framework validates against the rules defined in the JSON configuration; it does not infer validation rules from the data itself.

## Summary

The KPI Validation Automation Framework provides a reusable, configuration-driven approach to validating KPI Excel/CSV data. By keeping validation logic in Python and validation criteria in JSON, the framework can be extended to new datasets and rule sets without code changes, while consistently producing a structured, auditable validation report.
