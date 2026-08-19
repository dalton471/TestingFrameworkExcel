# Validation Automation Framework

## 1. Project Overview

A generic, configuration-driven Python framework that validates a
multi-sheet Excel workbook against a JSON-defined rule set and
produces a structured, four-sheet `Validation_Report.xlsx`.

The framework is built around the same separation of concerns as the
previous KPI Automation Framework:

- **Python** defines *how* a validation works (the validation engine
  and logic).
- **JSON** defines *what* needs to be validated (sheets, fields,
  rules, thresholds, date formats).

No sheet name or column name is hardcoded anywhere in the Python
code. Everything the engine checks - which sheets must exist, which
fields belong to which sheet, and which of Dalton's seven validation
types applies to each field - comes from `config/validation_config.json`.

## 2. Architecture

1. Parse command-line arguments (`--config`, `--input`, `--output`).
2. Load and validate the JSON configuration file.
3. Load the input workbook and read every sheet into a DataFrame.
4. Run three validation layers, in order, all driven by JSON:
   - **Sheet Existence Testing** - is each configured sheet present?
   - **Field Existence Testing** - does each configured column exist
     on its sheet?
   - **Data Validation** - Dalton's seven field-level checks, applied
     cell by cell to every configured field.
5. Collect every test result (pass and fail) as a flat list of
   summary rows, plus a separate list of individual cell-level
   failures.
6. Generate `Validation_Report.xlsx` via `report_generator.py`.
7. Print structured, readable results to the terminal as the run
   progresses, followed by a final run summary.
8. Exit with a status code reflecting whether any configured sheet
   was missing (data-quality FAILs in the input do not themselves
   fail the framework run - see section 19).

## 3. Folder Structure

```
Validation_Automation_Framework/
│
├── config/
│   └── validation_config.json
│
├── input/
│   └── Validation_Sample.xlsx
│
├── output/
│   └── Validation_Report.xlsx
│
├── validations/
│   └── validation.py
│
├── utils/
│   ├── helper.py
│   ├── excel_reader.py
│   └── report_generator.py
│
├── main.py
│
├── requirements.txt
│
└── README.md
```

## 4. Supported Validations

| Validation Type | What it checks                                       | JSON parameters            |
|------------------|--------------------------------------------------------|------------------------------|
| `numeric_only`   | Value contains only digits                              | none                          |
| `email`          | Value matches a basic email address pattern              | none                          |
| `phone_number`   | Value matches a common US phone number format             | none                          |
| `amount`         | Value is numeric and within `[min_value, max_value]`      | `min_value`, `max_value`      |
| `number`         | Value is numeric and within `[min_value, max_value]`      | `min_value`, `max_value`      |
| `date`           | Value parses under the configured `format`                | `format` (strptime pattern)   |
| `text_only`      | Value contains only letters and spaces                     | none                           |

**API validation is intentionally NOT implemented in this version.**
Dalton confirmed it is out of scope. No API URL, token,
`fetch_valid_codes()`, or API call exists anywhere in this codebase.

Blank/null cells are skipped for every validation type, matching
Dalton's original sample logic - they are excluded from both the
checked count and the failed count.

## 5. JSON Configuration

`config/validation_config.json` has four sections:

- `workbook`: informational metadata (expected workbook file name).
- `sheet_checks`: every sheet the framework expects, and whether it
  is `required`.
- `column_checks`: rule parameters per validation type (ranges, date
  format). Types with no parameters (`numeric_only`, `email`,
  `phone_number`, `text_only`) use an empty object `{}`.
- `fields`: for each sheet, maps each column name to the list of
  validation types that apply to it.

```json
{
  "sheet_checks": {
    "Employee Details": { "required": true }
  },
  "column_checks": {
    "amount": { "min_value": 20000, "max_value": 200000 },
    "date": { "format": "%Y-%m-%d" }
  },
  "fields": {
    "Employee Details": {
      "Joining_Date": ["date"]
    },
    "Compensation": {
      "Salary": ["amount"]
    }
  }
}
```

The demo sheet names, field names, ranges (`Salary`: 20,000-200,000;
`Performance_Score`: 0-100), and date format (`%Y-%m-%d`) are
illustrative assumptions made for this sample - not business
requirements. Edit the JSON to match a real workbook without touching
any Python code.

## 6. Input Workbook Structure

`input/Validation_Sample.xlsx` contains five sheets:

| Sheet             | Columns                                                    | Validations applied |
|--------------------|-------------------------------------------------------------|-----------------------|
| Read Me            | (descriptive text only - not data-validated)                 | Sheet existence only |
| Employee Details   | Employee_ID, Employee_Name, Department, Joining_Date          | numeric_only, text_only, text_only, date |
| Contact Details    | Employee_ID, Employee_Email, Employee_Phone                    | numeric_only, email, phone_number |
| Compensation       | Employee_ID, Salary, Performance_Score                         | numeric_only, amount, number |
| Manager Details    | Employee_ID, Manager_Name, Manager_Email                        | numeric_only, text_only, email |

Each data sheet has 10 rows, hand-built so that every configured
field has at least one deliberate, explainable failure, one field
has a deliberately blank cell to prove skip behavior, and every
sheet/field passes existence checks (so the framework's "everything
passed" and "data failed" paths are both visible in the same run).
See section 17 for the full expected-vs-actual failure table.

## 7. Validation Processing

For each sheet listed under `fields` in the JSON:

1. If the sheet itself is missing, it is reported once at the
   Sheet Existence layer and skipped entirely at the field and data
   layers (no false passes are ever recorded for a sheet that isn't
   there).
2. For each configured column: if the column is missing from that
   sheet, it is reported at the Field Existence layer and skipped at
   the data layer.
3. For each column that does exist, every configured validation type
   runs against every non-blank cell in that column, using Dalton's
   original validator functions.

## 8. Output Report Structure

`output/Validation_Report.xlsx` has four sheets:

- **Validation Summary** - every test that ran (sheet existence,
  field existence, and data validation), pass and fail alike, with
  columns: `Sheet Name`, `Field`, `Test Type`, `Test Name`, `Status`,
  `Failed Count`.
- **Validation Details** - the same tests with more depth: the rule
  parameters applied (from JSON), and `Total Checked` /
  `Passed Count` / `Failed Count` per test.
- **Failed Records** - one row per individual failing cell:
  `Sheet Name`, `Field`, `Row`, `Value`, `Validation Type`,
  `Required Datatype`, `Error Message`.
- **Execution Summary** - run-level totals: total sheets,
  sheets passed/failed, total rows processed, total validations
  executed, total passed/failed, total failed cells, and overall
  execution status.

Each report sheet uses a frozen header row, an Excel Table (built-in
column filters), auto-sized columns, and PASS/FAIL/SKIPPED row
shading for quick visual scanning.

## 9. Installation

```bash
pip install -r requirements.txt
```

## 10. Dependencies

- Python 3.8+
- pandas >= 2.0.0
- openpyxl >= 3.1.0

## 11. Execution Command

```bash
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"
```

Optional custom output path:

```bash
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx" --output output/MyReport.xlsx
```

## 12. Example Output

```
============================================================
VALIDATION FRAMEWORK
============================================================

JSON Loaded Successfully

============================================================
PROCESSING: input/Validation_Sample.xlsx
============================================================
Excel Loaded Successfully
Sheets found in workbook: Read Me, Employee Details, Contact Details, Compensation, Manager Details

============================================================
VALIDATION RESULTS
============================================================
Read Me | Sheet | Sheet Existence Testing | PASS | Failed Count: 0
Employee Details | Sheet | Sheet Existence Testing | PASS | Failed Count: 0
...
Employee Details | Employee_ID | numeric_only | FAIL | Failed Count: 1
...
============================================================
VALIDATION COMPLETED
============================================================

output/Validation_Report.xlsx created successfully.

============================================================
RUN SUMMARY
============================================================
Total Sheets: 5
Sheets Passed: 5
Sheets Failed: 0
Total Rows Processed: 40
Total Validations Executed: 31
Total Passed: 18
Total Failed: 13
Total Failed Cells: 19
Execution Status: SUCCESS
============================================================
```

This is the actual terminal output from a real run against the
included sample workbook (see section 17 for verification).

## 13. How to Modify Validations

Add or remove a validation type from a field's list under `fields` in
the JSON. Validation types are matched against the keys already
defined under `column_checks` - Python code never needs to change.

## 14. How to Add Another Sheet

1. Add the sheet name under `sheet_checks` with `"required": true/false`.
2. Add a matching entry under `fields` mapping its column names to
   validation types.

## 15. How to Add Another Field

Add a new key under the relevant sheet's block in `fields`, pointing
the column name at one or more validation types already defined
under `column_checks`.

## 16. How to Change the Amount Range

Edit `min_value` / `max_value` under `column_checks.amount` in the
JSON.

## 17. How to Change the Number Range

Edit `min_value` / `max_value` under `column_checks.number` in the
JSON.

## 18. How to Change the Date Format

Edit `format` under `column_checks.date` in the JSON, using a Python
`strptime`-compatible pattern (e.g. `%d-%m-%Y`).

## 19. Error Handling

The framework is designed to handle the following without crashing:

- Missing configuration file or invalid path
- Missing input file or invalid path
- Malformed/empty JSON configuration
- Unsupported input file extension
- Corrupt or unreadable workbook
- A configured sheet missing from the workbook (reported as a FAIL at
  the Sheet Existence layer; field and data checks for that sheet are
  skipped rather than silently passed)
- A configured column missing from an existing sheet (reported as a
  FAIL at the Field Existence layer; data validation for that column
  is skipped rather than silently passed)
- Unexpected exceptions during validation or report generation (full
  traceback printed, framework exits with a non-zero status)

**Important distinction:** the framework's own exit status reflects
whether it *ran successfully* (config/input loaded, all layers
executed, report generated) - it exits non-zero only if a required
sheet is missing. Data-quality FAILs found *within* valid sheets
(e.g. an invalid email) are expected findings the framework is
designed to surface, not framework defects, so they do not by
themselves cause a non-zero exit. This mirrors how the KPI framework
distinguishes "the pipeline ran" from "the data passed every rule."

## 20. Testing Strategy

The sample workbook was hand-built so every expected failure count
could be calculated manually before running the framework (see
section 17 below). The framework was then executed end-to-end against
that workbook, and the generated report's counts were compared
row-by-row against the manual calculation. All counts matched on the
first run; no expected values, code paths, or test data were adjusted
to force a match.

## 21. API Validation Limitation

API-based validation (cross-checking a value such as an email or code
against an external system of record) is intentionally excluded from
this version, per Dalton's explicit instruction that it is out of
scope. There is no API URL, token, `fetch_valid_codes()` function, or
network call anywhere in this codebase. If API-backed validation is
needed in a future version, it should be added as an additional,
separately-configured validation type rather than folded into the
existing seven.

## 22. Expected vs Actual Failure Counts (Verified)

Calculated manually from the deliberately invalid sample data, then
confirmed against the framework's actual output:

| Sheet             | Field              | Validation    | Expected | Actual |
|--------------------|---------------------|----------------|----------|--------|
| Employee Details   | Employee_ID         | numeric_only   | 1        | 1      |
| Employee Details   | Employee_Name       | text_only      | 1        | 1      |
| Employee Details   | Department          | text_only      | 1        | 1      |
| Employee Details   | Joining_Date        | date           | 1        | 1      |
| Contact Details    | Employee_ID         | numeric_only   | 1        | 1      |
| Contact Details    | Employee_Email      | email          | 2        | 2      |
| Contact Details    | Employee_Phone      | phone_number   | 2        | 2      |
| Compensation       | Employee_ID         | numeric_only   | 1        | 1      |
| Compensation       | Salary              | amount         | 2        | 2      |
| Compensation       | Performance_Score   | number         | 2        | 2      |
| Manager Details    | Employee_ID         | numeric_only   | 1        | 1      |
| Manager Details    | Manager_Name        | text_only      | 2        | 2      |
| Manager Details    | Manager_Email       | email          | 2        | 2      |

**Total expected failed cells = Total actual failed cells = 19.**

Sheet existence: 5/5 configured sheets present (0 failures).
Field existence: 13/13 configured fields present across all sheets
(0 failures). Total validations executed: 31 (5 sheet + 13 field +
13 data). Total passed: 18. Total failed: 13. All numbers above were
produced by the actual `main.py` run against the included sample
workbook - see section 12 for the real terminal output.

## 23. Known Limitations

- `phone_number` only supports common US-style formats
  (`123-456-7890`, `(123) 456-7890`, `123 456 7890`).
- `email` uses a basic regex, not full RFC 5322 validation.
- The `Read Me` sheet is intentionally excluded from `fields` (and
  therefore from data validation and row-processing totals) since it
  contains descriptive text, not tabular data to validate.
- API-based validation is out of scope for this version (see section 21).
- The framework validates against the rules defined in the JSON
  configuration; it does not infer validation rules from the data
  itself.
