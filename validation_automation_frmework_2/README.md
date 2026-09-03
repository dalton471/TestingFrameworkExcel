# Validation Automation Framework

## 1. Overview

The Validation Automation Framework is a Python tool that validates
the contents of a multi-sheet Excel workbook and produces a
structured Excel report describing what passed and what failed.

It checks two things about a workbook: that the sheets you expect are
actually present, and that the data inside each sheet's columns
matches one of seven supported validation rules (numeric values,
email addresses, phone numbers, monetary amounts, bounded numbers,
dates, and plain text). Every sheet and column is discovered directly
from the workbook at runtime - the framework does not require a
pre-defined list of column names to work.

The framework is configuration-driven: an optional JSON file can be
used to specify which sheets are expected and to tune how validation
rules are detected, while a built-in set of default rules lets the
framework run without any configuration file at all. Both modes use
the same validation engine and produce the same report format, making
the framework reusable across different workbooks without code
changes.

## 2. Key Features

- Multi-sheet Excel workbook validation
- Dynamic sheet discovery and sheet existence checking
- Dynamic column discovery - no manually maintained list of field names
- Content-based validation rule detection per column
- Seven supported validation rule types (see section 4)
- Optional JSON-based configuration for sheet expectations and rule tuning
- JSON-based validation selection, with automatic fallback to the framework's built-in validation set
- Built-in/default validation mode when no configuration file is supplied
- Automated three-sheet Excel validation report
- PASS/FAIL reporting at both the sheet level and the individual validation-rule level
- Cell-level failed-record detail, including sheet, column, row, value, and error message
- Command-line execution using `argparse`
- Structured, readable terminal output during execution

## 3. Validation Workflow

```
Input Excel Workbook
        ↓
Load Configuration (JSON if present, otherwise built-in defaults)
        ↓
Load Workbook and Discover Sheets
        ↓
Sheet Existence Validation
        ↓
Dynamic Column Discovery (per present sheet)
        ↓
Validation Rule Determination (content-based, per column)
        ↓
Data Validation (cell-by-cell, using the determined rule)
        ↓
Collect Validation Results
        ↓
Generate Validation Report
```

**Load Configuration** - If a JSON configuration file is found at the
path given by `--config`, it is loaded and used. If it is not found,
the framework builds an equivalent configuration internally from
built-in defaults, using the sheets it discovers in the workbook.

**Load Workbook and Discover Sheets** - The workbook is opened once
and every sheet name in it is read.

**Sheet Existence Validation** - Every sheet named in the
configuration's `sheet_checks` is checked against the sheets actually
found in the workbook. Each check reports PASS or FAIL.

**Dynamic Column Discovery** - For every sheet that is present, every
column (the header row) is read directly from the workbook. No
column name needs to be declared anywhere in advance.

**Validation Rule Determination** - For each discovered column, a
sample of its non-blank values is checked against each of the seven
supported validation rules. The first rule whose values match at or
above a configurable confidence threshold is assigned to that column.
If no rule matches confidently enough, the column is not assigned a
rule and is excluded from data validation.

**Data Validation** - Every non-blank cell in a column that was
assigned a rule is checked against that rule. Cells that fail are
recorded with their sheet, column, row number, value, and an error
message.

**Collect Validation Results** - All results (sheet checks, rule
determinations, and data validation outcomes) are gathered into a
single result set.

**Generate Validation Report** - The results are written to a
three-sheet Excel workbook (see section 9).

## 4. Supported Validations

| Validation | Description | Parameters |
|------------|-------------|------------|
| `numeric_only` | Value must consist only of digits (0-9) | None |
| `email` | Value must match a standard email address pattern | None |
| `phone_number` | Value must match a common US phone number format (e.g. `123-456-7890`, `(123) 456-7890`) | None |
| `amount` | Value must be a real number that is not negative | None (fixed minimum of 0) |
| `number` | Value must be a real number between 0 and 100 | None (fixed range of 0-100) |
| `date` | Value must parse under at least one of several common date formats (e.g. `YYYY-MM-DD`, `DD/MM/YYYY`, `MM/DD/YYYY`, `Month D, YYYY`) | None |
| `text_only` | Value must contain only alphabetic characters and spaces | None |

None of the seven validations take parameters from the JSON
configuration. Their acceptable conditions (ranges, formats, and
patterns) are fixed inside the Python validation engine
(`validations/validation.py`), so the same rule behaves identically
regardless of which workbook or configuration is used.

Blank or empty cells are skipped by every validation rule and are
never reported as failures.

## 5. Configuration

`config/validation_config.json` is entirely optional. When present,
it controls three things:

- **`sheet_checks`** - which sheet names are expected in the
  workbook, and whether each one is required.
- **`rule_detection`** - tuning for how confidently a column's
  content must match a rule before that rule is assigned to it
  (`match_threshold`) and how many non-blank values are sampled per
  column when making that determination (`sample_size`).
- **`validations`** - an optional list naming which of the seven
  supported validations to explicitly select. Any validation not
  named in this list is automatically added by the framework, so
  every supported validation always runs; this list only affects how
  the run is reported (which validations were explicitly selected
  versus picked up automatically).

The JSON file never contains regular expressions, numeric ranges,
date formats, or any other validation logic - only which sheets and
validations to select, and how to tune the detection mechanism. The
actual validation logic lives in `validations/validation.py`.

Example configuration:

```json
{
  "workbook": {
    "name": "Validation_Sample.xlsx"
  },
  "sheet_checks": {
    "Employee Details": { "required": true }
  },
  "rule_detection": {
    "match_threshold": 0.6,
    "sample_size": 200
  },
  "validations": ["numeric_only", "email", "amount", "date", "text_only"]
}
```

## 6. Default Validation Mode

If `--config` is not supplied, or the path it points to does not
exist, the framework uses its built-in default configuration instead
of stopping:

```bash
python main.py --input "input/Validation_Sample.xlsx"
```

In this mode, the set of "expected sheets" is built directly from the
sheets found in the input workbook, rule-detection tuning uses the
framework's built-in defaults, and no validations are pre-selected -
so every one of the seven supported validations is applied. The
result is the same three-sheet report described in section 9,
produced by the same validation engine used when a configuration file
is supplied.

## 7. Project Structure

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

- **`config/validation_config.json`** - optional configuration file
  (see section 5).
- **`input/`** - holds the workbook(s) to be validated.
- **`output/`** - the generated `Validation_Report.xlsx` is written here.
- **`validations/validation.py`** - the validation engine: the seven
  validation functions, rule-determination logic, sheet/column
  discovery, and the orchestration that runs all three validation
  stages.
- **`utils/helper.py`** - loads the JSON configuration file.
- **`utils/excel_reader.py`** - opens the input workbook and reads
  its sheets into data frames.
- **`utils/report_generator.py`** - builds the three-sheet
  `Validation_Report.xlsx` from the validation results.
- **`main.py`** - the command-line entry point that ties the above
  together and prints terminal output.

## 8. Input

The framework accepts a single Excel workbook file as input, supplied
via the `--input` command-line argument. Supported file extensions
are `.xlsx` and `.xls`. The workbook may contain any number of
sheets; every sheet found is read, and sheets named in the
configuration's `sheet_checks` (or discovered automatically in
default mode) are checked for existence.

The framework validates one workbook file per run - it does not
accept a folder of files as input.

## 9. Output

Running the framework produces `output/Validation_Report.xlsx` (or
the path given via `--output`), containing exactly three sheets:

- **Validation Summary** - one row per check performed: sheet
  existence checks, the validation rule determined for each column,
  and the data validation outcome for each column that received a
  rule. Each row shows the sheet name, field/column, test type, test
  name, PASS/FAIL status, and failed cell count.
- **Failed Records** - one row per individual cell that failed
  validation, showing the sheet name, column, row number, the actual
  value, the validation type applied, the required datatype, and a
  descriptive error message.
- **Execution Summary** - run-level totals, including sheet and
  column counts, how many validations were selected from JSON versus
  automatically added, rows processed, validations executed,
  pass/fail counts, total failed cells, and an overall execution
  status.

## 10. Installation

Install the required dependencies with:

```bash
pip install -r requirements.txt
```

The framework depends on `pandas` (>=2.0.0) and `openpyxl` (>=3.1.0),
and requires a Python 3 environment compatible with those versions.

## 11. Execution

Both execution modes use the same `main.py` entry point and the same
underlying validation engine.

### With JSON configuration

```bash
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"
```

### Without JSON configuration (built-in defaults)

```bash
python main.py --input "input/Validation_Sample.xlsx"
```

### Optional output path

Both commands accept an optional `--output` argument to control where
the report is written:

```bash
python main.py --input "input/Validation_Sample.xlsx" --output output/MyReport.xlsx
```

If `--output` is omitted, the report is written to
`output/Validation_Report.xlsx`.

