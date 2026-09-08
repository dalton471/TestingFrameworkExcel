# Validation Automation Framework

A generic, configuration-driven framework for validating the structure
and content of Excel and CSV datasets, producing professional Excel and
JSON reports.

## Overview

The Validation Automation Framework inspects an input file, dynamically
discovers its sheets and columns, applies a configurable set of
validation rules, and produces a structured validation report. It
requires no prior knowledge of a dataset's schema: sheets and columns
are discovered at runtime, and the rule that applies to each column is
determined from the column's own content.

The framework is designed to be reused across different datasets
without modifying any Python code - a new workbook with entirely
different sheets and columns can be validated immediately.

## Features

- Dynamic sheet and column discovery - no maintained schema or field list
- Content-based rule detection for data-quality columns
- A central rule registry resolving rules by name, with no core
  engine changes required to add a new rule
- Structural, data-quality, integrity, and relational (cross-field /
  cross-sheet) validation rule categories
- A single, common validation engine shared by every configuration mode
- Optional JSON configuration with an automatic built-in default rule
  set when no configuration is supplied
- Standardized validation results shared by every rule and every report
- A dedicated result collector producing run-level statistics
- Excel input (`.xlsx`, `.xls`) and CSV input (`.csv`)
- A professional, formatted Excel report and a machine-readable JSON
  report generated from the same result data
- Structured logging to console and file
- Clear, distinct exception types for configuration, input, and
  execution errors
- A command-line interface built with `argparse`
- An automated unit and integration test suite

## Architecture

```
Excel / CSV Input
        │
   Input Reader (format-specific, behind a common interface)
        │
   Normalized Sheet Data
        │
   Configuration (JSON file or built-in defaults)
        │
   Rule Selection (Rule Registry)
        │
   Validation Engine
        │
   Result Collector
        │
   Reporting (Excel + JSON)
```

The package is organized under `src/validation_framework/`:

| Module | Responsibility |
|---|---|
| `core` | The validation engine, rule registry, and application runner |
| `config` | Configuration models, loading, defaults, and validation |
| `readers` | File format abstraction (Excel, CSV) |
| `rules` | Rule base class and every rule implementation, by category |
| `results` | Standardized result model and the result collector |
| `reporting` | Excel and JSON report generation |
| `logging` | Logging configuration |
| `exceptions` | The framework's exception hierarchy |
| `utils` | Discovery helpers shared across the engine |

## Validation Workflow

1. **Input detection** - the input file's extension selects the
   appropriate reader (Excel or CSV) through a factory.
2. **Configuration resolution** - a JSON configuration file is loaded
   if present; otherwise, the framework's built-in default
   configuration is used. Both are validated before execution.
3. **Sheet discovery and validation** - every sheet in the input file
   is discovered; configured sheet expectations are checked for
   existence, and any undeclared sheets are reported informationally.
4. **Column discovery and rule determination** - every column in every
   sheet is discovered from its header row; a sample of each column's
   values is checked against the enabled data-quality rules, and the
   first rule whose match rate clears a configurable confidence
   threshold is assigned to that column.
5. **Data validation** - every non-blank cell in a column with an
   assigned rule is validated; any explicitly targeted rules
   (required, length, regex, uniqueness, cross-field, cross-sheet) are
   also executed according to their configuration.
6. **Result collection** - every rule's outcome is collected into a
   standardized result set, and a summary is derived showing PASS/FAIL
   status and failure counts per check.
7. **Reporting** - the collected results are written to a formatted
   Excel workbook and a machine-readable JSON file.

## Supported Validation Rules

| Rule | Category | Description | Parameters |
|---|---|---|---|
| `sheet_existence` | Structural | Confirms each configured sheet is present | `sheet_checks` |
| `unexpected_sheet` | Structural | Reports sheets present but not configured | `sheet_checks` |
| `column_existence` | Structural | Confirms configured columns exist in a sheet | `expected_columns` |
| `numeric_only` | Data Quality | Value contains only digits | none |
| `text_only` | Data Quality | Value contains only letters and spaces | none |
| `email` | Data Quality | Value matches a standard email address pattern | none |
| `phone` | Data Quality | Value matches a common phone number format | none |
| `date` | Data Quality | Value parses under a common date format | `formats` |
| `number` | Data Quality | Value is numeric within a range | `min_value`, `max_value` |
| `amount` | Data Quality | Value is a non-negative numeric amount | `min_value`, `max_value` |
| `required` | Data Quality | Value must not be blank | `columns` |
| `length` | Data Quality | String length within bounds | `columns`, `min_length`, `max_length` |
| `regex` | Data Quality | Value matches a configured pattern | `columns`, `pattern` |
| `data_type` | Data Quality | Value can be cast to a target type | `columns`, `expected_type` |
| `duplicate_rows` | Integrity | Flags exact duplicate rows within a sheet | `columns` (optional subset) |
| `unique_field` | Integrity | Flags repeated values within a column | `columns` |
| `cross_field` | Relational | Compares two columns of a sheet with an operator | `sheet`, `left_column`, `right_column`, `operator` |
| `cross_sheet` | Relational | Confirms a column's values exist in another sheet's column | `source_sheet`, `source_column`, `target_sheet`, `target_column` |

Blank and null cells are excluded from every content-based rule.

## Configuration

Configuration is supplied as a JSON file with three sections:

```json
{
  "sheet_checks": {
    "SheetName": { "required": true }
  },
  "rule_detection": {
    "match_threshold": 0.6,
    "sample_size": 200
  },
  "rules": [
    { "name": "email", "enabled": true, "parameters": {} },
    { "name": "amount", "enabled": true, "parameters": { "min_value": 0 } }
  ]
}
```

- `sheet_checks` declares which sheets are expected and whether each
  is required. Omitting a sheet from this section means its presence
  is never checked.
- `rule_detection` tunes the content-based rule detection mechanism:
  `match_threshold` is the minimum fraction of sampled values a rule
  must match to be assigned to a column, and `sample_size` caps how
  many values are sampled per column.
- `rules` lists which of the framework's registered rules are enabled
  and what parameters they use. Unknown rule names or invalid
  parameters are rejected before execution with a clear error message.

Configuration defines **what** runs and with **what parameters**; it
never contains executable code, regular expressions bound to specific
business meaning, or Python logic of any kind. How each rule works is
implemented entirely in the framework's Python code.

Supplying a configuration file is optional.

## Default Validation Mode

When `--config` is not supplied, or the given path does not exist, the
framework does not fail. It loads its built-in default rule
configuration and executes the identical validation engine used when a
configuration file is supplied. The default configuration enables the
framework's structural and data-quality rules, requires no declared
sheets (every sheet found in the input is treated as present), and
uses the same rule-detection tuning as the example configuration
above.

## Project Structure

```
Validation_Automation_Framework/
│
├── config/
│   ├── validation_config.json
│   └── default_config.json
│
├── input/
│
├── output/
│
├── logs/
│
├── src/
│   └── validation_framework/
│       ├── core/          engine, registry, runner
│       ├── config/        models, loader, defaults, validator
│       ├── readers/       base, excel_reader, csv_reader, factory
│       ├── rules/         base, detection, structural/, data_quality/, integrity/, business/
│       ├── results/       models, collector
│       ├── reporting/     excel_report, json_report, formatter
│       ├── logging/       logger
│       ├── exceptions/    exceptions
│       └── utils/         discovery
│
├── tests/
│   ├── unit/
│   └── integration/
│
├── main.py
├── requirements.txt
├── pyproject.toml
└── README.md
```

- `config/validation_config.json` - example configuration.
- `config/default_config.json` - the built-in default rule set, used
  automatically when no configuration is supplied.
- `input/` - place files to be validated here.
- `output/` - generated reports are written here.
- `logs/` - the framework's log file is written here.
- `main.py` - the command-line entry point.

## Input Formats

The framework accepts a single input file per run, supplied through
the `--input` argument. Supported formats:

- `.xlsx` and `.xls` - every sheet in the workbook is read.
- `.csv` - read as a single table, named after the file.

## Output

Each run produces two reports:

**`Validation_Report.xlsx`** - three sheets:

- **Validation Summary** - one row per check executed (sheet
  existence, rule determination, and data validation), with sheet
  name, field, test type, rule name, status, and failed count.
- **Failed Records** - one row per individual failing cell or row,
  with sheet name, field, row number, actual value, rule name,
  severity, and error message.
- **Execution Summary** - run-level totals: sheet and column counts,
  rows processed, rules enabled, checks executed/passed/failed, total
  failed records, execution duration, and overall status.

**`validation_results.json`** - a machine-readable file containing the
same execution summary, validation summary, and full result list as
the Excel report, generated from the identical result objects.

## Installation

```bash
pip install -r requirements.txt
```

Requires Python 3.10 or later, `pandas`, and `openpyxl`.

## CLI Usage

```bash
python main.py --help
```

With a JSON configuration file:

```bash
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"
```

Without a configuration file (built-in defaults are used automatically):

```bash
python main.py --input "input/Validation_Sample.xlsx"
```

Optional arguments:

| Argument | Description | Default |
|---|---|---|
| `--config` | Path to a JSON configuration file | `config/validation_config.json` |
| `--input` | Path to the input file to validate (required) | - |
| `--output` | Path for the generated Excel report | `output/Validation_Report.xlsx` |
| `--json-report` | Path for the generated JSON report | `output/validation_results.json` |
| `--log-dir` | Directory for log files | `logs` |

The process exits with `0` on success, or a non-zero code identifying
a configuration error (`2`), input error (`3`), execution error (`4`),
or other framework error (`1`), or when the input contains a failing
required-sheet check.

## Testing

Run the full test suite with:

```bash
python -m pytest
```

The suite includes unit tests for every rule, the rule registry, the
configuration loader and validator, the input readers, the validation
engine, the result model and collector, and the report generators, as
well as integration tests covering JSON-configured mode, default mode,
and error scenarios such as missing input, invalid configuration,
unknown rules, invalid parameters, missing sheets, missing columns,
empty datasets, and invalid data values.

## Extensibility

**Adding a rule:** implement a class inheriting from
`BaseValidationRule` in the appropriate category module under
`src/validation_framework/rules/`, decorate it with `@register_rule`,
and it becomes resolvable by name immediately - no change to the
validation engine is required.

**Adding an input format:** implement the `InputReader` interface in
`src/validation_framework/readers/` and register it in
`readers/factory.py`.

**Adding a report format:** implement a new writer in
`src/validation_framework/reporting/` that consumes the same
standardized result and summary objects already produced by the
engine.

**Changing default behavior:** edit `config/default_config.json` -
the file the framework loads automatically when no configuration is
supplied.
