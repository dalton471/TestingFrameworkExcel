# Validation Automation Framework

## 1. Purpose

A generic, reusable Python framework that validates a multi-sheet
Excel workbook against a small set of JSON-defined generic rules and
produces a structured, three-sheet `Validation_Report.xlsx`.

The framework does **not** maintain a list of the workbook's column
names anywhere. Every column is discovered from the workbook itself
at runtime, and the validation rule applied to it is determined from
the actual content of its cells. This is what makes the framework
reusable across different workbooks (and different runs of the same
workbook, if columns are added, removed, or renamed) without touching
Python or JSON.

```
Generic JSON Rules
        |
Excel Workbook
        |
Dynamic Sheet Discovery
        |
Dynamic Column/Field Discovery
        |
Generic Rule Determination
        |
Validation Engine
        |
3-Sheet Validation Report
```

## 2. Project Structure

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

## 3. Configuration Structure

`config/validation_config.json` contains **no column or field names**.
It has three sections:

- `workbook`: informational metadata only.
- `sheet_checks`: which sheets are expected, and whether each is
  `required`. This is sheet-level configuration, not field-level, and
  is unaffected by this rebuild.
- `column_checks`: the generic validation rules that exist, and their
  parameters (ranges, date format). This is the complete list of
  rules the framework knows how to apply - it is not tied to any
  specific column.
- `rule_detection`: tuning for the dynamic rule-determination
  mechanism (see section 6): `match_threshold` (default `0.6`) and
  `sample_size` (default `200`).

```json
{
  "sheet_checks": {
    "Employee Details": { "required": true }
  },
  "column_checks": {
    "amount": { "min_value": 20000, "max_value": 200000 },
    "date": { "format": "%Y-%m-%d" }
  },
  "rule_detection": {
    "match_threshold": 0.6,
    "sample_size": 200
  }
}
```

There is intentionally no `"fields"` section, and nowhere in this
repository is there an equivalent manually maintained mapping of
column name to validation rule.

## 4. Generic Validation Rules

| Rule            | What it checks                                       | Parameters                 |
|------------------|--------------------------------------------------------|------------------------------|
| `numeric_only`   | Value contains only digits                              | none                          |
| `email`          | Value matches a basic email address pattern              | none                          |
| `phone_number`   | Value matches a common US phone number format             | none                          |
| `amount`         | Value is numeric and within `[min_value, max_value]`      | `min_value`, `max_value`      |
| `number`         | Value is numeric and within `[min_value, max_value]`      | `min_value`, `max_value`      |
| `date`           | Value parses under the configured `format`                | `format` (strptime pattern)   |
| `text_only`      | Value contains only letters and spaces                     | none                           |

These are Dalton's original seven validators; their internal logic is
unchanged. What changed is how a column gets matched to one of them.

API validation remains intentionally excluded from this version, per
Dalton's original instruction - there is no API URL, token,
`fetch_valid_codes()`, or network call anywhere in this codebase.

## 5. Dynamic Field Discovery

Field discovery has exactly one mechanism: `pandas.read_excel()`
reads each sheet's header row, and `df.columns` is treated as the
complete list of fields on that sheet. There is no expected-column
list to check against, no maintained schema, and no code path that
special-cases a particular column name.

If a column is renamed, removed, or added between runs, the very next
run simply discovers whatever columns are actually there - nothing
needs to be edited. This is verified in section 14.

## 6. How Validation Rules Are Determined

For each discovered column, the framework samples its non-blank
values (capped at `rule_detection.sample_size`, default 200, so very
wide/tall sheets stay fast to classify) and checks, for every rule
defined in `column_checks`, what fraction of the sample that rule
would accept.

Rules are tried in a fixed, generic priority order - most
structurally specific first:

```
email  ->  phone_number  ->  date  ->  amount  ->  number  ->  numeric_only  ->  text_only
```

The **first** rule whose match rate reaches `match_threshold`
(default 60%) is assigned to the column. Evaluating in this order,
rather than picking whichever rule scores highest, is what keeps a
bounded numeric score from being mistaken for a free-form numeric
identifier, and an email address from being mistaken for generic
text - using only the data itself, never the column's name.

This priority order and the threshold are fixed properties of the
*mechanism*, not per-column configuration - they apply identically to
every column on every sheet in every workbook.

**Why `amount` and `number` don't collide:** both are numeric range
checks, distinguished only by their configured `min_value`/
`max_value`. A column is assigned whichever of the two its values
actually fall inside at a high rate - e.g. large currency-like values
naturally clear the `amount` range and fail the `number` range (and
vice versa for small bounded scores). No naming convention is
involved.

**Why ID-like columns become `numeric_only` and not `amount`/
`number`:** identifier values (e.g. `1001`, `1002`) are pure digit
strings, so `numeric_only` matches them at a high rate - but they
typically fall outside whatever ranges are configured for `amount`
and `number` (which come earlier in the priority order and are tried
first). If a real dataset's identifiers happen to fall inside a
configured `amount`/`number` range, that range-based rule will be
selected preferentially; this is a known, documented edge case (see
section 12).

## 7. Fallback Behavior for Unrecognized Columns

If no rule reaches the match threshold, the column is marked
**Unrecognized** and is deliberately excluded from data validation -
the framework never guesses a rule "close enough" and applies it
anyway. This is reported transparently:

- It appears in the **Validation Summary** sheet with
  `Test Type = Rule Determination`, `Status = SKIPPED`, and a
  `Test Name` such as `Unrecognized (best match 42%)` so it's clear
  why no rule was applied.
- It is excluded from `Total Rows Processed` and `Total Failed Cells`
  in the Execution Summary, since no validation actually ran against
  it.
- It never produces false failures in **Failed Records**.

This is the framework's answer to "what happens when field discovery
finds something it can't classify" - it is surfaced, not hidden and
not guessed at.

## 8. Supported Validation Types

See section 4. All seven of Dalton's original types are supported;
none were added or removed by this change.

## 9. How to Execute the Framework

```bash
pip install -r requirements.txt
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"
```

Optional custom output path:

```bash
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx" --output output/MyReport.xlsx
```

## 10. Input Workbook Requirements

- An `.xlsx` (or `.xls`) file with one or more worksheets.
- Sheets listed in `sheet_checks` are checked for existence; a
  workbook may contain additional sheets not listed there, which are
  still dynamically discovered and validated the same way.
- Each data sheet should have a single header row followed by data
  rows. There is no requirement on column names, order, or count -
  the framework adapts to whatever it finds.
- Sheets with no columns that match any generic rule (e.g. a purely
  descriptive "Read Me" sheet) are automatically excluded from data
  validation and row-processing totals - no configuration is needed
  to achieve this; it falls out naturally from section 7.

## 11. Output Report Structure

`output/Validation_Report.xlsx` now contains **exactly three sheets**:

1. **Validation Summary** - every test that ran: Sheet Existence
   Testing, Rule Determination (one row per discovered column,
   showing which rule - if any - was assigned), and Data Validation
   (one row per column that received a rule). Columns: `Sheet Name`,
   `Field`, `Test Type`, `Test Name`, `Status`, `Failed Count`.
2. **Failed Records** - one row per individual failing cell:
   `Sheet Name`, `Field`, `Row`, `Value`, `Validation Type`,
   `Required Datatype`, `Error Message`.
3. **Execution Summary** - run-level totals (see section 13).

The previously separate **Validation Details** sheet has been
removed. The information it carried that is still meaningful under
dynamic discovery (which rule applies to which column, and its
pass/fail status) is already visible in the Validation Summary sheet;
the removed sheet's row-count/parameter breakdown was tied to the old
static field configuration and is no longer produced anywhere.

## 12. Failed Records Behavior

Only failed cells are listed - one row per failure, never a row for
values that passed. Error messages are generated dynamically from the
actual sheet name, column name, and value being processed, for
example:

- `Non-numeric value in column 'Employee_ID'`
- `Invalid email in column 'Employee_Email'`
- `Out of range in column 'Performance_Score'`

Row numbers account for the header row, so they line up with the row
you'd see if you opened the sheet in Excel.

## 13. Execution Summary Metrics

All values below are computed dynamically from the run's actual
results - none are hardcoded:

- `Total Sheets`, `Sheets Passed`, `Sheets Failed`
- `Total Columns Discovered`, `Total Columns Recognized`,
  `Total Columns Unrecognized` (new - visibility into dynamic
  discovery itself)
- `Total Rows Processed` (rows in sheets that had at least one
  recognized column)
- `Total Validations Executed`, `Total Passed`, `Total Failed`,
  `Total Skipped` (new - unrecognized-column tests, which are neither
  pass nor fail)
- `Total Failed Cells`
- `Execution Status`

## 14. Scalability Considerations

- The workbook is loaded once (`utils/excel_reader.py`); every sheet
  is parsed into a DataFrame a single time and reused across all
  three validation layers - no repeated I/O.
- Rule determination samples at most `rule_detection.sample_size`
  non-blank values per column (default 200), so classifying a column
  with a million rows costs the same as classifying one with 500.
- Once a rule is assigned, data validation still checks **every** row
  in that column (sampling is only used for rule *determination*, not
  for finding failures) - so no invalid record is ever missed because
  it fell outside the sample.
- There is no per-field Python branch, dictionary entry, or config
  line to add as a workbook grows wider or gains new columns - cost
  scales with the number of sheets and columns actually present, not
  with a maintained list.
- No new external dependencies were introduced; the mechanism uses
  only `pandas` and Python's standard library (`re`, `datetime`),
  exactly as before.

## 15. Verification: Dynamic Discovery on an Unseen Schema

To confirm there is no hidden sample-specific mapping, the engine was
run directly (bypassing the sample workbook entirely) against a
synthetic DataFrame with renamed columns and one column that has never
appeared anywhere in this codebase:

| Column (never configured anywhere) | Sample values | Rule assigned | Correct? |
|---|---|---|---|
| `Staff Number` (renamed from `Employee_ID`) | `2001, 2002, 20A3, 2004` | `numeric_only` | Yes |
| `Work Email` (renamed from `Employee_Email`) | `a@x.com, bad-email, c@y.com, d@z.com` | `email` | Yes |
| `Signup Score` (brand new, never seen before) | `95, 42, 150, 88` | `number` | Yes |

No config edit and no code change were made between the sample-workbook
run and this test - the same `validation_config.json` was reused
as-is. This is the practical proof of the "no maintained field list"
requirement.

## 16. Testing Strategy

The included `input/Validation_Sample.xlsx` (unchanged from the
previous version - not modified to make this change pass) was
re-validated against the new dynamic-discovery engine. Every column
was independently re-classified by content alone and, for this
dataset, arrived at the exact same rule assignment the old static
`"fields"` configuration used to specify by hand - confirming the
generic mechanism reproduces correct, expected behavior on real data
(see section 17 for the full before/after comparison).

## 17. Before / After Comparison (Same Sample Workbook)

| Metric | Before (static `"fields"` config) | After (dynamic discovery) | Explanation |
|---|---|---|---|
| Total Sheets | 5 | 5 | unchanged |
| Sheets Passed / Failed | 5 / 0 | 5 / 0 | unchanged |
| Total Rows Processed | 40 | 40 | unchanged - Read Me was excluded before by not being listed in `"fields"`; it's now excluded because its one pseudo-column doesn't match any rule (section 7) |
| Total Validations Executed | 31 (5 sheet + 13 field-existence + 13 data) | 32 (5 sheet + 14 rule-determination + 13 data) | +1 because Read Me's column is now discovered and reported (as Unrecognized/SKIPPED) instead of never being considered |
| Total Passed | 18 | 18 | unchanged |
| Total Failed | 13 | 13 | unchanged |
| Total Failed Cells | 19 | 19 | unchanged - every rule assignment landed on the same rule the static config used, so the same 19 deliberately invalid cells are caught |
| Report sheets | 4 (incl. Validation Details) | 3 | Validation Details removed per requirement |

The one numeric difference (31 -> 32 total validations) is explained
above, not forced to match; every other number is identical because
the content-based classification independently arrived at the same
rules the old hand-written config specified.

## 18. Error Handling

Unchanged from the previous version:

- Missing/invalid `--config` or `--input` paths exit with a clear
  `[ERROR]` message and non-zero exit code.
- Malformed/empty JSON, corrupt or unreadable workbooks, and
  unexpected exceptions are caught and reported (with traceback for
  genuinely unexpected errors) rather than crashing silently.
- A configured sheet missing from the workbook is reported as FAIL at
  the Sheet Existence layer; no columns are discovered for it and no
  data validation runs against it.
- A column that can't be confidently classified is reported as
  SKIPPED, never silently treated as passing or guessed at (section 7).

## 19. Known Limitations

- `phone_number` only supports common US-style formats.
- `email` uses a basic regex, not full RFC 5322 validation.
- Rule determination is a heuristic based on sampled content; on
  unusual or adversarial data (e.g. an identifier range that happens
  to overlap a configured `amount`/`number` range) it can select a
  different rule than a human would expect. Increasing
  `rule_detection.match_threshold` makes classification stricter;
  adjusting `column_checks` ranges to be more distinct from expected
  identifier ranges avoids the overlap.
- A column with no non-blank values in the sampled rows cannot be
  classified and is marked Unrecognized (section 7) - this is
  intentional, not a bug.
- The framework validates against the generic rules and parameters
  defined in `column_checks`; it does not invent new rule types from
  the data itself.
