# Validation Automation Framework

## 1. Purpose

A generic, reusable Python framework that validates a multi-sheet
Excel workbook and produces a structured, three-sheet
`Validation_Report.xlsx` - without depending on a manually maintained
list of the workbook's fields, without depending on the JSON
configuration to supply validation conditions (ranges, formats,
etc.), and without requiring the JSON configuration file to exist at
all. Every column is discovered from the workbook at runtime, the
validation rule that applies to it is determined from its actual
content, and the rule itself (what counts as valid) is implemented as
ordinary Python logic - exactly like Dalton's original functions.
JSON's only remaining job is to optionally *select/highlight* which
of the framework's built-in validations were explicitly requested;
every validation the framework supports always runs regardless
(section 11).

```
Excel Workbook
        |
Dynamic Sheet Discovery        (sheet_checks - from JSON if present,
        |                        else discovered directly from workbook)
Dynamic Column/Field Discovery  (df.columns at runtime - no config)
        |
Validation Selection            (JSON may name a subset to highlight;
        |                        the Python registry always supplies
        |                        the rest - see section 11)
Generic Rule Determination      (content-based, in Python)
        |
Validation Logic                (conditions defined in Python)
        |
Validation Results (PASS/FAIL only - no SKIPPED)
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

## 3. What Changed: JSON vs. Python Responsibilities

This is the core of this update. Previously, `column_checks` in the
JSON supplied validation *conditions* (amount's min/max, number's
min/max, date's format string). That made the framework depend on
someone hand-tuning JSON for every new workbook - exactly the
dependency this change removes.

**JSON now contains only generic framework metadata:**

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
  }
}
```

- `workbook`: informational only.
- `sheet_checks`: which sheets are expected to exist, and whether
  each is required. This is sheet-level configuration (does the
  *workbook* have the right shape), not a validation condition on any
  field's *content* - it stays in JSON deliberately, as explicit
  framework-level configuration.
- `rule_detection`: tuning for the auto-detection mechanism itself
  (see section 5) - how confidently a column's content must match a
  rule before it's assigned, and how many values to sample per
  column. This is a knob on the *classifier*, not on what makes a
  value valid, so it also legitimately belongs in JSON.

There is **no `column_checks` section**, **no `fields` section**, and
no other place in JSON that describes what a valid value looks like.

**Python (`validations/validation.py`) now contains everything that
used to live in `column_checks`:**

```python
NUMBER_MIN_VALUE = 0
NUMBER_MAX_VALUE = 100
AMOUNT_MIN_VALUE = 0
GENERIC_DATE_FORMATS = ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", ...]
```

These are ordinary Python constants at the top of the validation
module - the same place Dalton's original seven functions live. A
deployment that needs different generic conventions edits these
constants directly; it does not edit a per-workbook JSON file.

## 4. Generic Validation Rules (conditions now defined in Python)

| Rule            | What it checks                                          | Where the condition lives |
|------------------|-------------------------------------------------------------|------------------------------|
| `numeric_only`   | Value is a plain positive-integer digit string                | Python (fixed regex)          |
| `email`          | Value matches a basic email address pattern                    | Python (fixed regex)          |
| `phone_number`   | Value matches a common US phone number format                   | Python (fixed regex)          |
| `amount`         | Value is a non-negative real number                              | Python (`AMOUNT_MIN_VALUE = 0`) |
| `number`         | Value is a real number within 0-100                               | Python (`NUMBER_MIN_VALUE`/`NUMBER_MAX_VALUE`) |
| `date`           | Value parses under any of several common date formats              | Python (`GENERIC_DATE_FORMATS` list) |
| `text_only`      | Value contains only letters and spaces                              | Python (fixed logic)          |

**Why these specific generic conditions:**

- `number` (0-100) follows the widely recognized convention for a
  bounded metric - a percentage, rating, or score out of 100 - since
  the framework has no way to know a workbook-specific scale.
- `amount` (>= 0) reflects the one universally true constraint on a
  monetary/quantity figure - it can't be negative - without inventing
  an arbitrary upper limit that no generic amount actually has.
- `date` tries several real-world formats (ISO 8601, `DD/MM/YYYY`,
  `MM/DD/YYYY`, `Month D, YYYY`, etc.) rather than assuming or
  requiring one fixed format, so a workbook doesn't need a
  configuration entry just to say which format its dates use.

API validation remains out of scope, as in every previous version -
no API URL, token, `fetch_valid_codes()`, or network call exists
anywhere in this codebase.

## 5. Dynamic Field Discovery

Unchanged in mechanism from the previous version: `pandas.read_excel()`
reads each sheet's header row, and `df.columns` is treated as the
complete list of fields on that sheet. There is no expected-column
list, no maintained schema, and no code path that special-cases a
particular column name.

## 6. How Validation Rules Are Determined

For each discovered column, the framework samples its non-blank
values (capped at `rule_detection.sample_size`, default 200) and
checks what fraction of the sample each of the seven rules would
accept, in a fixed, generic priority order:

```
email  ->  phone_number  ->  date  ->  numeric_only  ->  number  ->  amount  ->  text_only
```

The first rule whose match rate reaches `rule_detection.match_threshold`
(default 60%) is assigned.

**Why this exact order matters:** among the purely numeric rules,
`numeric_only` (a strict positive-integer digit string) is the
*narrowest* pattern, `number` (0-100) is broader, and `amount`
(any non-negative number) is broadest of all. Every value that
satisfies `numeric_only` also satisfies `number` and `amount`, and
every value that satisfies `number` also satisfies `amount`. If a
broad rule were checked before a narrow one, the broad rule would
always "win" and the narrow one would become unreachable. Checking
narrowest-first guarantees all three rules stay meaningfully
reachable, using only the data's own structure - never a column name.

If nothing reaches the threshold, the column is marked
**Unrecognized** and is deliberately excluded from data validation -
the framework never guesses. This is unchanged from the previous
version (see the Validation Summary sheet, `Status = SKIPPED`).

## 7. A Known, Honest Trade-off of Full Genericization

Because `amount`/`number` conditions are no longer tailored per
workbook, a column of **plain whole numbers** (no decimal point) is
structurally indistinguishable, by content alone, from a generic
identifier column - both are just digit strings. Given the
narrowest-first priority above, such a column is classified
`numeric_only`.

This is visible in the bundled sample workbook: `Compensation.Salary`
and `Compensation.Performance_Score` contain only whole numbers (e.g.
`50000`, `85`), so they are now classified `numeric_only` rather than
`amount`/`number`. `numeric_only` still correctly rejects genuinely
corrupt values (e.g. `"ABC"`, `"XYZ"`), but it does not enforce a
numeric range, so an out-of-range whole number (e.g. a `500000`
salary or a `150` score) is no longer flagged purely on range.

This is an inherent, documented consequence of removing all
per-workbook configuration - there is no generic, content-only signal
that can distinguish "this whole number is an ID" from "this whole
number is an amount/score" without either a naming convention or a
workbook-tailored range, both of which this requirement rules out.
`amount` and `number` remain fully implemented and reachable - they
correctly classify and validate any column containing decimal values
(see section 8 for a live demonstration), and would immediately apply
to `Salary`/`Performance_Score` too if those columns contained
decimals (e.g. `50000.00`) instead of whole numbers.

## 8. Verification: Rules Still Work on Data That Exercises Them

Since the sample workbook's numeric columns happen to be whole
numbers, `amount`, `number`, and the multi-format date detector were
additionally verified against a synthetic sheet designed to exercise
them, with columns never referenced anywhere in this codebase:

| Column | Sample values | Rule assigned | Failure caught |
|---|---|---|---|
| `Invoice Total` | `199.99, 4599.50, -50.00, 12000.00` | `amount` | `-50.00` (negative) |
| `Satisfaction Rating` | `88.5, 42.0, 150.0, 71.25` | `number` | `150.0` (out of 0-100) |
| `Signup Date` | `15/01/2023, 2023-06-10, 31/02/2023, March 3, 2023` | `date` | `31/02/2023` (not a real calendar date) |

Three different date formats (`DD/MM/YYYY`, ISO 8601, and
`Month D, YYYY`) were each correctly accepted, and the impossible
date (February 31st) was correctly rejected regardless of which
format it was checked against - confirming the multi-format date
detector works without any format being configured.

## 9. Dynamic Discovery on a Renamed/Unseen Schema

As previously verified, and still true after this change: renaming a
column, or adding one that has never appeared anywhere in this
codebase, requires no config or code edit. For example, `Employee_ID`
renamed to `Staff Number`, `Employee_Email` renamed to `Work Email`,
and a brand-new `Signup Score` column are all still correctly
classified (`numeric_only`, `email`, `number` respectively) using the
exact same `validation_config.json` shown in section 3.

## 10. Optional JSON Configuration with Built-in Fallback

`validation_config.json` is now **optional**. The framework never
stops just because the file is missing.

**If the file exists at the given `--config` path:** it is loaded and
used exactly as documented in section 3 - no behavior change from
before.

**If the file does not exist:** `main.py` calls
`validations.validation.build_default_config(workbook_sheet_names)`
instead of raising an error. This builds a configuration object with
the identical shape (`workbook`, `sheet_checks`, `rule_detection`),
except `sheet_checks` is populated directly from whatever sheets the
input workbook actually contains (each marked required, since it was
discovered from the workbook itself - there's no external
expectation to fall short of), and `rule_detection` falls back to the
same defaults (`match_threshold = 0.6`, `sample_size = 200`) used
when a JSON file simply omits that section.

**There is only one validation engine.** Both branches feed their
resulting configuration object into the exact same
`run_validations()` function - the only difference between the two
scenarios is which function produced the configuration object
(`load_json()` vs. `build_default_config()`); everything downstream
(rule determination, data validation, report generation) is
identical code, run identically.

```
JSON exists?
    |
   YES -> load_json(path)  ----\
    |                            \
   NO  -> build_default_config()  \--> run_validations() --> Report
             (built from workbook's
              own discovered sheets)
```

No sample-specific field or sheet name is hardcoded in this fallback
- `build_default_config()` takes the workbook's discovered sheet
names as a parameter and works identically for any workbook.

## 11. JSON Selects, Python Defines (Validation Selection Mechanism)

This is the final architectural refinement: JSON may optionally list
which validations to explicitly call out via a `"validations"` array,
but this **never changes what actually runs**. The framework's Python
registry (`SUPPORTED_VALIDATIONS` in `validations/validation.py`) is
the single authoritative list of every validation the framework
knows how to perform, and **every run executes all of them** -
JSON only affects how the selection is *reported*, distinguishing
"explicitly requested" from "automatically added by the framework."

```json
{
  "validations": ["numeric_only", "email", "amount", "date", "text_only"]
}
```

Given the framework's 7 supported validations, this selects 5 of
them explicitly; `resolve_active_validations()` in
`validations/validation.py` computes:

- **Selected from JSON**: `numeric_only, email, amount, date, text_only` (5)
- **Automatically added**: `phone_number, number` (2 - present in the
  framework's registry but not named in JSON)
- **Final validation list executed**: all 7, always

This mirrors the requirement's own worked example (5 of 10 selected,
5 auto-added, 10 executed) using this framework's actual 7-item
registry. If the `"validations"` key is omitted from JSON entirely,
or JSON is absent altogether (Mode 2), *nothing* is "selected", so
*every* supported validation is reported as automatically added - the
functional result (all 7 run) is identical either way, only the
"selected vs. auto-added" labeling in the terminal and Execution
Summary differs.

**Why JSON can only select, never narrow:** the requirement is
explicit that "the final execution should contain: JSON-selected
validations + built-in validations not specified in JSON" - i.e. the
union of the two groups, which is always the complete registry. This
also means JSON can never disable a validation the framework
supports, and an unrecognized name in `"validations"` (a typo, or a
validation the framework doesn't implement) is reported as a warning
and safely ignored rather than crashing the run.

**Why this satisfies "JSON must not contain validation logic":** the
list in JSON is just a list of *names* - strings that are looked up
against the Python registry. No regex, range, date format, or
conditional logic ever appears in JSON; all of that remains exactly
where section 3 describes, in `validations/validation.py`.

## 12. How to Execute the Framework

With JSON configuration:

```bash
pip install -r requirements.txt
python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"
```

Without JSON configuration (omit `--config`, or point it at a path
that doesn't exist - the built-in defaults are used automatically,
with no error):

```bash
python main.py --input "input/Validation_Sample.xlsx"
```

Optional custom output path (works in either scenario):

```bash
python main.py --input "input/Validation_Sample.xlsx" --output output/MyReport.xlsx
```

Both commands print a **VALIDATION SELECTION** block before running
validations, showing exactly which validations were selected from
JSON, which were automatically added, and the final list executed -
see section 11.

## 13. Input Workbook Requirements

An `.xlsx`/`.xls` file, one header row per data sheet. No requirement
on column names, order, or count. When JSON is present, sheets listed
in `sheet_checks` are checked for existence; when JSON is absent,
every sheet found in the workbook is automatically treated as
present (section 10).

## 14. Output Report Structure

`output/Validation_Report.xlsx` contains **exactly three sheets**,
identical whether JSON was present or the built-in fallback was used:

1. **Validation Summary** - every test that ran: Sheet Existence
   Testing, Rule Determination (per successfully classified column -
   see section 15), and Data Validation (per column that received a
   rule).
2. **Failed Records** - one row per individual failing cell.
3. **Execution Summary** - run-level totals.

There is no `Validation Details` sheet, and this update did not
reintroduce it.

## 15. Failed Records Behavior

Unchanged: only failed cells are listed, with dynamically generated
error messages (`Non-numeric value in column 'X'`,
`Invalid email in column 'Y'`, `Out of range in column 'Z'`, etc.),
using the actual sheet, column, and value being processed.

## 16. No SKIPPED Status - PASS/FAIL Only

Every reported result now uses only `PASS` or `FAIL` - `SKIPPED` was
removed from the report entirely:

- **Sheet Existence Testing**: a sheet configured as optional
  (`"required": false`) and not present is now reported `PASS` (its
  absence is an expected, non-failing outcome), not a separate
  status. A required-but-missing sheet is still correctly `FAIL`.
- **Rule Determination**: a column whose content doesn't confidently
  match any of the seven rules - most often because the sheet is
  documentation/reference content rather than tabular data (e.g. a
  "Read Me" sheet) - simply produces **no Rule Determination row at
  all**, for that column. It is never falsely marked `PASS` (no rule
  was actually confirmed for it) and never falsely marked `FAIL` (no
  data validation failure occurred - none ran). This is not
  hardcoded to any particular sheet name; it applies to any column on
  any sheet that the content-based classifier can't confidently place
  into one of the seven rules.
- Its existence is still fully visible in the **Execution Summary**
  via the `Total Columns Unrecognized` metric, which is computed
  directly from the discovered column counts rather than from any
  per-row status - so nothing is silently hidden, it just isn't
  reported as a misleading PASS or FAIL.
- Real data-validation results for actual data sheets/columns are
  completely unaffected by this change - a column that failed before
  (e.g. an invalid email) still fails now.

## 17. Scalability Considerations

Unchanged from the previous version:

- The workbook is loaded once; every sheet is parsed into a DataFrame
  a single time and reused across all three validation layers.
- Rule determination samples at most `rule_detection.sample_size`
  non-blank values per column, so classification cost doesn't grow
  with sheet size.
- Once a rule is assigned, every row in that column is still checked
  for the actual failure pass (sampling only affects rule
  *determination*, never failure *detection*).
- No per-field Python branch, dictionary entry, or config line needs
  to be added as a workbook grows wider or gains new columns.
- No new external dependencies were introduced.

## 18. Before / After Comparison (Same Sample Workbook)

| Metric | Before Task 2/2B | After Task 2/2B (JSON present) | After Task 2/2B (JSON absent - built-in fallback) |
|---|---|---|---|
| Total Sheets | 5 | 5 | 5 |
| Sheets Passed / Failed | 5 / 0 | 5 / 0 | 5 / 0 |
| Total Columns Discovered | 14 | 14 | 14 |
| Total Rows Processed | 40 | 40 | 40 |
| **Total Validations Executed** | **32** | **31** | **31** |
| Total Passed | 18 | 18 | 18 |
| Total Failed | 13 | 13 | 13 |
| Total Failed Cells | 17 | 17 | 17 |
| Report sheets | 3 | 3 | 3 |
| **SKIPPED status in report** | **1 row** (Read Me's column) | **0** | **0** |

The one number that changed (32 -> 31 total validations) is exactly
the removed Read Me `SKIPPED` row (section 15) - not forced, simply
the direct, expected effect of no longer emitting a row for
undetermined columns. Every other number is identical, and - most
importantly - the JSON-present and JSON-absent columns are
byte-for-byte identical, confirming both scenarios genuinely run the
same validation engine.

## 19. Error Handling

Unchanged from the previous version - missing/invalid paths, malformed
JSON, corrupt workbooks, missing required sheets, and unrecognized
columns are all handled explicitly (see the previous README revisions
for the full list); nothing here was affected by this change.

## 20. Known Limitations

- `phone_number` only supports common US-style formats.
- `email` uses a basic regex, not full RFC 5322 validation.
- Whole-number columns cannot be generically distinguished between
  "identifier" and "bounded amount/score" by content alone (section 7)
  - this is an inherent trade-off of having no per-workbook
    configuration, not a bug.
- `GENERIC_DATE_FORMATS` covers common real-world formats but is not
  exhaustive; a workbook using an unusual date format would need that
  format added to the Python list (a framework-level Python change,
  not a per-workbook JSON change).
- A column with no non-blank values in the sampled rows cannot be
  classified and is marked Unrecognized - intentional (section 6).
