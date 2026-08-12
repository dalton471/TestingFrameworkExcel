# KPI Validation Automation Framework

A generic, configuration-driven framework that validates KPI Excel/CSV
workbooks against a JSON rule set: sheet/column existence, duplicates,
nulls, numeric precision, formulas, business rules, reconciliation,
cross-sheet checks, and month-over-month trend checks.

**Architecture:** `INPUT FILE + JSON CONFIG -> generic Python validation
engine -> results -> report`. Python defines *how* a validation type
works; the JSON defines *what* to validate. Adding or changing a
business rule normally only requires editing the JSON.

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

```
pip install pandas openpyxl
```

## Command-line usage

```
python main.py --config config/Kpi_automation_phase1.json --input input/KPI_Sample.xlsx
python main.py --config config/Kpi_automation_phase1.json --input input/
python main.py --config config/Kpi_automation_phase1.json --input input/KPI_Sample.xlsx --output output/MyReport.xlsx
python main.py --help
```

- `--config` (required) — path to the JSON rule file.
- `--input` (required) — a single `.xlsx`/`.xls`/`.csv` file, or a
  folder. In folder mode, every supported file is discovered
  (alphabetically, skipping Excel lock files like `~$KPI.xlsx`),
  validated independently, and given its own report
  (`Validation_Report_<filename>.xlsx`). One bad file does not stop
  the others; a run summary is printed at the end and the process
  exits non-zero only if every file failed.
- `--output` (optional) — defaults to `output/Validation_Report.xlsx`.

## CSV behaviour

A CSV is a single flat table with no worksheets. It is loaded and
converted into a temporary single-sheet `.xlsx` (sheet name
`Sheet1`) so it can flow through the same engine. Any rule that
expects a different sheet name (sheet existence, cross-sheet,
lookups, reconciliation, etc.) will correctly report that sheet as
missing — this is expected, not a bug, and is not hidden or
suppressed. There is currently no config option to map a CSV onto a
different logical sheet name; if you need that, add a
`"csvsheetname"` key to the config and a couple of lines in
`prepare_excel_path()`.

## Configuration architecture

Every validator reads its parameters from JSON — no sheet names,
column names, operators, thresholds, practice codes, or expected
counts are hard-coded in Python.

- **sheetlist / sheetvalidation** — expected sheets/fields.
- **dataqualityvalidation** — `duplicatecheck`, `nullcheck`,
  `numericprecision` (`precision`, and `usecellformat` — see below).
- **businessrulevalidation** — each rule has a `"type"` that
  selects a generic engine:
  - `comparison` — `leftcolumn OPERATOR rightcolumn` (dates or
    numbers, via `datatype`), with `blankindicators` (a config list
    of strings such as `"not applicable"` that count as "blank").
  - `staledata` — a date column older than `referencecolumn -
    months` should be blank; if it isn't, that row fails.
  - `lookupmatch` — `leftcolumn` on this sheet should equal
    `rightcolumn` on `lookupsheet`, joined on
    `lookupkeyleft`/`lookupkeyright`.
- **formulavalidation** — `threshold` or `comparison` formulas
  evaluated with `pandas.DataFrame.eval`.
- **reconciliationvalidation** — source/target sheet + match
  columns + numeric `tolerance` (default `0.01`) instead of exact
  float equality.
- **crosssheetvalidation** — `valuecomparison` or `existencecheck`,
  with `dropnullkeys` to ignore blank/junk key rows.
- **trendvalidation** — `metric`/`previousmetric`/`operator`/
  `threshold`, respected exactly as configured (see fix below).

## Numeric precision — root cause and fix

**Problem:** the original code counted decimal digits in
`str(raw_float_value)`. Aggregated/derived KPI columns are stored at
full floating-point precision (e.g. `133062.020162`) even though
Excel *displays* them rounded via cell formatting (e.g.
`#,##0.00`). Counting digits on the raw value flagged nearly every
row as a false failure.

**Fix:** `validate_numeric_precision` now opens the workbook with
`openpyxl` and reads each cell's `number_format`. If the format
fixes the number of displayed decimals (e.g. `#,##0.00` → 2,
`#,##0.0000` → 4), that is what's used to decide pass/fail — it's
what a user actually sees. Only when a cell has no explicit format
(`General`, or a CSV with no styling) does the code fall back to the
raw value, using a tolerance so genuine float noise
(`...0000000004`) isn't mistaken for real extra precision.

Verified against `KPI_Sample.xlsx`: `AvgDailyChargesByPractice` has
one cell formatted to 4 decimals (practice `PC_002`) while every
other cell in that column — and every cell in
`AvgDailyChargesByProvider`/`AvgDailyChargesByLocation` — is
formatted to 2 decimals. Result: 1 failure / 0 / 0, matching Durga
Ma'am's expected report exactly (she had annotated `F=39→1`,
`F=31→P`, `F=36→P`).

Set `"usecellformat": false` on a `numericprecision` rule to force
the raw-value/tolerance check everywhere (useful for CSV input,
which has no cell formatting).

## Other bugs found and fixed (with data proof)

1. **Lookup-based business rules incorrectly required the lookup
   sheet's column to exist on the *source* sheet.** The
   `Practice Status ↔ ProblemReason` rule needs `ProblemReason` from
   `Practice Detail`, not `KPI`. The old required-columns check
   looked for `rightcolumn` on the KPI dataframe, so it always
   failed with "Missing Column: ProblemReason" before the actual
   comparison ever ran (masking the real result behind `F=1`). Fixed
   by only checking lookup-side columns against the lookup
   dataframe. Verified result: `F=4` (`PC_002, PC_020, PC_021,
   PC_022`), exactly matching Durga Ma'am's note.

2. **Cross-sheet `valuecomparison` inflated failures by not
   deduplicating on the match key.** `KPI` has multiple rows per
   `PracticeCode` (one per fiscal month); every one of those rows
   was counted as a separate failure, turning 4 genuinely mismatched
   practices into 12 counted failures. Fixed by deduplicating on the
   match key (and dropping rows with a null key, e.g. the stray
   blank trailing row found in the KPI sheet) before counting.
   Verified result: `F=4`, matching Durga Ma'am's note.

3. **Trend validation ignored the configured operator for
   non-aggregate rules.** The per-row branch always used
   `abs(deviation) > threshold`, even if the JSON specified `<`,
   `>=`, `==`, etc. Fixed to use the same generic operator dispatch
   as the aggregate branch.

4. **Reconciliation used exact float equality**, which is fragile
   against floating-point noise from prior aggregation. Now uses a
   configurable `tolerance` (default `0.01`).

## Discrepancy NOT force-matched (flagged for review)

Durga Ma'am's note flags `PC_010` as an expected `F=1` on "Latest
Transaction Date older than 3 months should display No
LatestTransactionDate or NULL". Investigating the actual data:
`PC_010`'s `LatestTransactionDate` (2026-04-02) is **67 days**
before its `LastLoadedDate` (2026-06-08) — the largest gap of any
practice in the sheet, but still under the literal 3-month
(~90-day) threshold configured in the JSON. Under the rule exactly
as configured, this row correctly passes.

**This was not hard-coded to force a match.** If the intended
business threshold is shorter than 3 months (e.g. 60 days), that is
a one-line change to `"months"` in the JSON — not a Python change —
and should be confirmed with Durga Ma'am rather than guessed.
**NOT VERIFIED** as a framework bug; flagged as a config/threshold
question.

## Error handling

- Missing config / missing input / config path is a directory /
  unsupported extension / malformed or empty JSON → clear message,
  exit code 1, nothing is silently swallowed.
- A missing config *section* (e.g. no `trendvalidation` key) no
  longer aborts the whole file's report — only that one validator is
  skipped, and a `Configuration Validation` row records it.
- A missing sheet, missing column, or corrupt sheet inside a single
  validator produces an `F` result row for that check and the
  pipeline continues.
- In folder/batch mode, one corrupt/unreadable file does not stop
  the others; a summary of succeeded/failed files is printed, and
  the process exits non-zero only if every file failed.

## Determinism

Same input + same config + same code → identical results, byte for
byte, on repeated runs (verified — see test log). No shared mutable
state or accumulation between files or runs.

## Testing performed

- `KPI_Sample.xlsx` — full pipeline run, 146 result rows (matches
  Durga Ma'am's report row count), report inspected against her
  annotated discrepancies (3 of 4 now match exactly; 1 flagged as a
  config-threshold question above).
- Numeric precision unit test against the spec's exact cases:
  `123, 123.4, 123.45, 123.00` → PASS; `123.456, 123.4567` → FAIL.
- Repeated identical run → byte-identical report (determinism).
- Missing config file, missing input path, malformed JSON, empty
  JSON, unsupported extension, `--help` → all handled with clear
  messages and correct exit codes.
- Missing config section (`trendvalidation` removed) → only that
  validator skipped, rest of the report still generated.
- CSV input → runs without crashing; sheet-based rules correctly
  report the sheets a CSV can't have.
- Folder/batch mode with one valid `.xlsx` and one corrupt `.xlsx` →
  the valid file's report is generated, the corrupt one is reported
  as failed, run summary shows `1 succeeded / 1 failed`.

## Known limitations

- **`KPI_Sample (2).xlsx` (the regression file) was not provided**
  in this session, so the cross-file regression comparison called
  for in the brief is **NOT VERIFIED**. The same command/config
  should be re-run against it before sign-off:
  `python main.py --config config/Kpi_automation_phase1.json --input input/"KPI_Sample (2).xlsx"`.
- CSV → single-sheet mapping is fixed to `Sheet1`; making that
  configurable per rule set is a small follow-up if you validate
  multi-table CSVs regularly.
- The `PC_010` staleness discrepancy above needs a business decision
  (see above), not a code fix.
