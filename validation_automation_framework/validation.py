"""
validation.py
--------------
Generic, self-contained validation engine for the Validation
Automation Framework.

ARCHITECTURE:

    Excel Workbook
        -> Dynamic Sheet Discovery        (sheet_checks in JSON)
        -> Dynamic Column Discovery       (df.columns at runtime)
        -> Generic Rule Determination     (content-based, this module)
        -> Validation Logic               (this module)
        -> Validation Results
        -> Report

JSON now supplies ONLY generic framework metadata: which sheets are
expected (sheet_checks) and how the auto-detection mechanism should
be tuned (rule_detection.match_threshold / sample_size). It contains
NO column names and NO validation conditions - no ranges, no date
formats, nothing that describes what a value must look like to be
valid. All seven validation rules, their acceptable ranges/formats,
and the order in which they are auto-detected are implemented here,
as ordinary Python, exactly like Dalton's original functions were.

This means the framework can be pointed at a completely different
workbook - different sheet names, different column names, different
column count - and it will still discover every column and apply the
same generic rules, without a single line of JSON or Python needing
to change.

Three validation layers, run in this order:
    1. Sheet Existence Testing     - is each configured sheet present?
    2. Rule Determination          - for every column discovered on
                                       every present sheet, which (if
                                       any) generic rule does its data
                                       match?
    3. Data Validation             - the seven field-level checks,
                                       applied cell by cell, using the
                                       rule assigned in step 2.

SEVEN GENERIC VALIDATION RULES (Dalton's original checks; behavior
preserved, acceptable conditions now defined as Python constants
below instead of being read from JSON):

    - numeric_only    : value is a plain positive-integer digit string
    - email           : value matches a basic email address pattern
    - phone_number     : value matches a common US phone number format
    - amount           : value is a non-negative real number
                          (generic monetary/quantity convention - no
                          fixed upper bound, since amounts don't have
                          one in general)
    - number           : value is a real number within 0-100
                          (generic bounded score/percentage/rating
                          convention)
    - date             : value parses under one of several common,
                          hardcoded date formats (no single format is
                          assumed or configured)
    - text_only        : value contains only letters and spaces

RULE DETERMINATION ALGORITHM (see README.md "How Validation Rules Are
Determined" for the full rationale, including why the priority order
below is deliberately narrowest-pattern-first):

    For each discovered column, take a sample of its non-blank values
    (capped by rule_detection.sample_size, default 200, purely for
    performance on very large sheets) and compute, for every rule in
    _RULE_PRIORITY, the fraction of sampled values that rule would
    accept. The FIRST rule (in priority order) whose match rate
    reaches rule_detection.match_threshold (default 0.6) is assigned
    to the column. If nothing reaches the threshold, the column is
    marked "Unrecognized" and is intentionally excluded from data
    validation rather than guessed at.

    This uses only generic signals (regex structure, numeric parse
    success, fixed ranges, date parseability) - never a column name
    and never a workbook-specific lookup table.

INTENTIONALLY EXCLUDED: API validation. Out of scope for this
framework - no API URL, token, fetch_valid_codes(), or network call
exists anywhere in this module.
"""

import re
from datetime import datetime

import pandas as pd

DEFAULT_MATCH_THRESHOLD = 0.6
DEFAULT_SAMPLE_SIZE = 200

UNRECOGNIZED = "Unrecognized"


# --------------------------------------------------------------------------
# Generic validation conditions - implemented in Python, not JSON.
#
# These are the ONLY place acceptable ranges/formats are defined. A
# different deployment that needs different conventions edits these
# constants, not a per-workbook config file.
# --------------------------------------------------------------------------

# "number": generic bounded metric convention (percentage, rating,
# score out of 100). This is a widely recognized default, not a value
# tuned to any particular workbook.
NUMBER_MIN_VALUE = 0
NUMBER_MAX_VALUE = 100

# "amount": generic non-negative monetary/quantity convention. Real
# amounts (salaries, prices, totals) don't have a single universal
# upper bound, so the only condition enforced generically is that the
# value is a real number and is not negative.
AMOUNT_MIN_VALUE = 0

# "date": rather than assuming or configuring one fixed format, the
# engine tries several common, real-world date formats and accepts
# the value if ANY of them parse it. This keeps date validation fully
# generic - a workbook using DD/MM/YYYY works exactly the same as one
# using ISO 8601, with no configuration change.
GENERIC_DATE_FORMATS = [
    "%Y-%m-%d",       # 2023-01-15  (ISO 8601)
    "%d-%m-%Y",       # 15-01-2023
    "%m/%d/%Y",       # 01/15/2023
    "%d/%m/%Y",       # 15/01/2023
    "%Y/%m/%d",       # 2023/01/15
    "%m-%d-%Y",       # 01-15-2023
    "%B %d, %Y",      # January 15, 2023
    "%d %B %Y",       # 15 January 2023
]


# --------------------------------------------------------------------------
# The seven field-level validators. Signatures take only the value -
# no externally supplied parameters - because their conditions are
# now fixed, generic Python logic rather than configuration.
# --------------------------------------------------------------------------

def validate_numeric_only(value):
    """Return True if value consists only of digits (0-9)."""
    value_str = str(value).strip()
    return bool(re.match(r'^\d+$', value_str))


def validate_email(value):
    """Return True if value matches a basic email address pattern."""
    value_str = str(value).strip()
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(email_regex, value_str))


def validate_phone_number(value):
    """Return True if value matches a common US phone number format."""
    value_str = str(value).strip()
    return bool(re.match(r'^(\(\d{3}\)\s?|\d{3}[-\s]?)\d{3}[-\s]?\d{4}$', value_str))


def validate_amount(value):
    """Return True if value is a non-negative real number (generic
    monetary/quantity convention - see AMOUNT_MIN_VALUE)."""
    try:
        num = float(value)
        return num >= AMOUNT_MIN_VALUE
    except (ValueError, TypeError):
        return False


def validate_number(value):
    """Return True if value is a real number within the generic
    bounded range (see NUMBER_MIN_VALUE / NUMBER_MAX_VALUE)."""
    try:
        num = float(value)
        return NUMBER_MIN_VALUE <= num <= NUMBER_MAX_VALUE
    except (ValueError, TypeError):
        return False


def validate_date(value):
    """Return True if value parses under any of the generic, commonly
    used date formats (see GENERIC_DATE_FORMATS)."""
    if isinstance(value, pd.Timestamp):
        # Already a real date/time - unambiguous, always valid.
        return True

    value_str = str(value).strip()
    for date_format in GENERIC_DATE_FORMATS:
        try:
            datetime.strptime(value_str, date_format)
            return True
        except ValueError:
            continue
    return False


def validate_text_only(value):
    """Return True if value contains only alphabetic characters and spaces."""
    value_str = str(value).strip()
    return value_str.isalpha() or all(
        c.isalpha() or c.isspace()
        for c in value_str
    )


_ERROR_MESSAGES = {
    "numeric_only": "Non-numeric value in column '{column}'",
    "email": "Invalid email in column '{column}'",
    "phone_number": "Invalid phone number in column '{column}'",
    "amount": "Invalid amount in column '{column}'",
    "number": "Out of range in column '{column}'",
    "date": "Invalid date in column '{column}'",
    "text_only": "Non-text value in column '{column}'",
}

_REQUIRED_DATATYPE_LABELS = {
    "numeric_only": "numeric",
    "email": "email",
    "phone_number": "phone_number",
    "amount": "amount",
    "number": "number",
    "date": "date",
    "text_only": "text_only",
}

# Generic, fixed evaluation order used for RULE DETERMINATION - not
# specific to any workbook. Structurally distinctive patterns (email,
# phone, date) are tried first since they're the least ambiguous.
# Among the purely numeric rules, the NARROWEST pattern is tried
# first (numeric_only: a strict positive-integer digit string) before
# broader ones (number: bounded 0-100, then amount: any non-negative
# number) - this keeps every rule reachable, since a broader numeric
# rule would otherwise always "win" over a narrower one it structurally
# subsumes. text_only is the final generic fallback for word-like data.
_RULE_PRIORITY = [
    "email", "phone_number", "date",
    "numeric_only", "number", "amount",
    "text_only",
]


def _is_blank(value):
    """Blank/null values are skipped, matching Dalton's original sample."""
    return pd.isna(value) or value == ""


def _summary_row(sheet_name, field, test_type, test_name, status, failed_count):
    return {
        "Sheet Name": sheet_name,
        "Field": field,
        "Test Type": test_type,
        "Test Name": test_name,
        "Status": status,
        "Failed Count": failed_count,
    }


def _check_value(value, rule_name):
    """Apply a single generic rule to a single value. Returns True/False."""
    if rule_name == "numeric_only":
        return validate_numeric_only(value)
    if rule_name == "email":
        return validate_email(value)
    if rule_name == "phone_number":
        return validate_phone_number(value)
    if rule_name == "text_only":
        return validate_text_only(value)
    if rule_name == "amount":
        return validate_amount(value)
    if rule_name == "number":
        return validate_number(value)
    if rule_name == "date":
        return validate_date(value)
    return False


# --------------------------------------------------------------------------
# Layer 1: Sheet Existence Testing
# --------------------------------------------------------------------------

def validate_sheet_existence(config, available_sheet_names):
    """
    Confirm every configured sheet is present in the workbook.

    Returns
    -------
    summary_rows : list[dict]
    missing_sheets : set[str]
    """
    sheet_checks = config.get("sheet_checks", {})
    summary_rows = []
    missing_sheets = set()

    for sheet_name, rules in sheet_checks.items():
        is_required = rules.get("required", True)
        present = sheet_name in available_sheet_names

        if present:
            status = "PASS"
            failed_count = 0
        elif is_required:
            status = "FAIL"
            failed_count = 1
            missing_sheets.add(sheet_name)
        else:
            status = "SKIPPED"
            failed_count = 0

        summary_rows.append(_summary_row(
            sheet_name, "Sheet", "Sheet Validation",
            "Sheet Existence Testing", status, failed_count,
        ))

    return summary_rows, missing_sheets


# --------------------------------------------------------------------------
# Layer 2: Dynamic Column Discovery + Generic Rule Determination
# --------------------------------------------------------------------------

def discover_columns(df):
    """
    Dynamically discover every column present on a sheet's DataFrame.

    This is the ENTIRE "field discovery" mechanism - it simply reads
    whatever header row pandas already parsed. No expected-column
    list is consulted or maintained anywhere.
    """
    return list(df.columns)


def determine_rule_for_column(series, match_threshold, sample_size):
    """
    Inspect the actual data in a column and determine which generic
    validation rule (if any) its content matches.

    Parameters
    ----------
    series : pandas.Series
        The column's data.
    match_threshold : float
        Minimum fraction of sampled non-blank values that must
        satisfy a rule for that rule to be assigned.
    sample_size : int
        Maximum number of non-blank values inspected per column, so
        very large sheets remain fast to classify.

    Returns
    -------
    rule_name : str or None
        The assigned rule, or None if nothing reached the threshold
        (column is then marked Unrecognized).
    match_rate : float
        The best match rate observed (0.0 if no non-blank data).
    sample_count : int
        How many non-blank values were actually sampled.
    """
    non_blank = [v for v in series if not _is_blank(v)]
    if not non_blank:
        return None, 0.0, 0

    sample = non_blank[:sample_size]
    sample_count = len(sample)

    best_rate = 0.0

    for rule_name in _RULE_PRIORITY:
        matches = sum(1 for value in sample if _check_value(value, rule_name))
        rate = matches / sample_count

        if rate >= match_threshold:
            # First rule (in priority order) to clear the threshold
            # wins - deterministic and generic, favoring the most
            # structurally specific/narrow rule that fits.
            return rule_name, rate, sample_count

        best_rate = max(best_rate, rate)

    return None, best_rate, sample_count


def discover_and_determine_rules(config, sheets_data, missing_sheets):
    """
    For every present sheet, dynamically discover its columns and
    determine which generic rule (if any) applies to each one.

    Returns
    -------
    summary_rows : list[dict]
        One "Rule Determination" test per discovered column.
    column_rules : dict[str, dict[str, str]]
        sheet_name -> {column_name: rule_name}, only for columns that
        were assigned a recognized rule.
    """
    detection_config = config.get("rule_detection", {})
    match_threshold = detection_config.get("match_threshold", DEFAULT_MATCH_THRESHOLD)
    sample_size = detection_config.get("sample_size", DEFAULT_SAMPLE_SIZE)

    summary_rows = []
    column_rules = {}

    for sheet_name, df in sheets_data.items():
        if sheet_name in missing_sheets:
            continue

        discovered_columns = discover_columns(df)

        for column_name in discovered_columns:
            rule_name, match_rate, sample_count = determine_rule_for_column(
                df[column_name], match_threshold, sample_size
            )

            if rule_name is not None:
                status = "PASS"
                test_name = rule_name
                column_rules.setdefault(sheet_name, {})[column_name] = rule_name
            elif sample_count == 0:
                status = "SKIPPED"
                test_name = f"{UNRECOGNIZED} (no data)"
            else:
                status = "SKIPPED"
                test_name = f"{UNRECOGNIZED} (best match {match_rate:.0%})"

            summary_rows.append(_summary_row(
                sheet_name, column_name, "Rule Determination",
                test_name, status, 0,
            ))

    return summary_rows, column_rules


# --------------------------------------------------------------------------
# Layer 3: Data Validation (the seven generic checks)
# --------------------------------------------------------------------------

def _run_field_check(df, sheet_name, column_name, rule_name):
    """
    Apply one validation rule to one column of one sheet's DataFrame.

    Returns
    -------
    total_checked : int
        Non-blank cells actually evaluated.
    failures : list[dict]
        Detailed failure records (one per failing cell).
    """
    total_checked = 0
    failures = []

    for row_index, value in df[column_name].items():
        if _is_blank(value):
            continue

        total_checked += 1

        if not _check_value(value, rule_name):
            failures.append({
                "Sheet Name": sheet_name,
                "Field": column_name,
                "Row": row_index + 2,  # +1 for 0-index, +1 for header row
                "Value": value,
                "Validation Type": rule_name,
                "Required Datatype": _REQUIRED_DATATYPE_LABELS[rule_name],
                "Error Message": _ERROR_MESSAGES[rule_name].format(column=column_name),
            })

    return total_checked, failures


def validate_data(sheets_data, column_rules):
    """
    Run data validation for every column that was assigned a rule in
    the discovery/determination layer.

    Returns
    -------
    summary_rows : list[dict]
    detailed_failures : list[dict]
    """
    summary_rows = []
    detailed_failures = []

    for sheet_name, field_rules in column_rules.items():
        df = sheets_data.get(sheet_name)
        if df is None:
            continue

        for column_name, rule_name in field_rules.items():
            total_checked, failures = _run_field_check(
                df, sheet_name, column_name, rule_name
            )

            failed_count = len(failures)
            status = "PASS" if failed_count == 0 else "FAIL"

            summary_rows.append({
                "Sheet Name": sheet_name,
                "Field": column_name,
                "Test Type": "Data Validation",
                "Test Name": rule_name,
                "Status": status,
                "Failed Count": failed_count,
                "Total Checked": total_checked,
                "Passed Count": total_checked - failed_count,
            })

            detailed_failures.extend(failures)

    return summary_rows, detailed_failures


# --------------------------------------------------------------------------
# Orchestrator
# --------------------------------------------------------------------------

def run_validations(config, workbook_sheet_names, sheets_data):
    """
    Run all three validation layers in order and return combined,
    structured results.

    Parameters
    ----------
    config : dict
        Full JSON configuration (sheet_checks, rule_detection only -
        no field/column mapping, no validation conditions).
    workbook_sheet_names : list[str]
        Sheet names actually present in the input workbook.
    sheets_data : dict[str, pandas.DataFrame]
        Sheet name -> DataFrame, for every sheet in the workbook. The
        columns of each DataFrame are exactly the header row pandas
        found in that sheet - this IS the dynamic field discovery.

    Returns
    -------
    summary_rows : list[dict]
        Every test that was executed (sheet existence, rule
        determination, and data validation), pass, fail, and skipped
        alike.
    detailed_failures : list[dict]
        One row per individual failing cell.
    column_rules : dict[str, dict[str, str]]
        sheet_name -> {column_name: rule_name}, the rules that were
        dynamically assigned - exposed so callers (e.g. main.py) can
        compute accurate row-processing metrics without needing a
        maintained field list of their own.
    """
    sheet_summary, missing_sheets = validate_sheet_existence(config, workbook_sheet_names)

    rule_summary, column_rules = discover_and_determine_rules(
        config, sheets_data, missing_sheets
    )

    data_summary, detailed_failures = validate_data(sheets_data, column_rules)

    summary_rows = sheet_summary + rule_summary + data_summary
    return summary_rows, detailed_failures, column_rules
