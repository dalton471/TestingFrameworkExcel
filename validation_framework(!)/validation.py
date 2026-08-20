"""
validation.py
--------------
Generic, configuration-driven validation engine for the Validation
Automation Framework.

ARCHITECTURE (dynamic discovery - no maintained field list):

    Generic JSON Rules  ->  Excel Workbook  ->  Dynamic Sheet Discovery
        ->  Dynamic Column Discovery  ->  Generic Rule Determination
        ->  Validation Engine  ->  Report

JSON (config/validation_config.json) defines only:
    - which sheets are expected (sheet_checks)
    - which generic validation rules exist and their parameters
      (column_checks: numeric_only, email, phone_number, amount,
      number, date, text_only)

There is NO configured list of column names anywhere. Every column in
every discovered sheet is found by reading the sheet's header row at
runtime (pandas.read_excel already does this), and the validation
rule that applies to that column is determined purely from the
CONTENT of its cells - never from the column's name and never from a
sample-specific lookup table. This means a renamed column, a new
column, or a completely different workbook layout is handled without
any code or JSON change.

Three validation layers, run in this order:
    1. Sheet Existence Testing     - is each configured sheet present?
    2. Rule Determination          - for every column discovered on
                                       every present sheet, which (if
                                       any) generic rule does its data
                                       match?
    3. Data Validation             - Dalton's seven field-level
                                       checks, applied cell by cell,
                                       using the rule assigned in
                                       step 2.

Field-level validators (Dalton's original logic, behavior unchanged):
    - validate_numeric_only()
    - validate_email()
    - validate_phone_number()
    - validate_amount()
    - validate_number()
    - validate_date()
    - validate_text_only()

RULE DETERMINATION ALGORITHM (see README.md "Dynamic Field Discovery"
and "How Validation Rules Are Determined" for the full explanation):

    For each discovered column, take a sample of its non-blank values
    (capped by config["rule_detection"]["sample_size"] for
    scalability on very large sheets) and compute, for every generic
    rule in column_checks, the fraction of sampled values that rule
    would accept. Rules are evaluated in a fixed, generic priority
    order (most structurally specific first: email, phone_number,
    date, amount, number, numeric_only, text_only) and the FIRST rule
    whose match rate reaches the configured threshold
    (rule_detection.match_threshold, default 0.6) is assigned to the
    column. If no rule reaches the threshold, the column is marked
    "Unrecognized" and is intentionally excluded from data validation
    - it is never guessed at.

    This uses only generic signals (regex structure, numeric parse
    success, configured min/max ranges, date parseability) - never a
    column name or a sample-workbook-specific mapping.

INTENTIONALLY EXCLUDED: API validation. Dalton confirmed it is out of
scope for this version - no API URL, token, fetch_valid_codes(), or
API calls exist anywhere in this module.
"""

import re
from datetime import datetime

import pandas as pd

DEFAULT_MATCH_THRESHOLD = 0.6
DEFAULT_SAMPLE_SIZE = 200

UNRECOGNIZED = "Unrecognized"


# --------------------------------------------------------------------------
# Dalton's original field-level validators - behavior unchanged
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


def validate_amount(value, min_value, max_value):
    """Return True if value is numeric and within [min_value, max_value]."""
    try:
        num = float(value)
        return min_value <= num <= max_value
    except ValueError:
        return False


def validate_number(value, min_value, max_value):
    """Return True if value is numeric and within [min_value, max_value]."""
    try:
        num = float(value)
        return min_value <= num <= max_value
    except ValueError:
        return False


def validate_date(value, date_format):
    """Return True if value can be parsed using the configured date_format."""
    if isinstance(value, pd.Timestamp):
        value = value.strftime('%Y-%m-%d %H:%M:%S')
    else:
        value = str(value)

    try:
        datetime.strptime(value, date_format)
        return True
    except ValueError:
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

# Generic, fixed evaluation order used for RULE DETERMINATION (not
# specific to any workbook). More structurally specific patterns are
# tried first so that, for example, an email address is never
# mistaken for generic text, and a bounded score column is not
# mistaken for a free-form numeric identifier. This order applies to
# ANY workbook - it contains no sample-specific field names.
_RULE_PRIORITY = [
    "email", "phone_number", "date", "amount", "number",
    "numeric_only", "text_only",
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


def _check_value(value, rule_name, rules):
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
        min_value = rules.get("min_value", float("-inf"))
        max_value = rules.get("max_value", float("inf"))
        return validate_amount(value, min_value, max_value)
    if rule_name == "number":
        min_value = rules.get("min_value", float("-inf"))
        max_value = rules.get("max_value", float("inf"))
        return validate_number(value, min_value, max_value)
    if rule_name == "date":
        date_format = rules.get("format")
        return validate_date(value, date_format)
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


def determine_rule_for_column(series, column_checks, match_threshold, sample_size):
    """
    Inspect the actual data in a column and determine which generic
    validation rule (if any) its content matches.

    Parameters
    ----------
    series : pandas.Series
        The column's data.
    column_checks : dict
        The "column_checks" section of the JSON config (rule name ->
        parameters).
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
        The match rate achieved by the assigned rule (0.0 if None).
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
        if rule_name not in column_checks:
            continue

        rules = column_checks[rule_name]
        matches = sum(1 for value in sample if _check_value(value, rule_name, rules))
        rate = matches / sample_count

        if rate >= match_threshold:
            # First rule (in priority order) to clear the threshold
            # wins - this keeps the mechanism deterministic and
            # generic, favoring more structurally specific rules.
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
    column_checks = config.get("column_checks", {})
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
                df[column_name], column_checks, match_threshold, sample_size
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
# Layer 3: Data Validation (Dalton's seven checks)
# --------------------------------------------------------------------------

def _run_field_check(df, sheet_name, column_name, rule_name, rules):
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

        if not _check_value(value, rule_name, rules):
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


def validate_data(config, sheets_data, column_rules):
    """
    Run data validation for every column that was assigned a rule in
    the discovery/determination layer.

    Returns
    -------
    summary_rows : list[dict]
    detailed_failures : list[dict]
    """
    column_checks = config.get("column_checks", {})

    summary_rows = []
    detailed_failures = []

    for sheet_name, field_rules in column_rules.items():
        df = sheets_data.get(sheet_name)
        if df is None:
            continue

        for column_name, rule_name in field_rules.items():
            rules = column_checks.get(rule_name, {})
            total_checked, failures = _run_field_check(
                df, sheet_name, column_name, rule_name, rules
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
        Full JSON configuration (sheet_checks, column_checks,
        rule_detection - no field/column mapping).
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

    data_summary, detailed_failures = validate_data(config, sheets_data, column_rules)

    summary_rows = sheet_summary + rule_summary + data_summary
    return summary_rows, detailed_failures, column_rules
