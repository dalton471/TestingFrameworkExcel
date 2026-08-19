"""
validation.py
--------------
Generic, configuration-driven validation engine for the Validation
Automation Framework.

Architecture mirrors the previous KPI Automation Framework:
    - Every validator reads its scope (sheets, fields, rules) from the
      JSON configuration. No sheet or column names are hardcoded here.
    - Every validator returns a list of flat "summary rows" using the
      same shape used across the whole framework:
          Sheet Name, Field, Test Type, Test Name, Status, Failed Count
      This means the report shows EVERY validation that ran, not only
      the ones that failed.
    - Cell-level failures are additionally captured as "detailed
      failure records" for the Failed Records report sheet.

Three validation layers, run in this order:
    1. Sheet Existence Testing   - is each configured sheet present?
    2. Field Existence Testing   - does each configured column exist
                                    on its sheet?
    3. Data Validation           - Dalton's seven field-level checks,
                                    applied cell by cell.

Field-level validators (Dalton's original logic, behavior unchanged):
    - validate_numeric_only()
    - validate_email()
    - validate_phone_number()
    - validate_amount()
    - validate_number()
    - validate_date()
    - validate_text_only()

INTENTIONALLY EXCLUDED: API validation. Dalton confirmed it is out of
scope for this version - no API URL, token, fetch_valid_codes(), or
API calls exist anywhere in this module.
"""

import re
from datetime import datetime

import pandas as pd



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


_FIELD_VALIDATORS = {
    "numeric_only": validate_numeric_only,
    "email": validate_email,
    "phone_number": validate_phone_number,
    "text_only": validate_text_only,
}

_RANGE_VALIDATORS = {
    "amount": validate_amount,
    "number": validate_number,
}

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
            # Configured but not required and not present - not a
            # failure, simply skipped.
            status = "SKIPPED"
            failed_count = 0

        summary_rows.append(_summary_row(
            sheet_name, "Sheet", "Sheet Validation",
            "Sheet Existence Testing", status, failed_count,
        ))

    return summary_rows, missing_sheets


# --------------------------------------------------------------------------
# Layer 2: Field Existence Testing
# --------------------------------------------------------------------------

def validate_field_existence(config, sheets_data, missing_sheets):
    """
    For every sheet with configured fields, confirm each configured
    column exists in that sheet's DataFrame.

    Returns
    -------
    summary_rows : list[dict]
    missing_fields : dict[str, set[str]]
        sheet_name -> set of column names missing from that sheet.
    """
    fields_config = config.get("fields", {})
    summary_rows = []
    missing_fields = {}

    for sheet_name, field_map in fields_config.items():
        if sheet_name in missing_sheets:
            # Sheet itself is missing - field existence cannot be
            # tested; already reported at the sheet layer.
            continue

        df = sheets_data.get(sheet_name)
        if df is None:
            continue

        for column_name in field_map.keys():
            present = column_name in df.columns
            status = "PASS" if present else "FAIL"
            failed_count = 0 if present else 1

            if not present:
                missing_fields.setdefault(sheet_name, set()).add(column_name)

            summary_rows.append(_summary_row(
                sheet_name, column_name, "Field Validation",
                "Field Existence Testing", status, failed_count,
            ))

    return summary_rows, missing_fields


# --------------------------------------------------------------------------
# Layer 3: Data Validation (Dalton's seven checks)
# --------------------------------------------------------------------------

def _run_field_check(df, sheet_name, column_name, datatype, rules):
    """
    Apply one validation type to one column of one sheet's DataFrame.

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
        is_valid = True

        if datatype in _FIELD_VALIDATORS:
            is_valid = _FIELD_VALIDATORS[datatype](value)

        elif datatype in _RANGE_VALIDATORS:
            min_value = rules.get("min_value", float("-inf"))
            max_value = rules.get("max_value", float("inf"))
            is_valid = _RANGE_VALIDATORS[datatype](value, min_value, max_value)

        elif datatype == "date":
            date_format = rules.get("format")
            is_valid = validate_date(value, date_format)

        else:
            # Unknown datatype - nothing to check against.
            continue

        if not is_valid:
            failures.append({
                "Sheet Name": sheet_name,
                "Field": column_name,
                "Row": row_index + 2,  # +1 for 0-index, +1 for header row
                "Value": value,
                "Validation Type": datatype,
                "Required Datatype": _REQUIRED_DATATYPE_LABELS[datatype],
                "Error Message": _ERROR_MESSAGES[datatype].format(column=column_name),
            })

    return total_checked, failures


def validate_data(config, sheets_data, missing_sheets, missing_fields):
    """
    Run every configured (sheet, field, datatype) data validation.

    Returns
    -------
    summary_rows : list[dict]
    detailed_failures : list[dict]
    """
    fields_config = config.get("fields", {})
    column_checks = config.get("column_checks", {})

    summary_rows = []
    detailed_failures = []

    for sheet_name, field_map in fields_config.items():
        if sheet_name in missing_sheets:
            continue

        df = sheets_data.get(sheet_name)
        if df is None:
            continue

        sheet_missing_fields = missing_fields.get(sheet_name, set())

        for column_name, datatypes in field_map.items():
            if column_name in sheet_missing_fields:
                # Column doesn't exist - already reported at the field
                # existence layer; data validation cannot run on it.
                continue

            for datatype in datatypes:
                if datatype not in column_checks:
                    continue

                rules = column_checks[datatype]
                total_checked, failures = _run_field_check(
                    df, sheet_name, column_name, datatype, rules
                )

                failed_count = len(failures)
                status = "PASS" if failed_count == 0 else "FAIL"

                summary_rows.append({
                    "Sheet Name": sheet_name,
                    "Field": column_name,
                    "Test Type": "Data Validation",
                    "Test Name": datatype,
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
        Full JSON configuration.
    workbook_sheet_names : list[str]
        Sheet names actually present in the input workbook.
    sheets_data : dict[str, pandas.DataFrame]
        Sheet name -> DataFrame, for every sheet in the workbook.

    Returns
    -------
    summary_rows : list[dict]
        Every test that was executed (sheet existence, field
        existence, and data validation), pass and fail alike.
    detailed_failures : list[dict]
        One row per individual failing cell.
    """
    sheet_summary, missing_sheets = validate_sheet_existence(config, workbook_sheet_names)

    field_summary, missing_fields = validate_field_existence(
        config, sheets_data, missing_sheets
    )

    data_summary, detailed_failures = validate_data(
        config, sheets_data, missing_sheets, missing_fields
    )

    summary_rows = sheet_summary + field_summary + data_summary
    return summary_rows, detailed_failures
