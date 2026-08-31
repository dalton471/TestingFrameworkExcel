"""
main.py
--------
Entry point of the Validation Automation Framework.

This file does NOT contain any validation logic itself. Its only job
is to ORCHESTRATE the pipeline:

    1. Parse command-line arguments.
    2. Validate the input workbook path (always required).
    3. Load the input workbook and every sheet inside it (this happens
       BEFORE config loading, because the built-in fallback config -
       used when JSON is absent - needs the workbook's own discovered
       sheet names).
    4. Load validation_config.json IF IT EXISTS. If it does not exist,
       build the built-in default configuration instead
       (validations.validation.build_default_config) rather than
       treating a missing JSON file as an error.
    5. Run all three validation layers on whichever configuration was
       obtained in step 4 - Sheet Existence Testing, Rule
       Determination, Data Validation - via the SAME
       validations.validation.run_validations() used in both cases.
       There is only one validation engine; JSON-present and
       JSON-absent runs both call it identically.
    6. Generate the three-sheet Validation_Report.xlsx.
    7. Print structured, readable terminal output throughout.
    8. Print a final run summary and exit with a meaningful status code.

Example (JSON present):
    python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"

Example (JSON absent or omitted - built-in defaults are used automatically):
    python main.py --input "input/Validation_Sample.xlsx"
"""

import argparse
import json
import os
import sys
import traceback

from utils.helper import load_json
from utils.excel_reader import load_workbook, load_all_sheets, get_sheet_names
from utils.report_generator import generate_report
from validations.validation import run_validations, build_default_config

DEFAULT_OUTPUT_PATH = os.path.join("output", "Validation_Report.xlsx")
DEFAULT_CONFIG_PATH = os.path.join("config", "validation_config.json")
LINE = "=" * 60


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Validation Automation Framework - dynamically "
                     "validates a multi-sheet Excel workbook and "
                     "produces a Validation_Report.xlsx. A JSON "
                     "configuration file is optional; if omitted or "
                     "not found, built-in default rules are used."
    )
    parser.add_argument(
        "--config", required=False, default=DEFAULT_CONFIG_PATH,
        help=f"Path to the JSON configuration file. Optional - if the "
             f"file is not found, the framework falls back to a "
             f"built-in default configuration. Defaults to "
             f"'{DEFAULT_CONFIG_PATH}'"
    )
    parser.add_argument(
        "--input", required=True,
        help="Path to the input Excel workbook "
             "(e.g. input/Validation_Sample.xlsx)"
    )
    parser.add_argument(
        "--output", required=False, default=DEFAULT_OUTPUT_PATH,
        help=f"Path for the generated report. Defaults to '{DEFAULT_OUTPUT_PATH}'"
    )
    return parser.parse_args()


def validate_input_path(input_path):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: '{input_path}'")
    if not os.path.isfile(input_path):
        raise ValueError(f"Input path is not a file: '{input_path}'")


def load_config_or_fallback(config_path, workbook_sheet_names):
    """
    Load validation_config.json if it exists; otherwise build the
    built-in default configuration from the workbook's own discovered
    sheets. Returns (config, used_fallback: bool).

    This is the ONLY branch point between the two scenarios in the
    entire program - everything downstream (run_validations, report
    generation, terminal output) is identical either way.
    """
    if os.path.exists(config_path) and os.path.isfile(config_path):
        return load_json(config_path), False

    return build_default_config(workbook_sheet_names), True


def print_validation_results(summary_rows):
    print(f"\n{LINE}")
    print("VALIDATION RESULTS")
    print(LINE)

    for row in summary_rows:
        print(
            f"{row['Sheet Name']} | {row['Field']} | "
            f"{row['Test Name']} | {row['Status']} | "
            f"Failed Count: {row['Failed Count']}"
        )

    print(LINE)
    print("VALIDATION COMPLETED")
    print(LINE)


def build_execution_stats(summary_rows, sheets_data, column_rules):
    """
    Compute run-level totals purely from what was dynamically
    discovered and validated this run - never from a maintained field
    list (there isn't one).

    Column discovery/recognition counts are computed directly from
    sheets_data and column_rules rather than from summary_rows, since
    unrecognized columns intentionally produce no summary row at all
    (Task 2B - no SKIPPED status anywhere in the report).
    """
    sheet_rows = [r for r in summary_rows if r["Test Name"] == "Sheet Existence Testing"]
    data_rows = [r for r in summary_rows if r["Test Type"] == "Data Validation"]

    total_sheets = len(sheet_rows)
    sheets_passed = sum(1 for r in sheet_rows if r["Status"] == "PASS")
    sheets_failed = sum(1 for r in sheet_rows if r["Status"] == "FAIL")

    total_columns_discovered = sum(len(df.columns) for df in sheets_data.values())
    total_columns_recognized = sum(len(fields) for fields in column_rules.values())
    total_columns_unrecognized = total_columns_discovered - total_columns_recognized

    # Total rows processed = rows in every sheet that had at least one
    # column dynamically assigned a recognized rule. Sheets that are
    # purely descriptive (no column matches any generic rule) are
    # naturally excluded, with no exclusion list to maintain.
    total_rows_processed = sum(
        len(sheets_data[sheet_name])
        for sheet_name in column_rules
        if sheet_name in sheets_data
    )

    total_validations = len(summary_rows)
    total_passed = sum(1 for r in summary_rows if r["Status"] == "PASS")
    total_failed = sum(1 for r in summary_rows if r["Status"] == "FAIL")
    total_failed_cells = sum(r["Failed Count"] for r in data_rows)

    execution_status = "SUCCESS" if sheets_failed == 0 else "COMPLETED WITH FAILURES"

    return {
        "Total Sheets": total_sheets,
        "Sheets Passed": sheets_passed,
        "Sheets Failed": sheets_failed,
        "Total Columns Discovered": total_columns_discovered,
        "Total Columns Recognized": total_columns_recognized,
        "Total Columns Unrecognized": total_columns_unrecognized,
        "Total Rows Processed": total_rows_processed,
        "Total Validations Executed": total_validations,
        "Total Passed": total_passed,
        "Total Failed": total_failed,
        "Total Failed Cells": total_failed_cells,
        "Execution Status": execution_status,
    }


def print_run_summary(stats):
    print(f"\n{LINE}")
    print("RUN SUMMARY")
    print(LINE)
    for label, value in stats.items():
        print(f"{label}: {value}")
    print(LINE)


def main():
    print(LINE)
    print("VALIDATION FRAMEWORK")
    print(LINE)

    args = parse_arguments()

    try:
        validate_input_path(args.input)
    except (FileNotFoundError, ValueError) as error:
        print(f"[ERROR] {error}")
        sys.exit(1)

    print(f"\n{LINE}")
    print(f"PROCESSING: {args.input}")
    print(LINE)

    workbook = None
    try:
        workbook = load_workbook(args.input)
        workbook_sheet_names = get_sheet_names(workbook)
        sheets_data = load_all_sheets(workbook)
        print("Excel Loaded Successfully")
        print(f"Sheets found in workbook: {', '.join(workbook_sheet_names)}")
    except (FileNotFoundError, ValueError) as error:
        print(f"[ERROR] {error}")
        sys.exit(1)
    except Exception as error:
        print(f"[ERROR] Unexpected error while loading input file: {error}")
        traceback.print_exc()
        sys.exit(1)
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass

    try:
        config, used_fallback = load_config_or_fallback(args.config, workbook_sheet_names)
    except json.JSONDecodeError as error:
        print(f"[ERROR] Config file is not valid JSON: '{args.config}' -> {error}")
        sys.exit(1)
    except Exception as error:
        print(f"[ERROR] Unexpected error while loading config: {error}")
        traceback.print_exc()
        sys.exit(1)

    if used_fallback:
        print(f"\nConfig file not found at '{args.config}'")
        print("Using built-in default configuration (no JSON required)")
        print(f"Sheets discovered directly from workbook: {', '.join(workbook_sheet_names)}")
    else:
        print(f"\nJSON configuration loaded from '{args.config}'")

    if not isinstance(config, dict) or not config:
        print(f"[ERROR] Config is empty or not a valid object.")
        sys.exit(1)

    try:
        summary_rows, detailed_failures, column_rules = run_validations(
            config, workbook_sheet_names, sheets_data
        )
    except Exception as error:
        print(f"[ERROR] Unexpected error during validation: {error}")
        traceback.print_exc()
        sys.exit(1)

    print_validation_results(summary_rows)

    execution_stats = build_execution_stats(summary_rows, sheets_data, column_rules)

    try:
        generate_report(config, summary_rows, detailed_failures, execution_stats, args.output)
        print(f"\n{args.output} created successfully.")
    except Exception as error:
        print(f"[ERROR] Unexpected error while generating report: {error}")
        traceback.print_exc()
        sys.exit(1)

    print_run_summary(execution_stats)

    if execution_stats["Sheets Failed"] > 0:
        sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    main()
