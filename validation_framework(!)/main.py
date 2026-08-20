"""
main.py
--------
Entry point of the Validation Automation Framework.

This file does NOT contain any validation logic itself. Its only job
is to ORCHESTRATE the pipeline, matching the architecture of the
previous KPI Automation Framework:

    1. Parse command-line arguments.
    2. Validate config/input paths.
    3. Load the JSON configuration.
    4. Load the input workbook and every sheet inside it.
    5. Run all three validation layers (validations/validation.py):
         - Sheet Existence Testing
         - Rule Determination (dynamic column discovery + generic
           content-based rule assignment - no maintained field list)
         - Data Validation (Dalton's seven checks)
    6. Generate the three-sheet Validation_Report.xlsx.
    7. Print structured, readable terminal output throughout.
    8. Print a final run summary and exit with a meaningful status code.

Example:
    python main.py --config config/validation_config.json --input "input/Validation_Sample.xlsx"
"""

import argparse
import json
import os
import sys
import traceback

from utils.helper import load_json
from utils.excel_reader import load_workbook, load_all_sheets, get_sheet_names
from utils.report_generator import generate_report
from validations.validation import run_validations

DEFAULT_OUTPUT_PATH = os.path.join("output", "Validation_Report.xlsx")
LINE = "=" * 60


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Validation Automation Framework - validates a "
                     "multi-sheet Excel workbook against a JSON rule "
                     "set and produces a Validation_Report.xlsx"
    )
    parser.add_argument(
        "--config", required=True,
        help="Path to the JSON configuration file "
             "(e.g. config/validation_config.json)"
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


def validate_config_path(config_path):
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found: '{config_path}'")
    if not os.path.isfile(config_path):
        raise ValueError(f"Config path is not a file: '{config_path}'")


def validate_input_path(input_path):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: '{input_path}'")
    if not os.path.isfile(input_path):
        raise ValueError(f"Input path is not a file: '{input_path}'")


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
    """
    sheet_rows = [r for r in summary_rows if r["Test Name"] == "Sheet Existence Testing"]
    rule_rows = [r for r in summary_rows if r["Test Type"] == "Rule Determination"]
    data_rows = [r for r in summary_rows if r["Test Type"] == "Data Validation"]

    total_sheets = len(sheet_rows)
    sheets_passed = sum(1 for r in sheet_rows if r["Status"] == "PASS")
    sheets_failed = sum(1 for r in sheet_rows if r["Status"] == "FAIL")

    total_columns_discovered = len(rule_rows)
    total_columns_recognized = sum(1 for r in rule_rows if r["Status"] == "PASS")
    total_columns_unrecognized = sum(1 for r in rule_rows if r["Status"] == "SKIPPED")

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
    total_skipped = sum(1 for r in summary_rows if r["Status"] == "SKIPPED")
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
        "Total Skipped": total_skipped,
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
        validate_config_path(args.config)
        validate_input_path(args.input)
    except (FileNotFoundError, ValueError) as error:
        print(f"[ERROR] {error}")
        sys.exit(1)

    try:
        config = load_json(args.config)
        print("\nJSON Loaded Successfully")
    except json.JSONDecodeError as error:
        print(f"[ERROR] Config file is not valid JSON: '{args.config}' -> {error}")
        sys.exit(1)
    except Exception as error:
        print(f"[ERROR] Unexpected error while loading config: {error}")
        traceback.print_exc()
        sys.exit(1)

    if not isinstance(config, dict) or not config:
        print(f"[ERROR] Config file '{args.config}' is empty or not a JSON object.")
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
