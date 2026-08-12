"""
main.py
--------
Entry point of the KPI Validation Automation Framework.

This file does NOT contain any validation logic itself. Its only job
is to ORCHESTRATE the pipeline, in this order:

    1. Load the JSON configuration file.
    2. Load the input file (single Excel/CSV file, or every supported
       file inside an input folder).
    3. Call every validation function (from validations/validations.py).
    4. Collect all the results together.
    5. Generate the Validation_Report.xlsx (via utils/report_generator.py).

All paths are supplied on the command line via argparse:

    python main.py --config config/Kpi_automation_phase1.json --input input/KPI.xlsx
    python main.py --config config/Kpi_automation_phase1.json --input input/ --output output/Validation_Report.xlsx
    python main.py --config config/rules.json --input input/data.csv

Run "python main.py --help" to see all available options.
"""

import os
import sys
import json
import argparse
import traceback

import pandas as pd

from utils.helper import load_json
from utils.excel_reader import load_excel
from utils.report_generator import generate_report

from validations.validations import (
    validate_sheet_list,
    validate_sheet_columns,
    validate_duplicate_check,
    validate_null_check,
    validate_formula_check,
    validate_numeric_precision,
    validate_business_rule,
    validate_reconciliation,
    validate_cross_sheet,
    validate_trend,
)


# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")
DEFAULT_OUTPUT_PATH = os.path.join("output", "Validation_Report.xlsx")

# Each validator is paired with the config section(s) it needs, so a
# missing/invalid section only disables THAT validator instead of
# aborting the whole file's report.
VALIDATOR_STEPS = [
    ("sheetlist", "validate_sheet_list", "Sheet Validation"),
    ("sheetvalidation", "validate_sheet_columns", "Field Validation"),
    ("dataqualityvalidation.duplicatecheck", "validate_duplicate_check", "Duplicate Check"),
    ("dataqualityvalidation.nullcheck", "validate_null_check", "Null Check"),
    ("formulavalidation", "validate_formula_check", "Formula Validation"),
    ("dataqualityvalidation.numericprecision", "validate_numeric_precision", "Numeric Precision"),
    ("businessrulevalidation", "validate_business_rule", "Business Rule Validation"),
    ("reconciliationvalidation", "validate_reconciliation", "Reconciliation Validation"),
    ("crosssheetvalidation", "validate_cross_sheet", "Cross Sheet Validation"),
    ("trendvalidation", "validate_trend", "Trend Validation"),
]


# --------------------------------------------------------------------------
# Argument Parsing
# --------------------------------------------------------------------------

def parse_arguments():
    parser = argparse.ArgumentParser(
        description="KPI Validation Automation Framework - validates "
                     "Excel/CSV data against a JSON rule set and "
                     "produces a Validation_Report.xlsx"
    )
    parser.add_argument(
        "--config", required=True,
        help="Path to the JSON configuration file "
             "(e.g. config/Kpi_automation_phase1.json)"
    )
    parser.add_argument(
        "--input", required=True,
        help="Path to a single input file (.xlsx, .xls, .csv) OR a "
             "folder containing multiple such files "
             "(e.g. input/KPI.xlsx or input/)"
    )
    parser.add_argument(
        "--output", required=False, default=DEFAULT_OUTPUT_PATH,
        help=f"Path for the generated report. Defaults to '{DEFAULT_OUTPUT_PATH}'"
    )
    return parser.parse_args()


# --------------------------------------------------------------------------
# Input Validation / Discovery Helpers
# --------------------------------------------------------------------------

def validate_config_path(config_path):
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found: '{config_path}'")
    if not os.path.isfile(config_path):
        raise ValueError(f"Config path is not a file: '{config_path}'")


def validate_input_path(input_path):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input path not found: '{input_path}'")


def collect_input_files(input_path):
    """
    Accept either:
    1. A single supported input file
    2. A folder containing one or more supported input files

    Folder discovery is recursive, so files inside subfolders are also
    discovered.
    """

    if os.path.isfile(input_path):
        ext = os.path.splitext(input_path)[1].lower()

        if ext not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Unsupported file type '{ext}' for '{input_path}'. "
                f"Supported types are: {', '.join(SUPPORTED_EXTENSIONS)}"
            )

        return [input_path]

    if os.path.isdir(input_path):
        found_files = []

        for root, _, filenames in os.walk(input_path):
            for filename in sorted(filenames):

                # Ignore temporary Excel lock files
                if filename.startswith("~$"):
                    continue

                ext = os.path.splitext(filename)[1].lower()

                if ext in SUPPORTED_EXTENSIONS:
                    found_files.append(
                        os.path.join(root, filename)
                    )

        found_files.sort()

        if not found_files:
            raise ValueError(
                f"No supported files ({', '.join(SUPPORTED_EXTENSIONS)}) "
                f"found inside folder: '{input_path}'"
            )

        return found_files

    raise ValueError(f"Invalid input path: '{input_path}'")


def prepare_excel_path(file_path):
    """
    Excel files pass through unchanged. A CSV is converted into a
    temporary single-sheet .xlsx so it can flow through the same
    validation functions. A CSV has no worksheets, so every
    sheet-based/cross-sheet rule will correctly report a missing
    sheet for anything other than that one logical sheet - this is
    expected, not a bug (documented in the README).
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in (".xlsx", ".xls"):
        return file_path, False

    if ext == ".csv":
        try:
            df = pd.read_csv(file_path)
        except Exception as error:
            raise ValueError(f"Could not read CSV file '{file_path}': {error}")

        os.makedirs("output", exist_ok=True)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        temp_path = os.path.join("output", f"_temp_{base_name}.xlsx")

        try:
            df.to_excel(temp_path, index=False, sheet_name="Sheet1")
        except Exception as error:
            raise ValueError(f"Could not convert CSV to Excel for validation: {error}")

        return temp_path, True

    raise ValueError(f"Unsupported file type '{ext}' for '{file_path}'")


def build_output_path_for_file(output_arg, input_file_path, is_batch):
    if not is_batch:
        return output_arg

    output_dir = os.path.dirname(output_arg) or "."
    base, ext = os.path.splitext(os.path.basename(output_arg))
    input_stem = os.path.splitext(os.path.basename(input_file_path))[0]
    return os.path.join(output_dir, f"{base}_{input_stem}{ext}")


def _get_config_section(config, dotted_path):
    """Walk a dotted path like 'dataqualityvalidation.nullcheck' through
    the config dict. Raises KeyError with the full path if missing."""
    node = config
    for part in dotted_path.split("."):
        node = node[part]  # KeyError propagates with the missing part
    return node


# --------------------------------------------------------------------------
# Core Orchestration
# --------------------------------------------------------------------------

VALIDATOR_FUNCTIONS = {
    "validate_sheet_list": validate_sheet_list,
    "validate_sheet_columns": validate_sheet_columns,
    "validate_duplicate_check": validate_duplicate_check,
    "validate_null_check": validate_null_check,
    "validate_formula_check": validate_formula_check,
    "validate_numeric_precision": validate_numeric_precision,
    "validate_business_rule": validate_business_rule,
    "validate_reconciliation": validate_reconciliation,
    "validate_cross_sheet": validate_cross_sheet,
    "validate_trend": validate_trend,
}

def print_validation_results(results):
    """Print a complete validation summary to the terminal."""

    print("\n" + "=" * 80)
    print("VALIDATION RESULTS")
    print("=" * 80)

    for result in results:
        sheet_name = result.get("Sheet Name", "")
        field = result.get("Field", "")
        test_name = result.get("Test Name", "")
        status = result.get("Status (P/F)", "")
        failed_count = result.get("Failed Count", 0)

        if status == "P":
            print(
                f"{sheet_name} | {field} | "
                f"{test_name} | PASS | Failed Count: {failed_count}"
            )
        else:
            print(
                f"{sheet_name} | {field} | "
                f"{test_name} | FAIL | Failed Count: {failed_count}"
            )

    print("=" * 80)
    print("VALIDATION COMPLETED")
    print("=" * 80)

def run_validation_pipeline(config, excel_file, output_file):
    """
    Runs every validator. A missing/invalid config section for ONE
    validator does not abort the others - it is recorded as a single
    "Configuration Error" row and the rest of the pipeline continues.
    """
    workbook = load_excel(excel_file)
    print("Excel Loaded Successfully")

    final_results = []

    for config_path, func_name, test_name in VALIDATOR_STEPS:
        func = VALIDATOR_FUNCTIONS[func_name]
        try:
            _get_config_section(config, config_path)  # existence check only
            if func_name == "validate_sheet_list":
                section_results = func(config, workbook)
            else:
                section_results = func(config, excel_file)
            final_results.extend(section_results)

        except KeyError as missing_key:
            print(f"[CONFIG ERROR] Missing configuration section '{config_path}' "
                  f"({missing_key}) - skipping {test_name}.")
            final_results.append({
                "Sheet Name": "Configuration",
                "Field": config_path,
                "Test Type": "Configuration Validation",
                "Test Name": test_name,
                "Status (P/F)": "F",
                "Failed Count": 1,
            })
        except Exception as error:
            print(f"[VALIDATOR ERROR] {test_name} failed: {error}")
            traceback.print_exc()
            final_results.append({
                "Sheet Name": "Configuration",
                "Field": test_name,
                "Test Type": "Configuration Validation",
                "Test Name": test_name,
                "Status (P/F)": "F",
                "Failed Count": 1,
            })

    # Print all validation results to terminal
    print_validation_results(final_results)

    output_dir = os.path.dirname(output_file)

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    generate_report(final_results, output_file)

def process_single_file(config, input_file, output_file):
    temp_path = None

    try:
        print(f"\n{'=' * 60}")
        print(f"Processing: {input_file}")
        print(f"{'=' * 60}")

        excel_path, is_temporary = prepare_excel_path(input_file)
        if is_temporary:
            temp_path = excel_path

        run_validation_pipeline(config, excel_path, output_file)

        print(f"Report generated: {output_file}")
        return True

    except FileNotFoundError as error:
        print(f"[FAILED] File not found while processing '{input_file}': {error}")
        return False

    except ValueError as error:
        print(f"[FAILED] Invalid file/path while processing '{input_file}': {error}")
        return False

    except (pd.errors.ParserError, pd.errors.EmptyDataError) as error:
        print(f"[FAILED] Invalid/corrupt Excel or CSV content in '{input_file}': {error}")
        return False

    except Exception as error:
        print(f"[FAILED] Unexpected error while processing '{input_file}': {error}")
        traceback.print_exc()
        return False

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


# --------------------------------------------------------------------------
# Main Entry Point
# --------------------------------------------------------------------------

def main():
    args = parse_arguments()

    try:
        validate_config_path(args.config)
        validate_input_path(args.input)
    except (FileNotFoundError, ValueError) as error:
        print(f"[ERROR] {error}")
        sys.exit(1)

    try:
        config = load_json(args.config)
        print("JSON Loaded Successfully")
    except json.JSONDecodeError as error:
        print(f"[ERROR] Config file is not valid JSON: '{args.config}' -> {error}")
        sys.exit(1)
    except FileNotFoundError:
        print(f"[ERROR] Config file not found: '{args.config}'")
        sys.exit(1)
    except Exception as error:
        print(f"[ERROR] Unexpected error while loading config: {error}")
        traceback.print_exc()
        sys.exit(1)

    if not isinstance(config, dict) or not config:
        print(f"[ERROR] Config file '{args.config}' is empty or not a JSON object.")
        sys.exit(1)

    try:
        input_files = collect_input_files(args.input)
    except ValueError as error:
        print(f"[ERROR] {error}")
        sys.exit(1)

    is_batch = len(input_files) > 1
    if is_batch:
        print(f"Found {len(input_files)} supported file(s) in folder '{args.input}'")

    success_count = 0
    failure_count = 0

    for input_file in input_files:
        output_file = build_output_path_for_file(args.output, input_file, is_batch)
        succeeded = process_single_file(config, input_file, output_file)

        if succeeded:
            success_count += 1
        else:
            failure_count += 1

    print(f"\n{'=' * 60}")
    print("Run Summary")
    print(f"{'=' * 60}")
    print(f"Total files:  {len(input_files)}")
    print(f"Succeeded:    {success_count}")
    print(f"Failed:       {failure_count}")

    if success_count == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
