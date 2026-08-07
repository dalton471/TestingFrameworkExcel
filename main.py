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

Unlike the previous version, NOTHING is hard-coded. All paths are
supplied on the command line via argparse:

    python main.py --config config/kpi_automation_phase1.json --input input/KPI.xlsx
    python main.py --config config/kpi_automation_phase1.json --input input/ --output output/Validation_Report.xlsx
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

# File extensions this framework knows how to validate.
SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")

# Default location for the report if the user does not pass --output.
DEFAULT_OUTPUT_PATH = os.path.join("output", "Validation_Report.xlsx")


# --------------------------------------------------------------------------
# Argument Parsing
# --------------------------------------------------------------------------

def parse_arguments():
    """
    Define and parse the command-line arguments for the framework.

    --config  : Path to the JSON configuration file (required).
    --input   : Path to a single input file (.xlsx/.xls/.csv) OR a
                folder containing multiple such files (required).
    --output  : Path where the report should be written. Optional -
                defaults to output/Validation_Report.xlsx.

    Returns
    -------
    argparse.Namespace
        The parsed --config, --input, and --output values.
    """
    parser = argparse.ArgumentParser(
        description="KPI Validation Automation Framework - validates "
                     "Excel/CSV data against a JSON rule set and "
                     "produces a Validation_Report.xlsx"
    )

    parser.add_argument(
        "--config",
        required=True,
        help="Path to the JSON configuration file "
             "(e.g. config/kpi_automation_phase1.json)"
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Path to a single input file (.xlsx, .xls, .csv) OR a "
             "folder containing multiple such files "
             "(e.g. input/KPI.xlsx or input/)"
    )

    parser.add_argument(
        "--output",
        required=False,
        default=DEFAULT_OUTPUT_PATH,
        help=f"Path for the generated report. "
             f"Defaults to '{DEFAULT_OUTPUT_PATH}'"
    )

    # argparse itself prints a clean usage message and exits(2) if a
    # required argument is missing, so no extra handling is needed here.
    return parser.parse_args()


# --------------------------------------------------------------------------
# Input Validation / Discovery Helpers
# --------------------------------------------------------------------------

def validate_config_path(config_path):
    """
    Confirm the config path exists and points to a real file.
    Raises FileNotFoundError / ValueError with a clear message if not.
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found: '{config_path}'")

    if not os.path.isfile(config_path):
        raise ValueError(f"Config path is not a file: '{config_path}'")


def validate_input_path(input_path):
    """
    Confirm the input path exists (as either a file or a folder).
    Raises FileNotFoundError if the path does not exist at all.
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input path not found: '{input_path}'")


def collect_input_files(input_path):
    """
    Given an input path that is either a single file or a folder,
    return a list of every supported file that needs to be validated.

    - If input_path is a file -> returns [input_path], provided its
      extension is supported.
    - If input_path is a folder -> returns every supported file found
      directly inside that folder (non-recursive), sorted alphabetically.

    Raises
    ------
    ValueError
        If a single file has an unsupported extension, or a folder
        contains no supported files at all.
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
        for filename in sorted(os.listdir(input_path)):
            # Skip Excel's temporary lock files (e.g. "~$KPI.xlsx")
            if filename.startswith("~$"):
                continue

            ext = os.path.splitext(filename)[1].lower()
            if ext in SUPPORTED_EXTENSIONS:
                found_files.append(os.path.join(input_path, filename))

        if not found_files:
            raise ValueError(
                f"No supported files ({', '.join(SUPPORTED_EXTENSIONS)}) "
                f"found inside folder: '{input_path}'"
            )
        return found_files

    # Path exists but is neither a regular file nor a directory
    # (e.g. a broken symlink or a special device file).
    raise ValueError(f"Invalid input path: '{input_path}'")


def prepare_excel_path(file_path):
    """
    The existing validation functions (in validations.py) always read
    data using pandas' Excel reader (pd.read_excel / pd.ExcelFile), so
    they expect an .xlsx/.xls file on disk.

    If the incoming file is already Excel, this simply returns it
    unchanged. If it is a .csv file, it is converted into a temporary
    single-sheet .xlsx file so that it can flow through the exact same
    validation functions without changing a single line of their code.

    Note: a CSV only contains one flat table, so multi-sheet checks
    (sheet existence, cross-sheet checks, etc.) will naturally report
    failures for missing sheets - this is expected, not a bug.

    Parameters
    ----------
    file_path : str
        Path to the original input file.

    Returns
    -------
    tuple(str, bool)
        (path_to_use, is_temporary) - is_temporary is True when a
        temp .xlsx file was created and should be deleted afterwards.
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext in (".xlsx", ".xls"):
        return file_path, False

    if ext == ".csv":
        try:
            df = pd.read_csv(file_path)
        except Exception as error:
            raise ValueError(
                f"Could not read CSV file '{file_path}': {error}"
            )

        # Build a temp .xlsx path next to the output folder so it's
        # easy to find/clean up, using the original filename as the
        # single sheet name.
        os.makedirs("output", exist_ok=True)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        temp_path = os.path.join("output", f"_temp_{base_name}.xlsx")

        try:
            df.to_excel(temp_path, index=False, sheet_name="Sheet1")
        except Exception as error:
            raise ValueError(
                f"Could not convert CSV to Excel for validation: {error}"
            )

        return temp_path, True

    # Should not happen because collect_input_files() already filtered
    # extensions, but guarded here for safety.
    raise ValueError(f"Unsupported file type '{ext}' for '{file_path}'")


def build_output_path_for_file(output_arg, input_file_path, is_batch):
    """
    Decide the final report path for a given input file.

    - Single-file mode: use the --output path exactly as given.
    - Batch (folder) mode: suffix the --output filename with the
      source file's name, so multiple reports don't overwrite
      each other, e.g.:
          output/Validation_Report.xlsx
              -> output/Validation_Report_KPI_March.xlsx
              -> output/Validation_Report_KPI_April.xlsx
    """
    if not is_batch:
        return output_arg

    output_dir = os.path.dirname(output_arg) or "."
    base, ext = os.path.splitext(os.path.basename(output_arg))
    input_stem = os.path.splitext(os.path.basename(input_file_path))[0]

    return os.path.join(output_dir, f"{base}_{input_stem}{ext}")


# --------------------------------------------------------------------------
# Core Orchestration (unchanged pipeline, now reusable per-file)
# --------------------------------------------------------------------------

def run_validation_pipeline(config, excel_file, output_file):
    """
    Run the exact same orchestration steps as the original framework:

        1. Load workbook.
        2. Execute all validation functions.
        3. Collect all results.
        4. Generate the report.

    Parameters
    ----------
    config : dict
        Already-loaded JSON configuration.
    excel_file : str
        Path to the .xlsx file to validate (CSV inputs are converted
        to a temporary .xlsx before reaching this function).
    output_file : str
        Path where the Validation_Report.xlsx should be written.
    """
    # ---- Step 2: Load input Excel ----
    workbook = load_excel(excel_file)
    print("Excel Loaded Successfully")

    # ---- Step 3: Execute all validation functions ----
    sheet_results = validate_sheet_list(config, workbook)
    column_results = validate_sheet_columns(config, excel_file)
    duplicate_results = validate_duplicate_check(config, excel_file)
    null_results = validate_null_check(config, excel_file)
    formula_results = validate_formula_check(config, excel_file)
    numeric_precision_results = validate_numeric_precision(config, excel_file)
    business_rule_results = validate_business_rule(config, excel_file)
    reconciliation_results = validate_reconciliation(config, excel_file)
    cross_sheet_results = validate_cross_sheet(config, excel_file)
    trend_results = validate_trend(config, excel_file)

    # ---- Step 4: Collect all results together ----
    final_results = (
        sheet_results
        + column_results
        + duplicate_results
        + null_results
        + formula_results
        + numeric_precision_results
        + business_rule_results
        + reconciliation_results
        + cross_sheet_results
        + trend_results
    )

    # ---- Step 5: Generate the Validation Report ----
    # Make sure the output folder exists before writing to it.
    output_dir = os.path.dirname(output_file)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    generate_report(final_results, output_file)


def process_single_file(config, input_file, output_file):
    """
    Wraps run_validation_pipeline() with per-file error handling and
    CSV-to-Excel conversion/cleanup, so that one bad file in a batch
    doesn't crash the whole run.

    Returns
    -------
    bool
        True if this file was validated and a report was generated
        successfully, False otherwise.
    """
    temp_path = None

    try:
        print(f"\n{'=' * 60}")
        print(f"Processing: {input_file}")
        print(f"{'=' * 60}")

        # Convert CSV to a temporary Excel file if needed; Excel files
        # pass through unchanged.
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
        # Catch-all for anything unexpected, without crashing the batch.
        print(f"[FAILED] Unexpected error while processing '{input_file}': {error}")
        traceback.print_exc()
        return False

    finally:
        # Clean up the temporary .xlsx file created from a CSV, if any.
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                # Non-fatal - leftover temp file isn't worth crashing over.
                pass


# --------------------------------------------------------------------------
# Main Entry Point
# --------------------------------------------------------------------------

def main():
    args = parse_arguments()

    # ---- Step 0: Validate the paths supplied on the command line ----
    try:
        validate_config_path(args.config)
        validate_input_path(args.input)
    except (FileNotFoundError, ValueError) as error:
        print(f"[ERROR] {error}")
        sys.exit(1)

    # ---- Step 1: Load JSON configuration (shared across all files) ----
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

    # ---- Discover which file(s) need to be validated ----
    try:
        input_files = collect_input_files(args.input)
    except ValueError as error:
        print(f"[ERROR] {error}")
        sys.exit(1)

    is_batch = len(input_files) > 1
    if is_batch:
        print(f"Found {len(input_files)} supported file(s) in folder '{args.input}'")

    # ---- Process every discovered file ----
    success_count = 0
    failure_count = 0

    for input_file in input_files:
        output_file = build_output_path_for_file(args.output, input_file, is_batch)
        succeeded = process_single_file(config, input_file, output_file)

        if succeeded:
            success_count += 1
        else:
            failure_count += 1

    # ---- Final summary ----
    print(f"\n{'=' * 60}")
    print("Run Summary")
    print(f"{'=' * 60}")
    print(f"Total files:  {len(input_files)}")
    print(f"Succeeded:    {success_count}")
    print(f"Failed:       {failure_count}")

    # Exit with a non-zero code if everything failed, so this plays
    # nicely with CI pipelines / task schedulers.
    if success_count == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
