"""Command-line entry point for the Validation Automation Framework.

This module is intentionally small: it parses arguments, configures
logging, delegates execution to ValidationRunner, prints a concise
terminal summary, and sets the process exit code. It contains no
validation logic, no Excel-specific code, and no hard-coded sheet or
column names.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import src.validation_framework.rules  # noqa: F401  (rule-registration side effect)
from src.validation_framework.core.runner import ValidationRunner
from src.validation_framework.exceptions.exceptions import (
    ConfigurationError, FrameworkError, InputError, ValidationEngineError,
)
from src.validation_framework.logging.logger import configure_logging

DEFAULT_CONFIG_PATH = Path("config") / "validation_config.json"
DEFAULT_OUTPUT_PATH = Path("output") / "Validation_Report.xlsx"
DEFAULT_JSON_REPORT_PATH = Path("output") / "validation_results.json"
DEFAULT_LOG_DIR = Path("logs")


def parse_arguments(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="Validate an Excel or CSV file against a configurable "
                     "or built-in set of data-quality rules.",
    )
    parser.add_argument(
        "--config", required=False, default=str(DEFAULT_CONFIG_PATH),
        help="Path to a JSON rule configuration file. Optional - if the "
             "file does not exist, the framework's built-in default rule "
             f"set is used instead. Defaults to '{DEFAULT_CONFIG_PATH}'.",
    )
    parser.add_argument(
        "--input", required=True,
        help="Path to the input file to validate (.xlsx, .xls, or .csv).",
    )
    parser.add_argument(
        "--output", required=False, default=str(DEFAULT_OUTPUT_PATH),
        help=f"Path for the generated Excel report. Defaults to '{DEFAULT_OUTPUT_PATH}'.",
    )
    parser.add_argument(
        "--json-report", required=False, default=str(DEFAULT_JSON_REPORT_PATH),
        help=f"Path for the generated machine-readable JSON report. Defaults to '{DEFAULT_JSON_REPORT_PATH}'.",
    )
    parser.add_argument(
        "--log-dir", required=False, default=str(DEFAULT_LOG_DIR),
        help=f"Directory for log files. Defaults to '{DEFAULT_LOG_DIR}'.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_arguments(argv)
    logger = configure_logging(Path(args.log_dir))
    logger.info("Validation framework starting")

    runner = ValidationRunner()
    config_path = Path(args.config) if args.config else None

    try:
        result = runner.run(
            input_path=Path(args.input),
            config_path=config_path,
            output_path=Path(args.output),
            json_report_path=Path(args.json_report),
        )
    except ConfigurationError as error:
        logger.error("Configuration error: %s", error)
        print(f"[CONFIGURATION ERROR] {error}")
        return 2
    except InputError as error:
        logger.error("Input error: %s", error)
        print(f"[INPUT ERROR] {error}")
        return 3
    except ValidationEngineError as error:
        logger.error("Validation execution error: %s", error)
        print(f"[EXECUTION ERROR] {error}")
        return 4
    except FrameworkError as error:
        logger.error("Framework error: %s", error)
        print(f"[FRAMEWORK ERROR] {error}")
        return 1

    _print_summary(result)
    logger.info("Validation framework completed")
    return 0 if result.succeeded else 1


def _print_summary(result) -> None:
    print("=" * 60)
    print("VALIDATION SUMMARY")
    print("=" * 60)
    print(f"Configuration source: {'built-in defaults' if result.used_default_config else result.config_source}")
    print(f"Rules executed: {', '.join(result.rule_names_used)}")
    print("-" * 60)
    for label, value in result.execution_stats.items():
        print(f"{label}: {value}")
    print("-" * 60)
    print(f"Excel report: {result.report_path}")
    print(f"JSON report:  {result.json_report_path}")
    print("=" * 60)


if __name__ == "__main__":
    sys.exit(main())
