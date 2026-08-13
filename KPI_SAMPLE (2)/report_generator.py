"""
report_generator.py
--------------------
Takes the combined list of validation result dictionaries and writes
them out to the final Validation_Report.xlsx file.

This is the exact same two lines that used to sit at the bottom of
main()'s old code -- just moved into its own reusable function.
"""

import pandas as pd


def generate_report(final_results, output_file="output/Validation_Report.xlsx"):
    """
    Convert the list of result dicts into a DataFrame and write it
    to an Excel file.

    Parameters
    ----------
    final_results : list[dict]
        Combined results from all validation functions.
    output_file : str
        Path where the Validation_Report.xlsx should be saved.
    """
    report = pd.DataFrame(final_results)
    report.to_excel(output_file, index=False)
    print(f"\n{output_file} created successfully.")
