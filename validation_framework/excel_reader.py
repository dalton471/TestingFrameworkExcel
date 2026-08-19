"""
excel_reader.py
----------------
Handles loading the input Excel workbook and exposing every sheet as
a pandas DataFrame.

This reader is purely mechanical - it does not know about validation
rules, sheet names, or column names configured in JSON. It only knows
how to open a workbook file and hand back its sheets.
"""

import os

import pandas as pd


def load_workbook(file_path):
    """
    Open the workbook and return a pandas ExcelFile, giving access to
    sheet names without loading every sheet into memory up front.

    Raises
    ------
    FileNotFoundError
        If file_path does not exist.
    ValueError
        If the path is not a file, has an unsupported extension, or
        cannot be parsed as an Excel workbook.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: '{file_path}'")

    if not os.path.isfile(file_path):
        raise ValueError(f"Input path is not a file: '{file_path}'")

    extension = os.path.splitext(file_path)[1].lower()
    if extension not in (".xlsx", ".xls"):
        raise ValueError(
            f"Unsupported input file type '{extension}'. "
            f"Supported types are: .xlsx, .xls"
        )

    try:
        workbook = pd.ExcelFile(file_path)
    except Exception as error:
        raise ValueError(f"Could not read Excel workbook '{file_path}': {error}")

    return workbook


def load_all_sheets(workbook):
    """
    Read every sheet in the workbook into a DataFrame.

    Parameters
    ----------
    workbook : pandas.ExcelFile
        The workbook returned by load_workbook().

    Returns
    -------
    dict[str, pandas.DataFrame]
        Mapping of sheet name -> DataFrame, in the workbook's own
        sheet order.
    """
    sheets = {}
    for sheet_name in workbook.sheet_names:
        try:
            sheets[sheet_name] = workbook.parse(sheet_name)
        except Exception as error:
            raise ValueError(
                f"Could not read sheet '{sheet_name}': {error}"
            )
    return sheets


def get_sheet_names(workbook):
    """Return the list of sheet names present in the workbook."""
    return list(workbook.sheet_names)
