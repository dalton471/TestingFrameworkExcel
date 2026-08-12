"""
excel_reader.py
----------------
Handles all logic related to opening/reading the input Excel workbook.

Currently contains:
    - load_excel(): opens the workbook and returns a pandas ExcelFile
      object (used to get sheet names etc.)
"""

import pandas as pd


def load_excel(excel_file):
    workbook = pd.ExcelFile(excel_file)
    return workbook

