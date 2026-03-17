"""
excel_reader.py
---------------
Utility functions for parsing Excel (.xlsx) files using pandas.
Returns column headers and row data to the certificate generator.
"""

import pandas as pd


def get_columns(filepath: str) -> list[str]:
    """
    Read the Excel file and return a list of column header names.

    Args:
        filepath: Absolute path to the .xlsx file.

    Returns:
        A list of column name strings, e.g. ['Name', 'College', 'Event', 'Date']
    """
    df = pd.read_excel(filepath, nrows=0)  # Only read the header row
    return list(df.columns)


def get_rows(filepath: str) -> list[dict]:
    """
    Read all data rows from the Excel file.

    Args:
        filepath: Absolute path to the .xlsx file.

    Returns:
        A list of dicts, one per participant row.
        Example: [{'Name': 'Arjun A', 'College': 'MIT', ...}, ...]
    """
    df = pd.read_excel(filepath)
    # Convert every cell to string (guards against int/float date values)
    df = df.astype(str)
    return df.to_dict(orient="records")
