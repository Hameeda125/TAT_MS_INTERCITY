from pathlib import Path

import pandas as pd


DEVICE_RAW_COLUMNS = {
    "speed",
    "lat-lng",
    "packet time",
}

MOBILE_RAW_COLUMNS = {
    "timestamp",
    "latitude",
    "longitude",
    "speed(km/h)",
}

MOBILE_PROCESSED_COLUMNS = {
    "timestamp",
    "latitude",
    "longitude",
    "speed(m/s)",
    "speed(km/h)",
    "time_diff_sec",
}


def normalize_column_name(column):
    """Normalize a column name for reliable schema comparison."""

    return " ".join(
        str(column).strip().lower().split()
    )


def read_excel_columns(file_path):
    """
    Read the first worksheet and return its normalized columns.

    The function does not modify the original Excel file.
    """

    file_path = Path(file_path)

    workbook = pd.ExcelFile(
        file_path,
        engine="openpyxl",
    )

    if not workbook.sheet_names:
        raise ValueError("Workbook does not contain any worksheets.")

    dataframe = pd.read_excel(
        file_path,
        sheet_name=workbook.sheet_names[0],
        engine="openpyxl",
        nrows=5,
    )

    normalized_columns = {
        normalize_column_name(column)
        for column in dataframe.columns
    }

    return dataframe, normalized_columns


def detect_schema_from_columns(columns):
    """Classify a normalized set of Excel column names."""

    columns = {
        normalize_column_name(column)
        for column in columns
    }

    # Check processed format before raw mobile format because
    # processed files also contain the raw mobile core columns.
    if MOBILE_PROCESSED_COLUMNS.issubset(columns):
        return "MOBILE_PROCESSED"

    if DEVICE_RAW_COLUMNS.issubset(columns):
        return "DEVICE_RAW"

    if MOBILE_RAW_COLUMNS.issubset(columns):
        return "MOBILE_RAW"

    return "UNSUPPORTED"


def detect_file_schema(file_path):
    """Detect the telemetry schema used by one Excel file."""

    file_path = Path(file_path)

    result = {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "schema": "UNSUPPORTED",
        "sheet_name": "",
        "columns": [],
        "error": "",
    }

    try:
        workbook = pd.ExcelFile(
            file_path,
            engine="openpyxl",
        )

        if not workbook.sheet_names:
            result["error"] = "Workbook contains no worksheets."
            return result

        first_sheet = workbook.sheet_names[0]

        dataframe = pd.read_excel(
            file_path,
            sheet_name=first_sheet,
            engine="openpyxl",
            nrows=5,
        )

        normalized_columns = [
            normalize_column_name(column)
            for column in dataframe.columns
        ]

        result["sheet_name"] = first_sheet
        result["columns"] = normalized_columns
        result["schema"] = detect_schema_from_columns(
            normalized_columns
        )

    except Exception as error:
        result["error"] = str(error)

    return result