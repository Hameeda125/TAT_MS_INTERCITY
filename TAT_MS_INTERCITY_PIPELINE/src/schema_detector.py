from pathlib import Path

import pandas as pd


DEVICE_RAW_REQUIRED = {
    "packet time",
    "lat-lng",
    "speed",
}

MOBILE_RAW_REQUIRED = {
    "timestamp",
    "latitude",
    "longitude",
    "speed(km/h)",
}

MOBILE_PROCESSED_REQUIRED = {
    "timestamp",
    "latitude",
    "longitude",
    "speed(km/h)",
    "speed(m/s)",
}


def _normalize_column_name(value):
    """Normalize one Excel header for schema comparison."""
    return " ".join(str(value).strip().lower().split())


def _normalized_columns(columns):
    """Return normalized worksheet headers as a set."""
    return {
        _normalize_column_name(column)
        for column in columns
    }


def _detect_from_columns(columns):
    """
    Detect one supported telemetry schema from normalized column names.

    MOBILE_PROCESSED is checked before MOBILE_RAW because processed mobile
    files contain the MOBILE_RAW core columns plus Speed(m/s).
    """
    normalized = _normalized_columns(columns)

    if DEVICE_RAW_REQUIRED.issubset(normalized):
        return "DEVICE_RAW"

    if MOBILE_PROCESSED_REQUIRED.issubset(normalized):
        return "MOBILE_PROCESSED"

    if MOBILE_RAW_REQUIRED.issubset(normalized):
        return "MOBILE_RAW"

    return None


def detect_file_schema(file_path):
    """
    Detect the supported telemetry schema and worksheet for one Excel file.

    This module only identifies the schema. It deliberately does not import
    telemetry_standardizer.py, preventing a circular dependency.

    Returns:
        {
            "schema": "DEVICE_RAW" | "MOBILE_RAW" |
                      "MOBILE_PROCESSED" | "UNSUPPORTED",
            "sheet_name": <worksheet name or None>,
            "error": <message or None>,
        }
    """

    file_path = Path(file_path)

    result = {
        "schema": "UNSUPPORTED",
        "sheet_name": None,
        "error": None,
    }

    if not file_path.exists():
        result["error"] = f"File does not exist: {file_path}"
        return result

    if not file_path.is_file():
        result["error"] = f"Path is not a file: {file_path}"
        return result

    try:
        with pd.ExcelFile(file_path, engine="openpyxl") as excel:
            if not excel.sheet_names:
                result["error"] = "Excel workbook contains no worksheets."
                return result

            inspected = []

            for sheet_name in excel.sheet_names:
                try:
                    header = pd.read_excel(
                        excel,
                        sheet_name=sheet_name,
                        nrows=0,
                    )
                except Exception as error:
                    inspected.append(
                        f"{sheet_name}: header could not be read ({error})"
                    )
                    continue

                schema = _detect_from_columns(header.columns)

                if schema is not None:
                    result["schema"] = schema
                    result["sheet_name"] = sheet_name
                    result["error"] = None
                    return result

                normalized = sorted(
                    _normalized_columns(header.columns)
                )
                inspected.append(
                    f"{sheet_name}: {', '.join(normalized)}"
                )

            result["error"] = (
                "No worksheet matches a supported telemetry schema. "
                "Expected either DEVICE_RAW columns "
                "(Packet Time, Lat-Lng, Speed), MOBILE_RAW columns "
                "(Timestamp, Latitude, Longitude, Speed(km/h)), or "
                "MOBILE_PROCESSED columns "
                "(Timestamp, Latitude, Longitude, Speed(km/h), Speed(m/s)). "
                "Inspected worksheets: "
                + " | ".join(inspected)
            )
            return result

    except Exception as error:
        result["error"] = (
            f"Could not inspect Excel workbook {file_path.name}: "
            f"{type(error).__name__}: {error}"
        )
        return result


if __name__ == "__main__":
    raise SystemExit(
        "schema_detector.py is a helper module. "
        "Run the pipeline or call detect_file_schema(file_path)."
    )
