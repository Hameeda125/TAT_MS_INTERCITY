from pathlib import Path

import pandas as pd

from schema_detector import detect_file_schema


STANDARD_COLUMNS = [
    "Timestamp",
    "Latitude",
    "Longitude",
    "Speed(m/s)",
    "Speed(km/h)",
]


def _normalize_columns(dataframe):
    """Return a copy with normalized column names."""

    dataframe = dataframe.copy()

    dataframe.columns = [
        " ".join(str(column).strip().lower().split())
        for column in dataframe.columns
    ]

    return dataframe


def _read_source_file(file_path, sheet_name):
    """Read the worksheet identified by the schema detector."""

    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        engine="openpyxl",
    )


def _standardize_device_raw(dataframe):
    """Extract core telemetry from a GPS-device export."""

    lat_lng = (
        dataframe["lat-lng"]
        .astype("string")
        .str.split(",", n=1, expand=True)
    )

    if lat_lng.shape[1] != 2:
        raise ValueError(
            "The Lat-Lng column could not be split into "
            "latitude and longitude."
        )

    standardized = pd.DataFrame(
        {
            "Timestamp": dataframe["packet time"],
            "Latitude": lat_lng[0].str.strip(),
            "Longitude": lat_lng[1].str.strip(),
            "Speed(km/h)": dataframe["speed"],
        }
    )

    standardized["Speed(m/s)"] = (
        pd.to_numeric(
            standardized["Speed(km/h)"],
            errors="coerce",
        )
        / 3.6
    )

    return standardized


def _standardize_mobile_raw(dataframe):
    """Extract core telemetry from a raw mobile-app export."""

    standardized = pd.DataFrame(
        {
            "Timestamp": dataframe["timestamp"],
            "Latitude": dataframe["latitude"],
            "Longitude": dataframe["longitude"],
            "Speed(km/h)": dataframe["speed(km/h)"],
        }
    )

    standardized["Speed(m/s)"] = (
        pd.to_numeric(
            standardized["Speed(km/h)"],
            errors="coerce",
        )
        / 3.6
    )

    return standardized


def _standardize_mobile_processed(dataframe):
    """
    Extract original telemetry from a previously processed mobile file.

    Existing derived columns such as acceleration, distance and phase
    are intentionally ignored and will be recalculated later.
    """

    standardized = pd.DataFrame(
        {
            "Timestamp": dataframe["timestamp"],
            "Latitude": dataframe["latitude"],
            "Longitude": dataframe["longitude"],
            "Speed(km/h)": dataframe["speed(km/h)"],
        }
    )

    speed_kmh = pd.to_numeric(
        standardized["Speed(km/h)"],
        errors="coerce",
    )

    speed_mps = pd.to_numeric(
        dataframe["speed(m/s)"],
        errors="coerce",
    )

    # Prefer km/h as the common source. If it is missing for an
    # individual row, recover it from the available m/s value.
    standardized["Speed(km/h)"] = speed_kmh.fillna(
        speed_mps * 3.6
    )

    standardized["Speed(m/s)"] = (
        standardized["Speed(km/h)"] / 3.6
    )

    return standardized


def standardize_telemetry_file(file_path):
    """
    Convert one supported telemetry file to the common core schema.

    The original Excel file is read only and is never modified.
    """

    file_path = Path(file_path)
    schema_result = detect_file_schema(file_path)

    schema = schema_result["schema"]
    sheet_name = schema_result["sheet_name"]

    if schema == "UNSUPPORTED":
        error = schema_result["error"] or (
            "The file does not match a supported telemetry schema."
        )
        raise ValueError(
            f"Unsupported file {file_path.name}: {error}"
        )

    dataframe = _read_source_file(
        file_path,
        sheet_name,
    )
    dataframe = _normalize_columns(dataframe)

    if schema == "DEVICE_RAW":
        standardized = _standardize_device_raw(dataframe)
    elif schema == "MOBILE_RAW":
        standardized = _standardize_mobile_raw(dataframe)
    elif schema == "MOBILE_PROCESSED":
        standardized = _standardize_mobile_processed(dataframe)
    else:
        raise ValueError(f"Unexpected schema: {schema}")

    standardized["Timestamp"] = pd.to_datetime(
        standardized["Timestamp"],
        errors="coerce",
        dayfirst=True,
    )

    for column in [
        "Latitude",
        "Longitude",
        "Speed(m/s)",
        "Speed(km/h)",
    ]:
        standardized[column] = pd.to_numeric(
            standardized[column],
            errors="coerce",
        )

    standardized = standardized[STANDARD_COLUMNS]

    metadata = {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "source_schema": schema,
        "sheet_name": sheet_name,
        "source_rows": len(dataframe),
        "standardized_rows": len(standardized),
    }

    return standardized, metadata