import sys
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC_FOLDER = PROJECT_ROOT / "src"

if str(SRC_FOLDER) not in sys.path:
    sys.path.insert(0, str(SRC_FOLDER))

from schema_detector import detect_file_schema
from telemetry_standardizer import standardize_telemetry_file


BUS_DATA_FOLDER = Path(
    r"C:\Users\JAVED\OneDrive\Desktop\Bus Data"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "tests"
    / "real_data_invalid_values_report.xlsx"
)


def normalize_columns(dataframe):
    """Normalize source column names."""
    dataframe = dataframe.copy()

    dataframe.columns = [
        " ".join(str(column).strip().lower().split())
        for column in dataframe.columns
    ]

    return dataframe


def original_value(dataframe, column, row_index):
    """Safely obtain an original value from the source file."""
    if column not in dataframe.columns:
        return ""

    value = dataframe.iloc[row_index][column]

    if pd.isna(value):
        return "<BLANK>"

    return value


def suggested_treatment(field, reason):
    """Return a preliminary treatment for review."""
    if field == "Timestamp":
        return (
            "Check source timestamp; exclude from time-dependent "
            "calculations if it cannot be recovered."
        )

    if field in ["Latitude", "Longitude"]:
        return (
            "Check neighbouring GPS observations; use short-gap "
            "interpolation only when spatially reliable."
        )

    if field == "Speed":
        if "negative" in reason.lower():
            return (
                "Treat as invalid; assess neighbouring timestamps "
                "for short-gap time interpolation."
            )

        return (
            "Recover from the alternative speed unit if available; "
            "otherwise assess short-gap time interpolation."
        )

    return "Review during preprocessing."


if not BUS_DATA_FOLDER.exists():
    raise FileNotFoundError(
        f"Bus Data folder was not found: {BUS_DATA_FOLDER}"
    )


excel_files = sorted(
    file_path
    for file_path in BUS_DATA_FOLDER.rglob("*.xlsx")
    if not file_path.name.startswith("~$")
)


print("Detailed invalid-value diagnosis started")
print(f"Source files found: {len(excel_files)}")


summary_records = []
invalid_records = []


for file_path in excel_files:
    summary = {
        "File Name": file_path.name,
        "File Path": str(file_path),
        "Source Schema": "",
        "Source Rows": 0,
        "Invalid Timestamps": 0,
        "Invalid Latitudes": 0,
        "Invalid Longitudes": 0,
        "Invalid Speeds": 0,
        "Total Invalid Cells": 0,
        "Status": "PASSED",
        "Error": "",
    }

    try:
        schema_result = detect_file_schema(file_path)
        schema = schema_result["schema"]
        sheet_name = schema_result["sheet_name"]

        source_data = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            engine="openpyxl",
        )

        source_data = normalize_columns(source_data)

        standardized, metadata = standardize_telemetry_file(
            file_path
        )

        summary["Source Schema"] = schema
        summary["Source Rows"] = len(source_data)

        for row_index in range(len(standardized)):
            excel_row_number = row_index + 2

            timestamp = standardized.iloc[row_index]["Timestamp"]
            latitude = standardized.iloc[row_index]["Latitude"]
            longitude = standardized.iloc[row_index]["Longitude"]
            speed_kmh = standardized.iloc[row_index]["Speed(km/h)"]
            speed_mps = standardized.iloc[row_index]["Speed(m/s)"]

            common_record = {
                "File Name": file_path.name,
                "File Path": str(file_path),
                "Source Schema": schema,
                "Sheet Name": sheet_name,
                "DataFrame Index": row_index,
                "Excel Row Number": excel_row_number,
            }

            if pd.isna(timestamp):
                summary["Invalid Timestamps"] += 1

                if schema == "DEVICE_RAW":
                    original_timestamp = original_value(
                        source_data,
                        "packet time",
                        row_index,
                    )
                else:
                    original_timestamp = original_value(
                        source_data,
                        "timestamp",
                        row_index,
                    )

                invalid_records.append(
                    {
                        **common_record,
                        "Invalid Field": "Timestamp",
                        "Original Value": original_timestamp,
                        "Original Speed(km/h)": "",
                        "Original Speed(m/s)": "",
                        "Reason": "Missing or unrecognized timestamp",
                        "Suggested Treatment": suggested_treatment(
                            "Timestamp",
                            "Missing or unrecognized timestamp",
                        ),
                    }
                )

            latitude_invalid = (
                pd.isna(latitude)
                or latitude < -90
                or latitude > 90
            )

            if latitude_invalid:
                summary["Invalid Latitudes"] += 1

                if schema == "DEVICE_RAW":
                    original_coordinate = original_value(
                        source_data,
                        "lat-lng",
                        row_index,
                    )
                else:
                    original_coordinate = original_value(
                        source_data,
                        "latitude",
                        row_index,
                    )

                invalid_records.append(
                    {
                        **common_record,
                        "Invalid Field": "Latitude",
                        "Original Value": original_coordinate,
                        "Original Speed(km/h)": "",
                        "Original Speed(m/s)": "",
                        "Reason": "Missing, non-numeric or out-of-range latitude",
                        "Suggested Treatment": suggested_treatment(
                            "Latitude",
                            "Invalid latitude",
                        ),
                    }
                )

            longitude_invalid = (
                pd.isna(longitude)
                or longitude < -180
                or longitude > 180
            )

            if longitude_invalid:
                summary["Invalid Longitudes"] += 1

                if schema == "DEVICE_RAW":
                    original_coordinate = original_value(
                        source_data,
                        "lat-lng",
                        row_index,
                    )
                else:
                    original_coordinate = original_value(
                        source_data,
                        "longitude",
                        row_index,
                    )

                invalid_records.append(
                    {
                        **common_record,
                        "Invalid Field": "Longitude",
                        "Original Value": original_coordinate,
                        "Original Speed(km/h)": "",
                        "Original Speed(m/s)": "",
                        "Reason": (
                            "Missing, non-numeric or out-of-range longitude"
                        ),
                        "Suggested Treatment": suggested_treatment(
                            "Longitude",
                            "Invalid longitude",
                        ),
                    }
                )

            speed_invalid = (
                pd.isna(speed_kmh)
                or pd.isna(speed_mps)
                or speed_kmh < 0
                or speed_mps < 0
            )

            if speed_invalid:
                summary["Invalid Speeds"] += 1

                if schema == "DEVICE_RAW":
                    original_kmh = original_value(
                        source_data,
                        "speed",
                        row_index,
                    )
                    original_mps = ""
                else:
                    original_kmh = original_value(
                        source_data,
                        "speed(km/h)",
                        row_index,
                    )
                    original_mps = original_value(
                        source_data,
                        "speed(m/s)",
                        row_index,
                    )

                if (
                    (not pd.isna(speed_kmh) and speed_kmh < 0)
                    or (not pd.isna(speed_mps) and speed_mps < 0)
                ):
                    reason = "Negative speed"
                else:
                    reason = (
                        "Speed is missing or non-numeric after "
                        "available unit recovery"
                    )

                invalid_records.append(
                    {
                        **common_record,
                        "Invalid Field": "Speed",
                        "Original Value": original_kmh,
                        "Original Speed(km/h)": original_kmh,
                        "Original Speed(m/s)": original_mps,
                        "Reason": reason,
                        "Suggested Treatment": suggested_treatment(
                            "Speed",
                            reason,
                        ),
                    }
                )

        total_invalid = (
            summary["Invalid Timestamps"]
            + summary["Invalid Latitudes"]
            + summary["Invalid Longitudes"]
            + summary["Invalid Speeds"]
        )

        summary["Total Invalid Cells"] = total_invalid

        if total_invalid > 0:
            summary["Status"] = "REVIEW"

    except Exception as error:
        summary["Status"] = "FAILED"
        summary["Error"] = str(error)

    summary_records.append(summary)


summary_dataframe = pd.DataFrame(summary_records)

invalid_columns = [
    "File Name",
    "File Path",
    "Source Schema",
    "Sheet Name",
    "DataFrame Index",
    "Excel Row Number",
    "Invalid Field",
    "Original Value",
    "Original Speed(km/h)",
    "Original Speed(m/s)",
    "Reason",
    "Suggested Treatment",
]

invalid_dataframe = pd.DataFrame(
    invalid_records,
    columns=invalid_columns,
)


with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
) as writer:
    summary_dataframe.to_excel(
        writer,
        sheet_name="File Summary",
        index=False,
    )

    invalid_dataframe.to_excel(
        writer,
        sheet_name="Invalid Values",
        index=False,
    )


print("\nDetailed diagnosis completed")
print(f"Files checked: {len(summary_dataframe)}")
print(f"Invalid cells found: {len(invalid_dataframe)}")
print(
    "Files requiring review: "
    f"{(summary_dataframe['Status'] == 'REVIEW').sum()}"
)
print(
    "Files failed: "
    f"{(summary_dataframe['Status'] == 'FAILED').sum()}"
)
print(f"\nReport saved at: {OUTPUT_FILE.resolve()}")