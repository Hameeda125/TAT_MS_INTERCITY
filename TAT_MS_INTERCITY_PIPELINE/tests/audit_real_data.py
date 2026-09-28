from collections import Counter
from pathlib import Path

import pandas as pd


BUS_DATA_FOLDER = Path(
    r"C:\Users\JAVED\OneDrive\Desktop\Bus Data"
)

AUDIT_OUTPUT = Path("tests/real_data_audit.csv")

EXPECTED_COLUMNS = [
    "Status",
    "Speed",
    "Lat-Lng",
    "Address",
    "Cumulative Distance",
    "Cumulative Time",
    "Packet Time",
    "Actual Received Time",
]


def classify_vehicle(file_name):
    """Identify the vehicle type from the filename."""

    name = file_name.lower()

    if "diesel" in name or "ice" in name:
        return "ICE"

    if "ev" in name or "electric" in name:
        return "EV"

    return "UNKNOWN"


def audit_excel_file(file_path):
    """Inspect one Excel file without changing it."""

    record = {
        "date_folder": file_path.parent.name,
        "file_name": file_path.name,
        "file_path": str(file_path),
        "vehicle_type": classify_vehicle(file_path.name),
        "status": "OK",
        "sheet_names": "",
        "row_count": 0,
        "column_count": 0,
        "missing_columns": "",
        "extra_columns": "",
        "invalid_speed_values": 0,
        "invalid_lat_lng_values": 0,
        "invalid_packet_times": 0,
        "duplicate_rows": 0,
        "error": "",
    }

    try:
        workbook = pd.ExcelFile(
            file_path,
            engine="openpyxl",
        )

        record["sheet_names"] = " | ".join(
            workbook.sheet_names
        )

        if not workbook.sheet_names:
            record["status"] = "ISSUE"
            record["error"] = "Workbook contains no sheets."
            return record

        dataframe = pd.read_excel(
            file_path,
            sheet_name=workbook.sheet_names[0],
            engine="openpyxl",
        )

        dataframe.columns = [
            str(column).strip()
            for column in dataframe.columns
        ]

        record["row_count"] = len(dataframe)
        record["column_count"] = len(dataframe.columns)

        missing_columns = [
            column
            for column in EXPECTED_COLUMNS
            if column not in dataframe.columns
        ]

        extra_columns = [
            column
            for column in dataframe.columns
            if column not in EXPECTED_COLUMNS
        ]

        record["missing_columns"] = " | ".join(
            missing_columns
        )
        record["extra_columns"] = " | ".join(
            extra_columns
        )

        if dataframe.empty:
            record["status"] = "ISSUE"
            record["error"] = "Worksheet is empty."
            return record

        if missing_columns:
            record["status"] = "ISSUE"
            return record

        speeds = pd.to_numeric(
            dataframe["Speed"],
            errors="coerce",
        )

        invalid_speed_mask = (
            speeds.isna()
            | (speeds < 0)
        )

        record["invalid_speed_values"] = int(
            invalid_speed_mask.sum()
        )

        lat_lng_parts = (
            dataframe["Lat-Lng"]
            .astype("string")
            .str.split(",", n=1, expand=True)
        )

        if lat_lng_parts.shape[1] != 2:
            record["invalid_lat_lng_values"] = len(
                dataframe
            )
        else:
            latitudes = pd.to_numeric(
                lat_lng_parts[0].str.strip(),
                errors="coerce",
            )

            longitudes = pd.to_numeric(
                lat_lng_parts[1].str.strip(),
                errors="coerce",
            )

            invalid_lat_lng_mask = (
                latitudes.isna()
                | longitudes.isna()
                | ~latitudes.between(-90, 90)
                | ~longitudes.between(-180, 180)
            )

            record["invalid_lat_lng_values"] = int(
                invalid_lat_lng_mask.sum()
            )

        packet_times = pd.to_datetime(
            dataframe["Packet Time"],
            errors="coerce",
            dayfirst=True,
        )

        record["invalid_packet_times"] = int(
            packet_times.isna().sum()
        )

        record["duplicate_rows"] = int(
            dataframe.duplicated().sum()
        )

        if (
            record["invalid_speed_values"] > 0
            or record["invalid_lat_lng_values"] > 0
            or record["invalid_packet_times"] > 0
        ):
            record["status"] = "ISSUE"

    except Exception as error:
        record["status"] = "ERROR"
        record["error"] = str(error)

    return record


if not BUS_DATA_FOLDER.exists():
    raise FileNotFoundError(
        f"Bus Data folder was not found: "
        f"{BUS_DATA_FOLDER}"
    )

excel_files = sorted(
    file_path
    for file_path in BUS_DATA_FOLDER.rglob("*.xlsx")
    if not file_path.name.startswith("~$")
)

print("Real-data audit started")
print(f"Excel files found: {len(excel_files)}")

audit_records = [
    audit_excel_file(file_path)
    for file_path in excel_files
]

audit_dataframe = pd.DataFrame(audit_records)
audit_dataframe.to_csv(AUDIT_OUTPUT, index=False)

status_counts = Counter(
    record["status"]
    for record in audit_records
)

vehicle_counts = Counter(
    record["vehicle_type"]
    for record in audit_records
)

date_folders = {
    record["date_folder"]
    for record in audit_records
}

print("\nAudit summary")
print(f"Date folders found: {len(date_folders)}")
print(f"Files checked: {len(audit_records)}")
print(f"Files OK: {status_counts['OK']}")
print(f"Files with issues: {status_counts['ISSUE']}")
print(f"Files with read errors: {status_counts['ERROR']}")
print(f"EV files: {vehicle_counts['EV']}")
print(f"ICE files: {vehicle_counts['ICE']}")
print(f"Unknown vehicle type: {vehicle_counts['UNKNOWN']}")

problem_records = [
    record
    for record in audit_records
    if record["status"] != "OK"
]

if problem_records:
    print("\nFiles requiring attention:")

    for record in problem_records:
        print(
            f"{record['status']} | "
            f"{record['date_folder']} | "
            f"{record['file_name']} | "
            f"Missing: {record['missing_columns'] or 'None'} | "
            f"Invalid speed: "
            f"{record['invalid_speed_values']} | "
            f"Invalid Lat-Lng: "
            f"{record['invalid_lat_lng_values']} | "
            f"Invalid Packet Time: "
            f"{record['invalid_packet_times']} | "
            f"Error: {record['error'] or 'None'}"
        )
else:
    print("\nAll files follow the expected structure.")

print(f"\nDetailed audit saved at: {AUDIT_OUTPUT.resolve()}")