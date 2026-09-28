import sys
from collections import Counter
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC_FOLDER = PROJECT_ROOT / "src"

if str(SRC_FOLDER) not in sys.path:
    sys.path.insert(0, str(SRC_FOLDER))

from telemetry_standardizer import standardize_telemetry_file


BUS_DATA_FOLDER = Path(
    r"C:\Users\JAVED\OneDrive\Desktop\Bus Data"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "tests"
    / "real_data_standardization.csv"
)


if not BUS_DATA_FOLDER.exists():
    raise FileNotFoundError(
        f"Bus Data folder was not found: {BUS_DATA_FOLDER}"
    )

excel_files = sorted(
    file_path
    for file_path in BUS_DATA_FOLDER.rglob("*.xlsx")
    if not file_path.name.startswith("~$")
)

print("Real-data standardization test started")
print(f"Source files found: {len(excel_files)}")

records = []

for file_path in excel_files:
    record = {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "source_schema": "",
        "source_rows": 0,
        "standardized_rows": 0,
        "invalid_timestamps": 0,
        "invalid_latitudes": 0,
        "invalid_longitudes": 0,
        "invalid_speeds": 0,
        "status": "PASSED",
        "error": "",
    }

    try:
        standardized, metadata = standardize_telemetry_file(
            file_path
        )

        record["source_schema"] = metadata["source_schema"]
        record["source_rows"] = metadata["source_rows"]
        record["standardized_rows"] = metadata[
            "standardized_rows"
        ]

        record["invalid_timestamps"] = int(
            standardized["Timestamp"].isna().sum()
        )

        record["invalid_latitudes"] = int(
            (
                standardized["Latitude"].isna()
                | ~standardized["Latitude"].between(-90, 90)
            ).sum()
        )

        record["invalid_longitudes"] = int(
            (
                standardized["Longitude"].isna()
                | ~standardized["Longitude"].between(-180, 180)
            ).sum()
        )

        record["invalid_speeds"] = int(
            (
                standardized["Speed(km/h)"].isna()
                | (standardized["Speed(km/h)"] < 0)
            ).sum()
        )

        invalid_total = (
            record["invalid_timestamps"]
            + record["invalid_latitudes"]
            + record["invalid_longitudes"]
            + record["invalid_speeds"]
        )

        if invalid_total > 0:
            record["status"] = "REVIEW"

    except Exception as error:
        record["status"] = "FAILED"
        record["error"] = str(error)

    records.append(record)

results = pd.DataFrame(records)

results.to_csv(
    OUTPUT_FILE,
    index=False,
)

status_counts = Counter(results["status"])

print("\nStandardization summary")
print(f"PASSED: {status_counts['PASSED']}")
print(f"REVIEW: {status_counts['REVIEW']}")
print(f"FAILED: {status_counts['FAILED']}")

review_or_failed = results[
    results["status"].isin(["REVIEW", "FAILED"])
]

if not review_or_failed.empty:
    print("\nFiles requiring attention:")

    for _, row in review_or_failed.iterrows():
        print(
            f"{row['status']} | {row['file_name']} | "
            f"Schema: {row['source_schema'] or 'Unknown'} | "
            f"Invalid timestamp: {row['invalid_timestamps']} | "
            f"Invalid latitude: {row['invalid_latitudes']} | "
            f"Invalid longitude: {row['invalid_longitudes']} | "
            f"Invalid speed: {row['invalid_speeds']} | "
            f"Error: {row['error'] or 'None'}"
        )
else:
    print("\nAll source files standardized without invalid values.")

print(
    f"\nDetailed results saved at: {OUTPUT_FILE.resolve()}"
)