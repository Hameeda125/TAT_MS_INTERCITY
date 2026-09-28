import sys
from collections import Counter
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC_FOLDER = PROJECT_ROOT / "src"

if str(SRC_FOLDER) not in sys.path:
    sys.path.insert(0, str(SRC_FOLDER))

from schema_detector import detect_file_schema

BUS_DATA_FOLDER = Path(
    r"C:\Users\JAVED\OneDrive\Desktop\Bus Data"
)

OUTPUT_FILE = Path(
    "tests/real_data_schema_detection.csv"
)


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

print("Multi-format schema detection started")
print(f"Excel files found: {len(excel_files)}")

results = [
    detect_file_schema(file_path)
    for file_path in excel_files
]

schema_counts = Counter(
    result["schema"]
    for result in results
)

output_records = []

for result in results:
    output_records.append(
        {
            "file_name": result["file_name"],
            "file_path": result["file_path"],
            "sheet_name": result["sheet_name"],
            "schema": result["schema"],
            "columns": " | ".join(result["columns"]),
            "error": result["error"],
        }
    )

pd.DataFrame(output_records).to_csv(
    OUTPUT_FILE,
    index=False,
)

print("\nSchema detection summary")
print(f"DEVICE_RAW: {schema_counts['DEVICE_RAW']}")
print(f"MOBILE_RAW: {schema_counts['MOBILE_RAW']}")
print(
    "MOBILE_PROCESSED: "
    f"{schema_counts['MOBILE_PROCESSED']}"
)
print(f"UNSUPPORTED: {schema_counts['UNSUPPORTED']}")

unsupported_results = [
    result
    for result in results
    if result["schema"] == "UNSUPPORTED"
]

if unsupported_results:
    print("\nUnsupported files:")

    for result in unsupported_results:
        print(
            f"{result['file_name']} | "
            f"Columns: {' | '.join(result['columns'])} | "
            f"Error: {result['error'] or 'None'}"
        )
else:
    print("\nAll files use a supported telemetry format.")

print(
    f"\nDetailed results saved at: "
    f"{OUTPUT_FILE.resolve()}"
)