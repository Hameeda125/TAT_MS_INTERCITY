from pathlib import Path

import pandas as pd


BUS_DATA_FOLDER = Path(
    r"C:\Users\JAVED\OneDrive\Desktop\Bus Data"
)

excel_files = sorted(BUS_DATA_FOLDER.rglob("*.xlsx"))

if not BUS_DATA_FOLDER.exists():
    raise FileNotFoundError(
        f"Bus Data folder was not found: {BUS_DATA_FOLDER}"
    )

if not excel_files:
    raise FileNotFoundError(
        f"No Excel files were found inside: {BUS_DATA_FOLDER}"
    )

sample_file = excel_files[0]

print("Real-data inspection started")
print(f"Excel files found: {len(excel_files)}")
print(f"File inspected: {sample_file}")

workbook = pd.ExcelFile(
    sample_file,
    engine="openpyxl",
)

print(f"Sheet names: {workbook.sheet_names}")

for sheet_name in workbook.sheet_names:
    dataframe = pd.read_excel(
        sample_file,
        sheet_name=sheet_name,
        engine="openpyxl",
    )

    print(f"\nSHEET: {sheet_name}")
    print(f"Rows: {len(dataframe)}")
    print(f"Columns: {len(dataframe.columns)}")

    print("\nColumn names:")
    for column_number, column_name in enumerate(
        dataframe.columns,
        start=1,
    ):
        print(f"{column_number}. {column_name}")

    print("\nFirst three records:")
    print(dataframe.head(3).to_string(index=False))

    print("\nDetected data types:")
    print(dataframe.dtypes.to_string())