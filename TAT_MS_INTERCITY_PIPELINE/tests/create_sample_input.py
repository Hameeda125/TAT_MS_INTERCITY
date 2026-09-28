from pathlib import Path

import pandas as pd


sample_folder = Path(
    "D:/TAT_MS/TAT_Data/Input/01-08-2026"
)
sample_folder.mkdir(parents=True, exist_ok=True)

sample_data = pd.DataFrame(
    {
        "lat": [
            13.0827,
            13.0828,
            13.0829,
            13.0830,
            13.0831,
            13.0832,
        ],
        "lon": [
            80.2707,
            80.2708,
            80.2709,
            80.2710,
            80.2711,
            80.2712,
        ],
        "timestamp": [
            "01-08-2026 08:00:00",
            "01-08-2026 08:00:04",
            "01-08-2026 08:00:09",
            "01-08-2026 08:00:15",
            "01-08-2026 08:00:18",
            "01-08-2026 08:00:25",
        ],
        "speed": [
            0.0,
            2.0,
            8.0,
            15.0,
            10.0,
            4.0,
        ],
    }
)

sample_file = sample_folder / "sample_ev_trip.xlsx"

sample_data.to_excel(
    sample_file,
    index=False,
    engine="openpyxl",
)

print("Sample test file created successfully")
print(f"File: {sample_file}")