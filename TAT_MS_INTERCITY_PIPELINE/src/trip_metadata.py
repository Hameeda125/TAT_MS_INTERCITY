import re

from datetime import datetime

from math import atan2, cos, radians, sin, sqrt

from pathlib import Path



import pandas as pd



from config_loader import load_config





CLEANED_SHEET_NAME = "Cleaned Data"



TIMESTAMP_COLUMN = "Timestamp"

LATITUDE_COLUMN = "Latitude"

LONGITUDE_COLUMN = "Longitude"

SPEED_KMH_COLUMN = "Speed(km/h)"

SPEED_MPS_COLUMN = "Speed(m/s)"



REQUIRED_CLEANED_COLUMNS = [

    TIMESTAMP_COLUMN,

    LATITUDE_COLUMN,

    LONGITUDE_COLUMN,

    SPEED_MPS_COLUMN,

    SPEED_KMH_COLUMN,

]



QUALITY_SUMMARY_FIELDS = [

    "source_schema",

    "source_sheet",

    "original_rows",

    "invalid_speeds_recovered",

    "invalid_speeds_unresolved",

    "unrecoverable_rows_removed",

    "duplicate_rows_removed",

    "total_rows_removed",

    "cleaned_rows",

]





def _normalise_path_key(path_value):

    """Return a case-insensitive path key for audit matching."""



    if path_value is None or pd.isna(path_value):

        return ""



    return str(path_value).replace("\\\\", "/").rstrip("/").lower()





def _safe_float(value):

    """Return a finite float, or None when no numeric value exists."""



    if value is None or pd.isna(value):

        return None



    return float(value)





def _rounded_float(value, decimals=3):

    """Return a rounded float while preserving unavailable values."""



    numeric_value = _safe_float(value)



    if numeric_value is None:

        return None



    return round(numeric_value, decimals)





def haversine_distance_km(lat1, lon1, lat2, lon2):

    """Calculate the great-circle distance between two GPS points."""



    earth_radius_km = 6371.0088



    latitude_1 = radians(float(lat1))

    latitude_2 = radians(float(lat2))

    latitude_difference = radians(float(lat2) - float(lat1))

    longitude_difference = radians(float(lon2) - float(lon1))



    haversine_value = (

        sin(latitude_difference / 2) ** 2

        + cos(latitude_1)

        * cos(latitude_2)

        * sin(longitude_difference / 2) ** 2

    )



    # Floating-point rounding can very slightly move the value outside

    # the valid [0, 1] range used by the square roots below.

    haversine_value = min(1.0, max(0.0, haversine_value))



    central_angle = 2 * atan2(

        sqrt(haversine_value),

        sqrt(1.0 - haversine_value),

    )



    return earth_radius_km * central_angle





def calculate_trip_distance_km(

    dataframe,

    latitude_column=LATITUDE_COLUMN,

    longitude_column=LONGITUDE_COLUMN,

):

    """Calculate cumulative point-to-point distance from cleaned GPS data."""



    coordinates = dataframe[

        [latitude_column, longitude_column]

    ].dropna()



    if len(coordinates) < 2:

        return 0.0



    total_distance_km = 0.0

    previous_latitude = None

    previous_longitude = None



    for latitude, longitude in coordinates.itertuples(

        index=False,

        name=None,

    ):

        if previous_latitude is not None:

            total_distance_km += haversine_distance_km(

                previous_latitude,

                previous_longitude,

                latitude,

                longitude,

            )



        previous_latitude = latitude

        previous_longitude = longitude



    return total_distance_km





def discover_cleaned_files(config):

    """Find cleaned Excel files while preserving their source paths."""



    input_root = Path(config["project"]["input_folder"])

    results_root = Path(config["project"]["results_folder"])

    cleaned_root = results_root / "Cleaned"

    date_format = config["files"]["date_folder_format"]



    accepted_extensions = {

        extension.lower()

        for extension in config["files"]["accepted_extensions"]

    }



    if not cleaned_root.exists():

        raise FileNotFoundError(

            f"Cleaned files folder does not exist: {cleaned_root}"

        )



    date_folders = []



    for folder in cleaned_root.iterdir():

        if not folder.is_dir():

            continue



        try:

            folder_date = datetime.strptime(folder.name, date_format)

        except ValueError:

            continue



        date_folders.append((folder_date, folder))



    date_folders.sort(key=lambda item: item[0])



    cleaned_files = []



    for _, date_folder in date_folders:

        for cleaned_path in sorted(date_folder.rglob("*")):

            if not cleaned_path.is_file():

                continue



            if cleaned_path.name.startswith("~$"):

                continue



            if cleaned_path.suffix.lower() not in accepted_extensions:

                continue



            relative_path = cleaned_path.relative_to(cleaned_root)

            source_path = input_root / relative_path



            cleaned_files.append(

                {

                    "date_folder": date_folder.name,

                    "file_name": cleaned_path.name,

                    "relative_path": str(relative_path),

                    "source_file_path": str(source_path),

                    "cleaned_file_path": str(cleaned_path),

                    "cleaned_file_size_bytes": (

                        cleaned_path.stat().st_size

                    ),

                }

            )



    return cleaned_files





def load_cleaned_telemetry(cleaned_file_path):

    """Load and validate one workbook produced by cleaned_file_writer.py."""



    cleaned_file_path = Path(cleaned_file_path)



    try:

        dataframe = pd.read_excel(

            cleaned_file_path,

            sheet_name=CLEANED_SHEET_NAME,

            engine="openpyxl",

        )

    except ValueError as error:

        raise ValueError(

            f"Worksheet '{CLEANED_SHEET_NAME}' was not found in "

            f"{cleaned_file_path.name}."

        ) from error



    missing_columns = [

        column

        for column in REQUIRED_CLEANED_COLUMNS

        if column not in dataframe.columns

    ]



    if missing_columns:

        raise ValueError(

            f"Cleaned file {cleaned_file_path.name} is missing: "

            f"{', '.join(missing_columns)}"

        )



    dataframe = dataframe.copy()

    dataframe[TIMESTAMP_COLUMN] = pd.to_datetime(

        dataframe[TIMESTAMP_COLUMN],

        errors="coerce",

        dayfirst=True,

    )



    for column in [

        LATITUDE_COLUMN,

        LONGITUDE_COLUMN,

        SPEED_MPS_COLUMN,

        SPEED_KMH_COLUMN,

    ]:

        dataframe[column] = pd.to_numeric(

            dataframe[column],

            errors="coerce",

        )



    invalid_core_mask = (

        dataframe[TIMESTAMP_COLUMN].isna()

        | dataframe[LATITUDE_COLUMN].isna()

        | dataframe[LONGITUDE_COLUMN].isna()

    )



    if bool(invalid_core_mask.any()):

        raise ValueError(

            f"Cleaned file {cleaned_file_path.name} still contains "

            f"{int(invalid_core_mask.sum())} rows with an invalid "

            "timestamp or coordinate. Rerun the cleaning stages."

        )



    dataframe = dataframe.sort_values(

        by=TIMESTAMP_COLUMN,

        kind="stable",

    ).reset_index(drop=True)



    return dataframe





def load_cleaning_quality_lookup(config):

    """Load per-file cleaning summaries from the cleaning audit workbook."""



    results_root = Path(config["project"]["results_folder"])

    report_name = config.get("outputs", {}).get(

        "cleaning_audit_report_name",

        "Data_Cleaning_Audit.xlsx",

    )

    report_path = results_root / report_name



    if not report_path.exists():

        return {

            "report_path": str(report_path),

            "report_found": False,

            "by_source_path": {},

            "by_unique_file_name": {},

        }



    try:

        audit_dataframe = pd.read_excel(

            report_path,

            sheet_name="File Summary",

            engine="openpyxl",

        )

    except (ValueError, OSError) as error:

        raise ValueError(

            f"Could not read the File Summary sheet from "

            f"{report_path}: {error}"

        ) from error



    required_audit_columns = {"file_name", "source_file"}

    missing_columns = required_audit_columns.difference(

        audit_dataframe.columns

    )



    if missing_columns:

        raise ValueError(

            "The cleaning audit File Summary sheet is missing: "

            + ", ".join(sorted(missing_columns))

        )



    by_source_path = {}

    records_by_file_name = {}



    for record in audit_dataframe.to_dict(orient="records"):

        source_key = _normalise_path_key(record.get("source_file"))



        if source_key:

            by_source_path[source_key] = record



        file_name_key = str(record.get("file_name", "")).lower()

        records_by_file_name.setdefault(file_name_key, []).append(record)



    by_unique_file_name = {

        file_name: records[0]

        for file_name, records in records_by_file_name.items()

        if file_name and len(records) == 1

    }



    return {

        "report_path": str(report_path),

        "report_found": True,

        "by_source_path": by_source_path,

        "by_unique_file_name": by_unique_file_name,

    }





def find_quality_summary(cleaned_item, quality_lookup):

    """Match a cleaned file to its cleaning-audit summary."""



    source_key = _normalise_path_key(

        cleaned_item["source_file_path"]

    )



    quality_summary = quality_lookup["by_source_path"].get(

        source_key

    )



    if quality_summary is None:

        quality_summary = quality_lookup[

            "by_unique_file_name"

        ].get(cleaned_item["file_name"].lower())



    if quality_summary is None:

        return {}



    return {

        field: quality_summary.get(field)

        for field in QUALITY_SUMMARY_FIELDS

    }





def infer_vehicle_type(file_name, relative_path=""):

    """Infer EV or ICE only from unambiguous filename/path tokens."""



    text = f"{relative_path} {file_name}".lower()

    tokens = {

        token

        for token in re.split(r"[^a-z0-9]+", text)

        if token

    }



    ev_detected = bool(tokens.intersection({"ev", "electric"}))

    ice_detected = bool(

        tokens.intersection({"ice", "diesel", "combustion"})

    )



    if ev_detected and not ice_detected:

        return "EV"



    if ice_detected and not ev_detected:

        return "ICE"



    return "UNKNOWN"





def infer_route_name(file_name):

    """Create a provisional route label from a structured filename."""



    route_text = Path(file_name).stem.lower()

    tokens = [

        token

        for token in re.split(r"[^a-z0-9]+", route_text)

        if token

    ]



    ignored_tokens = {

        "ev",

        "ice",

        "electric",

        "diesel",

        "bus",

        "truck",

        "raw",

        "processed",

        "mobile",

        "device",

    }



    route_tokens = []



    for token in tokens:

        if token in ignored_tokens:

            continue



        if re.fullmatch(r"day\d+", token):

            continue



        if re.fullmatch(r"trip\d+", token):

            continue



        if re.fullmatch(r"\d{1,4}", token):

            continue



        route_tokens.append(token)



    if not route_tokens:

        return "UNKNOWN"



    return "_".join(route_tokens)





def _quality_value(quality_summary, field):

    """Return an integer quality value or None when it is unavailable."""



    if not quality_summary:

        return None



    value = quality_summary.get(field)



    if value is None or pd.isna(value):

        return None



    return int(value)





def extract_trip_metadata(

    dataframe,

    file_name,

    date_folder,

    config=None,

    quality_summary=None,

    source_file_path=None,

    cleaned_file_path=None,

    relative_path=None,

):

    """Extract full-trip information from standardized cleaned telemetry."""



    del config  # Kept in the interface for compatibility with trip_index.py.



    if dataframe.empty:

        raise ValueError(

            f"No cleaned telemetry rows are available for {file_name}."

        )



    missing_columns = [

        column

        for column in REQUIRED_CLEANED_COLUMNS

        if column not in dataframe.columns

    ]



    if missing_columns:

        raise ValueError(

            f"Cannot extract metadata for {file_name}; missing "

            f"columns: {', '.join(missing_columns)}"

        )



    dataframe = dataframe.sort_values(

        by=TIMESTAMP_COLUMN,

        kind="stable",

    ).reset_index(drop=True)



    start_row = dataframe.iloc[0]

    end_row = dataframe.iloc[-1]



    start_time = start_row[TIMESTAMP_COLUMN]

    end_time = end_row[TIMESTAMP_COLUMN]

    duration_minutes = (

        end_time - start_time

    ).total_seconds() / 60.0



    time_gaps_seconds = (

        dataframe[TIMESTAMP_COLUMN]

        .diff()

        .dt.total_seconds()

        .dropna()

    )



    speeds = pd.to_numeric(

        dataframe[SPEED_KMH_COLUMN],

        errors="coerce",

    )

    valid_speeds = speeds.dropna()

    moving_speeds = valid_speeds[valid_speeds > 0]



    distance_km = calculate_trip_distance_km(dataframe)

    straight_line_distance_km = haversine_distance_km(

        start_row[LATITUDE_COLUMN],

        start_row[LONGITUDE_COLUMN],

        end_row[LATITUDE_COLUMN],

        end_row[LONGITUDE_COLUMN],

    )



    status_counts = pd.Series(dtype="int64")



    if "Speed Recovery Status" in dataframe.columns:

        status_counts = (

            dataframe["Speed Recovery Status"]

            .astype("string")

            .value_counts(dropna=False)

        )



    derived_recovered_speeds = int(

        status_counts.get("TIME_INTERPOLATED", 0)

    )

    derived_unresolved_speeds = int(

        status_counts.get("UNRESOLVED", 0)

    )



    quality_summary = quality_summary or {}

    audit_recovered_speeds = _quality_value(

        quality_summary,

        "invalid_speeds_recovered",

    )

    audit_unresolved_speeds = _quality_value(

        quality_summary,

        "invalid_speeds_unresolved",

    )



    recovered_speeds = (

        audit_recovered_speeds

        if audit_recovered_speeds is not None

        else derived_recovered_speeds

    )

    unresolved_speeds = (

        audit_unresolved_speeds

        if audit_unresolved_speeds is not None

        else derived_unresolved_speeds

    )



    quality_summary_available = bool(quality_summary)

    path_for_vehicle_detection = relative_path or ""



    metadata = {

        "date_folder": date_folder,

        "file_name": file_name,

        "relative_path": relative_path,

        "source_file_path": source_file_path,

        "cleaned_file_path": cleaned_file_path,

        "vehicle_type": infer_vehicle_type(

            file_name,

            path_for_vehicle_detection,

        ),

        "route_name": infer_route_name(file_name),

        "route_name_source": "FILENAME_PROVISIONAL",

        "start_time": start_time,

        "end_time": end_time,

        "start_latitude": round(

            float(start_row[LATITUDE_COLUMN]),

            6,

        ),

        "start_longitude": round(

            float(start_row[LONGITUDE_COLUMN]),

            6,

        ),

        "end_latitude": round(

            float(end_row[LATITUDE_COLUMN]),

            6,

        ),

        "end_longitude": round(

            float(end_row[LONGITUDE_COLUMN]),

            6,

        ),

        "duration_minutes": round(float(duration_minutes), 3),

        # Kept for compatibility with the existing pipeline terminology.

        "tat_minutes": round(float(duration_minutes), 3),

        "distance_km": round(float(distance_km), 3),

        "straight_line_distance_km": round(

            float(straight_line_distance_km),

            3,

        ),

        "minimum_speed_kmh": _rounded_float(valid_speeds.min()),

        "maximum_speed_kmh": _rounded_float(valid_speeds.max()),

        "average_speed_kmh": _rounded_float(valid_speeds.mean()),

        "average_moving_speed_kmh": _rounded_float(

            moving_speeds.mean()

        ),

        "telemetry_points": int(len(dataframe)),

        "valid_speed_points": int(valid_speeds.count()),

        "missing_speed_points": int(speeds.isna().sum()),

        "minimum_time_gap_seconds": _rounded_float(

            time_gaps_seconds.min()

        ),

        "maximum_time_gap_seconds": _rounded_float(

            time_gaps_seconds.max()

        ),

        "average_time_gap_seconds": _rounded_float(

            time_gaps_seconds.mean()

        ),

        "median_time_gap_seconds": _rounded_float(

            time_gaps_seconds.median()

        ),

        "source_schema": quality_summary.get("source_schema"),

        "source_sheet": quality_summary.get("source_sheet"),

        "original_rows": _quality_value(

            quality_summary,

            "original_rows",

        ),

        "invalid_speeds_recovered": recovered_speeds,

        "invalid_speeds_unresolved": unresolved_speeds,

        "unrecoverable_rows_removed": _quality_value(

            quality_summary,

            "unrecoverable_rows_removed",

        ),

        "duplicate_rows_removed": _quality_value(

            quality_summary,

            "duplicate_rows_removed",

        ),

        "total_rows_removed": _quality_value(

            quality_summary,

            "total_rows_removed",

        ),

        "cleaned_rows": int(len(dataframe)),

        "quality_summary_available": quality_summary_available,

    }



    return metadata





def extract_all_cleaned_trip_metadata(config):

    """Extract metadata for every saved cleaned workbook."""



    cleaned_files = discover_cleaned_files(config)

    quality_lookup = load_cleaning_quality_lookup(config)



    print("Cleaned-trip metadata extraction started")

    print(f"Cleaned files discovered: {len(cleaned_files)}")



    if not quality_lookup["report_found"]:

        print(

            "WARNING | Cleaning audit was not found. Removal counts "

            "will be unavailable."

        )



    records = []

    failed_count = 0



    for item in cleaned_files:

        try:

            dataframe = load_cleaned_telemetry(

                item["cleaned_file_path"]

            )

            quality_summary = find_quality_summary(

                item,

                quality_lookup,

            )



            metadata = extract_trip_metadata(

                dataframe=dataframe,

                file_name=item["file_name"],

                date_folder=item["date_folder"],

                config=config,

                quality_summary=quality_summary,

                source_file_path=item["source_file_path"],

                cleaned_file_path=item["cleaned_file_path"],

                relative_path=item["relative_path"],

            )



            records.append(metadata)



            print(

                f"EXTRACTED | {item['relative_path']} | "

                f"Type: {metadata['vehicle_type']} | "

                f"Duration: {metadata['duration_minutes']} min | "

                f"Distance: {metadata['distance_km']} km"

            )



        except Exception as error:

            failed_count += 1

            print(f"FAILED | {item['relative_path']} | {error}")



    return {

        "records": records,

        "files_discovered": len(cleaned_files),

        "metadata_extracted": len(records),

        "files_failed": failed_count,

        "quality_audit_found": quality_lookup["report_found"],

        "quality_audit_path": quality_lookup["report_path"],

    }





if __name__ == "__main__":

    settings = load_config()

    summary = extract_all_cleaned_trip_metadata(settings)



    print("\nTrip-metadata extraction completed")

    print(f"Cleaned files discovered: {summary['files_discovered']}")

    print(f"Metadata extracted: {summary['metadata_extracted']}")

    print(f"Files failed: {summary['files_failed']}")



    if summary["quality_audit_found"]:

        print(

            "Cleaning audit used: "

            f"{summary['quality_audit_path']}"

        )