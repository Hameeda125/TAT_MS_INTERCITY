from pathlib import Path

from config_loader import load_config
from file_discovery import discover_input_files
from telemetry_standardizer import standardize_telemetry_file


def validate_input_file(file_path, config):
    """
    Validate one telemetry file after converting it to the
    common standardized schema.

    The original Excel file is never modified.
    """

    file_path = Path(file_path)

    result = {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "status": "VALID",
        "source_schema": None,
        "row_count": 0,
        "valid_core_rows": 0,
        "missing_columns": [],
        "invalid_timestamps": 0,
        "invalid_latitudes": 0,
        "invalid_longitudes": 0,
        "invalid_speeds": 0,
        "issues": [],
    }

    try:
        dataframe, metadata = standardize_telemetry_file(
            file_path
        )

    except Exception as error:
        result["status"] = "INVALID"
        result["issues"].append(str(error))
        return result

    result["source_schema"] = metadata["source_schema"]
    result["row_count"] = len(dataframe)

    invalid_timestamp_mask = dataframe["Timestamp"].isna()

    invalid_latitude_mask = (
        dataframe["Latitude"].isna()
        | ~dataframe["Latitude"].between(
            config["cleaning"]["latitude_min"],
            config["cleaning"]["latitude_max"],
        )
    )

    invalid_longitude_mask = (
        dataframe["Longitude"].isna()
        | ~dataframe["Longitude"].between(
            config["cleaning"]["longitude_min"],
            config["cleaning"]["longitude_max"],
        )
    )

    invalid_speed_mask = (
        dataframe["Speed(km/h)"].isna()
        | dataframe["Speed(m/s)"].isna()
        | (dataframe["Speed(km/h)"] < 0)
        | (dataframe["Speed(m/s)"] < 0)
    )

    result["invalid_timestamps"] = int(
        invalid_timestamp_mask.sum()
    )

    result["invalid_latitudes"] = int(
        invalid_latitude_mask.sum()
    )

    result["invalid_longitudes"] = int(
        invalid_longitude_mask.sum()
    )

    result["invalid_speeds"] = int(
        invalid_speed_mask.sum()
    )

    valid_core_mask = ~(
        invalid_timestamp_mask
        | invalid_latitude_mask
        | invalid_longitude_mask
    )

    result["valid_core_rows"] = int(
        valid_core_mask.sum()
    )

    if result["invalid_timestamps"] > 0:
        result["issues"].append(
            f"{result['invalid_timestamps']} rows have "
            "invalid timestamps and will be removed "
            "during cleaning."
        )

    if result["invalid_latitudes"] > 0:
        result["issues"].append(
            f"{result['invalid_latitudes']} rows have "
            "invalid latitude values and will be removed "
            "during cleaning."
        )

    if result["invalid_longitudes"] > 0:
        result["issues"].append(
            f"{result['invalid_longitudes']} rows have "
            "invalid longitude values and will be removed "
            "during cleaning."
        )

    if result["invalid_speeds"] > 0:
        result["issues"].append(
            f"{result['invalid_speeds']} rows have missing "
            "or negative speeds and will be handled during "
            "cleaning."
        )

    minimum_rows = config["cleaning"][
        "minimum_valid_rows"
    ]

    if result["valid_core_rows"] < minimum_rows:
        result["status"] = "INVALID"

        result["issues"].append(
            f"Only {result['valid_core_rows']} rows have "
            "valid timestamps and coordinates; at least "
            f"{minimum_rows} are required."
        )

    return result


def validate_discovered_files(config):
    """Validate all discovered input files."""

    discovered_files = discover_input_files(config)

    return [
        validate_input_file(
            item["file_path"],
            config,
        )
        for item in discovered_files
    ]


if __name__ == "__main__":
    settings = load_config()

    validation_results = validate_discovered_files(
        settings
    )

    print("Input validation completed")
    print(
        f"Files checked: {len(validation_results)}"
    )

    valid_count = sum(
        result["status"] == "VALID"
        for result in validation_results
    )

    invalid_count = (
        len(validation_results) - valid_count
    )

    print(f"Valid files: {valid_count}")
    print(f"Invalid files: {invalid_count}")

    for result in validation_results:
        print(
            f"{result['status']} | "
            f"{result['file_name']} | "
            f"Schema: {result['source_schema']} | "
            f"Rows: {result['row_count']} | "
            f"Usable rows: {result['valid_core_rows']}"
        )

        for issue in result["issues"]:
            print(f"  - {issue}")
