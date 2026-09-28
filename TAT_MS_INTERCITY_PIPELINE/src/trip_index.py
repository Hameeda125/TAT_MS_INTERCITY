import hashlib
import os
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd

from config_loader import load_config
from trip_metadata import (
    extract_trip_metadata,
    find_quality_summary,
    load_cleaned_telemetry,
    load_cleaning_quality_lookup,
    discover_cleaned_files,
)


INDEX_SCHEMA_VERSION = "2.0"

REQUIRED_INDEX_COLUMNS = {
    "index_schema_version",
    "trip_id",
    "telemetry_checksum",
    "relative_path",
    "source_file_path",
    "cleaned_file_path",
    "vehicle_type",
    "start_time",
    "end_time",
}

CHECKSUM_COLUMNS = [
    "Timestamp",
    "Latitude",
    "Longitude",
    "Speed(m/s)",
    "Speed(km/h)",
]


def _timestamp_to_index_text(value):
    """Convert a timestamp to a CSV-stable string for the trip index."""

    if value is None or pd.isna(value):
        return None

    return pd.Timestamp(value).isoformat(sep=" ")


def _normalise_path_key(path_value):
    """Return a stable, case-insensitive key for a relative path."""

    if path_value is None or pd.isna(path_value):
        return ""

    return str(path_value).replace("\\", "/").rstrip("/").lower()


def calculate_file_checksum(file_path):
    """Calculate a SHA-256 checksum without changing the file."""

    file_path = Path(file_path)

    if not file_path.exists():
        return None

    checksum = hashlib.sha256()

    with file_path.open("rb") as file:
        for block in iter(lambda: file.read(65536), b""):
            checksum.update(block)

    return checksum.hexdigest()


def calculate_telemetry_checksum(dataframe):
    """Create a stable checksum from standardized cleaned telemetry."""

    missing_columns = [
        column
        for column in CHECKSUM_COLUMNS
        if column not in dataframe.columns
    ]

    if missing_columns:
        raise ValueError(
            "Cannot calculate telemetry checksum; missing columns: "
            + ", ".join(missing_columns)
        )

    canonical = dataframe[CHECKSUM_COLUMNS].copy()
    canonical["Timestamp"] = pd.to_datetime(
        canonical["Timestamp"],
        errors="coerce",
    ).dt.strftime("%Y-%m-%dT%H:%M:%S.%f")

    for column in CHECKSUM_COLUMNS[1:]:
        canonical[column] = pd.to_numeric(
            canonical[column],
            errors="coerce",
        ).map(
            lambda value: (
                ""
                if pd.isna(value)
                else format(float(value), ".12g")
            )
        )

    canonical_text = canonical.to_csv(
        index=False,
        header=False,
        lineterminator="\n",
    )

    return hashlib.sha256(
        canonical_text.encode("utf-8")
    ).hexdigest()


def get_trip_index_path(config):
    """Return the permanent trip-index CSV path."""

    results_folder = Path(config["project"]["results_folder"])
    system_folder = results_folder / "System"
    system_folder.mkdir(parents=True, exist_ok=True)

    return system_folder / "trip_index.csv"


def _backup_legacy_index(index_path):
    """Preserve an incompatible older index before rebuilding it."""

    backup_folder = index_path.parent / "Backups"
    backup_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_folder / (
        f"trip_index_legacy_{timestamp}.csv"
    )
    shutil.copy2(index_path, backup_path)

    return backup_path


def load_trip_index(index_path):
    """Load a compatible index or safely prepare to rebuild an old one."""

    index_path = Path(index_path)

    if not index_path.exists():
        return pd.DataFrame(), None

    try:
        existing_index = pd.read_csv(
            index_path,
            dtype={
                "trip_id": "string",
                "telemetry_checksum": "string",
                "relative_path": "string",
            },
        )
    except pd.errors.EmptyDataError:
        return pd.DataFrame(), None

    if not REQUIRED_INDEX_COLUMNS.issubset(
        existing_index.columns
    ):
        backup_path = _backup_legacy_index(index_path)
        return pd.DataFrame(), backup_path

    schema_versions = set(
        existing_index["index_schema_version"]
        .dropna()
        .astype(str)
    )

    if schema_versions and schema_versions != {INDEX_SCHEMA_VERSION}:
        backup_path = _backup_legacy_index(index_path)
        return pd.DataFrame(), backup_path

    # Pandas 3 uses strict inferred string dtypes when reading CSV files.
    # Object dtype keeps incremental updates safe when a record contains
    # Python scalars produced during fresh metadata extraction.
    return existing_index.astype(object), None


def _prepare_record(
    item,
    dataframe,
    config,
    quality_lookup,
    now_text,
):
    """Build one complete trip-index record."""

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

    telemetry_checksum = calculate_telemetry_checksum(dataframe)
    source_checksum = calculate_file_checksum(
        item["source_file_path"]
    )
    cleaned_checksum = calculate_file_checksum(
        item["cleaned_file_path"]
    )

    metadata.update(
        {
            "index_schema_version": INDEX_SCHEMA_VERSION,
            "trip_id": f"trip_{telemetry_checksum[:16]}",
            "telemetry_checksum": telemetry_checksum,
            "source_file_checksum": source_checksum,
            "cleaned_file_checksum": cleaned_checksum,
            "source_file_available": bool(source_checksum),
            "indexed_at": now_text,
            "updated_at": now_text,
        }
    )

    # Store timestamps as text, just as they will appear after the index is
    # read back from CSV.  This prevents pandas from trying to insert a
    # Timestamp object into an existing strict string column during updates.
    metadata["start_time"] = _timestamp_to_index_text(
        metadata.get("start_time")
    )
    metadata["end_time"] = _timestamp_to_index_text(
        metadata.get("end_time")
    )

    return metadata


def _records_equal(existing_record, new_record):
    """Check whether a stored record already has the current metadata."""

    ignored_fields = {"indexed_at", "updated_at"}

    for key, new_value in new_record.items():
        if key in ignored_fields:
            continue

        old_value = existing_record.get(key)

        if pd.isna(old_value) and pd.isna(new_value):
            continue

        if str(old_value) != str(new_value):
            return False

    return True


def _merge_record(existing_index, record):
    """Add, update, or identify one unchanged/duplicate trip."""

    if existing_index.empty:
        return pd.DataFrame([record]), "ADDED"

    telemetry_matches = existing_index.index[
        existing_index["telemetry_checksum"].astype(str)
        == str(record["telemetry_checksum"])
    ].tolist()

    path_key = _normalise_path_key(record["relative_path"])
    path_matches = existing_index.index[
        existing_index["relative_path"]
        .map(_normalise_path_key)
        == path_key
    ].tolist()

    if path_matches:
        target_index = path_matches[0]
        existing_record = existing_index.loc[target_index].to_dict()

        if _records_equal(existing_record, record):
            return existing_index, "UNCHANGED"

        record["indexed_at"] = existing_record.get(
            "indexed_at",
            record["indexed_at"],
        )

        for column, value in record.items():
            existing_index.at[target_index, column] = value

        return existing_index, "UPDATED"

    if telemetry_matches:
        target_index = telemetry_matches[0]
        existing_record = existing_index.loc[target_index].to_dict()
        old_source_path = Path(
            str(existing_record.get("source_file_path", ""))
        )
        old_cleaned_path = Path(
            str(existing_record.get("cleaned_file_path", ""))
        )

        # If the earlier path no longer exists, this is most likely a rename.
        # Refresh the stored paths and metadata while keeping the trip identity.
        if not old_source_path.exists() and not old_cleaned_path.exists():
            record["trip_id"] = existing_record.get(
                "trip_id",
                record["trip_id"],
            )
            record["indexed_at"] = existing_record.get(
                "indexed_at",
                record["indexed_at"],
            )

            for column, value in record.items():
                existing_index.at[target_index, column] = value

            return existing_index, "UPDATED"

        return existing_index, "DUPLICATE"

    updated_index = pd.concat(
        [existing_index, pd.DataFrame([record])],
        ignore_index=True,
    )

    return updated_index, "ADDED"


def save_trip_index(dataframe, index_path):
    """Write the trip index atomically so a failed save cannot corrupt it."""

    index_path = Path(index_path)
    index_path.parent.mkdir(parents=True, exist_ok=True)

    dataframe = dataframe.copy()
    dataframe = dataframe.sort_values(
        by=["date_folder", "start_time", "file_name"],
        kind="stable",
        na_position="last",
    ).reset_index(drop=True)

    with NamedTemporaryFile(
        prefix=".trip_index_",
        suffix=".csv",
        dir=index_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)

    try:
        dataframe.to_csv(
            temporary_path,
            index=False,
            encoding="utf-8-sig",
            date_format="%Y-%m-%d %H:%M:%S",
        )
        os.replace(temporary_path, index_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _remove_stale_index_records(dataframe, current_relative_paths):
    """Remove index rows whose cleaned/source path is no longer current."""

    if dataframe.empty:
        return dataframe, []

    current_keys = {
        _normalise_path_key(path)
        for path in current_relative_paths
    }
    active_mask = dataframe["relative_path"].map(
        _normalise_path_key
    ).isin(current_keys)
    removed_records = dataframe.loc[~active_mask].to_dict(
        orient="records"
    )
    synchronized_index = dataframe.loc[active_mask].copy()
    synchronized_index = synchronized_index.reset_index(drop=True)

    return synchronized_index, removed_records


def update_trip_index(config):
    """Create or incrementally update the permanent cleaned-trip index."""

    index_path = get_trip_index_path(config)
    existing_index, legacy_backup_path = load_trip_index(index_path)
    cleaned_files = discover_cleaned_files(config)
    quality_lookup = load_cleaning_quality_lookup(config)
    now_text = datetime.now().isoformat(timespec="seconds")

    print("Trip-index update started")
    print(f"Cleaned files discovered: {len(cleaned_files)}")

    if legacy_backup_path is not None:
        print(
            "OLD INDEX BACKED UP | Incompatible earlier index | "
            f"{legacy_backup_path}"
        )

    if not quality_lookup["report_found"]:
        print(
            "WARNING | Cleaning audit not found. The index will be "
            "created, but removal counts will be unavailable."
        )

    updated_index = existing_index.copy()
    added_count = 0
    updated_count = 0
    unchanged_count = 0
    duplicate_count = 0
    failed_count = 0
    unknown_vehicle_count = 0
    removed_count = 0

    if len(cleaned_files) == 0 and not existing_index.empty:
        failed_count = 1
        print(
            "INDEX SYNC BLOCKED | No cleaned files were discovered, "
            "but the existing index contains records. The index was "
            "preserved to prevent an accidental empty-folder reset."
        )

    for item in cleaned_files:
        try:
            dataframe = load_cleaned_telemetry(
                item["cleaned_file_path"]
            )
            record = _prepare_record(
                item=item,
                dataframe=dataframe,
                config=config,
                quality_lookup=quality_lookup,
                now_text=now_text,
            )

            if record["vehicle_type"] == "UNKNOWN":
                unknown_vehicle_count += 1

            updated_index, action = _merge_record(
                updated_index,
                record,
            )

            if action == "ADDED":
                added_count += 1
                print(
                    f"ADDED | {record['trip_id']} | "
                    f"{item['relative_path']}"
                )
            elif action == "UPDATED":
                updated_count += 1
                print(
                    f"UPDATED | {record['trip_id']} | "
                    f"{item['relative_path']}"
                )
            elif action == "DUPLICATE":
                duplicate_count += 1
                print(
                    f"DUPLICATE SKIPPED | {item['relative_path']}"
                )
            else:
                unchanged_count += 1
                print(f"UNCHANGED | {item['relative_path']}")

        except Exception as error:
            failed_count += 1
            print(f"FAILED | {item['relative_path']} | {error}")

    if failed_count > 0:
        print(
            "Index not saved because one or more cleaned files failed. "
            "Correct the reported files and run again."
        )
    else:
        current_relative_paths = {
            item["relative_path"]
            for item in cleaned_files
        }
        updated_index, removed_records = (
            _remove_stale_index_records(
                updated_index,
                current_relative_paths,
            )
        )
        removed_count = len(removed_records)

        for removed_record in removed_records:
            print(
                "REMOVED STALE INDEX RECORD | "
                f"{removed_record.get('relative_path')}"
            )

        save_trip_index(updated_index, index_path)

    return {
        "cleaned_files_discovered": len(cleaned_files),
        "new_trips_added": added_count,
        "trips_updated": updated_count,
        "unchanged_trips": unchanged_count,
        "duplicates_skipped": duplicate_count,
        "stale_records_removed": removed_count,
        "files_failed": failed_count,
        "unknown_vehicle_types": unknown_vehicle_count,
        "total_indexed_trips": len(updated_index),
        "index_saved": failed_count == 0,
        "index_path": str(index_path),
        "legacy_backup_path": (
            str(legacy_backup_path)
            if legacy_backup_path is not None
            else None
        ),
    }


if __name__ == "__main__":
    settings = load_config()
    summary = update_trip_index(settings)

    print("\nTrip-index update completed")
    print(
        "Cleaned files discovered: "
        f"{summary['cleaned_files_discovered']}"
    )
    print(f"New trips added: {summary['new_trips_added']}")
    print(f"Trips updated: {summary['trips_updated']}")
    print(f"Unchanged trips: {summary['unchanged_trips']}")
    print(f"Duplicates skipped: {summary['duplicates_skipped']}")
    print(
        "Stale index records removed: "
        f"{summary['stale_records_removed']}"
    )
    print(f"Files failed: {summary['files_failed']}")
    print(
        "Unknown vehicle types: "
        f"{summary['unknown_vehicle_types']}"
    )
    print(
        "Total indexed trips: "
        f"{summary['total_indexed_trips']}"
    )

    if summary["index_saved"]:
        print(f"Index saved at: {summary['index_path']}")