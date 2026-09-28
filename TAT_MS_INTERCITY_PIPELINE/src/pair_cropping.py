"""Validate EV-ICE routes and create pair-specific common-route workbooks.

This is Stage 3 of the TAT pipeline.  Run it after ``candidate_pairing.py``.
It deliberately does not calculate interval features, driving phases, TAT
totals, comparison results, or charts.  Those remain separate later stages.

Input
-----
``Results/Pairing/Pairing_Candidates.xlsx`` (sheet ``Candidate Pairs``)

Outputs
-------
``Results/Paired/<pair date folder>/PAIR_###_<class>_<route>.xlsx``
    01_EV_Cropped_Data
    02_ICE_Cropped_Data
    03_Pair_Crop_Info
    03A_Distance_Comparison

``Results/Logs/Pair_Cropping_Log.xlsx``
``Results/System/pair_index.csv``

Important scientific choices
----------------------------
* Route alignment is direction-aware and order-preserving.  Resampled points
  are used only to find the common section; they are never written as ordinary
  telemetry records.
* A candidate is accepted only when the EV and ICE share one continuous route
  section, satisfy the overlap/heading/coverage rules, and have comparable
  cropped distances.  There is no full-trip fallback for a weak route match.
* Original cleaned telemetry rows inside the accepted section are preserved.
  Only two explicitly labelled boundary rows may be interpolated.
* Both cropped sheets receive exactly the same start coordinate and exactly the
  same end coordinate.  Each boundary coordinate is the midpoint of a close
  EV/ICE matched pair, so the adjustment to either source route is small.
* Equal distance means the two independently GPS-calculated cropped distances
  differ by no more than max(0.5 km, 2% of the shorter crop).  Artificially
  moving all interior points merely to force a zero numerical difference would
  corrupt the observed routes and is deliberately avoided.
* A valid trip may appear in more than one distinct pair-specific workbook.
  Selection of a later one-to-one PRIMARY analytical sample remains a separate
  decision and is not performed by this stage.
* ``PAIR_001`` numbering is assigned after the best available crop has been
  identified.  Strict/non-strict classes from candidate pairing are preserved.
* A same-date pair is stored under its single input date folder.  A cross-date
  pair is stored once under a chronological ``date1_to_date2`` folder.  Both
  source dates are also recorded separately in the workbook and pair index.
"""

from __future__ import annotations

import bisect
import math
import os
import re
import shutil
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from config_loader import load_config


CANDIDATE_WORKBOOK_NAME = "Pairing_Candidates.xlsx"
CANDIDATE_SHEET_NAME = "Candidate Pairs"
SCRIPT_VERSION = "2026-08-09_STRICT_COMMON_ROUTE_DISTANCE_EVIDENCE"
VERIFIED_INDEX_NAME = "verified_trip_index.csv"
PAIR_INDEX_NAME = "pair_index.csv"
CROPPING_LOG_NAME = "Pair_Cropping_Log.xlsx"
CLEANED_SHEET_NAME = "Cleaned Data"

EV_SHEET_NAME = "01_EV_Cropped_Data"
ICE_SHEET_NAME = "02_ICE_Cropped_Data"
INFO_SHEET_NAME = "03_Pair_Crop_Info"
DISTANCE_SHEET_NAME = "03A_Distance_Comparison"

TIMESTAMP_COLUMN = "Timestamp"
LATITUDE_COLUMN = "Latitude"
LONGITUDE_COLUMN = "Longitude"
SPEED_MPS_COLUMN = "Speed(m/s)"
SPEED_KMH_COLUMN = "Speed(km/h)"
ADDRESS_COLUMN = "Address"

EARTH_RADIUS_M = 6_371_008.8
STRICT_CLASSES = {
    "STRICT",
    "STRICT_CANDIDATE",
    "PRIMARY_STRICT",
    "PRIMARY_STRICT_CANDIDATE",
}
NON_STRICT_CLASSES = {
    "NONSTRICT",
    "NON_STRICT",
    "NONSTRICT_CANDIDATE",
    "NON_STRICT_CANDIDATE",
    "SECONDARY_NONSTRICT",
    "SECONDARY_NON_STRICT",
}

PAIRABLE_ROUTE_STATUSES = {"PASSED_ORDERED_ROUTE_VALIDATION"}

AUDIT_COLUMNS = [
    "candidate_id",
    "candidate_class",
    "final_pair_class",
    "non_strict_category",
    "candidate_reason",
    "route_direction",
    "ev_verified_trip_id",
    "ice_verified_trip_id",
    "ev_file_name",
    "ice_file_name",
    "ev_cleaned_file_path",
    "ice_cleaned_file_path",
    "candidate_same_calendar_day",
    "candidate_same_day_type",
    "candidate_start_clock_difference_minutes",
    "ev_verified_start",
    "ev_verified_end",
    "ice_verified_start",
    "ice_verified_end",
    "gps_route_validation_status",
    "route_validation_quality",
    "route_validation_reason",
    "crop_selection_method",
    "route_point_tolerance_m",
    "route_resample_spacing_m",
    "ev_verified_datapoints_loaded",
    "ice_verified_datapoints_loaded",
    "ev_verified_full_haversine_distance_km",
    "ice_verified_full_haversine_distance_km",
    "ordered_match_count",
    "median_heading_difference_degrees",
    "ev_ordered_overlap_fraction_within_crop",
    "ice_ordered_overlap_fraction_within_crop",
    "minimum_ordered_overlap_fraction_within_crop",
    "maximum_internal_unmatched_distance_m",
    "common_route_length_km",
    "shorter_trip_common_coverage_fraction",
    "ev_crop_coverage_fraction_of_verified_trip",
    "ice_crop_coverage_fraction_of_verified_trip",
    "comparison_scope",
    "common_start_anchor_latitude",
    "common_start_anchor_longitude",
    "common_end_anchor_latitude",
    "common_end_anchor_longitude",
    "start_anchor_ev_ice_separation_m",
    "end_anchor_ev_ice_separation_m",
    "ev_crop_start_progress_m",
    "ev_crop_end_progress_m",
    "ice_crop_start_progress_m",
    "ice_crop_end_progress_m",
    "ev_cropped_datapoints",
    "ice_cropped_datapoints",
    "ev_cropped_start_timestamp",
    "ev_cropped_end_timestamp",
    "ice_cropped_start_timestamp",
    "ice_cropped_end_timestamp",
    "ev_cropped_elapsed_minutes",
    "ice_cropped_elapsed_minutes",
    "ev_cropped_haversine_distance_km",
    "ice_cropped_haversine_distance_km",
    "cropped_distance_difference_km",
    "cropped_distance_difference_percent_of_shorter_trip",
    "allowed_cropped_distance_difference_km",
    "cropped_distance_comparability_status",
    "sampling_quality_warning",
    "ev_input_date_folder",
    "ice_input_date_folder",
    "pair_creation_status",
    "selection_reason",
    "final_pair_status",
    "pair_id",
    "output_date_folder",
    "pair_workbook_name",
    "pair_workbook_path",
    "common_route_crop_status",
    "feature_engineering_status",
]


def _clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _normalise_key(value) -> str:
    text = unicodedata.normalize("NFKD", _clean_text(value))
    text = "".join(
        character for character in text if not unicodedata.combining(character)
    )
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def _normalise_pair_class(value) -> str:
    text = re.sub(r"[^A-Z0-9]+", "_", _clean_text(value).upper()).strip("_")
    if text in STRICT_CLASSES:
        return "STRICT"
    if text in NON_STRICT_CLASSES or "NON_STRICT" in text or "NONSTRICT" in text:
        return "NONSTRICT"
    raise ValueError(f"Unrecognised candidate class: {value!r}")


def _safe_bool(value, default=False) -> bool:
    if value is None or pd.isna(value):
        return bool(default)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return _clean_text(value).casefold() in {"true", "yes", "y", "1", "same"}


def _safe_float(value, default=np.nan) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError):
        return float(default)
    return result if math.isfinite(result) else float(default)


def _parse_timestamp(value):
    if value is None or pd.isna(value) or _clean_text(value) == "":
        return pd.NaT
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        parsed = pd.to_datetime(value, errors="coerce")
    return parsed


def _parse_datetime_series(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    unresolved = parsed.isna() & series.notna()
    if bool(unresolved.any()):
        parsed.loc[unresolved] = pd.to_datetime(series.loc[unresolved], errors="coerce")
    return parsed


def _read_csv_with_fallback(path: Path) -> pd.DataFrame:
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(path, encoding=encoding)
        except UnicodeDecodeError as error:
            last_error = error
    raise ValueError(f"Could not decode {path}: {last_error}")


def _column_lookup(dataframe: pd.DataFrame) -> dict[str, str]:
    return {_normalise_key(column): column for column in dataframe.columns}


def _series_from_aliases(
    dataframe: pd.DataFrame,
    aliases: list[str],
    default=pd.NA,
) -> pd.Series:
    lookup = _column_lookup(dataframe)
    for alias in aliases:
        actual = lookup.get(_normalise_key(alias))
        if actual is not None:
            return dataframe[actual]
    return pd.Series([default] * len(dataframe), index=dataframe.index)


def _first_present(row: pd.Series, aliases: list[str], default=pd.NA):
    lookup = {_normalise_key(column): column for column in row.index}
    for alias in aliases:
        actual = lookup.get(_normalise_key(alias))
        if actual is not None:
            value = row.get(actual, default)
            if value is not None and not pd.isna(value) and _clean_text(value) != "":
                return value
    return default


def _cropping_rules(config) -> dict:
    cropping = config.get("cropping", {})
    pairing = config.get("pairing", {})
    configured_tolerance = float(
        cropping.get(
            "common_point_tolerance_m",
            pairing.get("route_point_tolerance_m", 100.0),
        )
    )
    # The old configuration used 500 m.  That is too wide for a claim that the
    # vehicles followed the same route, particularly where parallel roads are
    # close together.  A smaller configured value is respected; a larger value
    # is capped at 100 m.
    tolerance = max(
        1.0,
        min(
            configured_tolerance,
            float(cropping.get("maximum_route_point_tolerance_m", 100.0)),
        ),
    )
    absolute_distance_tolerance_km = min(
        float(cropping.get("post_crop_distance_absolute_tolerance_km", 0.5)),
        0.5,
    )
    relative_distance_tolerance = min(
        float(cropping.get("post_crop_distance_relative_tolerance_fraction", 0.02)),
        0.02,
    )
    return {
        "route_point_tolerance_m": tolerance,
        "route_resample_spacing_m": float(
            cropping.get("route_resample_spacing_m", 50.0)
        ),
        "maximum_resampled_points": int(
            cropping.get("maximum_resampled_points", 1500)
        ),
        "maximum_heading_difference_degrees": max(
            1.0,
            min(
                float(cropping.get("maximum_heading_difference_degrees", 45.0)),
                45.0,
            ),
        ),
        "minimum_sustained_match_points": max(
            int(cropping.get("minimum_sustained_match_points", 3)),
            3,
        ),
        "maximum_internal_unmatched_distance_m": min(
            float(cropping.get("maximum_internal_unmatched_distance_m", 500.0)),
            500.0,
        ),
        "minimum_ordered_overlap_fraction": max(
            float(cropping.get("minimum_ordered_overlap_fraction", 0.80)),
            0.80,
        ),
        "minimum_common_route_distance_km": max(
            float(cropping.get("minimum_common_route_distance_km", 5.0)),
            5.0,
        ),
        "minimum_shorter_trip_coverage_fraction": max(
            float(cropping.get("minimum_shorter_trip_coverage_fraction", 0.70)),
            0.70,
        ),
        "full_route_coverage_threshold": float(
            cropping.get("full_route_coverage_threshold", 0.90)
        ),
        "post_crop_distance_absolute_tolerance_km": absolute_distance_tolerance_km,
        "post_crop_distance_relative_tolerance_fraction": relative_distance_tolerance,
        "maximum_average_sampling_interval_seconds": float(
            pairing.get("maximum_average_sampling_interval_seconds", 30.0)
        ),
        "minimum_verified_datapoints": int(
            cropping.get("minimum_verified_datapoints", 2)
        ),
        "backup_existing_pair_folder": bool(
            cropping.get("backup_existing_pair_folder", True)
        ),
    }


def _candidate_input_path(config) -> Path:
    return (
        Path(config["project"]["results_folder"])
        / "Pairing"
        / CANDIDATE_WORKBOOK_NAME
    )


def _verified_index_path(config) -> Path:
    return (
        Path(config["project"]["results_folder"])
        / "System"
        / VERIFIED_INDEX_NAME
    )


def _canonical_candidate_table(raw: pd.DataFrame) -> pd.DataFrame:
    """Accept both the earlier and revised candidate-workbook schemas."""

    aliases = {
        "candidate_id": ["candidate_id"],
        "candidate_class": ["candidate_class", "pair_class"],
        "route_direction": ["route_direction", "route"],
        "planned_date_folder": [
            "planned_output_date_folder_after_acceptance",
            "suggested_date_folder",
            "output_date_folder",
        ],
        "route_file_label": ["route_file_label"],
        "ev_verified_trip_id": ["ev_verified_trip_id", "ev_trip_id"],
        "ev_file_name": ["ev_file_name"],
        "ev_relative_path": ["ev_relative_path"],
        "ev_cleaned_file_path": [
            "ev_cleaned_file_path",
            "ev_source_cleaned_file_path",
        ],
        "ice_verified_trip_id": ["ice_verified_trip_id", "ice_trip_id"],
        "ice_file_name": ["ice_file_name"],
        "ice_relative_path": ["ice_relative_path"],
        "ice_cleaned_file_path": [
            "ice_cleaned_file_path",
            "ice_source_cleaned_file_path",
        ],
        "ev_verified_start": ["ev_verified_start", "ev_verified_start_timestamp"],
        "ev_verified_end": ["ev_verified_end", "ev_verified_end_timestamp"],
        "ice_verified_start": [
            "ice_verified_start",
            "ice_verified_start_timestamp",
        ],
        "ice_verified_end": ["ice_verified_end", "ice_verified_end_timestamp"],
        "same_calendar_day": ["same_calendar_day"],
        "start_clock_difference_minutes": ["start_clock_difference_minutes"],
        "same_day_type": ["same_day_type"],
        "ev_sampling_quality": ["ev_sampling_quality"],
        "ice_sampling_quality": ["ice_sampling_quality"],
        "pair_sampling_quality": ["pair_sampling_quality"],
        "candidate_reason": ["candidate_reason"],
        "non_strict_category": ["non_strict_category"],
    }

    canonical = pd.DataFrame(index=raw.index)
    for target, possible_names in aliases.items():
        canonical[target] = _series_from_aliases(raw, possible_names)

    if canonical["candidate_id"].isna().any() or canonical["candidate_id"].astype(str).str.strip().eq("").any():
        canonical["candidate_id"] = [
            f"CAND_{number:03d}" for number in range(1, len(canonical) + 1)
        ]

    canonical["candidate_id"] = canonical["candidate_id"].astype(str).str.strip()
    if canonical["candidate_id"].duplicated().any():
        duplicated = canonical.loc[
            canonical["candidate_id"].duplicated(False), "candidate_id"
        ].drop_duplicates()
        raise ValueError(
            "Candidate workbook contains duplicate candidate_id values: "
            + ", ".join(duplicated.head(10))
        )

    canonical["candidate_class"] = canonical["candidate_class"].map(
        _normalise_pair_class
    )
    return canonical.reset_index(drop=True)


def _verified_lookup(index_path: Path) -> dict[str, dict]:
    if not index_path.exists():
        return {}
    index = _read_csv_with_fallback(index_path)
    if "verified_trip_id" not in index.columns:
        return {}
    return {
        _clean_text(row["verified_trip_id"]): row.to_dict()
        for _, row in index.iterrows()
    }


def _fill_candidate_from_verified_index(
    candidates: pd.DataFrame,
    verified_lookup: dict[str, dict],
) -> pd.DataFrame:
    candidates = candidates.copy()
    mapping = {
        "file_name": "source_file_name",
        "relative_path": "source_relative_path",
        "cleaned_file_path": "source_cleaned_file_path",
        "verified_start": "verified_start_timestamp",
        "verified_end": "verified_end_timestamp",
    }
    for vehicle in ("ev", "ice"):
        trip_id_column = f"{vehicle}_verified_trip_id"
        for candidate_suffix, index_column in mapping.items():
            candidate_column = f"{vehicle}_{candidate_suffix}"
            values = []
            for _, row in candidates.iterrows():
                current = row.get(candidate_column, pd.NA)
                if current is not None and not pd.isna(current) and _clean_text(current):
                    values.append(current)
                    continue
                lookup_row = verified_lookup.get(_clean_text(row.get(trip_id_column)))
                values.append(
                    lookup_row.get(index_column, pd.NA) if lookup_row else pd.NA
                )
            candidates[candidate_column] = values
    return candidates


def load_candidates(candidate_path: Path, verified_index_path: Path) -> pd.DataFrame:
    if not candidate_path.exists():
        raise FileNotFoundError(
            f"Candidate workbook not found: {candidate_path}. "
            "Run candidate_pairing.py first."
        )
    try:
        raw = pd.read_excel(
            candidate_path,
            sheet_name=CANDIDATE_SHEET_NAME,
            engine="openpyxl",
        )
    except ValueError as error:
        raise ValueError(
            f"Sheet '{CANDIDATE_SHEET_NAME}' was not found in {candidate_path}."
        ) from error
    if raw.empty:
        raise ValueError("The Candidate Pairs sheet is empty.")

    candidates = _canonical_candidate_table(raw)
    candidates = _fill_candidate_from_verified_index(
        candidates, _verified_lookup(verified_index_path)
    )

    required = [
        "candidate_id",
        "candidate_class",
        "route_direction",
        "ev_verified_trip_id",
        "ice_verified_trip_id",
        "ev_cleaned_file_path",
        "ice_cleaned_file_path",
        "ev_verified_start",
        "ev_verified_end",
        "ice_verified_start",
        "ice_verified_end",
    ]
    missing = []
    for column in required:
        bad = candidates[column].isna() | candidates[column].astype(str).str.strip().eq("")
        if bool(bad.any()):
            missing.append(f"{column} ({int(bad.sum())} row(s))")
    if missing:
        raise ValueError(
            "Candidate data could not be completed from verified_trip_index.csv: "
            + ", ".join(missing)
        )

    for column in [
        "ev_verified_start",
        "ev_verified_end",
        "ice_verified_start",
        "ice_verified_end",
    ]:
        candidates[column] = _parse_datetime_series(candidates[column])
    invalid_time = candidates[
        [
            "ev_verified_start",
            "ev_verified_end",
            "ice_verified_start",
            "ice_verified_end",
        ]
    ].isna().any(axis=1)
    if bool(invalid_time.any()):
        ids = ", ".join(candidates.loc[invalid_time, "candidate_id"].head(10))
        raise ValueError(f"Invalid verified timestamps for candidate(s): {ids}")

    candidates["same_calendar_day"] = candidates.apply(
        lambda row: (
            _safe_bool(row["same_calendar_day"])
            if not pd.isna(row["same_calendar_day"])
            else row["ev_verified_start"].date()
            == row["ice_verified_start"].date()
        ),
        axis=1,
    )
    candidates["same_day_type"] = candidates.apply(
        lambda row: (
            _safe_bool(row["same_day_type"])
            if not pd.isna(row["same_day_type"])
            else (row["ev_verified_start"].weekday() >= 5)
            == (row["ice_verified_start"].weekday() >= 5)
        ),
        axis=1,
    )
    candidates["start_clock_difference_minutes"] = pd.to_numeric(
        candidates["start_clock_difference_minutes"], errors="coerce"
    )
    missing_clock = candidates["start_clock_difference_minutes"].isna()
    if bool(missing_clock.any()):
        differences = []
        for _, row in candidates.loc[missing_clock].iterrows():
            first = row["ev_verified_start"]
            second = row["ice_verified_start"]
            first_minutes = first.hour * 60 + first.minute + first.second / 60
            second_minutes = second.hour * 60 + second.minute + second.second / 60
            direct = abs(first_minutes - second_minutes)
            differences.append(min(direct, 1440 - direct))
        candidates.loc[missing_clock, "start_clock_difference_minutes"] = differences

    return candidates


def _logical_cleaned_file_key(value) -> tuple[str, str]:
    """Match a cleaned file before/after the CITY/INTERCITY rename.

    Examples treated as the same logical file:
    ``trip_ev.xlsx``, ``trip_ev_CITY.xlsx`` and
    ``trip_ev_INTERCITY.xlsx``.  The date folder is still checked first, so
    files from different trip dates are not silently mixed.
    """
    path = Path(_clean_text(value))
    stem = re.sub(
        r"_(?:CITY|INTERCITY)$",
        "",
        path.stem,
        flags=re.IGNORECASE,
    )
    return _normalise_key(stem), path.suffix.casefold()


def _logical_name_matches(
    directory: Path,
    requested_name: str,
    recursive: bool = False,
) -> list[Path]:
    if not directory.exists() or not directory.is_dir():
        return []

    requested_key = _logical_cleaned_file_key(requested_name)
    iterator = directory.rglob("*") if recursive else directory.iterdir()
    return sorted(
        (
            path
            for path in iterator
            if path.is_file()
            and _logical_cleaned_file_key(path.name) == requested_key
        ),
        key=lambda path: str(path).casefold(),
    )


def _resolve_cleaned_path(
    configured_path,
    relative_path,
    file_name,
    results_root: Path,
) -> Path:
    candidates = []
    if _clean_text(configured_path):
        candidates.append(Path(_clean_text(configured_path)))
    if _clean_text(relative_path):
        candidates.append(
            results_root / "Cleaned" / Path(_clean_text(relative_path))
        )
    if _clean_text(file_name):
        candidates.append(results_root / "Cleaned" / _clean_text(file_name))

    # First preserve the original exact-path behaviour.
    for path in candidates:
        if path.exists() and path.is_file():
            return path

    # The files may have been renamed after candidate creation by adding
    # _CITY or _INTERCITY.  Search the candidate's own date folder first.
    for candidate in candidates:
        matches = _logical_name_matches(
            candidate.parent,
            candidate.name,
            recursive=False,
        )
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise ValueError(
                "More than one CITY/INTERCITY cleaned file matches "
                f"{candidate.name} in {candidate.parent}: "
                + "; ".join(str(path) for path in matches)
            )

    # Last fallback: search all cleaned date folders.  Accept only a unique
    # result so a trip from the wrong date can never be chosen automatically.
    name = _clean_text(file_name)
    cleaned_root = results_root / "Cleaned"
    if name:
        matches = _logical_name_matches(cleaned_root, name, recursive=True)
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise ValueError(
                "More than one cleaned file matches the logical name "
                f"{name}. Use the full date-specific path: "
                + "; ".join(str(path) for path in matches)
            )

    attempted = "; ".join(str(path) for path in candidates) or "no usable path"
    raise FileNotFoundError(
        "Cleaned telemetry file not found, including CITY/INTERCITY variants. "
        f"Tried: {attempted}"
    )


def _read_cleaned_workbook(path: Path) -> pd.DataFrame:
    try:
        dataframe = pd.read_excel(
            path, sheet_name=CLEANED_SHEET_NAME, engine="openpyxl"
        )
    except ValueError:
        excel = pd.ExcelFile(path, engine="openpyxl")
        if not excel.sheet_names:
            raise ValueError(f"No worksheets found in {path}")
        dataframe = pd.read_excel(path, sheet_name=excel.sheet_names[0], engine="openpyxl")

    lookup = _column_lookup(dataframe)
    rename = {}
    aliases = {
        TIMESTAMP_COLUMN: ["Timestamp", "Date Time", "Datetime", "Packet Time"],
        LATITUDE_COLUMN: ["Latitude", "Lat"],
        LONGITUDE_COLUMN: ["Longitude", "Lon", "Lng", "Long"],
        SPEED_KMH_COLUMN: ["Speed(km/h)", "Speed km/h", "Speed", "Speed_kmh"],
        SPEED_MPS_COLUMN: ["Speed(m/s)", "Speed m/s", "Speed_mps"],
        ADDRESS_COLUMN: ["Address"],
    }
    for target, possible in aliases.items():
        if target in dataframe.columns:
            continue
        for alias in possible:
            actual = lookup.get(_normalise_key(alias))
            if actual is not None:
                rename[actual] = target
                break
    dataframe = dataframe.rename(columns=rename).copy()

    required = [TIMESTAMP_COLUMN, LATITUDE_COLUMN, LONGITUDE_COLUMN]
    missing = [column for column in required if column not in dataframe.columns]
    if missing:
        raise ValueError(
            f"Cleaned file {path.name} is missing: {', '.join(missing)}"
        )

    dataframe["Original Cleaned Workbook Row"] = np.arange(2, len(dataframe) + 2)
    dataframe[TIMESTAMP_COLUMN] = _parse_datetime_series(dataframe[TIMESTAMP_COLUMN])
    for column in [LATITUDE_COLUMN, LONGITUDE_COLUMN, SPEED_MPS_COLUMN, SPEED_KMH_COLUMN]:
        if column in dataframe.columns:
            dataframe[column] = pd.to_numeric(dataframe[column], errors="coerce")
    if SPEED_KMH_COLUMN not in dataframe.columns and SPEED_MPS_COLUMN in dataframe.columns:
        dataframe[SPEED_KMH_COLUMN] = dataframe[SPEED_MPS_COLUMN] * 3.6
    if SPEED_MPS_COLUMN not in dataframe.columns and SPEED_KMH_COLUMN in dataframe.columns:
        dataframe[SPEED_MPS_COLUMN] = dataframe[SPEED_KMH_COLUMN] / 3.6

    if ADDRESS_COLUMN not in dataframe.columns:
        dataframe[ADDRESS_COLUMN] = ""
    else:
        dataframe[ADDRESS_COLUMN] = (
            dataframe[ADDRESS_COLUMN]
            .astype("string")
            .fillna("")
            .str.strip()
        )

    invalid = dataframe[
        [TIMESTAMP_COLUMN, LATITUDE_COLUMN, LONGITUDE_COLUMN]
    ].isna().any(axis=1)
    if bool(invalid.any()):
        raise ValueError(
            f"Cleaned file {path.name} contains {int(invalid.sum())} invalid "
            "timestamp/coordinate row(s). Rerun cleaning."
        )
    dataframe = dataframe.sort_values(TIMESTAMP_COLUMN, kind="stable").reset_index(drop=True)
    return dataframe


def _verified_segment(
    full_data: pd.DataFrame,
    verified_start,
    verified_end,
    minimum_points: int,
) -> pd.DataFrame:
    start = pd.Timestamp(verified_start)
    end = pd.Timestamp(verified_end)
    if end < start:
        raise ValueError("Verified end timestamp is earlier than verified start timestamp")
    segment = full_data.loc[
        full_data[TIMESTAMP_COLUMN].between(start, end, inclusive="both")
    ].copy()
    if len(segment) < minimum_points:
        raise ValueError(
            f"Only {len(segment)} cleaned datapoint(s) fall inside the verified boundary"
        )
    return segment.reset_index(drop=True)


def _haversine_m(lat1, lon1, lat2, lon2) -> float:
    latitude_1 = math.radians(float(lat1))
    latitude_2 = math.radians(float(lat2))
    delta_latitude = math.radians(float(lat2) - float(lat1))
    delta_longitude = math.radians(float(lon2) - float(lon1))
    value = (
        math.sin(delta_latitude / 2) ** 2
        + math.cos(latitude_1)
        * math.cos(latitude_2)
        * math.sin(delta_longitude / 2) ** 2
    )
    value = min(1.0, max(0.0, value))
    return EARTH_RADIUS_M * 2 * math.atan2(math.sqrt(value), math.sqrt(1 - value))


def _segment_distances_m(dataframe: pd.DataFrame) -> np.ndarray:
    latitude = np.radians(dataframe[LATITUDE_COLUMN].to_numpy(dtype=float))
    longitude = np.radians(dataframe[LONGITUDE_COLUMN].to_numpy(dtype=float))
    if len(latitude) < 2:
        return np.array([], dtype=float)
    delta_latitude = np.diff(latitude)
    delta_longitude = np.diff(longitude)
    value = (
        np.sin(delta_latitude / 2) ** 2
        + np.cos(latitude[:-1])
        * np.cos(latitude[1:])
        * np.sin(delta_longitude / 2) ** 2
    )
    value = np.clip(value, 0.0, 1.0)
    return EARTH_RADIUS_M * 2 * np.arctan2(np.sqrt(value), np.sqrt(1 - value))


def _cumulative_distance_m(dataframe: pd.DataFrame) -> np.ndarray:
    return np.concatenate(([0.0], np.cumsum(_segment_distances_m(dataframe))))


def _unique_progress(dataframe: pd.DataFrame, cumulative: np.ndarray):
    keep = np.concatenate(([True], np.diff(cumulative) > 1e-6))
    unique_data = dataframe.loc[keep].reset_index(drop=True)
    unique_progress = cumulative[keep]
    if len(unique_progress) < 2 or unique_progress[-1] <= 0:
        raise ValueError("Verified trip has no usable movement distance")
    return unique_data, unique_progress


def _resample_route(
    dataframe: pd.DataFrame,
    configured_spacing_m: float,
    maximum_points: int,
) -> dict:
    cumulative = _cumulative_distance_m(dataframe)
    unique_data, unique_progress = _unique_progress(dataframe, cumulative)
    total = float(unique_progress[-1])
    spacing = max(float(configured_spacing_m), total / max(maximum_points - 1, 1))
    targets = np.arange(0.0, total, spacing)
    if len(targets) == 0 or not math.isclose(targets[-1], total, abs_tol=1e-6):
        targets = np.append(targets, total)
    latitude = np.interp(
        targets,
        unique_progress,
        unique_data[LATITUDE_COLUMN].to_numpy(dtype=float),
    )
    longitude = np.interp(
        targets,
        unique_progress,
        unique_data[LONGITUDE_COLUMN].to_numpy(dtype=float),
    )
    return {
        "progress_m": targets,
        "latitude": latitude,
        "longitude": longitude,
        "total_distance_m": total,
        "actual_spacing_m": spacing,
        "original_cumulative_m": cumulative,
    }


def _project_routes_xy(ev_route: dict, ice_route: dict):
    all_latitude = np.concatenate([ev_route["latitude"], ice_route["latitude"]])
    all_longitude = np.concatenate([ev_route["longitude"], ice_route["longitude"]])
    reference_latitude = math.radians(float(np.mean(all_latitude)))
    reference_longitude = math.radians(float(np.mean(all_longitude)))

    def project(route):
        latitude_radians = np.radians(route["latitude"])
        longitude_radians = np.radians(route["longitude"])
        x = (
            (longitude_radians - reference_longitude)
            * math.cos(reference_latitude)
            * EARTH_RADIUS_M
        )
        y = (latitude_radians - reference_latitude) * EARTH_RADIUS_M
        return np.column_stack([x, y])

    return project(ev_route), project(ice_route)


def _nearest_points(first_xy: np.ndarray, second_xy: np.ndarray):
    nearest_index = np.empty(len(first_xy), dtype=int)
    nearest_distance = np.empty(len(first_xy), dtype=float)
    chunk_size = 256
    for start in range(0, len(first_xy), chunk_size):
        end = min(start + chunk_size, len(first_xy))
        difference = first_xy[start:end, None, :] - second_xy[None, :, :]
        squared = np.einsum("ijk,ijk->ij", difference, difference)
        indices = np.argmin(squared, axis=1)
        nearest_index[start:end] = indices
        nearest_distance[start:end] = np.sqrt(
            squared[np.arange(end - start), indices]
        )
    return nearest_index, nearest_distance


def _route_headings_degrees(route_xy: np.ndarray) -> np.ndarray:
    """Estimate local travel direction at each resampled route point."""
    if len(route_xy) < 2:
        raise ValueError("At least two route points are required for heading checks")
    vectors = np.empty_like(route_xy, dtype=float)
    vectors[0] = route_xy[1] - route_xy[0]
    vectors[-1] = route_xy[-1] - route_xy[-2]
    if len(route_xy) > 2:
        vectors[1:-1] = route_xy[2:] - route_xy[:-2]
    # Zero degrees is north.  Only angular differences are used later.
    return (np.degrees(np.arctan2(vectors[:, 0], vectors[:, 1])) + 360.0) % 360.0


def _angular_difference_degrees(first, second) -> np.ndarray:
    first = np.asarray(first, dtype=float)
    second = np.asarray(second, dtype=float)
    return np.abs((first - second + 180.0) % 360.0 - 180.0)


def _continuous_ordered_match_runs(
    valid_ev_indices: np.ndarray,
    nearest_ice_indices: np.ndarray,
    ev_progress_m: np.ndarray,
    ice_progress_m: np.ndarray,
    maximum_internal_gap_m: float,
    minimum_points: int,
) -> list[list[tuple[int, int]]]:
    """Split matched points whenever order or spatial continuity is broken."""
    runs: list[list[tuple[int, int]]] = []
    current: list[tuple[int, int]] = []
    for ev_index in valid_ev_indices.tolist():
        ice_index = int(nearest_ice_indices[ev_index])
        if current:
            previous_ev, previous_ice = current[-1]
            ev_gap = ev_progress_m[ev_index] - ev_progress_m[previous_ev]
            ice_gap = ice_progress_m[ice_index] - ice_progress_m[previous_ice]
            continuous = (
                ev_index > previous_ev
                and ice_index >= previous_ice
                and ev_gap <= maximum_internal_gap_m
                and 0.0 <= ice_gap <= maximum_internal_gap_m
            )
            if not continuous:
                if len(current) >= minimum_points:
                    runs.append(current)
                current = []
        current.append((ev_index, ice_index))
    if len(current) >= minimum_points:
        runs.append(current)
    return runs


def _best_distance_balanced_span(
    run: list[tuple[int, int]],
    ev_progress_m: np.ndarray,
    ice_progress_m: np.ndarray,
    rules: dict,
):
    """Return the longest ordered span satisfying the hard distance rule."""
    minimum_length_m = rules["minimum_common_route_distance_km"] * 1000.0
    absolute_tolerance_m = (
        rules["post_crop_distance_absolute_tolerance_km"] * 1000.0
    )
    relative_tolerance = rules[
        "post_crop_distance_relative_tolerance_fraction"
    ]
    best = None
    best_length = -1.0

    for start_position in range(len(run) - 1):
        ev_start, ice_start = run[start_position]
        maximum_ev = ev_progress_m[run[-1][0]] - ev_progress_m[ev_start]
        maximum_ice = ice_progress_m[run[-1][1]] - ice_progress_m[ice_start]
        if min(maximum_ev, maximum_ice) <= best_length:
            continue
        for end_position in range(len(run) - 1, start_position, -1):
            ev_end, ice_end = run[end_position]
            ev_span = float(ev_progress_m[ev_end] - ev_progress_m[ev_start])
            ice_span = float(ice_progress_m[ice_end] - ice_progress_m[ice_start])
            shorter_span = min(ev_span, ice_span)
            if shorter_span < minimum_length_m:
                break
            allowed = max(
                absolute_tolerance_m,
                relative_tolerance * shorter_span,
            )
            if abs(ev_span - ice_span) <= allowed:
                if shorter_span > best_length:
                    best = {
                        "ev_start_index": ev_start,
                        "ev_end_index": ev_end,
                        "ice_start_index": ice_start,
                        "ice_end_index": ice_end,
                        "run": run[start_position : end_position + 1],
                        "ev_span_m": ev_span,
                        "ice_span_m": ice_span,
                        "allowed_span_difference_m": allowed,
                    }
                    best_length = shorter_span
                # Endpoints are being checked from farthest to nearest, so this
                # is the longest accepted endpoint for the present start.
                break
    return best


def _maximum_false_run_distance_m(mask: np.ndarray, spacing_m: float) -> float:
    maximum = 0
    current = 0
    for matched in np.asarray(mask, dtype=bool):
        if matched:
            maximum = max(maximum, current)
            current = 0
        else:
            current += 1
    maximum = max(maximum, current)
    return float(maximum * spacing_m)


def _longest_nondecreasing_subsequence(values: list[int]) -> list[int]:
    if not values:
        return []
    tails = []
    tail_positions = []
    previous = [-1] * len(values)
    for position, value in enumerate(values):
        insertion = bisect.bisect_right(tails, value)
        if insertion == len(tails):
            tails.append(value)
            tail_positions.append(position)
        else:
            tails[insertion] = value
            tail_positions[insertion] = position
        if insertion > 0:
            previous[position] = tail_positions[insertion - 1]
    result = []
    cursor = tail_positions[-1]
    while cursor >= 0:
        result.append(cursor)
        cursor = previous[cursor]
    return list(reversed(result))


def _anchor_midpoint(lat1, lon1, lat2, lon2):
    return (float(lat1) + float(lat2)) / 2.0, (float(lon1) + float(lon2)) / 2.0


def _sampling_warning(dataframe: pd.DataFrame, maximum_average_interval_s: float) -> str:
    if len(dataframe) < 2:
        return "INSUFFICIENT_DATAPOINTS_WARNING"
    elapsed = (
        dataframe[TIMESTAMP_COLUMN].iloc[-1] - dataframe[TIMESTAMP_COLUMN].iloc[0]
    ).total_seconds()
    average = elapsed / (len(dataframe) - 1)
    if average > maximum_average_interval_s:
        return "SPARSE_TELEMETRY_WARNING_DOES_NOT_BLOCK_CROPPING"
    return "NO_SAMPLING_WARNING"


def _route_validation(
    ev_data: pd.DataFrame,
    ice_data: pd.DataFrame,
    rules: dict,
) -> dict:
    """Find one continuous, direction-consistent, equal-distance common crop.

    A rejected candidate is retained in the audit workbook, but it is not
    written to ``Results/Paired`` and is not added to ``pair_index.csv``.
    """
    ev_route = _resample_route(
        ev_data,
        rules["route_resample_spacing_m"],
        rules["maximum_resampled_points"],
    )
    ice_route = _resample_route(
        ice_data,
        rules["route_resample_spacing_m"],
        rules["maximum_resampled_points"],
    )
    ev_xy, ice_xy = _project_routes_xy(ev_route, ice_route)
    ev_to_ice_index, ev_to_ice_distance = _nearest_points(ev_xy, ice_xy)
    ice_to_ev_index, ice_to_ev_distance = _nearest_points(ice_xy, ev_xy)

    tolerance = rules["route_point_tolerance_m"]
    heading_tolerance = rules["maximum_heading_difference_degrees"]
    ev_headings = _route_headings_degrees(ev_xy)
    ice_headings = _route_headings_degrees(ice_xy)
    ev_heading_difference = _angular_difference_degrees(
        ev_headings,
        ice_headings[ev_to_ice_index],
    )
    ice_heading_difference = _angular_difference_degrees(
        ice_headings,
        ev_headings[ice_to_ev_index],
    )
    valid_ev_mask = (
        (ev_to_ice_distance <= tolerance)
        & (ev_heading_difference <= heading_tolerance)
    )
    valid_ice_mask = (
        (ice_to_ev_distance <= tolerance)
        & (ice_heading_difference <= heading_tolerance)
    )
    # Requiring support in the reverse nearest-neighbour direction prevents a
    # dense route from matching a single isolated point on the other route.
    valid_ev_mask &= valid_ice_mask[ev_to_ice_index]
    valid_ice_mask &= valid_ev_mask[ice_to_ev_index]
    valid_ev_mask &= valid_ice_mask[ev_to_ice_index]
    valid_ev_indices = np.flatnonzero(valid_ev_mask)

    base = {
        "gps_route_validation_status": "REJECTED_ROUTE_VALIDATION",
        "route_validation_quality": "REJECTED_NO_COMPARABLE_COMMON_SECTION",
        "route_validation_reason": "",
        "crop_selection_method": "CONTINUOUS_ORDERED_DIRECTION_AWARE_MATCHES",
        "route_point_tolerance_m": tolerance,
        "route_resample_spacing_m": max(
            ev_route["actual_spacing_m"], ice_route["actual_spacing_m"]
        ),
        "ev_verified_datapoints_loaded": len(ev_data),
        "ice_verified_datapoints_loaded": len(ice_data),
        "ev_verified_full_haversine_distance_km": ev_route["total_distance_m"] / 1000,
        "ice_verified_full_haversine_distance_km": ice_route["total_distance_m"] / 1000,
        "ordered_match_count": 0,
    }
    runs = _continuous_ordered_match_runs(
        valid_ev_indices,
        ev_to_ice_index,
        ev_route["progress_m"],
        ice_route["progress_m"],
        rules["maximum_internal_unmatched_distance_m"],
        rules["minimum_sustained_match_points"],
    )
    if not runs:
        base["route_validation_reason"] = (
            "No sustained forward-ordered EV/ICE match was found within "
            f"{tolerance:.0f} m and {heading_tolerance:.0f} degrees."
        )
        return base

    spans = []
    for run in runs:
        span = _best_distance_balanced_span(
            run,
            ev_route["progress_m"],
            ice_route["progress_m"],
            rules,
        )
        if span is not None:
            spans.append(span)
    if not spans:
        base["route_validation_reason"] = (
            "Ordered route matches were found, but no continuous section was "
            f"at least {rules['minimum_common_route_distance_km']:.3f} km and "
            "also satisfied the EV/ICE distance-difference tolerance."
        )
        return base

    accepted_spans = []
    rejection_details = []
    for span in spans:
        ev_start_index = span["ev_start_index"]
        ev_end_index = span["ev_end_index"]
        ice_start_index = span["ice_start_index"]
        ice_end_index = span["ice_end_index"]

        ev_span_indices = np.arange(ev_start_index, ev_end_index + 1)
        ice_span_indices = np.arange(ice_start_index, ice_end_index + 1)
        ev_inside = (
            valid_ev_mask[ev_span_indices]
            & (ev_to_ice_index[ev_span_indices] >= ice_start_index)
            & (ev_to_ice_index[ev_span_indices] <= ice_end_index)
        )
        ice_inside = (
            valid_ice_mask[ice_span_indices]
            & (ice_to_ev_index[ice_span_indices] >= ev_start_index)
            & (ice_to_ev_index[ice_span_indices] <= ev_end_index)
        )
        ev_overlap = float(ev_inside.mean()) if len(ev_inside) else 0.0
        ice_overlap = float(ice_inside.mean()) if len(ice_inside) else 0.0
        minimum_overlap = min(ev_overlap, ice_overlap)
        maximum_unmatched_m = max(
            _maximum_false_run_distance_m(
                ev_inside, ev_route["actual_spacing_m"]
            ),
            _maximum_false_run_distance_m(
                ice_inside, ice_route["actual_spacing_m"]
            ),
        )
        common_length_m = min(span["ev_span_m"], span["ice_span_m"])
        shorter_trip_coverage = common_length_m / min(
            ev_route["total_distance_m"], ice_route["total_distance_m"]
        )

        failures = []
        if minimum_overlap < rules["minimum_ordered_overlap_fraction"]:
            failures.append(
                f"overlap {minimum_overlap:.1%} < "
                f"{rules['minimum_ordered_overlap_fraction']:.1%}"
            )
        if maximum_unmatched_m > rules["maximum_internal_unmatched_distance_m"]:
            failures.append(
                f"internal unmatched run {maximum_unmatched_m:.0f} m > "
                f"{rules['maximum_internal_unmatched_distance_m']:.0f} m"
            )
        if shorter_trip_coverage < rules["minimum_shorter_trip_coverage_fraction"]:
            failures.append(
                f"shorter-trip coverage {shorter_trip_coverage:.1%} < "
                f"{rules['minimum_shorter_trip_coverage_fraction']:.1%}"
            )

        start_latitude, start_longitude = _anchor_midpoint(
            ev_route["latitude"][ev_start_index],
            ev_route["longitude"][ev_start_index],
            ice_route["latitude"][ice_start_index],
            ice_route["longitude"][ice_start_index],
        )
        end_latitude, end_longitude = _anchor_midpoint(
            ev_route["latitude"][ev_end_index],
            ev_route["longitude"][ev_end_index],
            ice_route["latitude"][ice_end_index],
            ice_route["longitude"][ice_end_index],
        )
        ev_start_progress = float(ev_route["progress_m"][ev_start_index])
        ev_end_progress = float(ev_route["progress_m"][ev_end_index])
        ice_start_progress = float(ice_route["progress_m"][ice_start_index])
        ice_end_progress = float(ice_route["progress_m"][ice_end_index])

        # Validate the distance of the exact rows that would be written.  This
        # includes the very small midpoint adjustment at the two boundaries.
        ev_check_crop = _build_crop(
            ev_data,
            ev_start_progress,
            ev_end_progress,
            (start_latitude, start_longitude),
            (end_latitude, end_longitude),
            "VALIDATION_ONLY",
            "EV",
            "VALIDATION_ONLY",
        )
        ice_check_crop = _build_crop(
            ice_data,
            ice_start_progress,
            ice_end_progress,
            (start_latitude, start_longitude),
            (end_latitude, end_longitude),
            "VALIDATION_ONLY",
            "ICE",
            "VALIDATION_ONLY",
        )
        ev_check_distance = _crop_metrics(ev_check_crop)["distance_km"]
        ice_check_distance = _crop_metrics(ice_check_crop)["distance_km"]
        (
            exact_difference_km,
            exact_difference_percent,
            allowed_difference_km,
            exact_distance_status,
        ) = _distance_comparability(
            ev_check_distance,
            ice_check_distance,
            rules,
        )
        if exact_distance_status != "CROPPED_DISTANCES_WITHIN_TOLERANCE":
            failures.append(
                f"written-crop distance difference {exact_difference_km:.3f} km > "
                f"allowed {allowed_difference_km:.3f} km"
            )

        if failures:
            rejection_details.append("; ".join(failures))
            continue

        matched_pairs = span["run"]
        heading_differences = [
            ev_heading_difference[ev_index]
            for ev_index, _ in matched_pairs
        ]
        ev_coverage = span["ev_span_m"] / ev_route["total_distance_m"]
        ice_coverage = span["ice_span_m"] / ice_route["total_distance_m"]
        accepted_spans.append(
            {
                **span,
                "ev_overlap": ev_overlap,
                "ice_overlap": ice_overlap,
                "minimum_overlap": minimum_overlap,
                "maximum_unmatched_m": maximum_unmatched_m,
                "common_length_m": common_length_m,
                "shorter_trip_coverage": shorter_trip_coverage,
                "start_latitude": start_latitude,
                "start_longitude": start_longitude,
                "end_latitude": end_latitude,
                "end_longitude": end_longitude,
                "ev_start_progress": ev_start_progress,
                "ev_end_progress": ev_end_progress,
                "ice_start_progress": ice_start_progress,
                "ice_end_progress": ice_end_progress,
                "ev_coverage": ev_coverage,
                "ice_coverage": ice_coverage,
                "median_heading_difference": float(np.median(heading_differences)),
                "ev_check_distance": ev_check_distance,
                "ice_check_distance": ice_check_distance,
                "exact_difference_km": exact_difference_km,
                "exact_difference_percent": exact_difference_percent,
                "allowed_difference_km": allowed_difference_km,
            }
        )

    if not accepted_spans:
        detail = rejection_details[0] if rejection_details else "hard route checks failed"
        base["route_validation_reason"] = (
            "A possible common section was found but rejected: " + detail
        )
        return base

    best = max(
        accepted_spans,
        key=lambda item: (
            item["common_length_m"],
            item["minimum_overlap"],
            -item["exact_difference_km"],
        ),
    )
    comparison_scope = (
        "FULL_ROUTE_TAT_ELIGIBLE"
        if min(best["ev_coverage"], best["ice_coverage"])
        >= rules["full_route_coverage_threshold"]
        else "COMMON_ROUTE_SEGMENT_TAT"
    )
    start_separation = float(ev_to_ice_distance[best["ev_start_index"]])
    end_separation = float(ev_to_ice_distance[best["ev_end_index"]])
    sparse = (
        _sampling_warning(
            ev_data, rules["maximum_average_sampling_interval_seconds"]
        )
        != "NO_SAMPLING_WARNING"
        or _sampling_warning(
            ice_data, rules["maximum_average_sampling_interval_seconds"]
        )
        != "NO_SAMPLING_WARNING"
    )
    base.update(
        {
            "gps_route_validation_status": "PASSED_ORDERED_ROUTE_VALIDATION",
            "route_validation_quality": (
                "PASSED_WITH_SPARSE_TELEMETRY_WARNING" if sparse else "PASSED"
            ),
            "route_validation_reason": (
                "Continuous same-direction common route confirmed; shared start/end "
                "anchors and the hard cropped-distance tolerance are satisfied."
            ),
            "ordered_match_count": len(best["run"]),
            "median_heading_difference_degrees": best[
                "median_heading_difference"
            ],
            "ev_ordered_overlap_fraction_within_crop": best["ev_overlap"],
            "ice_ordered_overlap_fraction_within_crop": best["ice_overlap"],
            "minimum_ordered_overlap_fraction_within_crop": best[
                "minimum_overlap"
            ],
            "maximum_internal_unmatched_distance_m": best[
                "maximum_unmatched_m"
            ],
            "common_route_length_km": best["common_length_m"] / 1000.0,
            "shorter_trip_common_coverage_fraction": best[
                "shorter_trip_coverage"
            ],
            "ev_crop_coverage_fraction_of_verified_trip": best["ev_coverage"],
            "ice_crop_coverage_fraction_of_verified_trip": best["ice_coverage"],
            "comparison_scope": comparison_scope,
            "common_start_anchor_latitude": best["start_latitude"],
            "common_start_anchor_longitude": best["start_longitude"],
            "common_end_anchor_latitude": best["end_latitude"],
            "common_end_anchor_longitude": best["end_longitude"],
            "start_anchor_ev_ice_separation_m": start_separation,
            "end_anchor_ev_ice_separation_m": end_separation,
            "ev_crop_start_progress_m": best["ev_start_progress"],
            "ev_crop_end_progress_m": best["ev_end_progress"],
            "ice_crop_start_progress_m": best["ice_start_progress"],
            "ice_crop_end_progress_m": best["ice_end_progress"],
            "cropped_distance_difference_km": best["exact_difference_km"],
            "cropped_distance_difference_percent_of_shorter_trip": best[
                "exact_difference_percent"
            ],
            "allowed_cropped_distance_difference_km": best[
                "allowed_difference_km"
            ],
            "cropped_distance_comparability_status": (
                "CROPPED_DISTANCES_WITHIN_TOLERANCE"
            ),
        }
    )
    return base


def _interpolate_value(first, second, fraction):
    first_number = _safe_float(first)
    second_number = _safe_float(second)
    if math.isnan(first_number) and math.isnan(second_number):
        return pd.NA
    if math.isnan(first_number):
        return second_number
    if math.isnan(second_number):
        return first_number
    return first_number + fraction * (second_number - first_number)


def _interpolated_anchor_row(
    dataframe: pd.DataFrame,
    cumulative: np.ndarray,
    target_progress_m: float,
    anchor_latitude: float,
    anchor_longitude: float,
    anchor_label: str,
) -> dict:
    # ``side='left'`` chooses the first time a spatial progress value is
    # reached.  This retains any stationary time at a common start anchor but
    # stops at first arrival at the common end anchor, rather than silently
    # including a later dwell beyond the comparable route section.
    upper = int(np.searchsorted(cumulative, target_progress_m, side="left"))
    if upper <= 0:
        lower = 0
        upper = 1
        fraction = 0.0
    elif upper >= len(dataframe):
        lower = len(dataframe) - 2
        upper = len(dataframe) - 1
        fraction = 1.0
    else:
        lower = upper - 1
        denominator = cumulative[upper] - cumulative[lower]
        fraction = (
            1.0
            if denominator <= 1e-9
            else (target_progress_m - cumulative[lower]) / denominator
        )
    fraction = min(1.0, max(0.0, float(fraction)))
    first_time = pd.Timestamp(dataframe.iloc[lower][TIMESTAMP_COLUMN])
    second_time = pd.Timestamp(dataframe.iloc[upper][TIMESTAMP_COLUMN])
    timestamp = first_time + (second_time - first_time) * fraction

    row = {column: pd.NA for column in dataframe.columns}
    row[TIMESTAMP_COLUMN] = timestamp
    row[LATITUDE_COLUMN] = float(anchor_latitude)
    row[LONGITUDE_COLUMN] = float(anchor_longitude)
    for column in [SPEED_MPS_COLUMN, SPEED_KMH_COLUMN]:
        if column in dataframe.columns:
            row[column] = _interpolate_value(
                dataframe.iloc[lower][column], dataframe.iloc[upper][column], fraction
            )

    if ADDRESS_COLUMN in dataframe.columns:
        first_address = _clean_text(dataframe.iloc[lower][ADDRESS_COLUMN])
        second_address = _clean_text(dataframe.iloc[upper][ADDRESS_COLUMN])
        if first_address and second_address:
            row[ADDRESS_COLUMN] = first_address if fraction < 0.5 else second_address
        else:
            row[ADDRESS_COLUMN] = first_address or second_address

    row["Original Cleaned Workbook Row"] = pd.NA
    row["Crop Record Type"] = f"INTERPOLATED_COMMON_{anchor_label}_ANCHOR"
    row["Common Anchor ID"] = f"COMMON_{anchor_label}_ANCHOR"
    row["Is Interpolated Crop Anchor"] = True
    return row


def _build_crop(
    dataframe: pd.DataFrame,
    start_progress_m: float,
    end_progress_m: float,
    start_anchor: tuple[float, float],
    end_anchor: tuple[float, float],
    pair_id: str,
    vehicle_type: str,
    verified_trip_id: str,
) -> pd.DataFrame:
    cumulative = _cumulative_distance_m(dataframe)
    if end_progress_m <= start_progress_m:
        raise ValueError("Crop end progress must be greater than crop start progress")
    if end_progress_m > cumulative[-1] + 1e-6:
        raise ValueError("Crop end progress lies outside the verified trip")

    start_row = _interpolated_anchor_row(
        dataframe,
        cumulative,
        start_progress_m,
        start_anchor[0],
        start_anchor[1],
        "START",
    )
    end_row = _interpolated_anchor_row(
        dataframe,
        cumulative,
        end_progress_m,
        end_anchor[0],
        end_anchor[1],
        "END",
    )
    start_timestamp = pd.Timestamp(start_row[TIMESTAMP_COLUMN])
    end_timestamp = pd.Timestamp(end_row[TIMESTAMP_COLUMN])
    interior = dataframe.loc[
        (dataframe[TIMESTAMP_COLUMN] > start_timestamp)
        & (dataframe[TIMESTAMP_COLUMN] < end_timestamp)
    ].copy()
    interior["Crop Record Type"] = "ORIGINAL_CLEANED_DATAPOINT"
    interior["Common Anchor ID"] = ""
    interior["Is Interpolated Crop Anchor"] = False

    crop = pd.concat(
        [pd.DataFrame([start_row]), interior, pd.DataFrame([end_row])],
        ignore_index=True,
        sort=False,
    )
    crop = crop.sort_values(TIMESTAMP_COLUMN, kind="stable").reset_index(drop=True)
    crop["Pair ID"] = pair_id
    crop["Vehicle Type"] = vehicle_type
    crop["Verified Trip ID"] = verified_trip_id
    crop["Crop Datapoint Number"] = np.arange(1, len(crop) + 1)

    audit_columns = [
        "Pair ID",
        "Vehicle Type",
        "Verified Trip ID",
        "Crop Datapoint Number",
        "Original Cleaned Workbook Row",
        "Crop Record Type",
        "Common Anchor ID",
        "Is Interpolated Crop Anchor",
    ]
    original_columns = [
        column for column in dataframe.columns if column not in audit_columns
    ]
    return crop[original_columns + audit_columns]


def _crop_metrics(crop: pd.DataFrame) -> dict:
    distance_km = float(_segment_distances_m(crop).sum() / 1000)
    start = pd.Timestamp(crop[TIMESTAMP_COLUMN].iloc[0])
    end = pd.Timestamp(crop[TIMESTAMP_COLUMN].iloc[-1])
    return {
        "datapoints": len(crop),
        "start": start,
        "end": end,
        "elapsed_minutes": (end - start).total_seconds() / 60,
        "distance_km": distance_km,
    }


def _distance_comparability(ev_distance_km, ice_distance_km, rules):
    difference = abs(ev_distance_km - ice_distance_km)
    shorter = min(ev_distance_km, ice_distance_km)
    fraction = difference / shorter if shorter > 0 else float("inf")
    allowed = max(
        rules["post_crop_distance_absolute_tolerance_km"],
        rules["post_crop_distance_relative_tolerance_fraction"] * shorter,
    )
    if difference <= allowed:
        status = "CROPPED_DISTANCES_WITHIN_TOLERANCE"
    else:
        status = "REJECTED_CROPPED_DISTANCE_DIFFERENCE"
    return difference, fraction * 100.0, allowed, status


def _route_file_label(route_direction: str) -> str:
    text = unicodedata.normalize("NFKD", _clean_text(route_direction))
    text = "".join(
        character for character in text if not unicodedata.combining(character)
    )
    text = re.sub(r"(?i)\s+to\s+", "_to_", text)
    text = re.sub(r"[^A-Za-z0-9_-]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "Unknown_Route"


def _is_date_folder_name(value) -> bool:
    text = _clean_text(value)
    if not re.fullmatch(r"\d{2}-\d{2}-\d{4}", text):
        return False
    try:
        datetime.strptime(text, "%d-%m-%Y")
    except ValueError:
        return False
    return True


def _source_date_folder(relative_path, verified_start) -> str:
    """Return the original top-level input date folder for one trip."""

    for component in re.split(r"[\\/]+", _clean_text(relative_path)):
        if _is_date_folder_name(component):
            return component
    return pd.Timestamp(verified_start).strftime("%d-%m-%Y")


def _input_date_folders(config) -> list[str]:
    """Return individual date folders that should be mirrored under Paired."""

    input_root = Path(config["project"]["input_folder"])
    if not input_root.exists():
        return []
    folders = [
        path.name
        for path in input_root.iterdir()
        if path.is_dir() and _is_date_folder_name(path.name)
    ]
    return sorted(folders, key=lambda value: datetime.strptime(value, "%d-%m-%Y"))


def _pair_output_date_folder(ev_input_date_folder: str, ice_input_date_folder: str) -> str:
    """Return one unambiguous output folder for a same- or cross-date pair."""

    ev_date = datetime.strptime(ev_input_date_folder, "%d-%m-%Y")
    ice_date = datetime.strptime(ice_input_date_folder, "%d-%m-%Y")
    if ev_date == ice_date:
        return ev_input_date_folder
    first_date, second_date = sorted((ev_date, ice_date))
    return (
        f"{first_date.strftime('%d-%m-%Y')}_to_"
        f"{second_date.strftime('%d-%m-%Y')}"
    )


def _pairable_records(validated_records: list[dict]) -> list[dict]:
    """Keep only distinct candidates that pass every hard crop rule."""

    pairable = [
        record
        for record in validated_records
        if record.get("gps_route_validation_status") in PAIRABLE_ROUTE_STATUSES
    ]
    candidate_ids = [
        _clean_text(record.get("candidate_id")) for record in pairable
    ]
    if len(candidate_ids) != len(set(candidate_ids)):
        raise ValueError("Pairable records contain duplicate candidate IDs.")
    return pairable


def _format_worksheet(
    worksheet,
    dataframe: pd.DataFrame,
    info_sheet=False,
    distance_comparison_sheet=False,
    table_name=None,
):
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    worksheet.row_dimensions[1].height = 34

    if info_sheet:
        worksheet.column_dimensions["A"].width = 55
        worksheet.column_dimensions["B"].width = 85
        for row in worksheet.iter_rows(min_row=2):
            row[0].font = Font(bold=True, color="1F1F1F")
            row[0].alignment = Alignment(vertical="top", wrap_text=True)
            row[1].alignment = Alignment(vertical="top", wrap_text=True)
        return

    for column_number, column_name in enumerate(dataframe.columns, start=1):
        sample = [column_name] + dataframe[column_name].head(300).tolist()
        maximum = max(len(str(value)) if value is not None else 0 for value in sample)
        width = min(max(maximum + 2, 12), 38)
        if column_name == TIMESTAMP_COLUMN:
            width = 21
        worksheet.column_dimensions[get_column_letter(column_number)].width = width
        for row_number in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_number, column=column_number)
            cell.alignment = Alignment(vertical="top", wrap_text=maximum > 34)
            lowered = column_name.casefold()
            if "timestamp" in lowered or lowered.endswith("_start") or lowered.endswith("_end"):
                if isinstance(cell.value, datetime):
                    cell.number_format = "dd-mm-yyyy hh:mm:ss"
            elif "latitude" in lowered or "longitude" in lowered:
                cell.number_format = "0.000000"
            elif "fraction" in lowered:
                cell.number_format = "0.0%"
            elif "distance" in lowered:
                cell.number_format = "0.000000"

    if distance_comparison_sheet:
        for column_number, column_name in enumerate(dataframe.columns, start=1):
            lowered = column_name.casefold()
            if "distance (km)" in lowered or "difference (km)" in lowered:
                for row_number in range(2, worksheet.max_row + 1):
                    worksheet.cell(row=row_number, column=column_number).number_format = (
                        "0.000"
                    )
            elif "%" in lowered:
                for row_number in range(2, worksheet.max_row + 1):
                    worksheet.cell(row=row_number, column=column_number).number_format = (
                        "0.00"
                    )
            elif column_name == "Comparison Result":
                for row_number in range(2, worksheet.max_row + 1):
                    cell = worksheet.cell(row=row_number, column=column_number)
                    if cell.value == "PASS":
                        cell.fill = PatternFill(fill_type="solid", fgColor="70AD47")
                    else:
                        cell.fill = PatternFill(fill_type="solid", fgColor="C00000")
                    cell.font = Font(color="FFFFFF", bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        if table_name and worksheet.max_row >= 2 and worksheet.max_column >= 1:
            table_ref = (
                f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
            )
            table = Table(displayName=table_name, ref=table_ref)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            worksheet.add_table(table)


def _distance_comparison_table(records) -> pd.DataFrame:
    """Build a concise, human-readable EV-versus-ICE distance evidence table."""
    if isinstance(records, dict):
        records = [records]
    rows = []
    for record in records:
        status = _clean_text(record.get("cropped_distance_comparability_status"))
        passed = status == "CROPPED_DISTANCES_WITHIN_TOLERANCE"
        rows.append(
            {
                "Pair ID": record.get("pair_id", pd.NA),
                "Pair Class": record.get("final_pair_class", pd.NA),
                "Route Direction": record.get("route_direction", pd.NA),
                "EV Cropped Distance (km)": record.get(
                    "ev_cropped_haversine_distance_km", pd.NA
                ),
                "ICE Cropped Distance (km)": record.get(
                    "ice_cropped_haversine_distance_km", pd.NA
                ),
                "Distance Difference (km)": record.get(
                    "cropped_distance_difference_km", pd.NA
                ),
                "Distance Difference (% of Shorter)": record.get(
                    "cropped_distance_difference_percent_of_shorter_trip", pd.NA
                ),
                "Allowed Difference (km)": record.get(
                    "allowed_cropped_distance_difference_km", pd.NA
                ),
                "Comparison Result": "PASS" if passed else "FAIL",
                "Evidence Statement": (
                    "EV and ICE cropped distances are within the permitted tolerance."
                    if passed
                    else "EV and ICE cropped distances exceed the permitted tolerance."
                ),
                "Detailed Status": status or "NOT_AVAILABLE",
            }
        )
    return pd.DataFrame(
        rows,
        columns=[
            "Pair ID",
            "Pair Class",
            "Route Direction",
            "EV Cropped Distance (km)",
            "ICE Cropped Distance (km)",
            "Distance Difference (km)",
            "Distance Difference (% of Shorter)",
            "Allowed Difference (km)",
            "Comparison Result",
            "Evidence Statement",
            "Detailed Status",
        ],
    )


def _atomic_write_pair_workbook(
    output_path: Path,
    ev_crop: pd.DataFrame,
    ice_crop: pd.DataFrame,
    info: dict,
):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    info_frame = pd.DataFrame(
        [{"Parameter": key, "Value": value} for key, value in info.items()]
    )
    distance_frame = _distance_comparison_table(info)
    with NamedTemporaryFile(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)
    try:
        with pd.ExcelWriter(
            temporary_path,
            engine="openpyxl",
            datetime_format="dd-mm-yyyy hh:mm:ss",
        ) as writer:
            ev_crop.to_excel(writer, sheet_name=EV_SHEET_NAME, index=False)
            ice_crop.to_excel(writer, sheet_name=ICE_SHEET_NAME, index=False)
            info_frame.to_excel(writer, sheet_name=INFO_SHEET_NAME, index=False)
            distance_frame.to_excel(
                writer, sheet_name=DISTANCE_SHEET_NAME, index=False
            )
            _format_worksheet(writer.sheets[EV_SHEET_NAME], ev_crop)
            _format_worksheet(writer.sheets[ICE_SHEET_NAME], ice_crop)
            _format_worksheet(
                writer.sheets[INFO_SHEET_NAME], info_frame, info_sheet=True
            )
            _format_worksheet(
                writer.sheets[DISTANCE_SHEET_NAME],
                distance_frame,
                distance_comparison_sheet=True,
                table_name="PairDistanceComparison",
            )
        os.replace(temporary_path, output_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _format_audit_workbook(writer, tables):
    for sheet_name, dataframe in tables.items():
        is_distance_table = sheet_name == "Distance Comparison"
        _format_worksheet(
            writer.sheets[sheet_name],
            dataframe,
            distance_comparison_sheet=is_distance_table,
            table_name="AllPairDistanceComparison" if is_distance_table else None,
        )


def _atomic_write_audit_workbook(output_path: Path, tables: dict[str, pd.DataFrame]):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        prefix=".pair_cropping_log_",
        suffix=".xlsx",
        dir=output_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)
    try:
        with pd.ExcelWriter(
            temporary_path,
            engine="openpyxl",
            datetime_format="dd-mm-yyyy hh:mm:ss",
        ) as writer:
            for sheet_name, dataframe in tables.items():
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_audit_workbook(writer, tables)
        os.replace(temporary_path, output_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _atomic_write_csv(output_path: Path, dataframe: pd.DataFrame):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        prefix=".pair_index_",
        suffix=".csv",
        dir=output_path.parent,
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)
    try:
        dataframe.to_csv(temporary_path, index=False, encoding="utf-8-sig")
        os.replace(temporary_path, output_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def _rules_table(rules: dict) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Rule": "Candidate input",
                "Value": f"{CANDIDATE_WORKBOOK_NAME} / {CANDIDATE_SHEET_NAME}",
                "Purpose": "Every preliminary same-direction candidate is evaluated",
            },
            {
                "Rule": "Route-point tolerance",
                "Value": f"{rules['route_point_tolerance_m']:.0f} m",
                "Purpose": "Hard maximum separation for a same-route match",
            },
            {
                "Rule": "Direction agreement",
                "Value": f"<= {rules['maximum_heading_difference_degrees']:.0f} degrees",
                "Purpose": "Reject opposite-direction or cross-road matches",
            },
            {
                "Rule": "Route resampling",
                "Value": f"Nominal {rules['route_resample_spacing_m']:.0f} m; maximum {rules['maximum_resampled_points']} points per trip",
                "Purpose": "Compare routes even when original sampling intervals differ",
            },
            {
                "Rule": "Continuous ordered route overlap",
                "Value": f">= {rules['minimum_ordered_overlap_fraction']:.0%} inside the proposed crop for both trips",
                "Purpose": "Hard rule; both trips must follow the same route in order",
            },
            {
                "Rule": "Maximum internal unmatched run",
                "Value": f"<= {rules['maximum_internal_unmatched_distance_m']:.0f} m",
                "Purpose": "Reject a hidden detour or unmatched internal excursion",
            },
            {
                "Rule": "Minimum common route length",
                "Value": f">= {rules['minimum_common_route_distance_km']:.3f} km",
                "Purpose": "Hard minimum for an operationally meaningful comparison",
            },
            {
                "Rule": "Shorter-trip common coverage",
                "Value": f">= {rules['minimum_shorter_trip_coverage_fraction']:.0%}",
                "Purpose": "Hard rule that prevents a small accidental overlap from becoming a pair",
            },
            {
                "Rule": "Weak route match",
                "Value": "Reject; no full-trip fallback",
                "Purpose": "Only scientifically comparable route sections enter pair_index.csv",
            },
            {
                "Rule": "Common anchors",
                "Value": "Identical interpolated midpoint coordinates at common entry and exit",
                "Purpose": "Both cropped sheets represent the same spatial section; inserted rows are explicitly flagged",
            },
            {
                "Rule": "Full-route reporting threshold",
                "Value": f">= {rules['full_route_coverage_threshold']:.0%} coverage of both verified trips",
                "Purpose": "Below this threshold, report as COMMON_ROUTE_SEGMENT_TAT",
            },
            {
                "Rule": "Cropped-distance equality tolerance",
                "Value": (
                    "Difference <= max("
                    f"{rules['post_crop_distance_absolute_tolerance_km']:.1f} km, "
                    f"{rules['post_crop_distance_relative_tolerance_fraction']:.0%} "
                    "of shorter crop)"
                ),
                "Purpose": "Hard rule checked on the exact rows written to both sheets",
            },
            {
                "Rule": "Sampling warnings",
                "Value": "Recorded but do not block route validation or cropping",
                "Purpose": "A valid route can remain usable, but the limitation stays visible",
            },
            {
                "Rule": "Intercity pair class",
                "Value": "Same calendar day = STRICT; different calendar day = NONSTRICT",
                "Purpose": "Start-time similarity is used only for the four temporal subcategories",
            },
            {
                "Rule": "Trip reuse",
                "Value": "Allowed for distinct preliminary candidates",
                "Purpose": "One trip may have multiple pair-specific cropped copies",
            },
            {
                "Rule": "Pair acceptance",
                "Value": "Write only candidates passing every same-route, anchor, overlap, coverage, and distance rule",
                "Purpose": "Rejected candidates remain visible in the audit workbook only",
            },
            {
                "Rule": "Output date folder",
                "Value": "Same date: DD-MM-YYYY; different dates: DD-MM-YYYY_to_DD-MM-YYYY",
                "Purpose": "Stores each pair once while keeping same-date and cross-date comparisons clear",
            },
            {
                "Rule": "Later stages",
                "Value": "Feature engineering, phase segmentation, TAT, comparisons, and charts are not performed",
                "Purpose": "Keeps Stage 2 auditable and independently rerunnable",
            },
        ]
    )


def _run_summary(
    candidate_count: int,
    audit: pd.DataFrame,
    selected: pd.DataFrame,
    backup_path,
) -> pd.DataFrame:
    route_confirmed = int(
        (audit["gps_route_validation_status"] == "PASSED_ORDERED_ROUTE_VALIDATION").sum()
    )
    route_rejected = int(
        (audit["gps_route_validation_status"] == "REJECTED_ROUTE_VALIDATION").sum()
    )
    processing_failed = int(
        (
            audit["gps_route_validation_status"]
            == "FAILED_INPUT_OR_PROCESSING_ERROR"
        ).sum()
    )
    return pd.DataFrame(
        [
            {"Metric": "Generated at", "Value": datetime.now().isoformat(timespec="seconds")},
            {"Metric": "Preliminary candidates evaluated", "Value": candidate_count},
            {"Metric": "Candidates passing all common-route crop rules", "Value": route_confirmed},
            {"Metric": "Candidates rejected by route/distance rules", "Value": route_rejected},
            {"Metric": "Candidates not created due to input/processing error", "Value": processing_failed},
            {"Metric": "Final pair workbooks created", "Value": len(selected)},
            {"Metric": "Final strict pairs", "Value": int((selected["final_pair_class"] == "STRICT").sum()) if not selected.empty else 0},
            {"Metric": "Final non-strict pairs", "Value": int((selected["final_pair_class"] == "NONSTRICT").sum()) if not selected.empty else 0},
            {"Metric": "Candidates excluded due to route-quality rules", "Value": route_rejected},
            {"Metric": "Candidates excluded due to trip reuse", "Value": 0},
            {"Metric": "Previous Paired folder backup", "Value": str(backup_path) if backup_path else "No previous Paired folder"},
            {"Metric": "Next stage", "Value": "Run feature_engineering.py on the created pair workbooks"},
        ]
    )


def _make_unique_backup_path(results_root: Path) -> Path:
    base = (
        results_root
        / "Backups"
        / f"Pair_Cropping_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        / "Paired"
    )
    candidate = base
    suffix = 1
    while candidate.exists():
        candidate = base.parent.with_name(f"{base.parent.name}_{suffix}") / "Paired"
        suffix += 1
    return candidate


def _replace_paired_folder(
    staging_root: Path,
    paired_root: Path,
    results_root: Path,
    backup_existing: bool,
):
    backup_path = None
    if paired_root.exists():
        backup_path = _make_unique_backup_path(results_root)
        backup_path.parent.mkdir(parents=True, exist_ok=True)
        if backup_existing:
            shutil.move(str(paired_root), str(backup_path))
        else:
            shutil.rmtree(paired_root)
            backup_path = None
    try:
        os.replace(staging_root, paired_root)
    except Exception:
        if backup_path is not None and backup_path.exists() and not paired_root.exists():
            shutil.move(str(backup_path), str(paired_root))
        raise
    return backup_path


def _derive_intercity_pair_class_and_category(candidate: pd.Series) -> tuple[str, str]:
    """
    Intercity project class rule.

    CLASS is determined ONLY by calendar day:
        same day       -> STRICT
        different day  -> NONSTRICT

    The start-time relationship is retained only as a subtype:
        same day + similar time       -> SAME_DAY_SIMILAR_TIME
        same day + different time     -> SAME_DAY_DIFFERENT_TIME
        different day + similar time  -> DIFFERENT_DAY_SIMILAR_TIME
        different day + different time-> DIFFERENT_DAY_DIFFERENT_TIME

    This stage does not alter route validation or distance/QC acceptance.
    """
    same_day = _safe_bool(candidate.get("same_calendar_day"), default=False)
    start_difference = _safe_float(
        candidate.get("start_clock_difference_minutes"),
        default=np.nan,
    )

    if math.isnan(start_difference):
        ev_start = _parse_timestamp(candidate.get("ev_verified_start"))
        ice_start = _parse_timestamp(candidate.get("ice_verified_start"))
        if not pd.isna(ev_start) and not pd.isna(ice_start):
            ev_minutes = (
                ev_start.hour * 60
                + ev_start.minute
                + ev_start.second / 60
            )
            ice_minutes = (
                ice_start.hour * 60
                + ice_start.minute
                + ice_start.second / 60
            )
            direct = abs(ev_minutes - ice_minutes)
            start_difference = min(direct, 1440 - direct)

    # Use the same 30-minute threshold already used by candidate pairing.
    similar_time = (
        not math.isnan(start_difference)
        and start_difference <= 30.0
    )

    if same_day:
        pair_class = "STRICT"
        category = (
            "SAME_DAY_SIMILAR_TIME"
            if similar_time
            else "SAME_DAY_DIFFERENT_TIME"
        )
    else:
        pair_class = "NONSTRICT"
        category = (
            "DIFFERENT_DAY_SIMILAR_TIME"
            if similar_time
            else "DIFFERENT_DAY_DIFFERENT_TIME"
        )

    return pair_class, category


def _base_audit_record(candidate: pd.Series) -> dict:
    pair_sampling = _clean_text(candidate.get("pair_sampling_quality"))
    if not pair_sampling:
        warnings = [
            _clean_text(candidate.get("ev_sampling_quality")),
            _clean_text(candidate.get("ice_sampling_quality")),
        ]
        pair_sampling = "; ".join(value for value in warnings if value)
    ev_input_date_folder = _source_date_folder(
        candidate.get("ev_relative_path"), candidate["ev_verified_start"]
    )
    ice_input_date_folder = _source_date_folder(
        candidate.get("ice_relative_path"), candidate["ice_verified_start"]
    )
    output_date_folder = _pair_output_date_folder(
        ev_input_date_folder, ice_input_date_folder
    )
    # ------------------------------------------------------------
    # INTERCITY CLASSIFICATION
    # ------------------------------------------------------------
    # The final class is derived from the actual calendar dates here,
    # so Stage 3 cannot accidentally turn a same-day pair into NONSTRICT.
    #
    # Same day       -> STRICT
    # Different day  -> NONSTRICT
    #
    # Similar/different start time is only a subtype.
    # Route validation and distance rules below remain unchanged.
    # ------------------------------------------------------------
    derived_pair_class, derived_category = _derive_intercity_pair_class_and_category(
        candidate
    )

    return {
        "candidate_id": candidate["candidate_id"],
        "candidate_class": candidate["candidate_class"],
        "final_pair_class": derived_pair_class,
        "non_strict_category": derived_category,
        "candidate_reason": candidate.get("candidate_reason", pd.NA),
        "route_direction": candidate["route_direction"],
        "ev_verified_trip_id": candidate["ev_verified_trip_id"],
        "ice_verified_trip_id": candidate["ice_verified_trip_id"],
        "ev_file_name": candidate["ev_file_name"],
        "ice_file_name": candidate["ice_file_name"],
        "candidate_same_calendar_day": bool(candidate["same_calendar_day"]),
        "candidate_same_day_type": bool(candidate["same_day_type"]),
        "candidate_start_clock_difference_minutes": float(
            candidate["start_clock_difference_minutes"]
        ),
        "sampling_quality_warning": pair_sampling or "NOT_AVAILABLE",
        "ev_input_date_folder": ev_input_date_folder,
        "ice_input_date_folder": ice_input_date_folder,
        "pair_creation_status": "PENDING_ROUTE_VALIDATION",
        "selection_reason": "PENDING_ROUTE_VALIDATION",
        "final_pair_status": "NOT_FINAL",
        "pair_id": pd.NA,
        "output_date_folder": output_date_folder,
        "pair_workbook_name": pd.NA,
        "pair_workbook_path": pd.NA,
        "common_route_crop_status": "NOT_STARTED",
        "feature_engineering_status": "NOT_STARTED",
    }


def _info_dict(record: dict) -> dict:
    ordered_keys = [
        "pair_id",
        "final_pair_class",
        "candidate_id",
        "route_direction",
        "comparison_scope",
        "selection_reason",
        "final_pair_status",
        "ev_verified_trip_id",
        "ev_file_name",
        "ev_cleaned_file_path",
        "ice_verified_trip_id",
        "ice_file_name",
        "ice_cleaned_file_path",
        "ev_verified_start",
        "ev_verified_end",
        "ice_verified_start",
        "ice_verified_end",
        "ev_input_date_folder",
        "ice_input_date_folder",
        "output_date_folder",
        "ev_verified_full_haversine_distance_km",
        "ice_verified_full_haversine_distance_km",
        "common_start_anchor_latitude",
        "common_start_anchor_longitude",
        "common_end_anchor_latitude",
        "common_end_anchor_longitude",
        "start_anchor_ev_ice_separation_m",
        "end_anchor_ev_ice_separation_m",
        "ev_cropped_start_timestamp",
        "ev_cropped_end_timestamp",
        "ice_cropped_start_timestamp",
        "ice_cropped_end_timestamp",
        "ev_cropped_elapsed_minutes",
        "ice_cropped_elapsed_minutes",
        "ev_cropped_haversine_distance_km",
        "ice_cropped_haversine_distance_km",
        "cropped_distance_difference_km",
        "cropped_distance_difference_percent_of_shorter_trip",
        "allowed_cropped_distance_difference_km",
        "cropped_distance_comparability_status",
        "ev_crop_coverage_fraction_of_verified_trip",
        "ice_crop_coverage_fraction_of_verified_trip",
        "minimum_ordered_overlap_fraction_within_crop",
        "median_heading_difference_degrees",
        "maximum_internal_unmatched_distance_m",
        "common_route_length_km",
        "shorter_trip_common_coverage_fraction",
        "route_point_tolerance_m",
        "route_resample_spacing_m",
        "crop_selection_method",
        "gps_route_validation_status",
        "route_validation_quality",
        "route_validation_reason",
        "sampling_quality_warning",
        "common_route_crop_status",
        "feature_engineering_status",
        "phase_segmentation_status",
        "tat_computation_status",
        "comparison_results_status",
        "visualization_status",
        "phase_time_reconciliation_error_seconds",
    ]
    info = {}
    for key in ordered_keys:
        if key == "phase_segmentation_status":
            value = "NOT_STARTED"
        elif key == "tat_computation_status":
            value = "NOT_STARTED"
        elif key == "comparison_results_status":
            value = "NOT_STARTED"
        elif key == "visualization_status":
            value = "NOT_STARTED"
        elif key == "phase_time_reconciliation_error_seconds":
            value = "NOT_COMPUTED_UNTIL_TAT_STAGE"
        else:
            value = record.get(key, pd.NA)
        info[key] = value
    return info



def _format_duration(seconds: float) -> str:
    total_seconds = max(0, int(round(seconds)))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"



def create_pair_workbooks(
    config,
    candidate_path=None,
    verified_index_path=None,
):
    rules = _cropping_rules(config)
    results_root = Path(config["project"]["results_folder"])
    candidate_path = Path(candidate_path or _candidate_input_path(config))
    verified_index_path = Path(
        verified_index_path or _verified_index_path(config)
    )
    candidates = load_candidates(candidate_path, verified_index_path)

    full_file_cache: dict[str, pd.DataFrame] = {}
    trip_cache: dict[tuple, pd.DataFrame] = {}
    validated_records = []

    started_monotonic = time.monotonic()
    total_candidates = len(candidates)

    print("=" * 72)
    print("STAGE 3: PAIR CROPPING")
    print("=" * 72)
    print(f"Candidates found: {total_candidates}")
    print("Address column will be preserved.")
    print()

    for candidate_number, (_, candidate) in enumerate(
        candidates.iterrows(),
        start=1,
    ):
        candidate_started = time.monotonic()
        print(
            f"[Validation {candidate_number}/{total_candidates}] "
            f"{_clean_text(candidate.get('candidate_id'))} | "
            f"{_clean_text(candidate.get('route_direction'))}",
            flush=True,
        )
        record = _base_audit_record(candidate)
        try:
            ev_path = _resolve_cleaned_path(
                candidate["ev_cleaned_file_path"],
                candidate["ev_relative_path"],
                candidate["ev_file_name"],
                results_root,
            )
            ice_path = _resolve_cleaned_path(
                candidate["ice_cleaned_file_path"],
                candidate["ice_relative_path"],
                candidate["ice_file_name"],
                results_root,
            )
            record["ev_cleaned_file_path"] = str(ev_path)
            record["ice_cleaned_file_path"] = str(ice_path)

            for path in (ev_path, ice_path):
                key = str(path.resolve())
                if key not in full_file_cache:
                    full_file_cache[key] = _read_cleaned_workbook(path)

            ev_key = (
                str(ev_path.resolve()),
                str(candidate["ev_verified_start"]),
                str(candidate["ev_verified_end"]),
            )
            ice_key = (
                str(ice_path.resolve()),
                str(candidate["ice_verified_start"]),
                str(candidate["ice_verified_end"]),
            )
            if ev_key not in trip_cache:
                trip_cache[ev_key] = _verified_segment(
                    full_file_cache[str(ev_path.resolve())],
                    candidate["ev_verified_start"],
                    candidate["ev_verified_end"],
                    rules["minimum_verified_datapoints"],
                )
            if ice_key not in trip_cache:
                trip_cache[ice_key] = _verified_segment(
                    full_file_cache[str(ice_path.resolve())],
                    candidate["ice_verified_start"],
                    candidate["ice_verified_end"],
                    rules["minimum_verified_datapoints"],
                )

            route_result = _route_validation(
                trip_cache[ev_key], trip_cache[ice_key], rules
            )
            record.update(route_result)
            record["ev_verified_start"] = candidate["ev_verified_start"]
            record["ev_verified_end"] = candidate["ev_verified_end"]
            record["ice_verified_start"] = candidate["ice_verified_start"]
            record["ice_verified_end"] = candidate["ice_verified_end"]
            record["_ev_cache_key"] = ev_key
            record["_ice_cache_key"] = ice_key
            if record["gps_route_validation_status"] in PAIRABLE_ROUTE_STATUSES:
                record["pair_creation_status"] = "READY_FOR_PAIR_WORKBOOK"
                record["selection_reason"] = (
                    "PASSED_ALL_COMMON_ROUTE_AND_DISTANCE_RULES"
                )
                record["common_route_crop_status"] = "VALIDATED_NOT_YET_WRITTEN"
            else:
                record["pair_creation_status"] = "NOT_CREATED_ROUTE_REJECTED"
                record["selection_reason"] = "FAILED_HARD_ROUTE_OR_DISTANCE_RULE"
                record["final_pair_status"] = "REJECTED_ROUTE_VALIDATION"
                record["common_route_crop_status"] = "NOT_CREATED_ROUTE_REJECTED"
        except Exception as error:
            record.update(
                {
                    "gps_route_validation_status": "FAILED_INPUT_OR_PROCESSING_ERROR",
                    "route_validation_quality": "NOT_USABLE",
                    "route_validation_reason": str(error),
                    "pair_creation_status": "NOT_CREATED_PROCESSING_FAILED",
                    "selection_reason": "INPUT_OR_PROCESSING_ERROR",
                    "final_pair_status": "REJECTED_PROCESSING_ERROR",
                    "common_route_crop_status": "NOT_CREATED",
                }
            )
        validated_records.append(record)

        candidate_elapsed = time.monotonic() - candidate_started
        print(
            f"    Status: {record.get('gps_route_validation_status', 'UNKNOWN')} | "
            f"Time: {_format_duration(candidate_elapsed)}",
            flush=True,
        )

    selected_records = _pairable_records(validated_records)
    for record in validated_records:
        if record.get("gps_route_validation_status") not in PAIRABLE_ROUTE_STATUSES:
            continue
        record["pair_creation_status"] = "SELECTED_FOR_PAIR_WORKBOOK"
        route_basis = "CONFIRMED_SAME_ROUTE_SAME_ANCHORS_EQUAL_DISTANCE"
        record["selection_reason"] = (
            f"ACCEPTED_{record['final_pair_class']}_{route_basis}"
        )
        record["final_pair_status"] = "SELECTED_FINAL_PAIR"

    selected_records.sort(
        key=lambda record: (
            min(
                datetime.strptime(record["ev_input_date_folder"], "%d-%m-%Y"),
                datetime.strptime(record["ice_input_date_folder"], "%d-%m-%Y"),
            ),
            max(
                datetime.strptime(record["ev_input_date_folder"], "%d-%m-%Y"),
                datetime.strptime(record["ice_input_date_folder"], "%d-%m-%Y"),
            ),
            _clean_text(record["route_direction"]),
            0 if record["final_pair_class"] == "STRICT" else 1,
            pd.Timestamp(record["ev_verified_start"]),
            pd.Timestamp(record["ice_verified_start"]),
            _clean_text(record["candidate_id"]),
        )
    )
    if len(selected_records) > 999:
        raise ValueError(
            "More than 999 final pairs were selected; three-digit PAIR_001 numbering is insufficient."
        )

    staging_root = results_root / f".Paired_staging_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    staging_root.mkdir(parents=True, exist_ok=False)
    mirrored_input_date_folders = _input_date_folders(config)
    for input_date_folder in mirrored_input_date_folders:
        (staging_root / input_date_folder).mkdir(parents=True, exist_ok=True)
    paired_root = results_root / "Paired"
    selected_rows = []
    backup_path = None
    try:
        total_selected = len(selected_records)
        print()
        print(f"Creating {total_selected} pair workbooks...")
        for number, record in enumerate(selected_records, start=1):
            workbook_started = time.monotonic()
            pair_id = f"PAIR_{number:03d}"
            pair_class = record["final_pair_class"]
            date_folder = record["output_date_folder"]
            route_label = _route_file_label(record["route_direction"])
            workbook_name = f"{pair_id}_{pair_class}_{route_label}.xlsx"
            staging_path = staging_root / date_folder / workbook_name
            final_path = paired_root / date_folder / workbook_name

            ev_data = trip_cache[record["_ev_cache_key"]]
            ice_data = trip_cache[record["_ice_cache_key"]]
            start_anchor = (
                record["common_start_anchor_latitude"],
                record["common_start_anchor_longitude"],
            )
            end_anchor = (
                record["common_end_anchor_latitude"],
                record["common_end_anchor_longitude"],
            )
            ev_crop = _build_crop(
                ev_data,
                record["ev_crop_start_progress_m"],
                record["ev_crop_end_progress_m"],
                start_anchor,
                end_anchor,
                pair_id,
                "EV",
                _clean_text(record["ev_verified_trip_id"]),
            )
            ice_crop = _build_crop(
                ice_data,
                record["ice_crop_start_progress_m"],
                record["ice_crop_end_progress_m"],
                start_anchor,
                end_anchor,
                pair_id,
                "ICE",
                _clean_text(record["ice_verified_trip_id"]),
            )
            ev_metrics = _crop_metrics(ev_crop)
            ice_metrics = _crop_metrics(ice_crop)
            (
                difference,
                difference_percent,
                allowed_difference,
                distance_status,
            ) = _distance_comparability(
                ev_metrics["distance_km"], ice_metrics["distance_km"], rules
            )
            start_coordinate_difference = _haversine_m(
                ev_crop[LATITUDE_COLUMN].iloc[0],
                ev_crop[LONGITUDE_COLUMN].iloc[0],
                ice_crop[LATITUDE_COLUMN].iloc[0],
                ice_crop[LONGITUDE_COLUMN].iloc[0],
            )
            end_coordinate_difference = _haversine_m(
                ev_crop[LATITUDE_COLUMN].iloc[-1],
                ev_crop[LONGITUDE_COLUMN].iloc[-1],
                ice_crop[LATITUDE_COLUMN].iloc[-1],
                ice_crop[LONGITUDE_COLUMN].iloc[-1],
            )
            if distance_status != "CROPPED_DISTANCES_WITHIN_TOLERANCE":
                raise RuntimeError(
                    f"{record['candidate_id']} failed the final distance assertion: "
                    f"{difference:.3f} km > {allowed_difference:.3f} km"
                )
            if start_coordinate_difference > 0.001 or end_coordinate_difference > 0.001:
                raise RuntimeError(
                    f"{record['candidate_id']} failed the identical start/end "
                    "coordinate assertion"
                )

            record.update(
                {
                    "pair_id": pair_id,
                    "output_date_folder": date_folder,
                    "pair_workbook_name": workbook_name,
                    "pair_workbook_path": str(final_path),
                    "ev_cropped_datapoints": ev_metrics["datapoints"],
                    "ice_cropped_datapoints": ice_metrics["datapoints"],
                    "ev_cropped_start_timestamp": ev_metrics["start"],
                    "ev_cropped_end_timestamp": ev_metrics["end"],
                    "ice_cropped_start_timestamp": ice_metrics["start"],
                    "ice_cropped_end_timestamp": ice_metrics["end"],
                    "ev_cropped_elapsed_minutes": ev_metrics["elapsed_minutes"],
                    "ice_cropped_elapsed_minutes": ice_metrics["elapsed_minutes"],
                    "ev_cropped_haversine_distance_km": ev_metrics["distance_km"],
                    "ice_cropped_haversine_distance_km": ice_metrics["distance_km"],
                    "cropped_distance_difference_km": difference,
                    "cropped_distance_difference_percent_of_shorter_trip": difference_percent,
                    "allowed_cropped_distance_difference_km": allowed_difference,
                    "cropped_distance_comparability_status": distance_status,
                    "pair_creation_status": "PAIR_WORKBOOK_CREATED",
                    "common_route_crop_status": (
                        "CROPPED_PAIR_WORKBOOK_CREATED_ALL_RULES_PASSED"
                    ),
                    "feature_engineering_status": "NOT_STARTED",
                }
            )
            _atomic_write_pair_workbook(
                staging_path, ev_crop, ice_crop, _info_dict(record)
            )
            selected_rows.append(record.copy())

            elapsed = time.monotonic() - started_monotonic
            average_per_workbook = elapsed / number
            eta = average_per_workbook * (total_selected - number)
            workbook_elapsed = time.monotonic() - workbook_started

            print(
                f"[Workbook {number}/{total_selected}] {workbook_name} | "
                f"EV rows: {len(ev_crop)} | ICE rows: {len(ice_crop)} | "
                f"This file: {_format_duration(workbook_elapsed)} | "
                f"ETA: {_format_duration(eta)}",
                flush=True,
            )

        backup_path = _replace_paired_folder(
            staging_root,
            paired_root,
            results_root,
            rules["backup_existing_pair_folder"],
        )
    except Exception:
        if staging_root.exists():
            shutil.rmtree(staging_root)
        raise

    audit_frame = pd.DataFrame(validated_records)
    for row in selected_rows:
        mask = audit_frame["candidate_id"].astype(str) == _clean_text(row["candidate_id"])
        for key, value in row.items():
            if not key.startswith("_"):
                audit_frame.loc[mask, key] = value
    for column in AUDIT_COLUMNS:
        if column not in audit_frame.columns:
            audit_frame[column] = pd.NA
    audit_frame = audit_frame[AUDIT_COLUMNS]

    selected_frame = audit_frame[
        audit_frame["final_pair_status"] == "SELECTED_FINAL_PAIR"
    ].copy()
    selected_frame = selected_frame.sort_values("pair_id", kind="stable").reset_index(drop=True)
    excluded_frame = audit_frame[
        audit_frame["final_pair_status"] != "SELECTED_FINAL_PAIR"
    ].copy().reset_index(drop=True)

    pair_index_columns = [
        "pair_id",
        "final_pair_class",
        "candidate_id",
        "route_direction",
        "comparison_scope",
        "gps_route_validation_status",
        "route_validation_quality",
        "crop_selection_method",
        "ev_input_date_folder",
        "ice_input_date_folder",
        "output_date_folder",
        "pair_workbook_name",
        "pair_workbook_path",
        "ev_verified_trip_id",
        "ice_verified_trip_id",
        "ev_cropped_start_timestamp",
        "ev_cropped_end_timestamp",
        "ice_cropped_start_timestamp",
        "ice_cropped_end_timestamp",
        "ev_cropped_haversine_distance_km",
        "ice_cropped_haversine_distance_km",
        "cropped_distance_difference_km",
        "cropped_distance_difference_percent_of_shorter_trip",
        "allowed_cropped_distance_difference_km",
        "cropped_distance_comparability_status",
        "common_route_length_km",
        "minimum_ordered_overlap_fraction_within_crop",
        "shorter_trip_common_coverage_fraction",
        "common_start_anchor_latitude",
        "common_start_anchor_longitude",
        "common_end_anchor_latitude",
        "common_end_anchor_longitude",
        "common_route_crop_status",
        "feature_engineering_status",
    ]
    for column in pair_index_columns:
        if column not in selected_frame.columns:
            selected_frame[column] = pd.NA
    pair_index = selected_frame[pair_index_columns].copy()

    log_path = results_root / "Logs" / CROPPING_LOG_NAME
    pair_index_path = results_root / "System" / PAIR_INDEX_NAME
    tables = {
        "Run Summary": _run_summary(
            len(candidates), audit_frame, selected_frame, backup_path
        ),
        "Distance Comparison": _distance_comparison_table(
            selected_frame.to_dict(orient="records")
        ),
        "Selected Pairs": selected_frame,
        "Route Validation Audit": audit_frame,
        "Excluded Candidates": excluded_frame,
        "Rules Applied": _rules_table(rules),
    }
    _atomic_write_audit_workbook(log_path, tables)
    _atomic_write_csv(pair_index_path, pair_index)

    return {
        "candidate_count": len(candidates),
        "route_confirmed_count": int(
            (
                audit_frame["gps_route_validation_status"]
                == "PASSED_ORDERED_ROUTE_VALIDATION"
            ).sum()
        ),
        "route_rejected_count": int(
            (
                audit_frame["gps_route_validation_status"]
                == "REJECTED_ROUTE_VALIDATION"
            ).sum()
        ),
        "processing_failure_count": int(
            (
                audit_frame["gps_route_validation_status"]
                == "FAILED_INPUT_OR_PROCESSING_ERROR"
            ).sum()
        ),
        "selected_pair_count": len(selected_frame),
        "strict_pair_count": int(
            (selected_frame["final_pair_class"] == "STRICT").sum()
        ),
        "nonstrict_pair_count": int(
            (selected_frame["final_pair_class"] == "NONSTRICT").sum()
        ),
        "reused_ev_trip_count": int(
            selected_frame.loc[
                selected_frame["ev_verified_trip_id"].duplicated(False),
                "ev_verified_trip_id",
            ].nunique()
        ),
        "reused_ice_trip_count": int(
            selected_frame.loc[
                selected_frame["ice_verified_trip_id"].duplicated(False),
                "ice_verified_trip_id",
            ].nunique()
        ),
        "mirrored_input_date_folder_count": len(mirrored_input_date_folders),
        "paired_root": str(paired_root),
        "log_path": str(log_path),
        "pair_index_path": str(pair_index_path),
        "backup_path": str(backup_path) if backup_path else None,
    }


if __name__ == "__main__":
    try:
        settings = load_config()
        print(f"Script version: {SCRIPT_VERSION}")
        print("Pair route validation and common-section cropping started")
        result = create_pair_workbooks(settings)
        print("\nPair cropping completed")
        print(f"Preliminary candidates evaluated: {result['candidate_count']}")
        print(
            "Candidates passing all common-route crop rules: "
            f"{result['route_confirmed_count']}"
        )
        print(
            "Candidates rejected by route/distance rules: "
            f"{result['route_rejected_count']}"
        )
        print(
            "Candidates not created due to input/processing error: "
            f"{result['processing_failure_count']}"
        )
        print(
            "Final comparable pair workbooks created: "
            f"{result['selected_pair_count']}"
        )
        print(f"Strict final pairs: {result['strict_pair_count']}")
        print(f"Non-strict final pairs: {result['nonstrict_pair_count']}")
        print(
            "EV trips validly reused across more than one pair: "
            f"{result['reused_ev_trip_count']}"
        )
        print(
            "ICE trips validly reused across more than one pair: "
            f"{result['reused_ice_trip_count']}"
        )
        print(
            "Individual Input date folders mirrored under Paired: "
            f"{result['mirrored_input_date_folder_count']}"
        )
        print(
            "Date-folder rule: same-date pairs use DD-MM-YYYY; cross-date "
            "pairs use DD-MM-YYYY_to_DD-MM-YYYY"
        )
        print(f"Pair workbooks: {result['paired_root']}")
        print(f"Cropping audit: {result['log_path']}")
        print(
            "Distance comparison evidence: 03A_Distance_Comparison in every "
            "pair workbook and Distance Comparison in the cropping audit"
        )
        print(f"Pair index for later stages: {result['pair_index_path']}")
        if result["backup_path"]:
            print(f"Previous Paired folder backed up to: {result['backup_path']}")
        print("\nNext stage: feature_engineering.py")
    except Exception as error:
        print(f"\nPAIR CROPPING FAILED | {error}")
        raise SystemExit(1)