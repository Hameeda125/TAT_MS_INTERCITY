"""Create selective EV-ICE pairing candidates from verified trips.

Run this module only after ``trip_log_reconciliation.py`` reports that the
intended trips are ready for pairing.  It reads ``verified_trip_index.csv`` and
compares EV and ICE trips only when their canonical start and end locations
have the same direction.  It does not create final pairs, crop telemetry,
engineer datapoint features, or calculate TAT outcomes.

The resulting workbook is an audit gate for the next stage: ordered GPS route
validation and pair-specific common-section cropping.  Candidate IDs are
simple review identifiers (CAND_001, CAND_002, ...).  Pre-crop distance and
sampling checks are quality flags; they do not turn a candidate into a final
pair or use TAT outcomes during pairing.
"""

from __future__ import annotations

import os
import re
import unicodedata
from datetime import datetime
from itertools import product
from math import atan2, cos, radians, sin, sqrt
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config_loader import load_config


VERIFIED_INDEX_FILE_NAME = "verified_trip_index.csv"
OUTPUT_FILE_NAME = "Pairing_Candidates.xlsx"

READY_STATUS = "READY_FOR_PAIRING"
STRICT_CLASS = "STRICT_CANDIDATE"
NON_STRICT_CLASS = "NON_STRICT_CANDIDATE"

REQUIRED_COLUMNS = {
    "verified_trip_id",
    "reconciliation_status",
    "source_file_name",
    "source_relative_path",
    "source_cleaned_file_path",
    "vehicle_type",
    "log_start_location",
    "log_end_location",
    "verified_start_timestamp",
    "verified_end_timestamp",
    "verified_start_latitude",
    "verified_start_longitude",
    "verified_end_latitude",
    "verified_end_longitude",
}

# These aliases are intentionally aligned with trip_log_reconciliation.py.
# Add only auditable spelling variants; do not merge genuinely different places.
LOCATION_ALIASES = {
    "cmbt": "cmbt",
    "cbmt": "cmbt",
    "cbmtbusstand": "cmbt",
    "cmbtbusstand": "cmbt",
    "koyambedu": "cmbt",
    "koyambedubusstand": "cmbt",
    "kcbt": "kcbt",
    "kilambakkam": "kcbt",
    "sholinganallur": "sholinganallur",
    "shollinganallur": "sholinganallur",
    "sholliganallur": "sholinganallur",
    "silkboard": "silkboard",
    "maduravoyal": "maduravoyal",
    "madhuravoyal": "maduravoyal",
    "maduravayol": "maduravoyal",
    "tnagar": "tnagar",
    "thiruporur": "thiruporur",
    "thirupporur": "thiruporur",
    "thirupur": "thiruporur",
    "tiruporur": "thiruporur",
    "pondy": "pondicherry",
    "pondi": "pondicherry",
    "pondicherry": "pondicherry",
    "puducherry": "pondicherry",
    "kia": "kia",
    "kempegowdainternationalairport": "kia",
    "kadugodibusstand": "kadugodi",
    "kadugodi": "kadugodi",
    "whitefield": "whitefield",
    "hebbal": "hebbal",
    "bellandur": "bellandur",
    "avrroundanasalem": "avrroundanasalem",
    "avrroundana": "avrroundanasalem",
    "salem": "avrroundanasalem",
    "chennai": "chennai",
    "navalur": "navalur",
}

LOCATION_DISPLAY = {
    "cmbt": "CMBT",
    "kcbt": "KCBT",
    "sholinganallur": "Sholinganallur",
    "silkboard": "SilkBoard",
    "maduravoyal": "Madhuravoyal",
    "tnagar": "TNagar",
    "thiruporur": "Thiruporur",
    "pondicherry": "Pondicherry",
    "kia": "KIA",
    "kadugodi": "Kadugodi",
    "whitefield": "Whitefield",
    "hebbal": "Hebbal",
    "bellandur": "Bellandur",
    "avrroundanasalem": "AVRRoundanaSalem",
    "chennai": "Chennai",
    "navalur": "Navalur",
}

CANDIDATE_COLUMNS = [
    "candidate_id",
    "candidate_class",
    "non_strict_category",
    "candidate_reason",
    "candidate_processing_status",
    "final_pair_status",
    "gps_route_validation_status",
    "common_route_crop_status",
    "feature_engineering_status",
    "route_direction",
    "planned_output_date_folder_after_acceptance",
    "candidate_priority_within_route",
    "ev_verified_trip_id",
    "ev_file_name",
    "ev_relative_path",
    "ev_cleaned_file_path",
    "ice_verified_trip_id",
    "ice_file_name",
    "ice_relative_path",
    "ice_cleaned_file_path",
    "ev_log_start",
    "ice_log_start",
    "ev_verified_start",
    "ev_verified_end",
    "ice_verified_start",
    "ice_verified_end",
    "same_calendar_day",
    "date_difference_days",
    "start_clock_difference_minutes",
    "same_day_type",
    "ev_day_type",
    "ice_day_type",
    "start_separation_m",
    "end_separation_m",
    "start_endpoint_within_tolerance",
    "end_endpoint_within_tolerance",
    "both_endpoints_within_tolerance",
    "full_trip_endpoint_status",
    "ev_pre_crop_verified_distance_km",
    "ice_pre_crop_verified_distance_km",
    "pre_crop_distance_difference_km",
    "pre_crop_distance_difference_percent_of_shorter_trip",
    "pre_crop_distance_status",
    "ev_verified_datapoints",
    "ice_verified_datapoints",
    "ev_verified_duration_minutes",
    "ice_verified_duration_minutes",
    "ev_average_sampling_interval_seconds",
    "ice_average_sampling_interval_seconds",
    "ev_sampling_quality",
    "ice_sampling_quality",
    "pair_sampling_quality",
]


def _clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _compact_ascii(value) -> str:
    normalized = unicodedata.normalize("NFKD", _clean_text(value))
    ascii_text = "".join(
        character
        for character in normalized
        if not unicodedata.combining(character)
    )
    return re.sub(r"[^a-z0-9]+", "", ascii_text.casefold())


def canonical_location(value) -> str:
    compact = _compact_ascii(value)
    if not compact:
        return ""
    return LOCATION_ALIASES.get(compact, compact)


def _display_location(canonical_name: str) -> str:
    if canonical_name in LOCATION_DISPLAY:
        return LOCATION_DISPLAY[canonical_name]
    if not canonical_name:
        return "Unknown"
    return canonical_name[:1].upper() + canonical_name[1:]


def _normalise_vehicle_type(value) -> str:
    compact = _compact_ascii(value)
    if compact in {"ev", "electric", "electricvehicle", "bev"}:
        return "EV"
    if compact in {"ice", "diesel", "internalcombustion", "combustion"}:
        return "ICE"
    return "UNKNOWN"


def _read_csv_with_fallback(path: Path) -> pd.DataFrame:
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(path, encoding=encoding)
        except UnicodeDecodeError as error:
            last_error = error
    raise ValueError(f"Could not decode CSV file {path}: {last_error}")


def _parse_datetime_series(series: pd.Series) -> pd.Series:
    # Reconciliation commonly writes DD-MM-YYYY timestamps.  A second ISO pass
    # safely handles future versions that use YYYY-MM-DD.
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    unresolved = parsed.isna() & series.notna()
    if bool(unresolved.any()):
        parsed.loc[unresolved] = pd.to_datetime(
            series.loc[unresolved], errors="coerce"
        )
    return parsed


def _haversine_m(lat1, lon1, lat2, lon2) -> float:
    earth_radius_m = 6_371_008.8
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
    haversine_value = min(1.0, max(0.0, haversine_value))
    central_angle = 2 * atan2(
        sqrt(haversine_value), sqrt(1.0 - haversine_value)
    )
    return earth_radius_m * central_angle


def _circular_clock_difference_minutes(first, second) -> float:
    first_minutes = (
        first.hour * 60
        + first.minute
        + first.second / 60
        + first.microsecond / 60_000_000
    )
    second_minutes = (
        second.hour * 60
        + second.minute
        + second.second / 60
        + second.microsecond / 60_000_000
    )
    direct = abs(first_minutes - second_minutes)
    return min(direct, 1440.0 - direct)


def _day_type(timestamp) -> str:
    return "WEEKEND" if timestamp.weekday() >= 5 else "WEEKDAY"


def _rules_from_config(config) -> dict:
    configured = config.get("pairing", {})
    return {
        "start_location_tolerance_m": float(
            configured.get(
                "start_location_tolerance_m",
                configured.get("endpoint_tolerance_m", 500.0),
            )
        ),
        "end_location_tolerance_m": float(
            configured.get(
                "end_location_tolerance_m",
                configured.get("endpoint_tolerance_m", 500.0),
            )
        ),
        "require_same_route": bool(
            configured.get("require_same_route", True)
        ),
        "strict_pair_same_day": bool(
            configured.get("strict_pair_same_day", True)
        ),
        "strict_pair_similar_start_time": bool(
            configured.get("strict_pair_similar_start_time", True)
        ),
        "strict_start_time_tolerance_minutes": float(
            configured.get(
                "strict_start_time_tolerance_minutes",
                configured.get("strict_start_difference_minutes", 30.0),
            )
        ),
        "allow_non_strict_pairs": bool(
            configured.get(
                "allow_non_strict_pairs",
                configured.get("analyse_non_strict_pairs", True),
            )
        ),
        "route_point_tolerance_m": float(
            configured.get("route_point_tolerance_m", 100.0)
        ),
        "minimum_route_overlap_fraction": float(
            configured.get("minimum_route_overlap_fraction", 0.80)
        ),
        "pre_crop_distance_absolute_tolerance_km": float(
            configured.get(
                "pre_crop_distance_absolute_tolerance_km", 2.0
            )
        ),
        "pre_crop_distance_relative_tolerance_fraction": float(
            configured.get(
                "pre_crop_distance_relative_tolerance_fraction", 0.10
            )
        ),
        "maximum_average_sampling_interval_seconds": float(
            configured.get(
                "maximum_average_sampling_interval_seconds", 30.0
            )
        ),
    }


def _verified_index_path(config) -> Path:
    return (
        Path(config["project"]["results_folder"])
        / "System"
        / VERIFIED_INDEX_FILE_NAME
    )


def _output_path(config) -> Path:
    folder = Path(config["project"]["results_folder"]) / "Pairing"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / OUTPUT_FILE_NAME


def load_verified_trips(index_path):
    """Load the reconciliation output and prepare only ready trips."""

    index_path = Path(index_path)
    if not index_path.exists():
        raise FileNotFoundError(
            f"Verified trip index not found: {index_path}. "
            "Run trip_log_reconciliation.py first."
        )

    index = _read_csv_with_fallback(index_path)
    missing = sorted(REQUIRED_COLUMNS.difference(index.columns))
    if missing:
        raise ValueError(
            "Verified trip index is missing required fields: "
            + ", ".join(missing)
        )

    index = index.copy()
    index["vehicle_type"] = index["vehicle_type"].map(
        _normalise_vehicle_type
    )
    unknown = index[index["vehicle_type"] == "UNKNOWN"]
    if not unknown.empty:
        names = ", ".join(
            unknown["source_file_name"].astype(str).head(10).tolist()
        )
        raise ValueError(
            f"{len(unknown)} trip(s) have UNKNOWN vehicle type: {names}"
        )

    if index["verified_trip_id"].astype(str).duplicated().any():
        duplicates = index.loc[
            index["verified_trip_id"].astype(str).duplicated(False),
            "verified_trip_id",
        ].astype(str)
        raise ValueError(
            "Duplicate verified_trip_id values found: "
            + ", ".join(duplicates.drop_duplicates().head(10))
        )

    ready = index[index["reconciliation_status"] == READY_STATUS].copy()
    not_ready = index[index["reconciliation_status"] != READY_STATUS].copy()
    if ready.empty:
        raise ValueError(
            "No READY_FOR_PAIRING trips were found in the verified index."
        )

    datetime_columns = [
        "resolved_log_start",
        "resolved_log_end",
        "verified_start_timestamp",
        "verified_end_timestamp",
    ]
    for column in datetime_columns:
        if column in ready.columns:
            ready[column] = _parse_datetime_series(ready[column])

    ready["pair_start_timestamp"] = ready.get(
        "resolved_log_start",
        pd.Series(pd.NaT, index=ready.index),
    )
    ready["pair_start_timestamp"] = ready[
        "pair_start_timestamp"
    ].fillna(ready["verified_start_timestamp"])

    coordinate_columns = [
        "verified_start_latitude",
        "verified_start_longitude",
        "verified_end_latitude",
        "verified_end_longitude",
    ]
    for column in coordinate_columns:
        ready[column] = pd.to_numeric(ready[column], errors="coerce")

    invalid = ready["pair_start_timestamp"].isna()
    invalid |= ready["verified_start_timestamp"].isna()
    invalid |= ready["verified_end_timestamp"].isna()
    invalid |= ready[coordinate_columns].isna().any(axis=1)
    if bool(invalid.any()):
        names = ", ".join(
            ready.loc[invalid, "source_file_name"]
            .astype(str)
            .head(10)
            .tolist()
        )
        raise ValueError(
            f"{int(invalid.sum())} READY trip(s) still have invalid pairing "
            f"timestamps or coordinates: {names}"
        )

    ready["origin_key"] = ready["log_start_location"].map(
        canonical_location
    )
    ready["destination_key"] = ready["log_end_location"].map(
        canonical_location
    )
    missing_route = (
        ready["origin_key"].eq("") | ready["destination_key"].eq("")
    )
    if bool(missing_route.any()):
        names = ", ".join(
            ready.loc[missing_route, "source_file_name"]
            .astype(str)
            .head(10)
            .tolist()
        )
        raise ValueError(
            f"{int(missing_route.sum())} READY trip(s) have missing route "
            f"endpoints: {names}"
        )

    ready["route_key"] = (
        ready["origin_key"] + "__to__" + ready["destination_key"]
    )
    ready["route_direction"] = ready.apply(
        lambda row: (
            f"{_display_location(row['origin_key'])} to "
            f"{_display_location(row['destination_key'])}"
        ),
        axis=1,
    )
    ready["route_file_label"] = ready.apply(
        lambda row: (
            f"{_display_location(row['origin_key'])}_to_"
            f"{_display_location(row['destination_key'])}"
        ),
        axis=1,
    )

    for column in [
        "verified_distance_km",
        "verified_points",
        "verified_duration_minutes",
        "raw_recalculated_distance_km",
        "raw_index_distance_km",
        "raw_index_points",
    ]:
        if column not in ready.columns:
            ready[column] = pd.NA
        ready[column] = pd.to_numeric(ready[column], errors="coerce")

    calculated_duration = (
        ready["verified_end_timestamp"]
        - ready["verified_start_timestamp"]
    ).dt.total_seconds() / 60.0
    ready["verified_duration_minutes"] = ready[
        "verified_duration_minutes"
    ].fillna(calculated_duration)

    return index, ready.reset_index(drop=True), not_ready.reset_index(drop=True)


def _temporal_non_strict_category(same_day, similar_clock) -> str:
    if same_day and similar_clock:
        return "SAME_DAY_SIMILAR_TIME_COMMON_SECTION_REQUIRED"
    if same_day:
        return "SAME_DAY_DIFFERENT_TIME"
    if similar_clock:
        return "DIFFERENT_DAY_SIMILAR_TIME"
    return "DIFFERENT_DAY_DIFFERENT_TIME"


def _full_trip_endpoint_status(
    start_pass,
    end_pass,
    start_tolerance_m,
    end_tolerance_m,
) -> str:
    start_label = f"{int(round(start_tolerance_m))}M"
    end_label = f"{int(round(end_tolerance_m))}M"
    if start_pass and end_pass:
        return (
            f"START_WITHIN_{start_label}_AND_END_WITHIN_{end_label}"
        )
    if not start_pass and not end_pass:
        return (
            f"START_OUTSIDE_{start_label}_AND_END_OUTSIDE_{end_label}"
        )
    if not start_pass:
        return f"START_OUTSIDE_{start_label}_END_WITHIN_{end_label}"
    return f"START_WITHIN_{start_label}_END_OUTSIDE_{end_label}"


def _average_sampling_interval_seconds(duration_minutes, datapoints):
    if pd.isna(duration_minutes) or pd.isna(datapoints):
        return pd.NA
    datapoints = int(datapoints)
    if datapoints < 2 or float(duration_minutes) < 0:
        return pd.NA
    return float(duration_minutes) * 60.0 / (datapoints - 1)


def _sampling_quality(average_interval_seconds, datapoints, rules) -> str:
    if pd.isna(datapoints) or int(datapoints) < 2:
        return "INSUFFICIENT_DATAPOINTS"
    if pd.isna(average_interval_seconds):
        return "SAMPLING_INTERVAL_CHECK_REQUIRED"
    if (
        float(average_interval_seconds)
        > rules["maximum_average_sampling_interval_seconds"]
    ):
        return "SPARSE_TELEMETRY_REVIEW_REQUIRED"
    return "PRELIMINARILY_SUITABLE_FOR_PHASE_ANALYSIS"


def _pre_crop_distance_quality(
    ev_distance,
    ice_distance,
    both_endpoints_within_tolerance,
    rules,
):
    if pd.isna(ev_distance) or pd.isna(ice_distance):
        return pd.NA, pd.NA, "PRE_CROP_DISTANCE_MISSING_REVIEW_REQUIRED"

    ev_distance = float(ev_distance)
    ice_distance = float(ice_distance)
    difference_km = abs(ev_distance - ice_distance)
    shorter_distance = min(ev_distance, ice_distance)
    difference_fraction = (
        difference_km / shorter_distance
        if shorter_distance > 0
        else float("inf")
    )

    if not both_endpoints_within_tolerance:
        status = "INFORMATION_ONLY_COMMON_SECTION_CROP_PENDING"
    elif (
        difference_km
        <= rules["pre_crop_distance_absolute_tolerance_km"]
        or difference_fraction
        <= rules["pre_crop_distance_relative_tolerance_fraction"]
    ):
        status = "PRE_CROP_DISTANCE_WITHIN_TOLERANCE"
    else:
        status = "LARGE_PRE_CROP_DISTANCE_DIFFERENCE_REVIEW_REQUIRED"

    return difference_km, difference_fraction * 100.0, status


def _candidate_record(ev_trip, ice_trip, rules) -> dict | None:
    ev_start = pd.Timestamp(ev_trip["pair_start_timestamp"])
    ice_start = pd.Timestamp(ice_trip["pair_start_timestamp"])
    same_day = ev_start.date() == ice_start.date()
    date_difference_days = abs((ev_start.date() - ice_start.date()).days)
    clock_difference = _circular_clock_difference_minutes(
        ev_start, ice_start
    )
    similar_clock = (
        clock_difference
        <= rules["strict_start_time_tolerance_minutes"]
    )

    start_separation = _haversine_m(
        ev_trip["verified_start_latitude"],
        ev_trip["verified_start_longitude"],
        ice_trip["verified_start_latitude"],
        ice_trip["verified_start_longitude"],
    )
    end_separation = _haversine_m(
        ev_trip["verified_end_latitude"],
        ev_trip["verified_end_longitude"],
        ice_trip["verified_end_latitude"],
        ice_trip["verified_end_longitude"],
    )
    start_pass = start_separation <= rules["start_location_tolerance_m"]
    end_pass = end_separation <= rules["end_location_tolerance_m"]
    endpoint_pass = start_pass and end_pass

    # ------------------------------------------------------------
    # ORIGINAL PAIRING LOGIC
    # ------------------------------------------------------------
    # Keep the old eligibility logic unchanged:
    # STRICT requires:
    #   1. same canonical route/direction
    #   2. both endpoints within the configured tolerance
    #   3. same calendar day
    #   4. start-clock difference within the configured tolerance
    #
    # The new intercity pipeline only adds the requested temporal
    # category labels. It does not change candidate eligibility.
    # ------------------------------------------------------------
    # ------------------------------------------------------------
    # CLASSIFICATION FOR THE INTERCITY-ONLY PIPELINE
    # ------------------------------------------------------------
    # Pair eligibility remains based on the original endpoint/route
    # validation.  The CLASS is now determined by calendar day:
    #
    #   SAME DAY       -> STRICT
    #   DIFFERENT DAY  -> NON-STRICT
    #
    # Start-time similarity is retained only as a SUB-CATEGORY:
    #
    #   STRICT:
    #       SAME_DAY_SIMILAR_TIME
    #       SAME_DAY_DIFFERENT_TIME
    #
    #   NON-STRICT:
    #       DIFFERENT_DAY_SIMILAR_TIME
    #       DIFFERENT_DAY_DIFFERENT_TIME
    #
    # Therefore, a same-day pair is STRICT even when its start times
    # differ by more than the 30-minute similarity threshold.
    # ------------------------------------------------------------

    strict_day_pass = (
        same_day
        or not rules["strict_pair_same_day"]
    )

    # Keep the configured time comparison because it is still needed
    # for the four-way category.
    strict_time_pass = (
        similar_clock
        or not rules["strict_pair_similar_start_time"]
    )

    # IMPORTANT:
    # Start-time similarity is NO LONGER a class gate.
    # Same-day pairs are STRICT; different-day pairs are NON-STRICT.
    strict_pass = endpoint_pass and strict_day_pass

    if strict_pass:
        candidate_class = STRICT_CLASS

        if same_day and similar_clock:
            non_strict_category = "SAME_DAY_SIMILAR_TIME"
        else:
            non_strict_category = "SAME_DAY_DIFFERENT_TIME"

        reason = (
            "SAME_DIRECTION_ENDPOINT_AND_SAME_DAY_RULES_PASSED"
            if similar_clock
            else
            "SAME_DIRECTION_ENDPOINT_AND_SAME_DAY_RULES_PASSED_START_TIME_DIFFERENCE_ALLOWED"
        )

    elif rules["allow_non_strict_pairs"]:
        candidate_class = NON_STRICT_CLASS

        if not same_day and similar_clock:
            non_strict_category = "DIFFERENT_DAY_SIMILAR_TIME"
        else:
            non_strict_category = "DIFFERENT_DAY_DIFFERENT_TIME"

        reasons = []

        if not endpoint_pass:
            reasons.append(
                "COMMON_SECTION_OR_ENDPOINT_REVIEW_REQUIRED"
            )

        if not same_day:
            reasons.append(
                "DIFFERENT_CALENDAR_DAY"
            )

        reason = (
            "; ".join(reasons)
            if reasons
            else "DIFFERENT_DAY_NON_STRICT_PAIR"
        )

    else:
        return None

    ev_day_type = _day_type(ev_start)
    ice_day_type = _day_type(ice_start)
    ev_date = ev_start.strftime("%d-%m-%Y")
    ice_date = ice_start.strftime("%d-%m-%Y")
    date_folder = ev_date if same_day else f"{ev_date}_to_{ice_date}"

    ev_distance = ev_trip.get("verified_distance_km")
    ice_distance = ice_trip.get("verified_distance_km")
    (
        distance_difference,
        distance_difference_percent,
        distance_status,
    ) = _pre_crop_distance_quality(
        ev_distance,
        ice_distance,
        endpoint_pass,
        rules,
    )

    ev_points = ev_trip.get("verified_points", pd.NA)
    ice_points = ice_trip.get("verified_points", pd.NA)
    ev_duration = ev_trip.get("verified_duration_minutes", pd.NA)
    ice_duration = ice_trip.get("verified_duration_minutes", pd.NA)
    ev_average_interval = _average_sampling_interval_seconds(
        ev_duration, ev_points
    )
    ice_average_interval = _average_sampling_interval_seconds(
        ice_duration, ice_points
    )
    ev_sampling_quality = _sampling_quality(
        ev_average_interval, ev_points, rules
    )
    ice_sampling_quality = _sampling_quality(
        ice_average_interval, ice_points, rules
    )
    phase_ready_label = "PRELIMINARILY_SUITABLE_FOR_PHASE_ANALYSIS"
    if (
        ev_sampling_quality == phase_ready_label
        and ice_sampling_quality == phase_ready_label
    ):
        pair_sampling_quality = (
            "BOTH_TRIPS_PRELIMINARILY_SUITABLE_FOR_PHASE_ANALYSIS"
        )
    else:
        pair_sampling_quality = "TELEMETRY_QUALITY_REVIEW_REQUIRED"

    if pair_sampling_quality == "TELEMETRY_QUALITY_REVIEW_REQUIRED":
        processing_status = "TELEMETRY_QUALITY_REVIEW_REQUIRED"
    elif distance_status in {
        "PRE_CROP_DISTANCE_MISSING_REVIEW_REQUIRED",
        "LARGE_PRE_CROP_DISTANCE_DIFFERENCE_REVIEW_REQUIRED",
    }:
        processing_status = "PRE_CROP_DISTANCE_REVIEW_REQUIRED"
    else:
        processing_status = "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION"

    gps_validation_status = (
        "PENDING_ORDERED_GPS_ROUTE_VALIDATION"
        if processing_status == "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION"
        else "PENDING_AFTER_DATA_QUALITY_REVIEW"
    )

    return {
        "candidate_id": pd.NA,
        "candidate_class": candidate_class,
        "non_strict_category": non_strict_category,
        "candidate_reason": reason,
        "candidate_processing_status": processing_status,
        "final_pair_status": (
            "NOT_FINAL_UNTIL_GPS_ROUTE_AND_COMMON_CROP_QC_PASS"
        ),
        "gps_route_validation_status": gps_validation_status,
        "common_route_crop_status": "NOT_STARTED",
        "feature_engineering_status": (
            "NOT_STARTED_UNTIL_AFTER_COMMON_ROUTE_CROP"
        ),
        "route_direction": ev_trip["route_direction"],
        "planned_output_date_folder_after_acceptance": date_folder,
        "candidate_priority_within_route": pd.NA,
        "ev_verified_trip_id": ev_trip["verified_trip_id"],
        "ev_file_name": ev_trip["source_file_name"],
        "ev_relative_path": ev_trip["source_relative_path"],
        "ev_cleaned_file_path": ev_trip["source_cleaned_file_path"],
        "ice_verified_trip_id": ice_trip["verified_trip_id"],
        "ice_file_name": ice_trip["source_file_name"],
        "ice_relative_path": ice_trip["source_relative_path"],
        "ice_cleaned_file_path": ice_trip["source_cleaned_file_path"],
        "ev_log_start": ev_start,
        "ice_log_start": ice_start,
        "ev_verified_start": ev_trip["verified_start_timestamp"],
        "ev_verified_end": ev_trip["verified_end_timestamp"],
        "ice_verified_start": ice_trip["verified_start_timestamp"],
        "ice_verified_end": ice_trip["verified_end_timestamp"],
        "same_calendar_day": same_day,
        "date_difference_days": date_difference_days,
        "start_clock_difference_minutes": round(clock_difference, 3),
        "same_day_type": ev_day_type == ice_day_type,
        "ev_day_type": ev_day_type,
        "ice_day_type": ice_day_type,
        "start_separation_m": round(start_separation, 3),
        "end_separation_m": round(end_separation, 3),
        "start_endpoint_within_tolerance": start_pass,
        "end_endpoint_within_tolerance": end_pass,
        "both_endpoints_within_tolerance": endpoint_pass,
        "full_trip_endpoint_status": _full_trip_endpoint_status(
            start_pass,
            end_pass,
            rules["start_location_tolerance_m"],
            rules["end_location_tolerance_m"],
        ),
        "ev_pre_crop_verified_distance_km": (
            round(float(ev_distance), 6) if pd.notna(ev_distance) else pd.NA
        ),
        "ice_pre_crop_verified_distance_km": (
            round(float(ice_distance), 6) if pd.notna(ice_distance) else pd.NA
        ),
        "pre_crop_distance_difference_km": (
            round(float(distance_difference), 6)
            if pd.notna(distance_difference)
            else pd.NA
        ),
        "pre_crop_distance_difference_percent_of_shorter_trip": (
            round(float(distance_difference_percent), 3)
            if pd.notna(distance_difference_percent)
            else pd.NA
        ),
        "pre_crop_distance_status": distance_status,
        "ev_verified_datapoints": ev_points,
        "ice_verified_datapoints": ice_points,
        "ev_verified_duration_minutes": (
            round(float(ev_duration), 3)
            if pd.notna(ev_duration)
            else pd.NA
        ),
        "ice_verified_duration_minutes": (
            round(float(ice_duration), 3)
            if pd.notna(ice_duration)
            else pd.NA
        ),
        "ev_average_sampling_interval_seconds": (
            round(float(ev_average_interval), 3)
            if pd.notna(ev_average_interval)
            else pd.NA
        ),
        "ice_average_sampling_interval_seconds": (
            round(float(ice_average_interval), 3)
            if pd.notna(ice_average_interval)
            else pd.NA
        ),
        "ev_sampling_quality": ev_sampling_quality,
        "ice_sampling_quality": ice_sampling_quality,
        "pair_sampling_quality": pair_sampling_quality,
    }


def build_candidate_tables(ready_trips, rules):
    """Compare EV/ICE trips only inside the same canonical route group."""

    records = []
    route_rows = []

    for route_key, group in ready_trips.groupby("route_key", sort=True):
        ev_trips = group[group["vehicle_type"] == "EV"]
        ice_trips = group[group["vehicle_type"] == "ICE"]
        route_label = group["route_direction"].iloc[0]
        evaluated = len(ev_trips) * len(ice_trips)
        route_records = []

        for (_, ev_trip), (_, ice_trip) in product(
            ev_trips.iterrows(), ice_trips.iterrows()
        ):
            record = _candidate_record(ev_trip, ice_trip, rules)
            if record is not None:
                route_records.append(record)

        records.extend(route_records)
        route_strict_count = sum(
            record["candidate_class"] == STRICT_CLASS
            for record in route_records
        )
        route_non_strict_count = sum(
            record["candidate_class"] == NON_STRICT_CLASS
            for record in route_records
        )
        route_rows.append(
            {
                "route_key": route_key,
                "route_direction": route_label,
                "ev_trips": len(ev_trips),
                "ice_trips": len(ice_trips),
                "strict_preliminary_candidates": route_strict_count,
                "non_strict_preliminary_candidates": route_non_strict_count,
                "preliminary_candidates_for_review": len(route_records),
                "route_candidate_status": (
                    "HAS_EV_AND_ICE_CANDIDATES"
                    if evaluated > 0
                    else "NO_OPPOSITE_VEHICLE_SAME_DIRECTION"
                ),
            }
        )

    candidates = pd.DataFrame(records, columns=CANDIDATE_COLUMNS)
    route_summary = pd.DataFrame(route_rows)

    if not candidates.empty:
        candidates["_strict_order"] = candidates["candidate_class"].map(
            {STRICT_CLASS: 0, NON_STRICT_CLASS: 1}
        )
        candidates["_processing_order"] = candidates[
            "candidate_processing_status"
        ].map(
            {
                "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION": 0,
                "PRE_CROP_DISTANCE_REVIEW_REQUIRED": 1,
                "TELEMETRY_QUALITY_REVIEW_REQUIRED": 2,
            }
        ).fillna(3)
        candidates["_endpoint_order"] = (~candidates[
            "both_endpoints_within_tolerance"
        ]).astype(int)
        candidates = candidates.sort_values(
            [
                "route_direction",
                "_strict_order",
                "_processing_order",
                "_endpoint_order",
                "date_difference_days",
                "start_clock_difference_minutes",
                "start_separation_m",
                "end_separation_m",
                "ev_verified_trip_id",
                "ice_verified_trip_id",
            ],
            kind="stable",
        ).reset_index(drop=True)
        candidates["candidate_priority_within_route"] = (
            candidates.groupby("route_direction").cumcount() + 1
        )
        if len(candidates) > 999:
            raise ValueError(
                "More than 999 candidates were generated; the agreed "
                "three-digit CAND_001 numbering is insufficient."
            )
        candidates["candidate_id"] = [
            f"CAND_{number:03d}"
            for number in range(1, len(candidates) + 1)
        ]
        candidates = candidates.drop(
            columns=[
                "_strict_order",
                "_processing_order",
                "_endpoint_order",
            ]
        ).reindex(columns=CANDIDATE_COLUMNS)

    coverage_rows = []
    for _, trip in ready_trips.iterrows():
        if trip["vehicle_type"] == "EV":
            matches = candidates[
                candidates["ev_verified_trip_id"]
                == trip["verified_trip_id"]
            ]
        else:
            matches = candidates[
                candidates["ice_verified_trip_id"]
                == trip["verified_trip_id"]
            ]
        coverage_rows.append(
            {
                "verified_trip_id": trip["verified_trip_id"],
                "vehicle_type": trip["vehicle_type"],
                "file_name": trip["source_file_name"],
                "route_direction": trip["route_direction"],
                "verified_datapoints": trip.get("verified_points", pd.NA),
                "verified_duration_minutes": trip.get(
                    "verified_duration_minutes", pd.NA
                ),
                "average_sampling_interval_seconds": round(
                    float(
                        _average_sampling_interval_seconds(
                            trip.get("verified_duration_minutes", pd.NA),
                            trip.get("verified_points", pd.NA),
                        )
                    ),
                    3,
                )
                if pd.notna(
                    _average_sampling_interval_seconds(
                        trip.get("verified_duration_minutes", pd.NA),
                        trip.get("verified_points", pd.NA),
                    )
                )
                else pd.NA,
                "sampling_quality": _sampling_quality(
                    _average_sampling_interval_seconds(
                        trip.get("verified_duration_minutes", pd.NA),
                        trip.get("verified_points", pd.NA),
                    ),
                    trip.get("verified_points", pd.NA),
                    rules,
                ),
                "candidate_count": len(matches),
                "strict_candidate_count": int(
                    (matches["candidate_class"] == STRICT_CLASS).sum()
                ),
                "non_strict_candidate_count": int(
                    (matches["candidate_class"] == NON_STRICT_CLASS).sum()
                ),
                "coverage_status": (
                    "HAS_CANDIDATE"
                    if len(matches) > 0
                    else "NO_OPPOSITE_VEHICLE_SAME_DIRECTION"
                ),
            }
        )
    coverage = pd.DataFrame(coverage_rows).sort_values(
        ["coverage_status", "route_direction", "vehicle_type", "file_name"],
        kind="stable",
    )

    return candidates, route_summary, coverage.reset_index(drop=True)


def _trip_metadata_log(ready_trips, rules):
    """Create one readable metadata row for every verified trip."""

    rows = []
    for _, trip in ready_trips.iterrows():
        points = trip.get("verified_points", pd.NA)
        duration = trip.get("verified_duration_minutes", pd.NA)
        average_interval = _average_sampling_interval_seconds(
            duration, points
        )
        raw_distance = trip.get("raw_recalculated_distance_km", pd.NA)
        if pd.isna(raw_distance):
            raw_distance = trip.get("raw_index_distance_km", pd.NA)

        rows.append(
            {
                "verified_trip_id": trip["verified_trip_id"],
                "reconciliation_status": trip["reconciliation_status"],
                "vehicle_type": trip["vehicle_type"],
                "file_name": trip["source_file_name"],
                "source_relative_path": trip["source_relative_path"],
                "cleaned_file_path": trip["source_cleaned_file_path"],
                "logged_route_name": trip.get("log_route_name", pd.NA),
                "logged_start_location": trip["log_start_location"],
                "logged_end_location": trip["log_end_location"],
                "route_direction": trip["route_direction"],
                "raw_recording_start": trip.get(
                    "raw_first_timestamp", pd.NA
                ),
                "raw_recording_end": trip.get("raw_last_timestamp", pd.NA),
                "raw_recording_distance_km": (
                    round(float(raw_distance), 6)
                    if pd.notna(raw_distance)
                    else pd.NA
                ),
                "raw_recording_datapoints": trip.get(
                    "raw_index_points", pd.NA
                ),
                "verified_trip_start": trip["verified_start_timestamp"],
                "verified_trip_end": trip["verified_end_timestamp"],
                "verified_trip_duration_minutes": (
                    round(float(duration), 3)
                    if pd.notna(duration)
                    else pd.NA
                ),
                "verified_pre_crop_distance_km": (
                    round(float(trip["verified_distance_km"]), 6)
                    if pd.notna(trip.get("verified_distance_km", pd.NA))
                    else pd.NA
                ),
                "verified_datapoints": points,
                "average_sampling_interval_seconds": (
                    round(float(average_interval), 3)
                    if pd.notna(average_interval)
                    else pd.NA
                ),
                "sampling_quality": _sampling_quality(
                    average_interval, points, rules
                ),
            }
        )

    return pd.DataFrame(rows).sort_values(
        ["route_direction", "vehicle_type", "verified_trip_start"],
        kind="stable",
    ).reset_index(drop=True)


def _summary_table(index, ready, candidates, coverage):
    ev_count = int((ready["vehicle_type"] == "EV").sum())
    ice_count = int((ready["vehicle_type"] == "ICE").sum())
    strict_count = int(
        (candidates["candidate_class"] == STRICT_CLASS).sum()
    )
    non_strict_count = int(
        (candidates["candidate_class"] == NON_STRICT_CLASS).sum()
    )
    trips_without_candidates = int(
        (coverage["candidate_count"] == 0).sum()
    )
    candidates_ready_for_route_validation = int(
        (
            candidates["candidate_processing_status"]
            == "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION"
        ).sum()
    )
    candidates_requiring_quality_review = int(
        (
            candidates["candidate_processing_status"]
            != "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION"
        ).sum()
    )
    trips_requiring_sampling_review = int(
        (
            coverage["sampling_quality"]
            != "PRELIMINARILY_SUITABLE_FOR_PHASE_ANALYSIS"
        ).sum()
    )
    return pd.DataFrame(
        [
            {"Metric": "Generated at", "Value": datetime.now().isoformat(timespec="seconds")},
            {"Metric": "Verified-index rows", "Value": len(index)},
            {"Metric": "Trips ready for pairing", "Value": len(ready)},
            {"Metric": "Trips excluded by reconciliation", "Value": len(index) - len(ready)},
            {"Metric": "Ready EV trips", "Value": ev_count},
            {"Metric": "Ready ICE trips", "Value": ice_count},
            {"Metric": "Strict preliminary candidates", "Value": strict_count},
            {"Metric": "Non-strict preliminary candidates", "Value": non_strict_count},
            {"Metric": "Candidates ready for ordered GPS-route validation", "Value": candidates_ready_for_route_validation},
            {"Metric": "Candidates requiring telemetry/distance review", "Value": candidates_requiring_quality_review},
            {"Metric": "Trips requiring sampling-quality review", "Value": trips_requiring_sampling_review},
            {"Metric": "Trips with no same-direction opposite-vehicle candidate", "Value": trips_without_candidates},
            {
                "Metric": "Important interpretation",
                "Value": (
                    "CAND_001 etc. are preliminary candidates, not final "
                    "PAIR_001 files. Ordered GPS-route overlap, common-section "
                    "cropping, and pair QC are still pending."
                ),
            },
        ]
    )


def _rules_table(rules):
    return pd.DataFrame(
        [
            {"Rule": "Input", "Value": VERIFIED_INDEX_FILE_NAME, "Use": "Only reconciled trip records"},
            {"Rule": "Readiness", "Value": READY_STATUS, "Use": "Hard eligibility gate"},
            {"Rule": "Vehicle combination", "Value": "Exactly one EV and one ICE", "Use": "Hard candidate rule"},
            {"Rule": "Route direction", "Value": "Same canonical logged origin and destination", "Use": "Hard candidate rule; cross-route pairs are skipped"},
            {
                "Rule": "Strict preliminary candidate",
                "Value": (
                    "Same route/direction; same calendar day; start and end "
                    f"separations within {rules['start_location_tolerance_m']:.0f} m/"
                    f"{rules['end_location_tolerance_m']:.0f} m; start-clock "
                    "difference is used only for STRICT sub-category"
                ),
                "Use": "Preliminary class only; ordered GPS-route validation is still required",
            },
            {
                "Rule": "Non-strict preliminary candidate",
                "Value": (
                    "Same route/direction, but one or more strict date/time/"
                    "full-trip endpoint rules are not met"
                ),
                "Use": "Common GPS section must later be validated and cropped to common endpoints",
            },
            {
                "Rule": "End-time similarity",
                "Value": "Not used",
                "Use": "End time/TAT is an outcome and must not bias pairing",
            },
            {
                "Rule": "Pre-crop distance check",
                "Value": (
                    f"Review when endpoint-matched trips differ by more than "
                    f"{rules['pre_crop_distance_absolute_tolerance_km']:.1f} km "
                    f"and {rules['pre_crop_distance_relative_tolerance_fraction'] * 100:.0f}%"
                ),
                "Use": "Quality flag only; final comparable distance is calculated after common-section cropping",
            },
            {
                "Rule": "Sampling-quality screen",
                "Value": (
                    "Average verified sampling interval <= "
                    f"{rules['maximum_average_sampling_interval_seconds']:.0f} s"
                ),
                "Use": "Preliminary suitability check for later phase analysis",
            },
            {"Rule": "Pending GPS point tolerance (m)", "Value": rules["route_point_tolerance_m"], "Use": "Used by next route-validation module"},
            {"Rule": "Pending minimum ordered route overlap", "Value": rules["minimum_route_overlap_fraction"], "Use": "Used by next route-validation module"},
            {"Rule": "TAT/speed/phase outcomes", "Value": "Not read", "Use": "Prevents outcome leakage into pairing"},
            {"Rule": "Candidate numbering", "Value": "CAND_001, CAND_002, ...", "Use": "Simple identifiers for workbook review"},
            {"Rule": "Final three-digit PAIR number", "Value": "Not assigned at candidate stage", "Use": "PAIR_001 is assigned only after route/crop validation"},
            {"Rule": "Feature engineering", "Value": "Not performed here", "Use": "Runs after pair-specific common-route crop"},
        ]
    )


def _format_workbook(writer, tables):
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for sheet_name, dataframe in tables.items():
        worksheet = writer.sheets[sheet_name]
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

        for column_number, column_name in enumerate(
            dataframe.columns, start=1
        ):
            values = [column_name] + dataframe[column_name].head(250).tolist()
            maximum = max(
                len(str(value)) if value is not None else 0
                for value in values
            )
            worksheet.column_dimensions[
                get_column_letter(column_number)
            ].width = min(max(maximum + 2, 12), 44)
            for row_number in range(2, worksheet.max_row + 1):
                cell = worksheet.cell(row=row_number, column=column_number)
                cell.alignment = Alignment(
                    vertical="top", wrap_text=maximum > 35
                )
                if "start" in column_name.lower() and isinstance(
                    cell.value, datetime
                ):
                    cell.number_format = "dd-mm-yyyy hh:mm:ss"


def _atomic_save_workbook(output_path, tables):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        prefix=".pairing_candidates_",
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
            _format_workbook(writer, tables)
        os.replace(temporary_path, output_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def identify_candidate_pairs(config, index_path=None, output_path=None):
    index_path = Path(index_path or _verified_index_path(config))
    output_path = Path(output_path or _output_path(config))
    rules = _rules_from_config(config)

    index, ready, not_ready = load_verified_trips(index_path)
    candidates, route_summary, coverage = build_candidate_tables(ready, rules)
    summary = _summary_table(index, ready, candidates, coverage)
    tables = {
        "Run Summary": summary,
        "Candidate Pairs": candidates,
        "Trip Metadata Log": _trip_metadata_log(ready, rules),
        "Trip Coverage": coverage,
        "Route Summary": route_summary,
        "Reconciliation Excluded": not_ready,
        "Rules Applied": _rules_table(rules),
    }
    _atomic_save_workbook(output_path, tables)

    ev_count = int((ready["vehicle_type"] == "EV").sum())
    ice_count = int((ready["vehicle_type"] == "ICE").sum())
    return {
        "verified_index_rows": len(index),
        "ready_trips": len(ready),
        "not_ready_trips": len(not_ready),
        "ev_trips": ev_count,
        "ice_trips": ice_count,
        "strict_candidates": int(
            (candidates["candidate_class"] == STRICT_CLASS).sum()
        ),
        "non_strict_candidates": int(
            (candidates["candidate_class"] == NON_STRICT_CLASS).sum()
        ),
        "candidates_ready_for_route_validation": int(
            (
                candidates["candidate_processing_status"]
                == "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION"
            ).sum()
        ),
        "candidates_requiring_quality_review": int(
            (
                candidates["candidate_processing_status"]
                != "READY_FOR_ORDERED_GPS_ROUTE_VALIDATION"
            ).sum()
        ),
        "trips_requiring_sampling_review": int(
            (
                coverage["sampling_quality"]
                != "PRELIMINARILY_SUITABLE_FOR_PHASE_ANALYSIS"
            ).sum()
        ),
        "trips_without_candidates": int(
            (coverage["candidate_count"] == 0).sum()
        ),
        "strict_start_time_tolerance_minutes": rules[
            "strict_start_time_tolerance_minutes"
        ],
        "start_location_tolerance_m": rules[
            "start_location_tolerance_m"
        ],
        "end_location_tolerance_m": rules["end_location_tolerance_m"],
        "output_path": str(output_path),
    }


if __name__ == "__main__":
    try:
        settings = load_config()
        print("Candidate-pair identification started")
        result = identify_candidate_pairs(settings)
        print("\nCandidate-pair identification completed")
        print(f"Trips ready for pairing: {result['ready_trips']}")
        print("\nStrict preliminary-candidate rule:")
        print(
            "Same route/direction + same calendar day + start/end "
            "locations within "
            f"{result['start_location_tolerance_m']:.0f}/"
            f"{result['end_location_tolerance_m']:.0f} m; start-time "
            "similarity is used only for the STRICT sub-category"
        )
        print("\nNon-strict preliminary-candidate rule:")
        print(
            "Same route/direction, but one or more strict date/time/full-trip "
            "endpoint rules are not met; common GPS-section validation and "
            "cropping are required"
        )
        print(
            "\nStrict preliminary candidates: "
            f"{result['strict_candidates']}"
        )
        print(
            "Non-strict preliminary candidates: "
            f"{result['non_strict_candidates']}"
        )
        print(
            "Candidates ready for ordered GPS-route validation: "
            f"{result['candidates_ready_for_route_validation']}"
        )
        print(
            "Candidates requiring telemetry/distance review: "
            f"{result['candidates_requiring_quality_review']}"
        )
        print(
            "Trips requiring sampling-quality review: "
            f"{result['trips_requiring_sampling_review']}"
        )
        print(
            "Trips without a same-direction opposite-vehicle candidate: "
            f"{result['trips_without_candidates']}"
        )
        print(f"Candidate workbook: {result['output_path']}")
    except Exception as error:
        print(f"\nCANDIDATE PAIRING FAILED | {error}")
        raise SystemExit(1)