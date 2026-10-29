"""
TAT Stage 4 - Revised Phase Segmentation
========================================

Purpose
-------
Classify each telemetry interval into one of the seven project phases while
handling the actual data limitations:

1. The GPS device often does NOT record points while vehicle speed = 0.
2. Ignition status is unavailable.
3. Sampling frequency is irregular (commonly 5-30 s, sometimes much longer).
4. Manual trip logs contain stop timestamps and charging information.
5. Manual event times and GPS/device timestamps can be offset.
6. Mobile-app telemetry retains frequent zero-speed rows and therefore needs
   cumulative consecutive-interval stationary detection instead of relying on
   one large missing-record gap.

This script is standalone and INTERCITY-only. It does NOT depend on
config_loader or earlier phase-segmentation scripts.

---------------------------------------------------------------------------
FINAL CLASSIFICATION LOGIC
---------------------------------------------------------------------------

A. Route-specific stationary indicator (used only internally)
--------------------------------------------------------------
Device / missing-zero path - gap-equivalent speed:

    gap_equivalent_speed_kmh =
        Haversine Distance (m) / Time Difference (s) * 3.6

This is NOT the device Speed column.

It estimates the net movement rate across a telemetry gap. Because the ideal
stationary rule was approximately Speed < 1 km/h, it is used as stationary
evidence for short/fallback intervals when:

    gap_equivalent_speed_kmh < 1.0

For an INTERCITY sparse-device gap longer than 180 seconds, the established
project time-gap rule is applied even when implied speed is higher. A gap is
CHARGING/REFUELING only when it is matched to charging/refuelling evidence in
the manual trip log; otherwise it is STOP and the high implied speed is flagged
for review. This avoids erasing a real omitted stop merely because the points
immediately before and after it contain movement.

The implied-speed test also replaces the old fixed "Distance <= 150 m" test
for shorter gaps, which incorrectly classified moving intervals as IDLE.

INTERCITY recorded-zero episode support:

    recorded Speed <= 0.5 km/h
    for at least two consecutive positive-duration intervals

The complete episode duration is the sum of its consecutive time gaps.  This
path remains valid even when small GPS jitter makes gap-equivalent speed exceed
1 km/h.  An isolated zero-speed reading is not enough to create an episode.

B. Phase rules
--------------
Priority 1 - CHARGING / REFUELING
    Manual charging/refuelling evidence is applied first. The manual timestamp is
    treated as an approximate anchor and matched to the most plausible stationary
    telemetry episode within the existing tolerance. The complete matched episode
    is reserved as CHARGING/REFUELING before STOP/IDLE logic runs.

    Station/operator address text and manual charging location are matching/ranking
    aids only. They never create CHARGING/REFUELING without manual log evidence.

Priority 2 - STOP
    INTERCITY: qualifying stationary duration > 180 s -> STOP unless already
    reserved as CHARGING/REFUELING.


    Gap-equivalent speed remains an audit/confidence signal. The separator is
    manual charging/refuelling evidence: without it, stationary telemetry remains
    STOP/IDLE according to duration.

Priority 3 - IDLE
    INTERCITY: qualifying stationary duration <= 180 s
    CITY:      consecutive recorded-zero episode <= 30 s

Priority 4 - INCH
    1 <= device Speed <= 5 km/h

Priority 5 - ACC
    device Speed > 5 km/h
    AND Acceleration > +0.14 m/s^2

Priority 6 - DEC
    device Speed > 5 km/h
    AND Acceleration < -0.14 m/s^2

Priority 7 - CRUISE
    device Speed > 5 km/h
    AND -0.14 <= Acceleration <= +0.14 m/s^2

C. Manual stop start timestamps
-------------------------------
The maintained manual log records the START of each normal stationary event.
It does not usually contain a stop end time or duration.

For every start-only manual stop timestamp, the code goes to the matched
EV/ICE telemetry worksheet and infers the duration forward from that start:

    1. A telemetry row represents [Timestamp - Gap, Timestamp].
    2. An interval ending exactly at the manual timestamp is BEFORE the stop
       and is never selected as the stop interval.
    3. If the timestamp falls at the beginning of a consecutive zero-speed or
       stationary-like episode, use the full observed episode after the start.
    4. Otherwise, use the next telemetry interval whose time gap begins at or
       contains the manual start timestamp.
    5. If the exact row is unavailable because of device delay, search forward
       within the existing 15-minute manual-event tolerance.
    6. Charging/refuelling rows are reserved first and can never be reused as
       normal STOP/IDLE rows.
    7. One inferred telemetry episode can be matched to only one manual stop.

The complete inferred telemetry duration uses the route boundary:
    CITY:      Duration > 30 s  -> STOP; otherwise IDLE
    INTERCITY: Duration > 180 s -> STOP; otherwise IDLE

Explicit start-end stop intervals remain supported. Automatic device-gap and
mobile zero-speed rules still run for all stops missing from the manual log.

D. Manual charging/refuelling
-----------------------------
Primary manual trip log:
    MANUAL_TRIP_LOG_PATH below (currently D:/My Desktop/TAT Data Collection Intercity.xlsx)

The path can also be supplied through the TAT_MANUAL_TRIP_LOG environment
variable when the workbook is intentionally moved. The script does not use a
different project copy when the authoritative file is missing.

The manual record is matched using:
    vehicle type + source trip date + route direction + closest trip start time

Source trip dates / route are taken preferentially from:
    03_Pair_Crop_Info

Charging/refuelling is matched BEFORE normal stop events.

Authoritative maintained columns:
    Charging time (min) -> charging/refuelling duration only
    charging timestamp  -> charging/refuelling start time only
    stop timestamps     -> normal STOP/IDLE timestamps only

Numbers in Additional or stop timestamps are never used as charging duration.
When stop timestamps is populated, Additional is not parsed as a second stop
source.  Older Additional-only records remain supported as a fallback.

Manual charging matching:
    - charging timestamp is an approximate event anchor, not an exact telemetry key
    - Charging time (min) is the target event duration when it is available
    - telemetry around the anchor is searched for a contiguous stationary/very-low-speed
      region; INCH-like manoeuvring (<=5 km/h) may belong to that region
    - when the region is longer than the manual duration, the contiguous sub-window whose
      observed telemetry duration best matches the manual target is reserved
    - clearly moving rows are never pulled into CHARGING merely to force the target duration
    - if manual duration is unavailable, the complete matched low-movement episode is used
    - if manual location is available, matching address is used only as supporting evidence
    - if no charging timestamp exists but a manual charging location exists, that
      location can identify the low-movement episode
    - if the manual log contains a positive Charging time (min) but neither timestamp
      nor location, the duration itself remains authoritative manual charging evidence:
      the complete trip is searched for the safest low-movement window matching that
      duration; charging-station address evidence and separation from logged normal-stop
      timestamps are used as ranking evidence
    - sparse-device long gaps remain eligible because a stopped vehicle may emit no points

The manual charging duration is stored separately in the audit/details sheet and used as a
target for selecting telemetry rows. It is NOT written into or substituted for the original
Time Difference column.

E. Address station evidence
---------------------------
The address evidence uses BOTH ends of the gap interval:
    previous row Address
    current row Address
    previous/current Reverse Geocoded Address (if present)

This evidence is used only to rank/confirm a manually recorded charging or
refuelling event. Address evidence alone never creates a replenishment phase.

F. Boundary / data-quality scenarios
------------------------------------
A very long stationary-looking gap at the beginning or end of a cropped trip
can be caused by crop/device boundaries.

Boundary long gaps (>180 s):
    - if supported by manual stop, manual charging, or station evidence:
          classify, but flag as BOUNDARY_GAP_SUPPORTED
    - if unsupported:
          leave phase blank and flag BOUNDARY_DATA_GAP_UNCONFIRMED

Irregular moving gaps (>30 s with gap-equivalent speed >=1 km/h):
    - classify using the normal device Speed/Acceleration rules
    - flag as IRREGULAR_MOVING_GAP because acceleration over a long interval
      has lower confidence

G. Exhaustive fallback for short / inconsistent telemetry
------------------------------------------------------------
TAT computation requires every positive-duration interval to belong to one of
the seven phases. In the real data there are occasional rows where the device
reports Speed = 0 even though GPS displacement proves movement, and there are
also consecutive short stationary-like rows.

To keep the seven phases exhaustive without corrupting STOP/IDLE:

1. Consecutive stationary-like rows that remain after manual events, device
   gaps, and recorded-zero episodes are grouped using the same route boundary:
       INTERCITY >180 s -> STOP
       otherwise                    -> IDLE

   A stationary-like episode is NEVER classified as INCH.
   INCH is reserved for low-speed movement only.

2. INCH is a MOVING phase:
       primary rule: 1 <= recorded device Speed <= 5 km/h.

   When the device Speed is zero/missing but GPS displacement proves that the
   interval was moving, Gap Equivalent Speed may be used only as a low-
   confidence proxy because the speed field is inconsistent for those rows:
       1 <= proxy speed <=5 -> INCH
       proxy speed >5       -> ACC/DEC/CRUISE using existing acceleration.

   This proxy is never applied to a stationary-like interval.

3. If acceleration is unavailable for a clearly moving fallback interval,
   classify it as CRUISE with LOW confidence rather than leaving phase blank.

4. Any extremely rare positive-duration row still unresolved is assigned INCH
   with VERY_LOW confidence and is retained in the audit.

This fallback is deliberately applied only AFTER the primary rules.  The
moving-phase thresholds and classification order are unchanged.

---------------------------------------------------------------------------
WORKBOOK OUTPUT
---------------------------------------------------------------------------

Existing data sheets:
    01_EV_Cropped_Data
    02_ICE_Cropped_Data

Only the `phase` column is added/updated in these two sheets.

Created/replaced sheets:
    06_Phase_Summary
    07_Classification_Audit
    08_Manual_Event_Matching
    charging time details

Old downstream TAT/visualization sheets are NOT recalculated by this script.
After Stage 4 finishes successfully, rerun tat_computation.py and the
visualization stage to avoid looking at stale phase/TAT results.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import datetime, time as datetime_time
from difflib import SequenceMatcher
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Iterable
import math
import os
import re
import time

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================================
# PROJECT PATHS
# ============================================================================

PAIRED_ROOT = Path(r"D:\Hameeda_Projects\TAT_MS_INTERCITY\Data\Results\Paired")

# The maintained manual log is authoritative for manual event evidence. An
# environment-variable override is available for an intentional future move,
# but the script never silently substitutes another workbook.
MANUAL_TRIP_LOG_PATH = Path(
    os.environ.get(
        "TAT_MANUAL_TRIP_LOG",
        r"D:\TAT_Project\TAT Data Collection Intercity.xlsx",
    )
)

SYSTEM_OUTPUT_DIR = Path(r"D:\Hameeda_Projects\TAT_MS_INTERCITY\Data\Results\System")

# Optional safety filter.
# Leave as None to process all pair workbooks.
# For a first test, set:
#     PAIR_NAME_FILTER = "PAIR_124"
PAIR_NAME_FILTER: str | None = None

EV_SHEET = "01_EV_Cropped_Data"
ICE_SHEET = "02_ICE_Cropped_Data"
PAIR_INFO_SHEET = "03_Pair_Crop_Info"

PHASE_SUMMARY_SHEET = "06_Phase_Summary"
CLASSIFICATION_AUDIT_SHEET = "07_Classification_Audit"
MANUAL_EVENT_AUDIT_SHEET = "08_Manual_Event_Matching"
CHARGING_DETAILS_SHEET = "charging time details"

GLOBAL_LOG_NAME = "phase_segmentation_revised_log.xlsx"
SCRIPT_VERSION = f"{datetime.now():%Y-%m-%d}_INTERCITY_MANUAL_FIRST_HYBRID_V8_DURATION_ONLY_CHARGING"


# ============================================================================
# COLUMN NAMES
# ============================================================================

TIMESTAMP = "Timestamp"
TIME_GAP = "Time Difference (s)"
DISTANCE = "Haversine Distance (m)"
SPEED_KMH = "Speed (km/h)"
ACCELERATION = "Acceleration (m/s^2)"
ADDRESS = "Address"
REVERSE_ADDRESS = "Reverse Geocoded Address"
PHASE = "phase"


# ============================================================================
# THRESHOLDS
# ============================================================================

# Original ideal stationary rule was approximately speed < 1 km/h.
STATIONARY_GAP_SPEED_KMH = 1.0

# Mobile-app telemetry records true stationary rows instead of omitting them.
# A small tolerance is used because exported speed values may contain minor
# floating-point noise around zero.  Isolated zero-speed values are not enough
# to create a stationary episode; at least two consecutive positive-duration
# intervals are required.
RECORDED_ZERO_SPEED_MAX_KMH = 0.5
MIN_ZERO_SPEED_EPISODE_INTERVALS = 2

# INTERCITY IDLE / STOP duration threshold.  This is the previously approved
# project rule and remains unchanged for INTERCITY routes.
IDLE_MIN_SECONDS = 30.0
STOP_THRESHOLD_SECONDS = 180.0


# Moving rules.
INCH_MIN_KMH = 1.0
INCH_MAX_KMH = 5.0
ACCELERATION_THRESHOLD = 0.14

# Manual vs device timestamp lag tolerance.
MANUAL_EVENT_TOLERANCE_SECONDS = 15.0 * 60.0

# A manual clock is normally recorded only to the minute, whereas telemetry
# contains seconds.  Treat a forward interval beginning within two minutes as
# the timestamp-aligned interval before considering a later stationary event.
MANUAL_STOP_START_ALIGNMENT_SECONDS = 2.0 * 60.0

# Gaps longer than this are explicitly flagged for review.
VERY_LONG_GAP_SECONDS = 15.0 * 60.0

# Boundary rows considered susceptible to crop/device-start/end effects.
BOUNDARY_ROW_COUNT = 2


# ============================================================================
# ADDRESS KEYWORDS
# ============================================================================

EV_STATION_KEYWORDS = (
    "charging station",
    "ev charging",
    "ev charger",
    "electric vehicle charging",
    "charger",
    "chargezone",
    "charge zone",
    "statiq",
    "zeon",
    "tata power ez charge",
    "shell recharge",
    "relux electric",
    # Confirmed project charging zones appearing in the maintained manual
    # trip log.  These are deliberately specific place/operator labels; the
    # generic word "krishnagiri" is not used because an ordinary halt
    # elsewhere in the city must not become CHARGING.
    "krishnagiri bypass",
    "azad india",
)

ICE_STATION_KEYWORDS = (
    "petrol pump",
    "petrol bunk",
    "petrol station",
    "fuel station",
    "fuel pump",
    "refuelling station",
    "refueling station",
    "diesel station",
    "indian oil",
    "iocl",
    "bharat petroleum",
    "bpcl",
    "hindustan petroleum",
    "hpcl",
    "shell petrol",
)


# ============================================================================
# ROUTE / PLACE ALIASES
# ============================================================================

PLACE_ALIASES = {
    "cmbt": {
        "cmbt",
        "cbmt",
        "cmbt bus stand",
        "cbmt bus stand",
        "koyambedu",
    },
    "navalur": {"navalur"},
    "tnagar": {
        "t nagar",
        "tnagar",
        "t-nagar",
        "thyagaraya nagar",
    },
    "thiruporur": {
        "thiruporur",
        "thirupporur",
        "tiruporur",
        "thirupur",
    },
    "sholinganallur": {
        "sholinganallur",
        "shollinganallur",
    },
    "kcbt": {"kcbt"},
    "hebbal": {"hebbal", "hebal"},
    "silkboard": {"silk board", "silkboard"},
    "bellandur": {"bellandur"},
    "kia": {
        "kia",
        "kempegowda international airport",
        "kempegowda airport",
    },
    "kadugodi": {
        "kadugodi",
        "kadugodi bus stand",
    },
    "whitefield": {"whitefield"},
    "chennai": {"chennai"},
    "pondicherry": {
        "pondicherry",
        "pondi",
        "pondy",
        "puducherry",
    },
    "maduravoyal": {
        "maduravoyal",
        "madhuravoyal",
        "maduravayol",
        "maduravoyil",
    },
    "avrroundana": {
        "avr roundana",
        "avr roundana salem",
        "avr",
    },
    "salem": {"salem"},
    "krishnagiri": {
        "krishnagiri",
        "krishnagiri bypass",
    },
}


# Route-direction pairs are unordered: both directions receive the same
# CITY/INTERCITY category.  These definitions mirror the route groups already
# used by the project's master-summary stage.
CITY_ROUTE_PAIRS = {
    frozenset(("tnagar", "thiruporur")),
    frozenset(("cmbt", "navalur")),
    frozenset(("kcbt", "sholinganallur")),
    frozenset(("hebbal", "silkboard")),
    frozenset(("bellandur", "kia")),
    frozenset(("kadugodi", "silkboard")),
    frozenset(("whitefield", "silkboard")),
}

INTERCITY_ROUTE_PAIRS = {
    frozenset(("chennai", "pondicherry")),
    frozenset(("maduravoyal", "silkboard")),
    frozenset(("avrroundana", "silkboard")),
    frozenset(("salem", "silkboard")),
}


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class PairVehicleContext:
    vehicle: str
    source_date: pd.Timestamp | None
    source_trip_start: pd.Timestamp | None
    source_trip_end: pd.Timestamp | None
    route_origin: str
    route_destination: str
    source_file_name: str
    # Default preserves the previous 180-second behaviour for callers/tests
    # that construct PairVehicleContext directly without route metadata.
    route_category: str = "INTERCITY"


@dataclass
class ManualEvent:
    vehicle: str
    event_kind: str
    event_index: int
    raw_text: str
    start: pd.Timestamp | None
    end: pd.Timestamp | None
    duration_seconds: float | None
    location: str
    source_column: str
    is_interval: bool


@dataclass
class EventMatch:
    vehicle: str
    event_kind: str
    event_index: int
    manual_start: pd.Timestamp | None
    manual_end: pd.Timestamp | None
    manual_duration_seconds: float | None
    manual_location: str
    raw_text: str
    match_status: str
    match_method: str
    telemetry_row_index: int | None
    excel_row: int | None
    interval_start: pd.Timestamp | None
    interval_end: pd.Timestamp | None
    telemetry_gap_seconds: float | None
    telemetry_distance_m: float | None
    gap_equivalent_speed_kmh: float | None
    time_offset_seconds: float | None
    boundary_flag: bool
    address_station_evidence: bool
    notes: str


# ============================================================================
# BASIC HELPERS
# ============================================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return str(value).strip()


def parse_datetime_value(value: Any) -> pd.Timestamp | None:
    """Parse one datetime without swapping month/day in ISO-style dates."""
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (pd.Timestamp, datetime)):
        timestamp = pd.Timestamp(value)
        return None if pd.isna(timestamp) else timestamp

    if isinstance(value, (int, float, np.integer, np.floating)):
        number = float(value)
        if math.isfinite(number) and 20_000 <= number <= 80_000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(number, unit="D")

    raw = clean_text(value)
    if not raw:
        return None

    iso_style = bool(re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:[ T]|$)", raw))

    try:
        parsed = pd.to_datetime(
            raw,
            errors="coerce",
            yearfirst=iso_style,
            dayfirst=not iso_style,
            format="mixed",
        )
    except (TypeError, ValueError):
        parsed = pd.to_datetime(
            raw,
            errors="coerce",
            yearfirst=iso_style,
            dayfirst=not iso_style,
        )

    return None if pd.isna(parsed) else pd.Timestamp(parsed)


def parse_datetime_series(series: pd.Series) -> pd.Series:
    """Parse telemetry timestamps with ISO-safe and DD-MM-YYYY-safe rules."""
    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return pd.to_datetime(series, errors="coerce")

    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    numeric = pd.to_numeric(series, errors="coerce")
    excel_serial = numeric.between(20_000, 80_000, inclusive="both")

    if bool(excel_serial.any()):
        result.loc[excel_serial] = pd.to_datetime(
            numeric.loc[excel_serial],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    raw = series.astype("string").str.strip()
    unresolved = result.isna() & series.notna() & raw.ne("")
    iso_style = unresolved & raw.str.match(
        r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:[ T]|$)",
        na=False,
    )

    if bool(iso_style.any()):
        try:
            parsed_iso = pd.to_datetime(
                series.loc[iso_style],
                errors="coerce",
                yearfirst=True,
                dayfirst=False,
                format="mixed",
            )
        except (TypeError, ValueError):
            parsed_iso = pd.to_datetime(
                series.loc[iso_style],
                errors="coerce",
                yearfirst=True,
                dayfirst=False,
            )
        result.loc[iso_style] = parsed_iso

    unresolved = result.isna() & series.notna() & raw.ne("")
    if bool(unresolved.any()):
        try:
            parsed_dayfirst = pd.to_datetime(
                series.loc[unresolved],
                errors="coerce",
                dayfirst=True,
                format="mixed",
            )
        except (TypeError, ValueError):
            parsed_dayfirst = pd.to_datetime(
                series.loc[unresolved],
                errors="coerce",
                dayfirst=True,
            )
        result.loc[unresolved] = parsed_dayfirst

    unresolved = result.isna() & series.notna() & raw.ne("")
    if bool(unresolved.any()):
        try:
            parsed_fallback = pd.to_datetime(
                series.loc[unresolved],
                errors="coerce",
                dayfirst=False,
                format="mixed",
            )
        except (TypeError, ValueError):
            parsed_fallback = pd.to_datetime(
                series.loc[unresolved],
                errors="coerce",
                dayfirst=False,
            )
        result.loc[unresolved] = parsed_fallback

    return result


def concat_frames_without_futurewarning(
    frames: Iterable[pd.DataFrame | None],
) -> pd.DataFrame:
    """Concatenate audit frames without pandas empty/all-NA concat warnings."""
    usable = [
        frame.copy()
        for frame in frames
        if frame is not None and not frame.empty
    ]

    if not usable:
        return pd.DataFrame()

    columns: list[str] = []
    for frame in usable:
        for column in frame.columns:
            if column not in columns:
                columns.append(column)

    trimmed = [frame.dropna(axis=1, how="all") for frame in usable]
    combined = pd.concat(trimmed, ignore_index=True, sort=False)
    return combined.reindex(columns=columns)


def safe_timedelta(
    *,
    days: float = 0.0,
    hours: float = 0.0,
    minutes: float = 0.0,
    seconds: float = 0.0,
) -> pd.Timedelta:
    """Return a unit-explicit pandas Timedelta.

    NumPy has deprecated its old generic timedelta unit.  Building every
    duration as an explicit number of seconds avoids that warning and keeps
    the arithmetic unchanged.
    """

    total_seconds = (
        float(days) * 86400.0
        + float(hours) * 3600.0
        + float(minutes) * 60.0
        + float(seconds)
    )

    return pd.to_timedelta(
        total_seconds,
        unit="s",
    )


def normalise_text(value: Any) -> str:
    text = clean_text(value).casefold()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalised_column_lookup(frame: pd.DataFrame) -> dict[str, str]:
    return {
        normalise_text(column): str(column)
        for column in frame.columns
    }


def first_existing_column(
    frame: pd.DataFrame,
    candidates: Iterable[str],
) -> str | None:
    lookup = normalised_column_lookup(frame)

    for candidate in candidates:
        actual = lookup.get(
            normalise_text(candidate)
        )
        if actual is not None:
            return actual

    return None


def is_manual_placeholder(value: Any) -> bool:
    """
    Return True only for genuinely empty/manual placeholder cells.

    Numeric values, including 20, are deliberately not placeholders.  In the
    maintained trip log the dedicated ``Charging time (min)`` column is an
    authoritative numeric duration field, so a value such as 20 must remain a
    valid 20-minute charging duration.
    """
    return clean_text(value).casefold() in {
        "",
        "-",
        "--",
        "—",
        "na",
        "n/a",
        "none",
        "nan",
    }


def vehicle_type(value: Any) -> str:
    text = normalise_text(value)

    if text == "ev" or "electric" in text:
        return "EV"

    if (
        "diesel" in text
        or "petrol" in text
        or text == "ice"
        or "internal combustion" in text
    ):
        return "ICE"

    return ""


def parse_date(value: Any) -> pd.Timestamp | None:
    parsed = parse_datetime_value(value)
    return None if parsed is None else parsed.normalize()


def excel_serial_to_timestamp(value: Any) -> pd.Timestamp | None:
    """
    Pair_Crop_Info stores many datetimes as Excel serial numbers.
    """
    if value is None:
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value)

    try:
        number = float(value)
    except (TypeError, ValueError):
        return parse_datetime_value(value)

    if not math.isfinite(number):
        return None

    # Excel's standard serial-date origin.
    return pd.Timestamp("1899-12-30") + pd.to_timedelta(
        number,
        unit="D",
    )


def parse_clock(value: Any) -> tuple[int, int, int, str] | None:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (datetime, pd.Timestamp)):
        timestamp = pd.Timestamp(value)
        return (
            timestamp.hour,
            timestamp.minute,
            timestamp.second,
            "",
        )

    if isinstance(value, datetime_time):
        return (
            value.hour,
            value.minute,
            value.second,
            "",
        )

    match = re.search(
        r"(?i)(?<!\d)"
        r"(\d{1,2}):(\d{2})"
        r"(?::(\d{2}))?"
        r"\s*(am|pm)?"
        r"(?!\d)",
        clean_text(value),
    )

    if match is None:
        return None

    hour = int(match.group(1))
    minute = int(match.group(2))
    second = int(match.group(3) or 0)
    ampm = (match.group(4) or "").casefold()

    if minute > 59 or second > 59:
        return None

    if ampm:
        if hour < 1 or hour > 12:
            return None

        if ampm == "pm" and hour != 12:
            hour += 12
        elif ampm == "am" and hour == 12:
            hour = 0

    elif hour > 23:
        return None

    return (
        hour,
        minute,
        second,
        ampm,
    )


def parse_positive_minutes(value: Any) -> float | None:
    """Parse a positive duration expressed in minutes.

    The updated manual workbook mixes numeric Excel cells (for example 29)
    with text such as ``30 min``. A clock such as ``08:34`` must not be
    mistaken for an 8-minute duration.
    """
    if is_manual_placeholder(value):
        return None

    if isinstance(value, (int, float, np.integer, np.floating)):
        number = float(value)
        return number if math.isfinite(number) and number > 0 else None

    text = clean_text(value).replace(",", "")

    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
        number = float(text)
        return number if math.isfinite(number) and number > 0 else None

    match = re.search(
        r"(?i)(\d+(?:\.\d+)?)\s*(?:min|mins|minute|minutes)\b",
        text,
    )

    if match is None:
        return None

    number = float(match.group(1))
    return number if math.isfinite(number) and number > 0 else None


def canonical_place(value: Any) -> str:
    text = normalise_text(value)

    if not text:
        return ""

    for canonical, aliases in PLACE_ALIASES.items():
        values = {
            normalise_text(canonical),
            *(normalise_text(alias) for alias in aliases),
        }

        for alias in values:
            if (
                text == alias
                or alias in text
                or text in alias
            ):
                return canonical

    return text.replace(" ", "")


def text_similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0

    if left == right:
        return 1.0

    return SequenceMatcher(
        None,
        left,
        right,
    ).ratio()


def format_duration(seconds: float) -> str:
    seconds = max(0, int(round(seconds)))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


# ============================================================================
# PAIR WORKBOOK METADATA
# ============================================================================

def read_pair_info(
    workbook_path: Path,
) -> dict[str, Any]:
    try:
        frame = pd.read_excel(
            workbook_path,
            sheet_name=PAIR_INFO_SHEET,
            engine="openpyxl",
        )
    except ValueError:
        return {}

    if frame.empty or len(frame.columns) < 2:
        return {}

    parameter_column = frame.columns[0]
    value_column = frame.columns[1]

    return {
        clean_text(parameter): value
        for parameter, value in zip(
            frame[parameter_column],
            frame[value_column],
        )
        if clean_text(parameter)
    }


def parse_route_direction(
    value: Any,
) -> tuple[str, str]:
    text = clean_text(value)

    match = re.search(
        r"(?i)^\s*(.+?)\s+(?:to|→|-|–)\s+(.+?)\s*$",
        text,
    )

    if match is None:
        return "", ""

    return (
        canonical_place(match.group(1)),
        canonical_place(match.group(2)),
    )


def route_category_from_context(
    route_origin: str,
    route_destination: str,
    pair_info: dict[str, Any],
    workbook_path: Path,
    source_file_name: str,
) -> str:
    """Return INTERCITY for this dedicated intercity-only pipeline.

    Route direction and filename metadata are retained for audit/context,
    but this Stage 4 execution must never switch to CITY rules.
    """
    return "INTERCITY"


def applied_stop_threshold_seconds(
    route_category: str,
) -> float:
    """Return the approved INTERCITY stationary duration boundary."""
    return STOP_THRESHOLD_SECONDS


def pair_vehicle_context(
    pair_info: dict[str, Any],
    vehicle: str,
    fallback_frame: pd.DataFrame,
    workbook_path: Path,
) -> PairVehicleContext:
    prefix = "ev" if vehicle == "EV" else "ice"

    route_origin, route_destination = parse_route_direction(
        pair_info.get("route_direction", "")
    )

    if not route_origin or not route_destination:
        match = re.search(
            r"(?i)(?:STRICT|NONSTRICT)_(.+?)_to_(.+)$",
            workbook_path.stem,
        )
        if match:
            route_origin = canonical_place(
                match.group(1)
            )
            route_destination = canonical_place(
                match.group(2)
            )

    source_date = parse_date(
        pair_info.get(
            f"{prefix}_input_date_folder"
        )
    )

    source_trip_start = excel_serial_to_timestamp(
        pair_info.get(
            f"{prefix}_verified_start"
        )
    )

    source_trip_end = excel_serial_to_timestamp(
        pair_info.get(
            f"{prefix}_verified_end"
        )
    )

    fallback_times = parse_datetime_series(
        fallback_frame[TIMESTAMP]
    ).dropna()

    if source_trip_start is None and not fallback_times.empty:
        source_trip_start = pd.Timestamp(
            fallback_times.iloc[0]
        )

    if source_trip_end is None and not fallback_times.empty:
        source_trip_end = pd.Timestamp(
            fallback_times.iloc[-1]
        )

    if source_date is None and source_trip_start is not None:
        source_date = source_trip_start.normalize()

    source_file_name = clean_text(
        pair_info.get(
            f"{prefix}_file_name",
            "",
        )
    )

    route_category = route_category_from_context(
        route_origin=route_origin,
        route_destination=route_destination,
        pair_info=pair_info,
        workbook_path=workbook_path,
        source_file_name=source_file_name,
    )

    return PairVehicleContext(
        vehicle=vehicle,
        source_date=source_date,
        source_trip_start=source_trip_start,
        source_trip_end=source_trip_end,
        route_origin=route_origin,
        route_destination=route_destination,
        source_file_name=source_file_name,
        route_category=route_category,
    )


# ============================================================================
# MANUAL TRIP LOG LOADING / MATCHING
# ============================================================================

def resolve_manual_trip_log_path(
    configured_path: Path,
) -> Path:
    """Require the explicitly configured authoritative manual workbook."""
    path = configured_path.expanduser()

    if path.is_file():
        return path

    raise FileNotFoundError(
        "Authoritative manual trip log was not found at:\n"
        f"  {path}\n"
        "Keep the updated workbook at this exact path, or intentionally set "
        "the TAT_MANUAL_TRIP_LOG environment variable to its new full path."
    )


def load_manual_trip_log(
    path: Path,
) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(
            f"Manual trip log not found: {path}"
        )

    with pd.ExcelFile(
        path,
        engine="openpyxl",
    ) as excel:
        sheet_name = (
            "Sheet1"
            if "Sheet1" in excel.sheet_names
            else excel.sheet_names[0]
        )

        raw = pd.read_excel(
            excel,
            sheet_name=sheet_name,
        )

    aliases = {
        "Date": ["Date", "Trip Date"],
        "Vehicle Type": [
            "Vehicle Type",
            "Vehicle",
        ],
        "Route Name": [
            "Route Name",
            "Route",
        ],
        "Start Location": [
            "Start Location",
            "Origin",
        ],
        "End Location": [
            "End Location",
            "Destination",
        ],
        "Start Time": [
            "Start time",
            "Start Time",
        ],
        "End Time": [
            "End Time",
            "End time",
        ],
        "Charging Time": [
            "Charging time (min)",
            "Charging Time (min)",
            "Charging time",
            "Charging Time",
            "Refuelling time (min)",
            "Refueling time (min)",
            "Refuelling time",
            "Refueling time",
        ],
        "Charging Timestamp": [
            "charging timestamp",
            "Charging timestamp",
            "Charging Timestamp",
            "charging time stamp",
            "Charging time stamp",
        ],
        # Optional dedicated stop-time column(s).  If the current workbook
        # does not contain one, the existing Additional-column parser below
        # is still used exactly as before.
        "Stop Timestamps": [
            "Stationary timestamps",
            "Stationary timestamp",
            "Stationary times",
            "Stationary time",
            "Stop timestamps",
            "Stop timestamp",
            "stop timestamps",
            "stop timestamp",
            "Stop times",
            "Stop time",
            "stop times",
            "stop time",
            "Intermediate stop timestamps",
            "Intermediate timestamps",
            "Stop Record",
            "Stops",
        ],
        "Charging Type": [
            "Charging type",
            "Charging Type",
        ],
        "Additional": [
            "Additional",
            "Notes",
            "Remarks",
        ],
    }

    frame = pd.DataFrame(
        index=raw.index
    )

    for target, candidates in aliases.items():
        actual = first_existing_column(
            raw,
            candidates,
        )

        frame[target] = (
            raw[actual]
            if actual is not None
            else pd.Series(
                pd.NA,
                index=raw.index,
            )
        )

    frame["_vehicle"] = frame[
        "Vehicle Type"
    ].map(vehicle_type)

    frame["_date"] = frame[
        "Date"
    ].map(parse_date)

    frame["_origin"] = frame[
        "Start Location"
    ].map(canonical_place)

    frame["_destination"] = frame[
        "End Location"
    ].map(canonical_place)

    frame["_source_excel_row"] = (
        frame.index
        + 2
    )

    usable = frame[
        frame["_vehicle"].isin(
            {"EV", "ICE"}
        )
        & frame["_date"].notna()
    ].copy()

    return usable.reset_index(
        drop=True
    )


def resolve_manual_start_time(
    record: pd.Series,
    context: PairVehicleContext,
) -> pd.Timestamp | None:
    date = record.get("_date")

    if date is None or pd.isna(date):
        return None

    clock = parse_clock(
        record.get(
            "Start Time"
        )
    )

    if clock is None:
        return None

    hour, minute, second, ampm = clock

    base = (
        pd.Timestamp(date)
        + safe_timedelta(
            hours=hour,
            minutes=minute,
            seconds=second,
        )
    )

    if ampm:
        return base

    candidates = [
        base,
        base + safe_timedelta(hours=12),
        base - safe_timedelta(hours=12),
        base + safe_timedelta(days=1),
    ]

    if context.source_trip_start is not None:
        return min(
            candidates,
            key=lambda candidate: abs(
                (
                    candidate
                    - context.source_trip_start
                ).total_seconds()
            ),
        )

    return base


def resolve_manual_end_time(
    record: pd.Series,
    manual_start: pd.Timestamp | None,
) -> pd.Timestamp | None:
    date = record.get("_date")

    if date is None or pd.isna(date):
        return None

    raw_end = record.get(
        "End Time"
    )

    if is_manual_placeholder(
        raw_end
    ):
        return None

    # Sometimes the end cell includes a date in parentheses.
    full_parsed = parse_datetime_value(
        clean_text(raw_end)
    )

    if full_parsed is not None:
        parsed_timestamp = pd.Timestamp(
            full_parsed
        )

        # A time-only value parsed by pandas may use today's date;
        # only accept a parsed explicit date when the string appears to
        # actually contain a date separator/year.
        if re.search(
            r"\d{1,2}[-/]\d{1,2}",
            clean_text(raw_end),
        ):
            return parsed_timestamp

    clock = parse_clock(
        raw_end
    )

    if clock is None:
        return None

    hour, minute, second, ampm = clock

    base = (
        pd.Timestamp(date)
        + safe_timedelta(
            hours=hour,
            minutes=minute,
            seconds=second,
        )
    )

    if ampm:
        candidates = [
            base,
            base + safe_timedelta(days=1),
        ]
    else:
        candidates = [
            base,
            base + safe_timedelta(hours=12),
            base + safe_timedelta(days=1),
            base + safe_timedelta(days=1, hours=12),
        ]

    if manual_start is None:
        return base

    valid = [
        candidate
        for candidate in candidates
        if candidate >= manual_start
        and (
            candidate
            - manual_start
        ).total_seconds()
        <= 24 * 3600
    ]

    if not valid:
        return None

    # Prefer a plausible shortest forward trip duration.
    return min(
        valid,
        key=lambda candidate: (
            candidate
            - manual_start
        ).total_seconds(),
    )


def match_manual_record(
    manual_log: pd.DataFrame,
    context: PairVehicleContext,
) -> pd.Series | None:
    candidates = manual_log[
        manual_log["_vehicle"]
        == context.vehicle
    ].copy()

    if candidates.empty:
        return None

    # Date is the strongest signal.
    if context.source_date is not None:
        same_date = candidates[
            candidates["_date"].map(
                lambda value: (
                    pd.Timestamp(value).date()
                    == context.source_date.date()
                    if value is not None
                    and pd.notna(value)
                    else False
                )
            )
        ].copy()

        # Never apply a manual event from another calendar date to the
        # current telemetry trip. A missing date match is safer than a false
        # CHARGING/STOP/IDLE classification.
        if same_date.empty:
            return None

        candidates = same_date

    scored: list[
        tuple[
            float,
            float,
            float,
            int,
        ]
    ] = []

    for index, record in candidates.iterrows():
        manual_origin = canonical_place(
            record.get(
                "Start Location"
            )
        )

        manual_destination = canonical_place(
            record.get(
                "End Location"
            )
        )

        route_score = (
            text_similarity(
                context.route_origin,
                manual_origin,
            )
            + text_similarity(
                context.route_destination,
                manual_destination,
            )
        )

        manual_start = resolve_manual_start_time(
            record,
            context,
        )

        start_difference = float(
            "inf"
        )

        if (
            manual_start is not None
            and context.source_trip_start is not None
        ):
            start_difference = abs(
                (
                    manual_start
                    - context.source_trip_start
                ).total_seconds()
            )

        date_bonus = 0.0

        if (
            context.source_date is not None
            and record.get("_date") is not None
            and pd.notna(
                record.get("_date")
            )
            and pd.Timestamp(
                record.get("_date")
            ).date()
            == context.source_date.date()
        ):
            date_bonus = 1000.0

        score = (
            route_score * 1000.0
            + date_bonus
            - min(
                start_difference,
                24 * 3600,
            )
            / 60.0
        )

        scored.append(
            (
                score,
                route_score,
                start_difference,
                index,
            )
        )

    if not scored:
        return None

    scored.sort(
        reverse=True
    )

    _, route_score, _, best_index = scored[0]

    # Do not fuzzy-match an unrelated route.
    if (
        context.route_origin
        and context.route_destination
        and route_score < 1.10
    ):
        return None

    return candidates.loc[
        best_index
    ]


# ============================================================================
# MANUAL EVENT PARSING
# ============================================================================

def resolve_event_clock(
    record: pd.Series,
    context: PairVehicleContext,
    clock_value: Any,
) -> pd.Timestamp | None:
    clock = parse_clock(
        clock_value
    )

    if clock is None:
        return None

    date = record.get("_date")

    if date is None or pd.isna(date):
        return None

    manual_start = resolve_manual_start_time(
        record,
        context,
    )

    manual_end = resolve_manual_end_time(
        record,
        manual_start,
    )

    hour, minute, second, ampm = clock

    base = (
        pd.Timestamp(date)
        + safe_timedelta(
            hours=hour,
            minutes=minute,
            seconds=second,
        )
    )

    if ampm:
        candidates = [
            base,
            base + safe_timedelta(days=1),
        ]
    else:
        candidates = [
            base,
            base + safe_timedelta(hours=12),
            base + safe_timedelta(days=1),
            base + safe_timedelta(days=1, hours=12),
        ]

    if (
        manual_start is not None
        and manual_end is not None
    ):
        inside = [
            candidate
            for candidate in candidates
            if manual_start
            - safe_timedelta(minutes=30)
            <= candidate
            <= manual_end
            + safe_timedelta(minutes=30)
        ]

        if inside:
            return min(
                inside,
                key=lambda candidate: abs(
                    (
                        candidate
                        - manual_start
                    ).total_seconds()
                ),
            )

    if context.source_trip_start is not None:
        return min(
            candidates,
            key=lambda candidate: abs(
                (
                    candidate
                    - context.source_trip_start
                ).total_seconds()
            ),
        )

    return base


def extract_manual_stop_events(
    record: pd.Series | None,
    context: PairVehicleContext,
) -> list[ManualEvent]:
    """
    Extract normal stop timestamps from the matched manual trip record.

    Sources, in priority order:
        1. Optional dedicated Stop Timestamps / Stop time column.
        2. Existing Additional column (backward compatibility).

    IMPORTANT:
    - Charging Timestamp is handled separately by
      extract_manual_replenishment_event() and is NOT treated as a normal stop.
    - "Trip ended at ..." is not treated as a stop.
    - "Total stops - 29" contains a count but no usable timestamp and therefore
      does not invent stop events.
    """
    if record is None:
        return []

    dedicated = clean_text(
        record.get(
            "Stop Timestamps"
        )
    )

    additional = clean_text(
        record.get(
            "Additional"
        )
    )

    sources: list[tuple[str, str]] = []

    if (
        dedicated
        and not is_manual_placeholder(
            dedicated
        )
    ):
        # The maintained workbook has a dedicated stop field.  Once that
        # field is populated it is the sole source of manual STOP/IDLE
        # timestamps.  This prevents charging notes or unrelated numbers in
        # Additional from being parsed as a second set of stop events.
        sources.append(
            (
                "Stop Timestamps",
                dedicated,
            )
        )

    elif (
        additional
        and not is_manual_placeholder(
            additional
        )
    ):
        # Backward compatibility for old records that pre-date the dedicated
        # stop-timestamps column.
        sources.append(
            (
                "Additional",
                additional,
            )
        )

    if not sources:
        return []

    events: list[ManualEvent] = []
    event_index = 1
    seen_keys: set[tuple] = set()

    for source_column, raw_source in sources:
        # Do not treat end-of-trip clock values as stops.
        cleaned = re.sub(
            r"(?i)trip\s+ended\s+at\s+"
            r"\d{1,2}:\d{2}(?::\d{2})?",
            " ",
            raw_source,
        )

        consumed_spans: list[
            tuple[int, int]
        ] = []

        # -------------------------------------------------------------
        # A. Explicit interval in parentheses:
        #       5 (03:13–03:29)
        # -------------------------------------------------------------
        interval_pattern = re.compile(
            r"(?i)"
            r"(?:\b\d+\s*)?"
            r"\(\s*"
            r"(\d{1,2}:\d{2}(?::\d{2})?)"
            r"\s*[–—-]\s*"
            r"(\d{1,2}:\d{2}(?::\d{2})?)"
            r"\s*\)"
        )

        for match in interval_pattern.finditer(
            cleaned
        ):
            start = resolve_event_clock(
                record,
                context,
                match.group(1),
            )

            finish = resolve_event_clock(
                record,
                context,
                match.group(2),
            )

            if (
                start is not None
                and finish is not None
                and finish < start
            ):
                # Try 12-hour ambiguity first, then next day.
                finish += safe_timedelta(
                    hours=12
                )

                if finish < start:
                    finish += safe_timedelta(
                        hours=12
                    )

            duration = (
                (
                    finish
                    - start
                ).total_seconds()
                if (
                    start is not None
                    and finish is not None
                )
                else None
            )

            key = (
                "STOP_INTERVAL",
                start,
                finish,
            )

            if key not in seen_keys:
                events.append(
                    ManualEvent(
                        vehicle=context.vehicle,
                        event_kind="STOP_INTERVAL",
                        event_index=event_index,
                        raw_text=match.group(0),
                        start=start,
                        end=finish,
                        duration_seconds=duration,
                        location="",
                        source_column=source_column,
                        is_interval=True,
                    )
                )

                seen_keys.add(
                    key
                )
                event_index += 1

            consumed_spans.append(
                match.span()
            )

        # -------------------------------------------------------------
        # B. "10:32 to 10:37 break/stop/halt"
        # -------------------------------------------------------------
        break_pattern = re.compile(
            r"(?i)"
            r"(\d{1,2}:\d{2}(?::\d{2})?)"
            r"\s+to\s+"
            r"(\d{1,2}:\d{2}(?::\d{2})?)"
            r"\s*(?:break|stop|halt)"
        )

        for match in break_pattern.finditer(
            cleaned
        ):
            start = resolve_event_clock(
                record,
                context,
                match.group(1),
            )

            finish = resolve_event_clock(
                record,
                context,
                match.group(2),
            )

            if (
                start is not None
                and finish is not None
                and finish < start
            ):
                finish += safe_timedelta(
                    hours=12
                )

                if finish < start:
                    finish += safe_timedelta(
                        hours=12
                    )

            duration = (
                (
                    finish
                    - start
                ).total_seconds()
                if (
                    start is not None
                    and finish is not None
                )
                else None
            )

            key = (
                "STOP_INTERVAL",
                start,
                finish,
            )

            if key not in seen_keys:
                events.append(
                    ManualEvent(
                        vehicle=context.vehicle,
                        event_kind="STOP_INTERVAL",
                        event_index=event_index,
                        raw_text=match.group(0),
                        start=start,
                        end=finish,
                        duration_seconds=duration,
                        location="",
                        source_column=source_column,
                        is_interval=True,
                    )
                )

                seen_keys.add(
                    key
                )
                event_index += 1

            consumed_spans.append(
                match.span()
            )

        def span_consumed(
            span: tuple[int, int],
        ) -> bool:
            return any(
                span[0] >= used[0]
                and span[1] <= used[1]
                for used in consumed_spans
            )

        # -------------------------------------------------------------
        # C. Numbered timestamps:
        #       1 (07:44), 2 (07:50), ...
        # -------------------------------------------------------------
        point_pattern = re.compile(
            r"(?i)"
            r"\b\d+\s*"
            r"\(\s*"
            r"(\d{1,2}:\d{2}(?::\d{2})?)"
            r"\s*\)"
        )

        for match in point_pattern.finditer(
            cleaned
        ):
            if span_consumed(
                match.span()
            ):
                continue

            start = resolve_event_clock(
                record,
                context,
                match.group(1),
            )

            key = (
                "STOP_POINT",
                start,
            )

            if key in seen_keys:
                continue

            events.append(
                ManualEvent(
                    vehicle=context.vehicle,
                    event_kind="STOP_POINT",
                    event_index=event_index,
                    raw_text=match.group(0),
                    start=start,
                    end=None,
                    duration_seconds=None,
                    location="",
                    source_column=source_column,
                    is_interval=False,
                )
            )

            seen_keys.add(
                key
            )
            event_index += 1

        # -------------------------------------------------------------
        # D. Plain / comma-separated stop timestamps in a dedicated
        #    stop-time field, e.g.:
        #       07:44, 07:50, 08:00
        #
        #    Also accepted in Additional only when the text explicitly says
        #    intermediate timestamps/stops.
        # -------------------------------------------------------------
        allow_generic_points = (
            source_column
            == "Stop Timestamps"
            or bool(
                re.search(
                    r"(?i)"
                    r"intermediate\s+timestamps?"
                    r"|intermediate\s+stops?"
                    r"|stop\s+timestamps?",
                    cleaned,
                )
            )
        )

        if allow_generic_points:
            generic_time_pattern = re.compile(
                r"(?<!\d)"
                r"(\d{1,2}:\d{2}(?::\d{2})?)"
                r"(?!\d)"
            )

            for match in generic_time_pattern.finditer(
                cleaned
            ):
                if span_consumed(
                    match.span()
                ):
                    continue

                start = resolve_event_clock(
                    record,
                    context,
                    match.group(1),
                )

                key = (
                    "STOP_POINT",
                    start,
                )

                if (
                    start is None
                    or key in seen_keys
                ):
                    continue

                events.append(
                    ManualEvent(
                        vehicle=context.vehicle,
                        event_kind="STOP_POINT",
                        event_index=event_index,
                        raw_text=match.group(0),
                        start=start,
                        end=None,
                        duration_seconds=None,
                        location="",
                        source_column=source_column,
                        is_interval=False,
                    )
                )

                seen_keys.add(
                    key
                )
                event_index += 1

    return sorted(
        events,
        key=lambda event: (
            pd.Timestamp.max
            if event.start is None
            else event.start
        ),
    )


def extract_manual_replenishment_event(
    record: pd.Series | None,
    context: PairVehicleContext,
) -> ManualEvent | None:
    if record is None:
        return None

    charging_value = clean_text(
        record.get(
            "Charging Time"
        )
    )

    # NEW MANUAL-LOG COLUMN:
    # dedicated charging/refuelling start timestamp.
    charging_timestamp = clean_text(
        record.get(
            "Charging Timestamp"
        )
    )

    charging_type = clean_text(
        record.get(
            "Charging Type"
        )
    )

    additional = clean_text(
        record.get(
            "Additional"
        )
    )

    combined = " | ".join(
        value
        for value in (
            charging_value,
            charging_type,
            additional,
        )
        if value
    )

    if context.vehicle == "EV":
        keyword_pattern = (
            r"charg(?:e|ed|er|ing)?"
        )
        event_kind = "CHARGING"
    else:
        keyword_pattern = (
            r"refuel(?:ing|ling)?|"
            r"fuel(?:ing|ling)?"
        )
        event_kind = "REFUELING"

    # Authoritative duration source: Charging time (min).
    # Never infer charging duration from Additional or stop text; those cells
    # can contain stop counts and unrelated numbers.
    duration_minutes = parse_positive_minutes(
        charging_value
    )

    duration_seconds = (
        duration_minutes * 60.0
        if duration_minutes is not None
        else None
    )

    # ------------------------------------------------------------------
    # NEW DIRECT CHARGING/REFUELLING ACCESS
    # ------------------------------------------------------------------
    # New manual-log columns:
    #     Charging time (min)  -> duration in minutes
    #     charging timestamp   -> charging/refuelling start time
    #
    # If either dedicated field is present, use the maintained column layout
    # directly.  ``Charging time (min)`` is a duration and is never parsed as
    # a clock.  ``charging timestamp`` is a start clock and is never parsed as
    # a stop.  Additional may provide location text only.
    has_dedicated_duration = (
        duration_minutes is not None
    )

    has_dedicated_timestamp = (
        not is_manual_placeholder(
            charging_timestamp
        )
    )

    if (
        has_dedicated_duration
        or has_dedicated_timestamp
    ):
        event_start = (
            resolve_event_clock(
                record,
                context,
                charging_timestamp,
            )
            if has_dedicated_timestamp
            else (
                resolve_event_clock(
                    record,
                    context,
                    charging_value,
                )
                if parse_clock(
                    charging_value
                )
                is not None
                else None
            )
        )

        event_end = (
            event_start
            + safe_timedelta(
                seconds=duration_seconds
            )
            if (
                event_start is not None
                and duration_seconds is not None
            )
            else None
        )

        # Keep the previous location extraction logic conceptually unchanged:
        # location still comes from Additional when it is available.
        location = ""

        location_match = re.search(
            rf"(?i)"
            rf"(?:{keyword_pattern})"
            r"[^.;]{0,60}?"
            r"(?:"
            r"location\s*[:\-]?\s*|"
            r"completed\s+at\s+|"
            r"\bat\s+"
            r")"
            r"([A-Za-z][A-Za-z0-9 .,&()'/-]{2,80})",
            additional,
        )

        if location_match:
            location = (
                location_match.group(1)
                .strip(" ,.;:-")
            )

        direct_raw_text = (
            f"Charging time (min): {charging_value} | "
            f"charging timestamp: {charging_timestamp}"
        )

        if charging_type:
            direct_raw_text += (
                f" | Charging type: {charging_type}"
            )

        if additional:
            direct_raw_text += (
                f" | Additional: {additional}"
            )

        return ManualEvent(
            vehicle=context.vehicle,
            event_kind=event_kind,
            event_index=0,
            raw_text=direct_raw_text,
            start=event_start,
            end=event_end,
            duration_seconds=duration_seconds,
            location=location,
            source_column=(
                "Charging time (min) + charging timestamp"
            ),
            is_interval=(
                event_start is not None
                and event_end is not None
            ),
        )

    # Legacy fallback: used only when the maintained dedicated columns do not
    # contain a usable duration or timestamp.  This preserves support for old
    # free-text records without allowing current stop text to redefine a
    # charging event.
    has_value = not is_manual_placeholder(
        charging_value
    )

    has_keyword = bool(
        re.search(
            rf"(?i)\b(?:{keyword_pattern})\b",
            combined,
        )
    )

    if not has_value and not has_keyword:
        return None

    if (
        duration_seconds is None
        and has_keyword
    ):
        legacy_duration_minutes = parse_positive_minutes(
            additional
        )

        duration_seconds = (
            legacy_duration_minutes * 60.0
            if legacy_duration_minutes is not None
            else None
        )

    event_start = None

    if (
        has_value
        and parse_clock(
            charging_value
        )
        is not None
    ):
        event_start = resolve_event_clock(
            record,
            context,
            charging_value,
        )

    if event_start is None:
        contextual_time = re.search(
            rf"(?i)"
            rf"(?:{keyword_pattern})"
            r"[^.;]{0,100}?"
            r"(?<!\d)"
            r"(\d{1,2}:\d{2}(?::\d{2})?"
            r"\s*(?:am|pm)?)"
            r"(?!\d)",
            additional,
        )

        if contextual_time:
            event_start = resolve_event_clock(
                record,
                context,
                contextual_time.group(1),
            )

    event_end = (
        event_start
        + safe_timedelta(
            seconds=duration_seconds
        )
        if (
            event_start is not None
            and duration_seconds is not None
        )
        else None
    )

    location = ""

    location_match = re.search(
        rf"(?i)"
        rf"(?:{keyword_pattern})"
        r"[^.;]{0,60}?"
        r"(?:"
        r"location\s*[:\-]?\s*|"
        r"completed\s+at\s+|"
        r"\bat\s+"
        r")"
        r"([A-Za-z][A-Za-z0-9 .,&()'/-]{2,80})",
        additional,
    )

    if location_match:
        location = (
            location_match.group(1)
            .strip(" ,.;:-")
        )

    if not location and has_value:
        parenthetical = re.search(
            r"\(([^()]*)\)",
            charging_value,
        )

        if parenthetical:
            candidate = parenthetical.group(1).strip()

            if not re.fullmatch(
                r"(?i)"
                r"\d+(?:\.\d+)?\s*"
                r"(?:min|mins|minute|minutes)",
                candidate,
            ):
                location = candidate

    return ManualEvent(
        vehicle=context.vehicle,
        event_kind=event_kind,
        event_index=0,
        raw_text=combined,
        start=event_start,
        end=event_end,
        duration_seconds=duration_seconds,
        location=location,
        source_column="Charging Time / Additional",
        is_interval=(
            event_start is not None
            and event_end is not None
        ),
    )


# ============================================================================
# TELEMETRY PREPARATION
# ============================================================================

def prepare_telemetry(
    source: pd.DataFrame,
) -> pd.DataFrame:
    required = [
        TIMESTAMP,
        TIME_GAP,
        DISTANCE,
        SPEED_KMH,
        ACCELERATION,
    ]

    missing = [
        column
        for column in required
        if column not in source.columns
    ]

    if missing:
        raise ValueError(
            "Missing required column(s): "
            + ", ".join(
                missing
            )
        )

    frame = source.copy()

    frame[TIMESTAMP] = parse_datetime_series(
        frame[TIMESTAMP]
    )

    for column in (
        TIME_GAP,
        DISTANCE,
        SPEED_KMH,
        ACCELERATION,
    ):
        frame[column] = pd.to_numeric(
            frame[column],
            errors="coerce",
        )

    if ADDRESS not in frame.columns:
        frame[ADDRESS] = ""

    if REVERSE_ADDRESS not in frame.columns:
        frame[
            REVERSE_ADDRESS
        ] = ""

    frame[ADDRESS] = (
        frame[ADDRESS]
        .fillna("")
        .astype(str)
    )

    frame[REVERSE_ADDRESS] = (
        frame[REVERSE_ADDRESS]
        .fillna("")
        .astype(str)
    )

    return frame.reset_index(
        drop=True
    )


def gap_equivalent_speed(
    gap_seconds: np.ndarray,
    distance_m: np.ndarray,
) -> np.ndarray:
    result = np.full(
        len(gap_seconds),
        np.nan,
        dtype=float,
    )

    valid = (
        np.isfinite(
            gap_seconds
        )
        & (
            gap_seconds > 0
        )
        & np.isfinite(
            distance_m
        )
        & (
            distance_m >= 0
        )
    )

    result[valid] = (
        distance_m[valid]
        / gap_seconds[valid]
        * 3.6
    )

    return result



def consecutive_true_runs(
    mask: np.ndarray,
) -> list[np.ndarray]:
    """
    Return consecutive True index runs.

    Used only for the fallback treatment of consecutive short
    stationary-like telemetry intervals.
    """
    indices = np.flatnonzero(
        mask
    )

    if len(indices) == 0:
        return []

    split_points = np.where(
        np.diff(indices) > 1
    )[0] + 1

    return [
        run.astype(int)
        for run in np.split(
            indices,
            split_points,
        )
        if len(run)
    ]


def interval_start_times(
    frame: pd.DataFrame,
) -> pd.Series:
    return (
        frame[TIMESTAMP]
        - pd.to_timedelta(
            frame[TIME_GAP],
            unit="s",
            errors="coerce",
        )
    )


def manual_event_overlap_mask(
    frame: pd.DataFrame,
    event: ManualEvent | None,
) -> np.ndarray:
    """
    Return telemetry intervals that overlap an explicit manual event window.

    This is used only when the manual log provides both a start time and an
    end time (or a start time plus duration).  The ordinary telemetry rules are
    unchanged for every other row.
    """
    mask = np.zeros(
        len(frame),
        dtype=bool,
    )

    if (
        event is None
        or event.start is None
        or event.end is None
        or event.end <= event.start
    ):
        return mask

    starts = interval_start_times(
        frame
    )

    ends = frame[
        TIMESTAMP
    ]

    positive = (
        pd.to_numeric(
            frame[TIME_GAP],
            errors="coerce",
        ).fillna(0.0)
        > 0
    ).to_numpy(dtype=bool)

    overlap = (
        starts.notna()
        & ends.notna()
        & (
            starts
            < pd.Timestamp(event.end)
        )
        & (
            ends
            > pd.Timestamp(event.start)
        )
    ).to_numpy(dtype=bool)

    return (
        overlap
        & positive
    )


def event_match_from_rows(
    frame: pd.DataFrame,
    event: ManualEvent,
    rows: np.ndarray,
    method: str,
    station_evidence: np.ndarray,
    notes: str,
) -> EventMatch:
    """Build one audit record for a manual event applied to one or more rows."""
    rows = np.asarray(rows, dtype=int)
    if len(rows) == 0:
        raise ValueError("event_match_from_rows requires at least one telemetry row")

    starts = interval_start_times(frame)
    first = int(rows[0])
    last = int(rows[-1])
    gap = pd.to_numeric(frame[TIME_GAP], errors="coerce").to_numpy(dtype=float)
    distance = pd.to_numeric(frame[DISTANCE], errors="coerce").to_numpy(dtype=float)

    total_gap = float(np.nansum(gap[rows]))
    total_distance = float(np.nansum(distance[rows]))
    aggregate_speed = (
        total_distance / total_gap * 3.6
        if total_gap > 0
        else None
    )

    interval_start = (
        pd.Timestamp(starts.iloc[first])
        if pd.notna(starts.iloc[first])
        else None
    )
    interval_end = (
        pd.Timestamp(frame.at[last, TIMESTAMP])
        if pd.notna(frame.at[last, TIMESTAMP])
        else None
    )

    offset = None
    if event.start is not None and interval_start is not None and interval_end is not None:
        direct, computed_offset = candidate_time_offset(
            event, interval_start, interval_end
        )
        offset = 0.0 if direct else float(computed_offset)

    return EventMatch(
        vehicle=event.vehicle,
        event_kind=event.event_kind,
        event_index=event.event_index,
        manual_start=event.start,
        manual_end=event.end,
        manual_duration_seconds=event.duration_seconds,
        manual_location=event.location,
        raw_text=event.raw_text,
        match_status="MATCHED",
        match_method=method,
        telemetry_row_index=first,
        excel_row=first + 2,
        interval_start=interval_start,
        interval_end=interval_end,
        telemetry_gap_seconds=total_gap,
        telemetry_distance_m=total_distance,
        gap_equivalent_speed_kmh=aggregate_speed,
        time_offset_seconds=offset,
        boundary_flag=bool(boundary_mask(len(frame))[rows].any()),
        address_station_evidence=bool(station_evidence[rows].any()),
        notes=notes,
    )


def recorded_zero_speed_episodes(
    positive: np.ndarray,
    gap_seconds: np.ndarray,
    device_speed_kmh: np.ndarray,
) -> tuple[list[np.ndarray], np.ndarray, np.ndarray]:
    """Identify cumulative stationary episodes in dense/mobile telemetry.

    Device datasets often omit zero-speed rows, so their stationary periods
    continue to be inferred from a single large time gap.  Mobile-app datasets
    usually retain many short consecutive rows with recorded speed equal (or
    very close) to zero.  For those files, the correct duration is the sum of
    the consecutive interval gaps, not the duration of any one row.

    Returns
    -------
    episodes:
        Consecutive row-index arrays with at least
        ``MIN_ZERO_SPEED_EPISODE_INTERVALS`` intervals.
    episode_id:
        Integer episode identifier per row, or -1 outside an episode.
    episode_duration_seconds:
        Total cumulative duration repeated on every member row, or NaN.
    """
    zero_candidate = (
        positive
        & np.isfinite(device_speed_kmh)
        & (device_speed_kmh >= 0.0)
        & (
            device_speed_kmh
            <= RECORDED_ZERO_SPEED_MAX_KMH
        )
    )

    episodes = [
        run
        for run in consecutive_true_runs(zero_candidate)
        if len(run) >= MIN_ZERO_SPEED_EPISODE_INTERVALS
    ]

    episode_id = np.full(
        len(gap_seconds),
        -1,
        dtype=int,
    )
    episode_duration_seconds = np.full(
        len(gap_seconds),
        np.nan,
        dtype=float,
    )

    for identifier, run in enumerate(episodes, start=1):
        duration = float(
            np.nansum(
                gap_seconds[run]
            )
        )
        episode_id[run] = identifier
        episode_duration_seconds[run] = duration

    return (
        episodes,
        episode_id,
        episode_duration_seconds,
    )


def station_keyword_mask(
    frame: pd.DataFrame,
    vehicle: str,
) -> np.ndarray:
    """
    Evidence can occur at either end of the gap interval.
    """
    keywords = (
        EV_STATION_KEYWORDS
        if vehicle == "EV"
        else ICE_STATION_KEYWORDS
    )

    current = (
        frame[ADDRESS]
        .fillna("")
        .astype(str)
        + " "
        + frame[
            REVERSE_ADDRESS
        ]
        .fillna("")
        .astype(str)
    )

    previous = current.shift(
        1,
        fill_value="",
    )

    interval_text = (
        previous
        + " "
        + current
    ).str.casefold()

    mask = np.zeros(
        len(frame),
        dtype=bool,
    )

    for keyword in keywords:
        mask |= interval_text.str.contains(
            re.escape(
                keyword.casefold()
            ),
            regex=True,
            na=False,
        ).to_numpy(
            dtype=bool
        )

    return mask


def manual_location_mask(
    frame: pd.DataFrame,
    location: str,
) -> np.ndarray:
    if not location:
        return np.zeros(
            len(frame),
            dtype=bool,
        )

    current = (
        frame[ADDRESS]
        .fillna("")
        .astype(str)
        + " "
        + frame[
            REVERSE_ADDRESS
        ]
        .fillna("")
        .astype(str)
    )

    previous = current.shift(
        1,
        fill_value="",
    )

    interval_text = (
        previous
        + " "
        + current
    )

    wanted = normalise_text(
        location
    )

    wanted_place = canonical_place(
        location
    )

    wanted_tokens = {
        token
        for token in wanted.split()
        if len(token) >= 4
    }

    matches = []

    for value in interval_text:
        actual = normalise_text(
            value
        )

        actual_place = canonical_place(
            value
        )

        actual_tokens = set(
            actual.split()
        )

        exact = (
            bool(wanted)
            and wanted in actual
        )

        place_match = (
            bool(wanted_place)
            and wanted_place == actual_place
        )

        token_match = (
            bool(wanted_tokens)
            and len(
                wanted_tokens
                & actual_tokens
            )
            >= min(
                2,
                len(
                    wanted_tokens
                ),
            )
        )

        matches.append(
            exact
            or place_match
            or token_match
        )

    return np.array(
        matches,
        dtype=bool,
    )


def boundary_mask(
    length: int,
) -> np.ndarray:
    mask = np.zeros(
        length,
        dtype=bool,
    )

    if length == 0:
        return mask

    first_end = min(
        BOUNDARY_ROW_COUNT,
        length,
    )

    mask[
        :first_end
    ] = True

    last_start = max(
        0,
        length
        - BOUNDARY_ROW_COUNT
    )

    mask[
        last_start:
    ] = True

    return mask


# ============================================================================
# MANUAL EVENT MATCHING
# ============================================================================

def candidate_time_offset(
    event: ManualEvent,
    interval_start: pd.Timestamp,
    interval_end: pd.Timestamp,
) -> tuple[
    bool,
    float,
]:
    if event.start is None:
        return False, float("inf")

    manual_time = pd.Timestamp(
        event.start
    )

    if (
        interval_start
        <= manual_time
        <= interval_end
    ):
        return True, 0.0

    return (
        False,
        min(
            abs(
                (
                    manual_time
                    - interval_start
                ).total_seconds()
            ),
            abs(
                (
                    manual_time
                    - interval_end
                ).total_seconds()
            ),
        ),
    )


def match_event_to_stationary_gap(
    frame: pd.DataFrame,
    event: ManualEvent,
    eligible_mask: np.ndarray,
    address_station_evidence: np.ndarray,
    used_rows: set[int],
    allow_location_only: bool,
    relaxed_mask: np.ndarray | None = None,
) -> EventMatch:
    """
    Match a manual STOP / CHARGING / REFUELING event to telemetry.

    PASS 1 (strict):
        Uses the caller's existing stationary-like eligibility mask.

    PASS 2 (manual-evidence fallback):
        If no strict match exists, use relaxed_mask and search the same
        +/-15-minute window.

    Why the second pass is necessary:
        the device does not log true zero-speed points and timestamps can lag;
        therefore a genuine manually recorded stop can sit next to a telemetry
        gap whose GPS displacement makes gap-equivalent speed slightly >1.

    The relaxed pass is used ONLY because an explicit manual event exists.
    """
    gap = frame[
        TIME_GAP
    ].to_numpy(
        dtype=float
    )

    distance = frame[
        DISTANCE
    ].to_numpy(
        dtype=float
    )

    gap_speed = gap_equivalent_speed(
        gap,
        distance,
    )

    starts = interval_start_times(
        frame
    )

    ends = frame[
        TIMESTAMP
    ]

    boundaries = boundary_mask(
        len(frame)
    )

    location_evidence = manual_location_mask(
        frame,
        event.location,
    )

    def build_ranked(
        candidate_mask: np.ndarray,
        relaxed: bool,
    ) -> list[
        tuple[
            tuple[Any, ...],
            int,
            bool,
            float,
        ]
    ]:
        ranked: list[
            tuple[
                tuple[Any, ...],
                int,
                bool,
                float,
            ]
        ] = []

        candidate_rows = [
            int(index)
            for index in np.flatnonzero(
                candidate_mask
            )
            if int(index) not in used_rows
        ]

        for row in candidate_rows:
            if (
                pd.isna(
                    starts.iloc[row]
                )
                or pd.isna(
                    ends.iloc[row]
                )
            ):
                continue

            interval_start = pd.Timestamp(
                starts.iloc[row]
            )

            interval_end = pd.Timestamp(
                ends.iloc[row]
            )

            direct, offset = candidate_time_offset(
                event,
                interval_start,
                interval_end,
            )

            # Explicit manual interval overlap is additional strong evidence.
            manual_interval_overlap = 0.0

            if (
                event.start is not None
                and event.end is not None
            ):
                overlap_start = max(
                    interval_start,
                    pd.Timestamp(
                        event.start
                    ),
                )

                overlap_end = min(
                    interval_end,
                    pd.Timestamp(
                        event.end
                    ),
                )

                manual_interval_overlap = max(
                    0.0,
                    (
                        overlap_end
                        - overlap_start
                    ).total_seconds(),
                )

            time_eligible = (
                direct
                or manual_interval_overlap
                > 0
                or (
                    math.isfinite(
                        offset
                    )
                    and offset
                    <= MANUAL_EVENT_TOLERANCE_SECONDS
                )
            )

            if event.start is None:
                time_eligible = (
                    allow_location_only
                    and bool(
                        location_evidence[
                            row
                        ]
                    )
                )

            if not time_eligible:
                continue

            duration_difference = float(
                "inf"
            )

            if (
                event.duration_seconds
                is not None
                and math.isfinite(
                    event.duration_seconds
                )
            ):
                duration_difference = abs(
                    gap[
                        row
                    ]
                    - event.duration_seconds
                )

            current_gap_speed = (
                float(
                    gap_speed[
                        row
                    ]
                )
                if np.isfinite(
                    gap_speed[
                        row
                    ]
                )
                else float(
                    "inf"
                )
            )

            # Strict pass keeps the old ranking philosophy.
            if not relaxed:
                score = (
                    0
                    if direct
                    else 1,
                    0
                    if location_evidence[
                        row
                    ]
                    else 1,
                    0
                    if address_station_evidence[
                        row
                    ]
                    else 1,
                    -manual_interval_overlap,
                    offset,
                    duration_difference,
                )

            else:
                # In the fallback pass, manual evidence is already known.
                # Prefer a plausible missing-data/stationary interval:
                #   - overlapping manual interval
                #   - location/station support
                #   - lower displacement rate
                #   - longer telemetry gap
                #   - closer timestamp
                #   - duration agreement
                score = (
                    0
                    if manual_interval_overlap
                    > 0
                    else 1,
                    0
                    if location_evidence[
                        row
                    ]
                    else 1,
                    0
                    if address_station_evidence[
                        row
                    ]
                    else 1,
                    current_gap_speed,
                    -float(
                        gap[
                            row
                        ]
                    ),
                    offset,
                    duration_difference,
                    0
                    if direct
                    else 1,
                )

            ranked.append(
                (
                    score,
                    row,
                    direct,
                    offset,
                )
            )

        ranked.sort(
            key=lambda item: item[
                0
            ]
        )

        return ranked

    # -------------------------------------------------------------
    # PASS 1 - original strict stationary-like match.
    # -------------------------------------------------------------
    ranked = build_ranked(
        eligible_mask,
        relaxed=False,
    )

    match_mode = (
        "STRICT_STATIONARY"
    )

    # -------------------------------------------------------------
    # PASS 2 - explicit manual evidence fallback.
    # -------------------------------------------------------------
    if (
        not ranked
        and relaxed_mask is not None
    ):
        ranked = build_ranked(
            relaxed_mask,
            relaxed=True,
        )

        match_mode = (
            "MANUAL_EVIDENCE_RELAXED"
        )

    if not ranked:
        return EventMatch(
            vehicle=event.vehicle,
            event_kind=event.event_kind,
            event_index=event.event_index,
            manual_start=event.start,
            manual_end=event.end,
            manual_duration_seconds=event.duration_seconds,
            manual_location=event.location,
            raw_text=event.raw_text,
            match_status="NOT_MATCHED",
            match_method=(
                "NO_TELEMETRY_GAP_WITHIN_15_MIN"
            ),
            telemetry_row_index=None,
            excel_row=None,
            interval_start=None,
            interval_end=None,
            telemetry_gap_seconds=None,
            telemetry_distance_m=None,
            gap_equivalent_speed_kmh=None,
            time_offset_seconds=None,
            boundary_flag=False,
            address_station_evidence=False,
            notes=(
                "No compatible telemetry interval was found within "
                "+/-15 minutes of the explicit manual event."
            ),
        )

    _, row, direct, offset = ranked[
        0
    ]

    if event.start is None:
        method = (
            "MANUAL_LOCATION_MATCH"
        )

    elif direct:
        method = (
            "MANUAL_TIME_INSIDE_GAP"
        )

    else:
        method = (
            "MANUAL_TIME_WITHIN_15_MIN"
        )

    method = (
        f"{match_mode} + "
        f"{method}"
    )

    if (
        event.location
        and location_evidence[
            row
        ]
    ):
        method += (
            " + LOCATION_SUPPORTED"
        )

    if event.duration_seconds is not None:
        method += (
            " + DURATION_RANKING"
        )

    return EventMatch(
        vehicle=event.vehicle,
        event_kind=event.event_kind,
        event_index=event.event_index,
        manual_start=event.start,
        manual_end=event.end,
        manual_duration_seconds=event.duration_seconds,
        manual_location=event.location,
        raw_text=event.raw_text,
        match_status="MATCHED",
        match_method=method,
        telemetry_row_index=row,
        excel_row=row + 2,
        interval_start=pd.Timestamp(
            starts.iloc[
                row
            ]
        ),
        interval_end=pd.Timestamp(
            ends.iloc[
                row
            ]
        ),
        telemetry_gap_seconds=float(
            gap[
                row
            ]
        ),
        telemetry_distance_m=(
            float(
                distance[
                    row
                ]
            )
            if np.isfinite(
                distance[
                    row
                ]
            )
            else None
        ),
        gap_equivalent_speed_kmh=(
            float(
                gap_speed[
                    row
                ]
            )
            if np.isfinite(
                gap_speed[
                    row
                ]
            )
            else None
        ),
        time_offset_seconds=(
            float(
                offset
            )
            if math.isfinite(
                offset
            )
            else None
        ),
        boundary_flag=bool(
            boundaries[
                row
            ]
        ),
        address_station_evidence=bool(
            address_station_evidence[
                row
            ]
        ),
        notes=(
            "Strict telemetry-stationary match."
            if match_mode
            == "STRICT_STATIONARY"
            else (
                "Matched using explicit manual-event evidence because "
                "the strict stationary-like telemetry test did not match."
            )
        ),
    )



def select_manual_replenishment_rows(
    frame: pd.DataFrame,
    event: ManualEvent,
    positive: np.ndarray,
    stationary_like: np.ndarray,
    device_gap_long: np.ndarray,
    device_speed_kmh: np.ndarray,
    gap_speed_kmh: np.ndarray,
    address_station_evidence: np.ndarray,
    used_rows: set[int],
    normal_stop_events: list[ManualEvent] | None = None,
) -> tuple[
    EventMatch,
    np.ndarray,
]:
    """Select the telemetry rows belonging to one manual charging/refuelling event.

    The manual charging timestamp is an approximate anchor.  When the maintained
    log also provides ``Charging time (min)``, that duration is treated as the
    TARGET duration rather than as an exact start->end window.  This is important
    because the hand-recorded timestamp may be a little early/late while the
    telemetry may contain a longer low-speed dwell around the charger.

    Selection is deliberately conservative:

    1. Build contiguous low-movement runs using established stationary evidence,
       sparse-device long gaps, or very-low-speed (<=5 km/h) telemetry.
    2. If a timestamp exists, keep runs containing the manual anchor or lying within
       the existing +/-15-minute tolerance. If timestamp is absent but a manual
       location exists, require that location. If both timestamp and location are
       absent but a positive manual duration exists, search the whole trip and rank
       low-movement candidates using duration agreement, charging-station evidence,
       and separation from logged normal-stop timestamps.
    3. When a manual duration exists and the run is longer than that duration,
       choose the contiguous sub-window whose ORIGINAL telemetry Time Difference
       total is closest to the target.  Ties prefer a start closer to the manual
       anchor and lower GPS displacement.
    4. Never include clearly moving rows solely to make the duration equal the
       manual value.  If the available low-movement evidence is shorter than the
       manual duration, keep the observed low-movement rows and expose the shortfall
       in the existing audit/details output instead of corrupting motion phases.

    No telemetry values are edited and no row is split.  The returned rows are
    reserved before STOP/IDLE/INCH classification.
    """
    empty_rows = np.array([], dtype=int)

    gap = frame[TIME_GAP].to_numpy(dtype=float)
    distance = frame[DISTANCE].to_numpy(dtype=float)
    starts = interval_start_times(frame)
    ends = frame[TIMESTAMP]

    manual_location_evidence = manual_location_mask(
        frame,
        event.location,
    )

    # Normal manual-stop timestamps are NEGATIVE ranking evidence for a
    # duration-only charging event. They do not override charging priority, but
    # when the charging timestamp/location is missing they help us avoid choosing
    # a stationary episode that the log explicitly identified as an ordinary stop.
    manual_normal_stop_evidence = np.zeros(
        len(frame),
        dtype=bool,
    )

    if normal_stop_events:
        alignment = safe_timedelta(
            seconds=MANUAL_STOP_START_ALIGNMENT_SECONDS
        )

        for stop_event in normal_stop_events:
            if stop_event.start is None:
                continue

            if (
                stop_event.is_interval
                and stop_event.end is not None
                and stop_event.end > stop_event.start
            ):
                manual_normal_stop_evidence |= manual_event_overlap_mask(
                    frame,
                    stop_event,
                )
                continue

            stop_time = pd.Timestamp(
                stop_event.start
            )

            nearby = (
                starts.notna()
                & ends.notna()
                & (
                    starts
                    <= stop_time
                    + alignment
                )
                & (
                    ends
                    >= stop_time
                    - alignment
                )
            ).to_numpy(
                dtype=bool
            )

            manual_normal_stop_evidence |= nearby

    low_device_speed = (
        positive
        & np.isfinite(device_speed_kmh)
        & (device_speed_kmh >= 0.0)
        & (device_speed_kmh <= INCH_MAX_KMH)
    )

    low_gap_speed = (
        positive
        & np.isfinite(gap_speed_kmh)
        & (gap_speed_kmh >= 0.0)
        & (gap_speed_kmh <= INCH_MAX_KMH)
    )

    # Charging can contain genuine zero-speed telemetry plus very slow plug-in /
    # parking manoeuvring.  Either device speed or GPS-equivalent speed being low
    # is sufficient evidence here because explicit MANUAL charging evidence is
    # already required before this function is called.
    low_movement = (
        positive
        & (
            stationary_like
            | device_gap_long
            | low_device_speed
            | low_gap_speed
        )
    )

    for row in used_rows:
        if 0 <= row < len(low_movement):
            low_movement[row] = False

    target_duration = (
        float(event.duration_seconds)
        if event.duration_seconds is not None
        and math.isfinite(event.duration_seconds)
        and event.duration_seconds > 0
        else None
    )

    anchor = (
        pd.Timestamp(event.start)
        if event.start is not None
        else None
    )

    tolerance = float(MANUAL_EVENT_TOLERANCE_SECONDS)

    def temporal_relation(
        window_start: pd.Timestamp,
        window_end: pd.Timestamp,
    ) -> tuple[bool, float, float]:
        """Return (eligible, boundary_offset_s, start_offset_s)."""
        if anchor is None:
            return True, 0.0, 0.0

        contains = (
            window_start
            <= anchor
            <= window_end
        )

        boundary_offset = (
            0.0
            if contains
            else min(
                abs((anchor - window_start).total_seconds()),
                abs((anchor - window_end).total_seconds()),
            )
        )

        start_offset = abs(
            (window_start - anchor).total_seconds()
        )

        return (
            contains
            or boundary_offset <= tolerance,
            float(boundary_offset),
            float(start_offset),
        )

    candidate_runs: list[
        tuple[
            tuple[Any, ...],
            np.ndarray,
            pd.Timestamp,
            pd.Timestamp,
            float,
        ]
    ] = []

    for run in consecutive_true_runs(low_movement):
        run = np.asarray(run, dtype=int)
        if len(run) == 0:
            continue

        first = int(run[0])
        last = int(run[-1])

        if (
            pd.isna(starts.iloc[first])
            or pd.isna(ends.iloc[last])
        ):
            continue

        run_start = pd.Timestamp(starts.iloc[first])
        run_end = pd.Timestamp(ends.iloc[last])

        eligible, boundary_offset, start_offset = temporal_relation(
            run_start,
            run_end,
        )

        location_supported = bool(
            manual_location_evidence[run].any()
        )

        # No timestamp:
        #   * location present  -> require the maintained manual location;
        #   * duration only     -> the positive manual duration itself is explicit
        #                          charging/refuelling evidence, so search all safe
        #                          low-movement runs in the trip.
        # Address/station text still remains ranking evidence only; it cannot create
        # charging when the manual event itself is absent.
        if anchor is None:
            if event.location:
                eligible = location_supported
            elif target_duration is not None:
                eligible = True
            else:
                eligible = False

        if not eligible:
            continue

        run_duration = float(
            np.nansum(gap[run])
        )

        run_distance = float(
            np.nansum(distance[run])
        )

        run_speed = (
            run_distance / run_duration * 3.6
            if run_duration > 0
            else float("inf")
        )

        duration_shortfall = (
            max(0.0, target_duration - run_duration)
            if target_duration is not None
            else 0.0
        )

        duration_difference = (
            abs(run_duration - target_duration)
            if target_duration is not None
            else 0.0
        )

        contains_anchor = (
            anchor is not None
            and run_start <= anchor <= run_end
        )

        station_supported = bool(
            address_station_evidence[run].any()
        )

        normal_stop_overlap_seconds = float(
            np.nansum(
                gap[
                    run[
                        manual_normal_stop_evidence[
                            run
                        ]
                    ]
                ]
            )
        )

        # Duration feasibility is intentionally considered before exact anchor
        # containment.  A five-minute stationary blip exactly at a manual
        # 30-minute charging timestamp must not beat a nearby 30-minute dwell.
        score = (
            0
            if target_duration is None
            or duration_shortfall <= max(60.0, target_duration * 0.10)
            else 1,
            duration_shortfall,
            0 if contains_anchor else (1 if anchor is not None else 0),
            boundary_offset,
            0 if location_supported else (1 if event.location else 0),
            0 if station_supported else 1,
            normal_stop_overlap_seconds,
            duration_difference,
            run_speed,
            start_offset,
            -run_duration,
        )

        candidate_runs.append(
            (
                score,
                run,
                run_start,
                run_end,
                run_duration,
            )
        )

    # Sparse-device safety fallback.  If no contiguous low-movement run was
    # available, retain the earlier manual-event matcher so a genuine charging
    # stop represented by one long missing-record gap can still be captured.
    if not candidate_runs:
        fallback_match = match_event_to_stationary_gap(
            frame=frame,
            event=event,
            eligible_mask=(
                positive
                & (
                    stationary_like
                    | device_gap_long
                )
            ),
            address_station_evidence=address_station_evidence,
            used_rows=used_rows,
            allow_location_only=True,
            relaxed_mask=(
                positive
                & np.isfinite(gap)
                & (gap > IDLE_MIN_SECONDS)
            ),
        )

        if (
            fallback_match.match_status == "MATCHED"
            and fallback_match.telemetry_row_index is not None
        ):
            row = int(
                fallback_match.telemetry_row_index
            )
            rows = np.array(
                [row],
                dtype=int,
            )

            rebuilt = event_match_from_rows(
                frame=frame,
                event=event,
                rows=rows,
                method=(
                    f"{fallback_match.match_method} + "
                    "SPARSE_DEVICE_REPLENISHMENT_FALLBACK"
                ),
                station_evidence=address_station_evidence,
                notes=(
                    "No contiguous low-movement charging episode was available; "
                    "the manually recorded event was matched to the safest "
                    "compatible sparse-device telemetry gap."
                ),
            )
            return rebuilt, rows

        return fallback_match, empty_rows

    candidate_runs.sort(
        key=lambda item: item[0]
    )

    _, best_run, run_start, run_end, run_duration = candidate_runs[0]
    selected_rows = best_run.copy()

    duration_only_manual = (
        anchor is None
        and not event.location
        and target_duration is not None
    )

    selection_method = (
        "MANUAL_DURATION_ONLY_MATCHED_LOW_MOVEMENT_EPISODE"
        if duration_only_manual
        else "MANUAL_MATCHED_LOW_MOVEMENT_EPISODE"
    )

    # If the telemetry dwell is longer than the manual duration, choose a
    # contiguous row window whose observed duration is as close as possible to
    # the manual target.  Prefix sums keep this O(n log n), even for dense data.
    if (
        target_duration is not None
        and run_duration > target_duration
        and len(best_run) > 1
    ):
        run_gaps = np.nan_to_num(
            gap[best_run],
            nan=0.0,
            posinf=0.0,
            neginf=0.0,
        )

        prefix = np.concatenate(
            ([0.0], np.cumsum(run_gaps))
        )

        window_candidates: list[
            tuple[
                tuple[Any, ...],
                np.ndarray,
            ]
        ] = []

        for left in range(len(best_run)):
            desired = (
                prefix[left]
                + target_duration
            )

            right_position = int(
                np.searchsorted(
                    prefix,
                    desired,
                    side="left",
                )
            )

            # Test the first window meeting the target and the immediately
            # shorter one.  Whole telemetry rows are preserved; no row splitting.
            for right in {
                right_position,
                right_position - 1,
            }:
                if (
                    right <= left
                    or right > len(best_run)
                ):
                    continue

                rows = best_run[left:right]
                if len(rows) == 0:
                    continue

                first = int(rows[0])
                last = int(rows[-1])

                if (
                    pd.isna(starts.iloc[first])
                    or pd.isna(ends.iloc[last])
                ):
                    continue

                window_start = pd.Timestamp(
                    starts.iloc[first]
                )
                window_end = pd.Timestamp(
                    ends.iloc[last]
                )

                eligible, boundary_offset, start_offset = temporal_relation(
                    window_start,
                    window_end,
                )

                if not eligible:
                    continue

                observed_duration = float(
                    prefix[right]
                    - prefix[left]
                )

                observed_distance = float(
                    np.nansum(
                        distance[rows]
                    )
                )

                observed_speed = (
                    observed_distance
                    / observed_duration
                    * 3.6
                    if observed_duration > 0
                    else float("inf")
                )

                contains_anchor = (
                    anchor is not None
                    and window_start
                    <= anchor
                    <= window_end
                )

                location_supported = bool(
                    manual_location_evidence[rows].any()
                )

                station_supported = bool(
                    address_station_evidence[rows].any()
                )

                normal_stop_overlap_seconds = float(
                    np.nansum(
                        gap[
                            rows[
                                manual_normal_stop_evidence[
                                    rows
                                ]
                            ]
                        ]
                    )
                )

                score = (
                    abs(
                        observed_duration
                        - target_duration
                    ),
                    0 if contains_anchor else (1 if anchor is not None else 0),
                    boundary_offset,
                    start_offset,
                    0 if location_supported else (1 if event.location else 0),
                    0 if station_supported else 1,
                    normal_stop_overlap_seconds,
                    observed_speed,
                    observed_distance,
                    left,
                )

                window_candidates.append(
                    (
                        score,
                        np.asarray(
                            rows,
                            dtype=int,
                        ),
                    )
                )

        if window_candidates:
            window_candidates.sort(
                key=lambda item: item[0]
            )
            selected_rows = (
                window_candidates[0][1]
            )
            selection_method = (
                "MANUAL_DURATION_ONLY_TARGETED_LOW_MOVEMENT_WINDOW"
                if duration_only_manual
                else "MANUAL_DURATION_TARGETED_LOW_MOVEMENT_WINDOW"
            )

    selected_duration = float(
        np.nansum(
            gap[selected_rows]
        )
    )

    selected_distance = float(
        np.nansum(
            distance[selected_rows]
        )
    )

    duration_note = (
        f"Manual target={target_duration:.1f} s; "
        f"selected telemetry={selected_duration:.1f} s. "
        if target_duration is not None
        else (
            f"No manual duration; selected complete low-movement episode "
            f"of {selected_duration:.1f} s. "
        )
    )

    match = event_match_from_rows(
        frame=frame,
        event=event,
        rows=selected_rows,
        method=selection_method,
        station_evidence=address_station_evidence,
        notes=(
            duration_note
            + "Charging/refuelling rows were selected from contiguous "
            "stationary/very-low-speed telemetry around the approximate manual "
            "anchor. Clearly moving rows were not added merely to force the "
            "manual duration. "
            + f"Selected distance={selected_distance:.1f} m."
        ),
    )

    return match, selected_rows


def infer_point_stop_from_telemetry(
    frame: pd.DataFrame,
    event: ManualEvent,
    positive: np.ndarray,
    stationary_like: np.ndarray,
    zero_speed_episode_id: np.ndarray,
    excluded_rows: np.ndarray,
    used_rows: set[int],
    address_station_evidence: np.ndarray,
) -> tuple[
    EventMatch,
    np.ndarray,
    float | None,
]:
    """Infer a start-only manual stop's duration from forward telemetry.

    The value in ``stop timestamps`` is the start of the event.  A telemetry
    row describes the interval ``[Timestamp - Time Gap, Timestamp]``.  The
    inference therefore looks *forward* from the manual time and never assigns
    an interval whose end is exactly equal to the manual start.

    Selection order
    ---------------
    1. The complete consecutive stationary/recorded-zero episode containing
       or starting shortly after the manual start.
    2. If the data source omitted stationary rows, the next telemetry interval
       beginning at, containing, or occurring shortly after the manual start.

    Charging/refuelling intervals and rows already used by another manual
    event are excluded before the search.  The returned duration is the sum of
    the original telemetry ``Time Difference (s)`` values for the selected
    rows; the telemetry table itself is never edited or duration-inflated.
    """
    empty_rows = np.array(
        [],
        dtype=int,
    )

    if event.start is None:
        return (
            EventMatch(
                vehicle=event.vehicle,
                event_kind=event.event_kind,
                event_index=event.event_index,
                manual_start=event.start,
                manual_end=event.end,
                manual_duration_seconds=event.duration_seconds,
                manual_location=event.location,
                raw_text=event.raw_text,
                match_status="NOT_MATCHED",
                match_method="MANUAL_STOP_START_MISSING",
                telemetry_row_index=None,
                excel_row=None,
                interval_start=None,
                interval_end=None,
                telemetry_gap_seconds=None,
                telemetry_distance_m=None,
                gap_equivalent_speed_kmh=None,
                time_offset_seconds=None,
                boundary_flag=False,
                address_station_evidence=False,
                notes="The manual stop entry did not contain a usable start time.",
            ),
            empty_rows,
            None,
        )

    manual_start = pd.Timestamp(
        event.start
    )

    gap = frame[
        TIME_GAP
    ].to_numpy(
        dtype=float
    )

    distance = frame[
        DISTANCE
    ].to_numpy(
        dtype=float
    )

    starts = interval_start_times(
        frame
    )

    ends = frame[
        TIMESTAMP
    ]

    # If an earlier manual event already owns the telemetry interval
    # containing this timestamp, the present point is already represented in
    # the phase totals.  Searching farther forward would incorrectly attach it
    # to an unrelated later stop.
    for used_row in sorted(
        used_rows
    ):
        if not (
            0
            <= used_row
            < len(frame)
        ):
            continue

        if (
            pd.isna(
                starts.iloc[used_row]
            )
            or pd.isna(
                ends.iloc[used_row]
            )
        ):
            continue

        used_start = pd.Timestamp(
            starts.iloc[used_row]
        )
        used_end = pd.Timestamp(
            ends.iloc[used_row]
        )

        if (
            used_start
            <= manual_start
            < used_end
        ):
            return (
                EventMatch(
                    vehicle=event.vehicle,
                    event_kind=event.event_kind,
                    event_index=event.event_index,
                    manual_start=event.start,
                    manual_end=event.end,
                    manual_duration_seconds=event.duration_seconds,
                    manual_location=event.location,
                    raw_text=event.raw_text,
                    match_status="NOT_MATCHED",
                    match_method=(
                        "STOP_START_ALREADY_COVERED_BY_RESERVED_INTERVAL"
                    ),
                    telemetry_row_index=used_row,
                    excel_row=used_row + 2,
                    interval_start=used_start,
                    interval_end=used_end,
                    telemetry_gap_seconds=(
                        float(
                            gap[used_row]
                        )
                        if np.isfinite(
                            gap[used_row]
                        )
                        else None
                    ),
                    telemetry_distance_m=(
                        float(
                            distance[used_row]
                        )
                        if np.isfinite(
                            distance[used_row]
                        )
                        else None
                    ),
                    gap_equivalent_speed_kmh=None,
                    time_offset_seconds=0.0,
                    boundary_flag=bool(
                        boundary_mask(
                            len(frame)
                        )[
                            used_row
                        ]
                    ),
                    address_station_evidence=bool(
                        address_station_evidence[
                            used_row
                        ]
                    ),
                    notes=(
                        "This timestamp lies inside a telemetry interval "
                        "already assigned to an earlier manual event; it was "
                        "not shifted to an unrelated later interval."
                    ),
                ),
                empty_rows,
                None,
            )

    available = (
        positive
        & ~excluded_rows
    ).copy()

    for row in used_rows:
        if 0 <= row < len(
            available
        ):
            available[row] = False

    tolerance = float(
        MANUAL_EVENT_TOLERANCE_SECONDS
    )

    # Identify the telemetry interval aligned to the manual start before
    # searching for any later stationary episode.  This prevents a verified
    # 10:31 stop, for example, from being shifted to an unrelated 10:40 stop
    # merely because the immediate device gap contains some GPS movement.
    aligned_candidates: list[
        tuple[
            tuple[Any, ...],
            int,
        ]
    ] = []

    for row in np.flatnonzero(
        available
    ):
        row = int(row)

        if (
            pd.isna(
                starts.iloc[row]
            )
            or pd.isna(
                ends.iloc[row]
            )
        ):
            continue

        interval_start = pd.Timestamp(
            starts.iloc[row]
        )
        interval_end = pd.Timestamp(
            ends.iloc[row]
        )

        if interval_end <= manual_start:
            continue

        contains_start = (
            interval_start
            <= manual_start
            < interval_end
        )

        forward_offset = max(
            0.0,
            (
                interval_start
                - manual_start
            ).total_seconds(),
        )

        if (
            not contains_start
            and forward_offset
            > MANUAL_STOP_START_ALIGNMENT_SECONDS
        ):
            continue

        aligned_candidates.append(
            (
                (
                    0
                    if contains_start
                    else 1,
                    forward_offset,
                    -float(
                        gap[row]
                    ),
                ),
                row,
            )
        )

    aligned_row: int | None = None

    if aligned_candidates:
        aligned_candidates.sort(
            key=lambda item: item[0]
        )
        aligned_row = int(
            aligned_candidates[0][1]
        )

    # A stationary run can contain one long device gap or many short mobile
    # rows.  Keep only intervals whose END is strictly after the manual start;
    # an interval ending at the manual timestamp belongs to the pre-stop trip.
    ranked_runs: list[
        tuple[
            tuple[Any, ...],
            np.ndarray,
            str,
        ]
    ] = []

    for run in consecutive_true_runs(
        stationary_like
        & available
    ):
        forward_rows = np.array(
            [
                int(row)
                for row in run
                if (
                    pd.notna(
                        ends.iloc[
                            int(row)
                        ]
                    )
                    and pd.Timestamp(
                        ends.iloc[
                            int(row)
                        ]
                    )
                    > manual_start
                )
            ],
            dtype=int,
        )

        if len(
            forward_rows
        ) == 0:
            continue

        first = int(
            forward_rows[0]
        )
        last = int(
            forward_rows[-1]
        )

        if (
            pd.isna(
                starts.iloc[first]
            )
            or pd.isna(
                ends.iloc[last]
            )
        ):
            continue

        run_start = pd.Timestamp(
            starts.iloc[first]
        )
        run_end = pd.Timestamp(
            ends.iloc[last]
        )

        contains_start = (
            run_start
            <= manual_start
            < run_end
        )

        forward_offset = max(
            0.0,
            (
                run_start
                - manual_start
            ).total_seconds(),
        )

        if (
            not contains_start
            and forward_offset
            > tolerance
        ):
            continue

        observed_duration = float(
            np.nansum(
                gap[
                    forward_rows
                ]
            )
        )

        has_recorded_zero = bool(
            (
                zero_speed_episode_id[
                    forward_rows
                ]
                >= 0
            ).any()
        )

        if has_recorded_zero:
            method = (
                "MANUAL_STOP_START_TO_CONSECUTIVE_ZERO_SPEED_EPISODE"
            )
        elif len(
            forward_rows
        ) > 1:
            method = (
                "MANUAL_STOP_START_TO_CONSECUTIVE_STATIONARY_EPISODE"
            )
        else:
            method = (
                "MANUAL_STOP_START_TO_DEVICE_TIME_GAP"
            )

        score = (
            0
            if contains_start
            else 1,
            forward_offset,
            0
            if has_recorded_zero
            else 1,
            -observed_duration,
        )

        ranked_runs.append(
            (
                score,
                forward_rows,
                method,
            )
        )

    selected_rows = empty_rows
    selected_method = ""

    if ranked_runs:
        ranked_runs.sort(
            key=lambda item: item[0]
        )

    aligned_stationary_run = next(
        (
            item
            for item in ranked_runs
            if (
                aligned_row is not None
                and aligned_row in item[1]
            )
        ),
        None,
    )

    if aligned_stationary_run is not None:
        (
            _,
            selected_rows,
            selected_method,
        ) = aligned_stationary_run

    elif aligned_row is not None:
        selected_rows = np.array(
            [aligned_row],
            dtype=int,
        )
        selected_method = (
            "MANUAL_STOP_START_TO_NEXT_TELEMETRY_TIME_GAP"
        )

    elif ranked_runs:
        _, selected_rows, selected_method = ranked_runs[0]

    else:
        # No stationary episode was recorded.  This is common in device data:
        # while speed is zero the device emits nothing, so the stop is visible
        # only as the next large time gap.  Use the next forward interval and
        # retain an explicit low-evidence method in the audit when GPS suggests
        # movement.
        ranked_intervals: list[
            tuple[
                tuple[Any, ...],
                int,
            ]
        ] = []

        for row in np.flatnonzero(
            available
        ):
            row = int(row)

            if (
                pd.isna(
                    starts.iloc[row]
                )
                or pd.isna(
                    ends.iloc[row]
                )
            ):
                continue

            interval_start = pd.Timestamp(
                starts.iloc[row]
            )
            interval_end = pd.Timestamp(
                ends.iloc[row]
            )

            # Strictly forward: do not take the interval ending at the manual
            # timestamp, because that is the interval before the recorded stop.
            if interval_end <= manual_start:
                continue

            contains_start = (
                interval_start
                <= manual_start
                < interval_end
            )

            forward_offset = max(
                0.0,
                (
                    interval_start
                    - manual_start
                ).total_seconds(),
            )

            if (
                not contains_start
                and forward_offset
                > tolerance
            ):
                continue

            current_gap = (
                float(
                    gap[row]
                )
                if np.isfinite(
                    gap[row]
                )
                else 0.0
            )

            ranked_intervals.append(
                (
                    (
                        0
                        if contains_start
                        else 1,
                        forward_offset,
                        -current_gap,
                    ),
                    row,
                )
            )

        if ranked_intervals:
            ranked_intervals.sort(
                key=lambda item: item[0]
            )
            _, row = ranked_intervals[0]
            selected_rows = np.array(
                [row],
                dtype=int,
            )
            selected_method = (
                "MANUAL_STOP_START_TO_NEXT_TELEMETRY_TIME_GAP"
            )

    if len(
        selected_rows
    ) == 0:
        return (
            EventMatch(
                vehicle=event.vehicle,
                event_kind=event.event_kind,
                event_index=event.event_index,
                manual_start=event.start,
                manual_end=event.end,
                manual_duration_seconds=event.duration_seconds,
                manual_location=event.location,
                raw_text=event.raw_text,
                match_status="NOT_MATCHED",
                match_method="NO_FORWARD_TELEMETRY_WITHIN_15_MIN",
                telemetry_row_index=None,
                excel_row=None,
                interval_start=None,
                interval_end=None,
                telemetry_gap_seconds=None,
                telemetry_distance_m=None,
                gap_equivalent_speed_kmh=None,
                time_offset_seconds=None,
                boundary_flag=False,
                address_station_evidence=False,
                notes=(
                    "No telemetry interval after the manual stop start was "
                    "available within the matching tolerance."
                ),
            ),
            empty_rows,
            None,
        )

    first = int(
        selected_rows[0]
    )
    last = int(
        selected_rows[-1]
    )

    inferred_duration = float(
        np.nansum(
            gap[
                selected_rows
            ]
        )
    )

    inferred_distance = float(
        np.nansum(
            distance[
                selected_rows
            ]
        )
    )

    inferred_speed = (
        inferred_distance
        / inferred_duration
        * 3.6
        if inferred_duration > 0
        else None
    )

    selected_start = pd.Timestamp(
        starts.iloc[first]
    )
    selected_end = pd.Timestamp(
        ends.iloc[last]
    )

    contains_start = (
        selected_start
        <= manual_start
        < selected_end
    )

    time_offset = (
        0.0
        if contains_start
        else max(
            0.0,
            (
                selected_start
                - manual_start
            ).total_seconds(),
        )
    )

    has_stationary_evidence = bool(
        stationary_like[
            selected_rows
        ].any()
    )

    evidence_note = (
        "The complete forward stationary episode was used."
        if has_stationary_evidence
        else (
            "No zero-speed/stationary episode was recorded; the next device "
            "time gap was used because the manual timestamp is direct stop "
            "start evidence. Review this row if GPS displacement is large."
        )
    )

    return (
        EventMatch(
            vehicle=event.vehicle,
            event_kind=event.event_kind,
            event_index=event.event_index,
            manual_start=event.start,
            manual_end=None,
            manual_duration_seconds=None,
            manual_location=event.location,
            raw_text=event.raw_text,
            match_status="MATCHED",
            match_method=selected_method,
            telemetry_row_index=first,
            excel_row=first + 2,
            interval_start=selected_start,
            interval_end=selected_end,
            telemetry_gap_seconds=inferred_duration,
            telemetry_distance_m=inferred_distance,
            gap_equivalent_speed_kmh=inferred_speed,
            time_offset_seconds=time_offset,
            boundary_flag=bool(
                boundary_mask(
                    len(frame)
                )[
                    selected_rows
                ].any()
            ),
            address_station_evidence=bool(
                address_station_evidence[
                    selected_rows
                ].any()
            ),
            notes=(
                f"Manual timestamp treated as stop START; telemetry inferred "
                f"{inferred_duration:.1f} s across {len(selected_rows)} "
                f"interval(s). {evidence_note}"
            ),
        ),
        selected_rows,
        inferred_duration,
    )


# ============================================================================
# SEGMENTATION
# ============================================================================

def segment_vehicle(
    source: pd.DataFrame,
    vehicle: str,
    context: PairVehicleContext,
    manual_record: pd.Series | None,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    dict[str, Any],
]:
    frame = prepare_telemetry(
        source
    )

    route_category = clean_text(
        context.route_category
    ).upper()

    stop_threshold_seconds = applied_stop_threshold_seconds(
        route_category
    )

    # Dedicated intercity pipeline: route_category is forced to INTERCITY.
    is_city = False

    gap = frame[
        TIME_GAP
    ].to_numpy(
        dtype=float
    )

    distance = frame[
        DISTANCE
    ].to_numpy(
        dtype=float
    )

    device_speed = frame[
        SPEED_KMH
    ].to_numpy(
        dtype=float
    )

    acceleration = frame[
        ACCELERATION
    ].to_numpy(
        dtype=float
    )

    gap_speed = gap_equivalent_speed(
        gap,
        distance,
    )

    positive = (
        np.isfinite(gap)
        & (
            gap > 0
        )
    )

    gap_stationary_like = (
        positive
        & np.isfinite(
            gap_speed
        )
        & (
            gap_speed
            < STATIONARY_GAP_SPEED_KMH
        )
    )

    (
        zero_speed_episodes,
        zero_speed_episode_id,
        zero_speed_episode_duration,
    ) = recorded_zero_speed_episodes(
        positive=positive,
        gap_seconds=gap,
        device_speed_kmh=device_speed,
    )

    recorded_zero_stationary = (
        zero_speed_episode_id
        >= 0
    )

    # Route-specific stationary evidence:
    #   CITY      -> a complete consecutive recorded-zero-speed episode;
    #   INTERCITY -> preserve the earlier hybrid device-gap + zero-speed rule.
    # Manual stop/charging evidence is processed separately below and can
    # still locate a device gap even when CITY telemetry omitted zero rows.
    stationary_like = (
        recorded_zero_stationary
        if is_city
        else (
            gap_stationary_like
            | recorded_zero_stationary
        )
    )

    boundaries = boundary_mask(
        len(frame)
    )

    station_address = station_keyword_mask(
        frame,
        vehicle,
    )

    # Parse the dedicated manual charging/refuelling record before any
    # stationary phase is assigned.  This establishes one authoritative
    # replenishment source for the complete function and prevents a later
    # STOP pass from claiming the same event.
    replenishment_event = extract_manual_replenishment_event(
        manual_record,
        context,
    )

    manual_charge_location = manual_location_mask(
        frame,
        (
            replenishment_event.location
            if replenishment_event is not None
            else ""
        ),
    )

    # Station/operator text and a charging location from the manual record are
    # supporting evidence for matching a MANUALLY RECORDED replenishment event.
    # They are deliberately not allowed to create CHARGING/REFUELING on their own.
    charge_zone_evidence = (
        station_address
        | manual_charge_location
    )

    # Preserve the established sparse-device rule for INTERCITY only.  CITY's
    # new automatic rule explicitly requires consecutive Speed=0 telemetry,
    # so an ordinary CITY sampling gap is not converted into STOP.  Manual
    # CITY stop timestamps remain able to identify a missing-record gap.
    device_gap_long = (
        (not is_city)
        & positive
        & (
            gap
            > STOP_THRESHOLD_SECONDS
        )
    )

    phase = np.full(
        len(frame),
        "",
        dtype=object,
    )

    reason = np.full(
        len(frame),
        "",
        dtype=object,
    )

    confidence = np.full(
        len(frame),
        "",
        dtype=object,
    )

    manual_stop_support = np.zeros(
        len(frame),
        dtype=bool,
    )

    # If the manual record contains an explicit stop interval, keep its
    # duration as a hint for STOP-vs-IDLE classification.
    manual_stop_duration_hint = np.full(
        len(frame),
        np.nan,
        dtype=float,
    )

    # Identifies rows whose STOP/IDLE duration was inferred forward from a
    # start-only value in the manual ``stop timestamps`` column.
    manual_stop_start_inferred = np.zeros(
        len(frame),
        dtype=bool,
    )

    manual_stop_start_inferred_events = 0

    manual_replenishment_support = np.zeros(
        len(frame),
        dtype=bool,
    )

    manual_matches: list[
        EventMatch
    ] = []

    used_manual_rows: set[int] = set()

    # Manual-first / telemetry-fallback segmentation state.
    unsupported_boundary_stop = np.zeros(len(frame), dtype=bool)
    manual_source_conflict = np.zeros(len(frame), dtype=bool)
    zero_episode_replenishment = np.zeros(len(frame), dtype=bool)
    mobile_episode_stop = np.zeros(len(frame), dtype=bool)
    mobile_episode_idle = np.zeros(len(frame), dtype=bool)
    high_implied_speed_stop = np.zeros(len(frame), dtype=bool)

    replenishment_phase = (
        "CHARGING"
        if vehicle.upper() == "EV"
        else "REFUELING"
    )

    # Parse normal-stop evidence before charging selection, but DO NOT assign
    # STOP/IDLE yet. For duration-only charging records these timestamps are used
    # only as negative ranking evidence so a known normal stop is not preferred
    # over an otherwise plausible charging dwell. Charging still keeps phase
    # priority and is reserved first.
    stop_events = extract_manual_stop_events(
        manual_record,
        context,
    )

    # ------------------------------------------------------------------
    # 1. MANUAL CHARGING / REFUELLING — HIGHEST PRIORITY
    # ------------------------------------------------------------------
    # Manual charging/refuelling is the only source allowed to create these
    # phases.  The manual timestamp is an approximate anchor; when a manual
    # duration exists it is used as a TARGET to choose the best contiguous
    # stationary/very-low-speed telemetry window.  This prevents a 30-minute
    # charging event from being split across CHARGING + IDLE + INCH merely
    # because only its zero-speed core satisfied the stricter stationary rule.
    #
    # Rows are reserved before normal STOP/IDLE logic.  They can therefore never
    # be counted again as a stop or idle event.
    if replenishment_event is not None:
        (
            replenishment_match,
            selected_replenishment_rows,
        ) = select_manual_replenishment_rows(
            frame=frame,
            event=replenishment_event,
            positive=positive,
            stationary_like=stationary_like,
            device_gap_long=device_gap_long,
            device_speed_kmh=device_speed,
            gap_speed_kmh=gap_speed,
            address_station_evidence=station_address,
            used_rows=used_manual_rows,
            normal_stop_events=stop_events,
        )

        if len(selected_replenishment_rows) > 0:
            rows = np.asarray(
                selected_replenishment_rows,
                dtype=int,
            )

            manual_replenishment_support[rows] = True
            phase[rows] = replenishment_phase
            duration_only_replenishment = (
                replenishment_event.start is None
                and not replenishment_event.location
                and replenishment_event.duration_seconds is not None
                and math.isfinite(
                    replenishment_event.duration_seconds
                )
            )

            reason[rows] = (
                "MANUAL_REPLENISHMENT_DURATION_ONLY_TARGETED_LOW_MOVEMENT"
                if duration_only_replenishment
                else (
                    "MANUAL_REPLENISHMENT_DURATION_TARGETED_LOW_MOVEMENT"
                    if (
                        replenishment_event.duration_seconds is not None
                        and math.isfinite(
                            replenishment_event.duration_seconds
                        )
                    )
                    else "MANUAL_REPLENISHMENT_MATCHED_LOW_MOVEMENT_EPISODE"
                )
            )

            confidence[rows] = (
                "MEDIUM_MANUAL_DURATION_ONLY"
                if duration_only_replenishment
                else "HIGH_MANUAL_EVENT"
            )
            used_manual_rows.update(
                map(int, rows)
            )

        manual_matches.append(
            replenishment_match
        )

    # ------------------------------------------------------------------
    # 2. MANUAL NORMAL STOPS — AFTER REPLENISHMENT RESERVATION
    # ------------------------------------------------------------------
    for event in stop_events:
        if event.is_interval and event.start is not None and event.end is not None:
            overlap = manual_event_overlap_mask(frame, event)
            conflicts = overlap & manual_replenishment_support
            manual_source_conflict |= conflicts

            available_overlap = overlap.copy()
            if used_manual_rows:
                for used_row in used_manual_rows:
                    if 0 <= used_row < len(available_overlap):
                        available_overlap[used_row] = False

            if bool(available_overlap.any()):
                rows = np.flatnonzero(available_overlap).astype(int)
                observed_duration = float(np.nansum(gap[rows]))
                event_duration = (
                    float(event.duration_seconds)
                    if event.duration_seconds is not None
                    and math.isfinite(event.duration_seconds)
                    else observed_duration
                )
                stop_phase = (
                    "STOP"
                    if event_duration > stop_threshold_seconds
                    else "IDLE"
                )
                manual_stop_support[rows] = True
                manual_stop_duration_hint[rows] = event_duration
                phase[rows] = stop_phase
                reason[rows] = "MANUAL_EXPLICIT_STOP_INTERVAL"
                confidence[rows] = "HIGH_MANUAL_EVENT"
                used_manual_rows.update(map(int, rows))
                manual_matches.append(
                    event_match_from_rows(
                        frame=frame,
                        event=event,
                        rows=rows,
                        method="MANUAL_EXPLICIT_STOP_INTERVAL_OVERLAP",
                        station_evidence=station_address,
                        notes=(
                            f"Manual stop interval assigned as {stop_phase}; "
                            f"duration used={event_duration:.1f} s."
                        ),
                    )
                )
                continue

            # If the explicit interval does not overlap telemetry exactly, use
            # the existing +/-15-minute manual-event matcher as a controlled
            # fallback.
            match = match_event_to_stationary_gap(
                frame=frame,
                event=event,
                eligible_mask=(
                    positive
                    & ~manual_replenishment_support
                    & (stationary_like | device_gap_long)
                ),
                address_station_evidence=station_address,
                used_rows=used_manual_rows,
                allow_location_only=False,
                relaxed_mask=(
                    positive
                    & ~manual_replenishment_support
                    & np.isfinite(gap)
                    & (gap > IDLE_MIN_SECONDS)
                ),
            )
            manual_matches.append(match)

            if (
                match.match_status == "MATCHED"
                and match.telemetry_row_index is not None
            ):
                row = int(match.telemetry_row_index)
                event_duration = (
                    float(event.duration_seconds)
                    if event.duration_seconds is not None
                    and math.isfinite(event.duration_seconds)
                    else float(gap[row])
                )
                stop_phase = (
                    "STOP"
                    if event_duration > stop_threshold_seconds
                    else "IDLE"
                )
                manual_stop_support[row] = True
                manual_stop_duration_hint[row] = event_duration
                phase[row] = stop_phase
                reason[row] = "MANUAL_STOP_INTERVAL_MATCH"
                confidence[row] = "HIGH_MANUAL_EVENT"
                used_manual_rows.add(row)

        else:
            match, selected_rows, inferred_duration = infer_point_stop_from_telemetry(
                frame=frame,
                event=event,
                positive=positive,
                stationary_like=stationary_like,
                zero_speed_episode_id=zero_speed_episode_id,
                excluded_rows=manual_replenishment_support,
                used_rows=used_manual_rows,
                address_station_evidence=station_address,
            )
            manual_matches.append(match)

            if (
                match.match_status == "MATCHED"
                and len(selected_rows) > 0
                and inferred_duration is not None
            ):
                stop_phase = (
                    "STOP"
                    if inferred_duration > stop_threshold_seconds
                    else "IDLE"
                )
                manual_stop_support[selected_rows] = True
                manual_stop_start_inferred[selected_rows] = True
                manual_stop_duration_hint[selected_rows] = inferred_duration
                phase[selected_rows] = stop_phase
                reason[selected_rows] = "MANUAL_STOP_START_TELEMETRY_DURATION"
                confidence[selected_rows] = "HIGH_MANUAL_EVENT_INFERRED_DURATION"
                used_manual_rows.update(map(int, selected_rows))
                manual_stop_start_inferred_events += 1

    # ------------------------------------------------------------------
    # 3. AUTOMATIC RECORDED-ZERO EPISODES (MANUAL EVENTS ALREADY RESERVED)
    # ------------------------------------------------------------------
    # Any stationary event not already reserved by a manual charging/refuelling
    # event is STOP/IDLE only. Station-address text is audit/matching evidence and
    # must never turn an unlogged stationary event into CHARGING/REFUELING.
    for run in zero_speed_episodes:
        unassigned = run[phase[run] == ""]
        if len(unassigned) == 0:
            continue

        episode_time = float(zero_speed_episode_duration[run[0]])
        if not np.isfinite(episode_time):
            continue

        if episode_time > stop_threshold_seconds:
            phase[unassigned] = "STOP"
            reason[unassigned] = "ZERO_SPEED_EPISODE_DURATION>180S"
            confidence[unassigned] = "HIGH_TELEMETRY_STATIONARY"
            mobile_episode_stop[unassigned] = True
        else:
            phase[unassigned] = "IDLE"
            reason[unassigned] = "ZERO_SPEED_EPISODE_DURATION<=180S"
            confidence[unassigned] = "HIGH_TELEMETRY_STATIONARY"
            mobile_episode_idle[unassigned] = True

    # ------------------------------------------------------------------
    # 4. SPARSE-DEVICE LONG GAPS (>180 s)
    # ------------------------------------------------------------------
    remaining = positive & (phase == "")
    long_gap = remaining & device_gap_long

    unsupported_boundary_stop = long_gap & boundaries
    interior_long_stop = long_gap & ~boundaries

    # A long gap with no matching manual charging/refuelling event is STOP.
    # Boundary gaps remain explicitly downgraded for review.
    phase[unsupported_boundary_stop] = "STOP"
    reason[unsupported_boundary_stop] = "BOUNDARY_STATIONARY_GAP_FALLBACK_STOP"
    confidence[unsupported_boundary_stop] = "VERY_LOW_BOUNDARY_FALLBACK"

    phase[interior_long_stop] = "STOP"
    reason[interior_long_stop] = "INTERCITY_DEVICE_GAP>180S"
    confidence[interior_long_stop] = "HIGH_TELEMETRY_TIME_GAP"

    high_implied_speed_stop = (
        interior_long_stop
        & np.isfinite(gap_speed)
        & (gap_speed >= STATIONARY_GAP_SPEED_KMH)
    )
    confidence[high_implied_speed_stop] = "LOW_LONG_GAP_HIGH_IMPLIED_SPEED"
    reason[high_implied_speed_stop] = (
        "INTERCITY_DEVICE_GAP>180S + HIGH_IMPLIED_SPEED_REVIEW"
    )

    # ------------------------------------------------------------------
    # 5. REMAINING STATIONARY-LIKE RUNS
    # ------------------------------------------------------------------
    # Rule-based fallback for events absent from the manual sheet.
    remaining_stationary = positive & (phase == "") & stationary_like

    for run in consecutive_true_runs(remaining_stationary):
        episode_time = float(np.nansum(gap[run]))

        if episode_time > stop_threshold_seconds:
            phase[run] = "STOP"
            reason[run] = "STATIONARY_RUN_DURATION>180S"
            confidence[run] = "HIGH_TELEMETRY_STATIONARY"
        else:
            phase[run] = "IDLE"
            reason[run] = "STATIONARY_RUN_DURATION<=180S"
            confidence[run] = "HIGH_TELEMETRY_STATIONARY"

    # ------------------------------------------------------------------
    # 7. PRIMARY INCH FROM DEVICE SPEED
    # ------------------------------------------------------------------
    remaining = (
        positive
        & (
            phase == ""
        )
    )

    inch = (
        remaining
        & np.isfinite(
            device_speed
        )
        & (
            device_speed
            >= INCH_MIN_KMH
        )
        & (
            device_speed
            <= INCH_MAX_KMH
        )
    )

    phase[
        inch
    ] = "INCH"

    reason[
        inch
    ] = (
        "1<=DEVICE_SPEED<=5"
    )

    confidence[
        inch
    ] = "HIGH"

    # ------------------------------------------------------------------
    # 8. PRIMARY ACC / DEC / CRUISE FROM DEVICE SPEED
    # ------------------------------------------------------------------
    remaining = (
        positive
        & (
            phase == ""
        )
    )

    normal_motion = (
        remaining
        & np.isfinite(
            device_speed
        )
        & (
            device_speed
            > INCH_MAX_KMH
        )
        & np.isfinite(
            acceleration
        )
    )

    acc_mask = (
        normal_motion
        & (
            acceleration
            > ACCELERATION_THRESHOLD
        )
    )

    dec_mask = (
        normal_motion
        & (
            acceleration
            < -ACCELERATION_THRESHOLD
        )
    )

    cruise_mask = (
        normal_motion
        & (
            acceleration
            >= -ACCELERATION_THRESHOLD
        )
        & (
            acceleration
            <= ACCELERATION_THRESHOLD
        )
    )

    phase[
        acc_mask
    ] = "ACC"

    phase[
        dec_mask
    ] = "DEC"

    phase[
        cruise_mask
    ] = "CRUISE"

    reason[
        acc_mask
    ] = (
        "DEVICE_SPEED>5 + ACC>0.14"
    )

    reason[
        dec_mask
    ] = (
        "DEVICE_SPEED>5 + ACC<-0.14"
    )

    reason[
        cruise_mask
    ] = (
        "DEVICE_SPEED>5 + -0.14<=ACC<=0.14"
    )

    confidence[
        acc_mask
        | dec_mask
        | cruise_mask
    ] = "HIGH"

    # If the moving record itself has an irregular time gap, acceleration
    # derived over that long interval is less reliable.
    irregular_moving = (
        (
            acc_mask
            | dec_mask
            | cruise_mask
        )
        & (
            gap
            > IDLE_MIN_SECONDS
        )
        & ~stationary_like
    )

    confidence[
        irregular_moving
    ] = "LOW_IRREGULAR_MOVING_GAP"

    # ------------------------------------------------------------------
    # 9. MOVEMENT FALLBACK WHEN DEVICE SPEED IS ZERO / MISSING
    #
    # The old errors are mainly rows where device Speed = 0 but GPS
    # displacement clearly shows movement. In those rows the zero device
    # speed cannot be used as the moving-speed gate.
    #
    # Use Gap Equivalent Speed ONLY when:
    #   - the row is still unclassified;
    #   - GPS movement is >=1 km/h equivalent;
    #   - device speed is <1 km/h or missing.
    #
    # This fallback does not change the Speed column.
    # ------------------------------------------------------------------
    remaining = (
        positive
        & (
            phase == ""
        )
    )

    device_speed_unreliable = (
        ~np.isfinite(
            device_speed
        )
        | (
            device_speed
            < INCH_MIN_KMH
        )
    )

    fallback_moving = (
        remaining
        & device_speed_unreliable
        & np.isfinite(
            gap_speed
        )
        & (
            gap_speed
            >= INCH_MIN_KMH
        )
    )

    fallback_inch = (
        fallback_moving
        & (
            gap_speed
            <= INCH_MAX_KMH
        )
    )

    phase[
        fallback_inch
    ] = "INCH"

    reason[
        fallback_inch
    ] = (
        "DEVICE_SPEED<1_OR_MISSING + 1<=GAP_EQ_SPEED<=5"
    )

    confidence[
        fallback_inch
    ] = (
        "LOW_GPS_MOVEMENT_FALLBACK"
    )

    remaining = (
        positive
        & (
            phase == ""
        )
    )

    fallback_fast = (
        remaining
        & device_speed_unreliable
        & np.isfinite(
            gap_speed
        )
        & (
            gap_speed
            > INCH_MAX_KMH
        )
    )

    fallback_acc = (
        fallback_fast
        & np.isfinite(
            acceleration
        )
        & (
            acceleration
            > ACCELERATION_THRESHOLD
        )
    )

    fallback_dec = (
        fallback_fast
        & np.isfinite(
            acceleration
        )
        & (
            acceleration
            < -ACCELERATION_THRESHOLD
        )
    )

    fallback_cruise = (
        fallback_fast
        & np.isfinite(
            acceleration
        )
        & (
            acceleration
            >= -ACCELERATION_THRESHOLD
        )
        & (
            acceleration
            <= ACCELERATION_THRESHOLD
        )
    )

    phase[
        fallback_acc
    ] = "ACC"

    phase[
        fallback_dec
    ] = "DEC"

    phase[
        fallback_cruise
    ] = "CRUISE"

    reason[
        fallback_acc
    ] = (
        "GAP_EQ_SPEED>5_FALLBACK + ACC>0.14"
    )

    reason[
        fallback_dec
    ] = (
        "GAP_EQ_SPEED>5_FALLBACK + ACC<-0.14"
    )

    reason[
        fallback_cruise
    ] = (
        "GAP_EQ_SPEED>5_FALLBACK + -0.14<=ACC<=0.14"
    )

    confidence[
        fallback_acc
        | fallback_dec
        | fallback_cruise
    ] = (
        "LOW_GPS_MOVEMENT_FALLBACK"
    )

    # Clearly moving but acceleration is missing: keep it in the moving
    # family as CRUISE, with explicit low confidence.
    remaining = (
        positive
        & (
            phase == ""
        )
    )

    fallback_missing_acc = (
        remaining
        & np.isfinite(
            gap_speed
        )
        & (
            gap_speed
            > INCH_MAX_KMH
        )
        & ~np.isfinite(
            acceleration
        )
    )

    phase[
        fallback_missing_acc
    ] = "CRUISE"

    reason[
        fallback_missing_acc
    ] = (
        "GAP_EQ_SPEED>5 + ACC_MISSING_FALLBACK"
    )

    confidence[
        fallback_missing_acc
    ] = (
        "VERY_LOW_MOVING_FALLBACK"
    )

    # ------------------------------------------------------------------
    # 10. EXHAUSTIVE FINAL FALLBACK
    #
    # TAT computation requires all positive-duration rows to have one of
    # the seven phases. Anything surviving this far has insufficient
    # telemetry evidence.
    #
    # Before this point, stationary-like rows have already been assigned
    # STOP/IDLE. Therefore anything remaining here is treated as a moving-
    # data anomaly. INCH is used only as the final low-speed MOVEMENT
    # fallback, never as a stationary phase.
    # ------------------------------------------------------------------
    remaining = (
        positive
        & (
            phase == ""
        )
    )

    # Do not leave an unsupported boundary long gap blank, because that would
    # break TAT computation. Keep it as STOP but mark it very low confidence.
    boundary_fallback_stop = (
        remaining
        & unsupported_boundary_stop
    )

    phase[
        boundary_fallback_stop
    ] = "STOP"

    reason[
        boundary_fallback_stop
    ] = (
        "BOUNDARY_STATIONARY_GAP_FALLBACK_STOP"
    )

    confidence[
        boundary_fallback_stop
    ] = (
        "VERY_LOW_BOUNDARY_FALLBACK"
    )

    remaining = (
        positive
        & (
            phase == ""
        )
    )

    phase[
        remaining
    ] = "INCH"

    reason[
        remaining
    ] = (
        "FINAL_EXHAUSTIVE_FALLBACK_TO_INCH"
    )

    confidence[
        remaining
    ] = (
        "VERY_LOW_FINAL_FALLBACK"
    )

    # ------------------------------------------------------------------
    # 11. First row / zero-duration rows
    # ------------------------------------------------------------------
    phase[
        ~positive
    ] = ""

    reason[
        ~positive
    ] = (
        "NO_POSITIVE_INTERVAL"
    )

    confidence[
        ~positive
    ] = "N/A"

    # Hard segregation invariant: a manually reserved charging/refuelling row
    # must never be reused by normal STOP/IDLE logic.  Fail loudly if a future
    # edit breaks this precedence instead of silently adding charging time into
    # stop time.
    overlap_manual_stationary = (
        manual_replenishment_support
        & manual_stop_support
    )

    if bool(overlap_manual_stationary.any()):
        rows = (
            np.flatnonzero(overlap_manual_stationary)[:20]
            + 2
        ).tolist()
        raise RuntimeError(
            f"{vehicle} manual charging/refuelling rows were also assigned "
            f"as manual STOP/IDLE rows. Excel rows: {rows}"
        )

    wrong_replenishment_phase = (
        manual_replenishment_support
        & (phase != replenishment_phase)
    )

    if bool(wrong_replenishment_phase.any()):
        rows = (
            np.flatnonzero(wrong_replenishment_phase)[:20]
            + 2
        ).tolist()
        raise RuntimeError(
            f"{vehicle} reserved charging/refuelling rows lost their phase "
            f"priority. Excel rows: {rows}"
        )

    # ------------------------------------------------------------------
    # 12. Classification audit
    # ------------------------------------------------------------------
    interval_starts = interval_start_times(
        frame
    )

    audit_rows: list[
        dict[str, Any]
    ] = []

    for index in range(
        len(frame)
    ):
        if not positive[index]:
            continue

        special = (
            phase[index] == ""
            or clean_text(
                confidence[
                    index
                ]
            ).startswith(
                "LOW"
            )
            or clean_text(
                confidence[
                    index
                ]
            ).startswith(
                "VERY_LOW"
            )
            or bool(
                manual_stop_support[
                    index
                ]
            )
            or bool(
                manual_replenishment_support[
                    index
                ]
            )
            or bool(
                charge_zone_evidence[
                    index
                ]
            )
            or bool(
                recorded_zero_stationary[
                    index
                ]
            )
            or bool(
                boundaries[
                    index
                ]
                and gap[index]
                > stop_threshold_seconds
            )
            or bool(
                gap[index]
                > IDLE_MIN_SECONDS
            )
        )

        if not special:
            continue

        scenario = ""

        if manual_stop_start_inferred[index]:
            scenario = (
                "MANUAL_STOP_START_TELEMETRY_DURATION"
            )

        elif recorded_zero_stationary[index]:
            scenario = (
                "CUMULATIVE_RECORDED_ZERO_SPEED_EPISODE"
            )

        elif confidence[index] == "LOW_SHORT_STATIONARY_FALLBACK_IDLE":
            scenario = (
                "SHORT_STATIONARY_FALLBACK_IDLE"
            )

        elif confidence[index] == "LOW_LONG_GAP_HIGH_IMPLIED_SPEED":
            scenario = (
                "LONG_DEVICE_GAP_HIGH_IMPLIED_SPEED_REVIEW"
            )

        elif confidence[index] == "LOW_GPS_MOVEMENT_FALLBACK":
            scenario = (
                "DEVICE_SPEED_ZERO_OR_MISSING_BUT_GPS_MOVING"
            )

        elif confidence[index] == "VERY_LOW_MOVING_FALLBACK":
            scenario = (
                "MOVING_WITH_MISSING_ACCELERATION"
            )

        elif confidence[index] == "VERY_LOW_BOUNDARY_FALLBACK":
            scenario = (
                "BOUNDARY_DATA_GAP_FALLBACK"
            )

        elif confidence[index] == "VERY_LOW_FINAL_FALLBACK":
            scenario = (
                "FINAL_EXHAUSTIVE_FALLBACK"
            )

        elif irregular_moving[
            index
        ]:
            scenario = (
                "IRREGULAR_MOVING_GAP"
            )

        elif (
            phase[index]
            in {
                "STOP",
                "IDLE",
                "CHARGING",
                "REFUELING",
            }
        ):
            scenario = (
                "STATIONARY_LIKE_INTERVAL"
            )

        audit_rows.append(
            {
                "Vehicle": vehicle,
                "Route Category": route_category,
                "Applied Stop Threshold (s)": stop_threshold_seconds,
                "Excel Row": index + 2,
                "Timestamp": frame.at[
                    index,
                    TIMESTAMP,
                ],
                "Interval Start": interval_starts.iloc[
                    index
                ],
                "Time Gap (s)": (
                    float(
                        gap[index]
                    )
                    if np.isfinite(
                        gap[index]
                    )
                    else None
                ),
                "Distance (m)": (
                    float(
                        distance[index]
                    )
                    if np.isfinite(
                        distance[index]
                    )
                    else None
                ),
                "Gap Equivalent Speed (km/h)": (
                    float(
                        gap_speed[index]
                    )
                    if np.isfinite(
                        gap_speed[index]
                    )
                    else None
                ),
                "Device Speed (km/h)": (
                    float(
                        device_speed[
                            index
                        ]
                    )
                    if np.isfinite(
                        device_speed[
                            index
                        ]
                    )
                    else None
                ),
                "Recorded Zero-Speed Episode": bool(
                    recorded_zero_stationary[
                        index
                    ]
                ),
                "Zero-Speed Episode ID": (
                    int(
                        zero_speed_episode_id[
                            index
                        ]
                    )
                    if recorded_zero_stationary[
                        index
                    ]
                    else None
                ),
                "Zero-Speed Episode Duration (s)": (
                    float(
                        zero_speed_episode_duration[
                            index
                        ]
                    )
                    if np.isfinite(
                        zero_speed_episode_duration[
                            index
                        ]
                    )
                    else None
                ),
                "Acceleration (m/s^2)": (
                    float(
                        acceleration[
                            index
                        ]
                    )
                    if np.isfinite(
                        acceleration[
                            index
                        ]
                    )
                    else None
                ),
                "Phase": clean_text(
                    phase[index]
                ),
                "Classification Reason": clean_text(
                    reason[index]
                ),
                "Confidence": clean_text(
                    confidence[index]
                ),
                "Manual Stop Support": bool(
                    manual_stop_support[
                        index
                    ]
                ),
                "Manual Stop Start-Only Inference": bool(
                    manual_stop_start_inferred[
                        index
                    ]
                ),
                "Manual Stop Duration Used (s)": (
                    float(
                        manual_stop_duration_hint[
                            index
                        ]
                    )
                    if (
                        manual_stop_support[
                            index
                        ]
                        and np.isfinite(
                            manual_stop_duration_hint[
                                index
                            ]
                        )
                    )
                    else None
                ),
                "Manual Charging/Refuelling Support": bool(
                    manual_replenishment_support[
                        index
                    ]
                ),
                "Charging-Stop Conflict Resolved": bool(
                    manual_source_conflict[
                        index
                    ]
                ),
                "Station Address Evidence": bool(
                    station_address[
                        index
                    ]
                ),
                "Manual Charge Location Evidence": bool(
                    manual_charge_location[
                        index
                    ]
                ),
                "Charge Zone Evidence": bool(
                    charge_zone_evidence[
                        index
                    ]
                ),
                "Boundary Gap": bool(
                    boundaries[
                        index
                    ]
                ),
                "Very Long Gap": bool(
                    gap[index]
                    > VERY_LONG_GAP_SECONDS
                ),
                "Scenario": scenario,
                "Address": frame.at[
                    index,
                    ADDRESS,
                ],
                "Reverse Geocoded Address": frame.at[
                    index,
                    REVERSE_ADDRESS,
                ],
            }
        )

    # ------------------------------------------------------------------
    # 10. Manual event audit
    # ------------------------------------------------------------------
    manual_event_frame = pd.DataFrame(
        [
            asdict(
                match
            )
            for match in manual_matches
        ]
    )

    if not manual_event_frame.empty:
        manual_event_frame = manual_event_frame.rename(
            columns={
                "vehicle": "Vehicle",
                "event_kind": "Event Type",
                "event_index": "Event Index",
                "manual_start": "Manual Start",
                "manual_end": "Manual End",
                "manual_duration_seconds": "Manual Duration (s)",
                "manual_location": "Manual Location",
                "raw_text": "Raw Manual Text",
                "match_status": "Match Status",
                "match_method": "Match Method",
                "telemetry_row_index": "Telemetry Row Index",
                "excel_row": "Excel Row",
                "interval_start": "Telemetry Interval Start",
                "interval_end": "Telemetry Interval End",
                "telemetry_gap_seconds": "Telemetry Gap (s)",
                "telemetry_distance_m": "Telemetry Distance (m)",
                "gap_equivalent_speed_kmh": "Gap Equivalent Speed (km/h)",
                "time_offset_seconds": "Time Offset (s)",
                "boundary_flag": "Boundary Gap",
                "address_station_evidence": "Station Address Evidence",
                "notes": "Notes",
            }
        )

    # Final replenishment mask after manual-first matching plus telemetry
    # fallback classification.
    replenishment = phase == replenishment_phase
    zero_episode_replenishment = (
        replenishment
        & recorded_zero_stationary
    )

    # ------------------------------------------------------------------
    # 11. Charging/refuelling details
    # ------------------------------------------------------------------
    charging_details: dict[
        str,
        Any,
    ] = {
        "Vehicle": vehicle,
        "Route Category": route_category,
        "Applied Stop Threshold (s)": stop_threshold_seconds,
        "Manual Trip Match": (
            "MATCHED"
            if manual_record is not None
            else "NOT MATCHED"
        ),
        "Source Date": context.source_date,
        "Source File": context.source_file_name,
        "Route Origin": context.route_origin,
        "Route Destination": context.route_destination,
        "Manual Log Excel Row": (
            int(
                manual_record.get(
                    "_source_excel_row"
                )
            )
            if manual_record is not None
            and pd.notna(
                manual_record.get(
                    "_source_excel_row"
                )
            )
            else None
        ),
        "Raw Charging/Refuelling Time": (
            replenishment_event.raw_text
            if replenishment_event is not None
            else ""
        ),
        "Manual Charging Time (min) Field": (
            clean_text(
                manual_record.get(
                    "Charging Time"
                )
            )
            if manual_record is not None
            else ""
        ),
        "Manual Charging Timestamp Field": (
            clean_text(
                manual_record.get(
                    "Charging Timestamp"
                )
            )
            if manual_record is not None
            else ""
        ),
        "Manual Stop Timestamp Field": (
            clean_text(
                manual_record.get(
                    "Stop Timestamps"
                )
            )
            if manual_record is not None
            else ""
        ),
        "Parsed Charging/Refuelling Start": (
            replenishment_event.start
            if replenishment_event is not None
            else None
        ),
        "Manual Duration (min)": (
            replenishment_event.duration_seconds
            / 60.0
            if replenishment_event is not None
            and replenishment_event.duration_seconds
            is not None
            else None
        ),
        "Parsed Charging/Refuelling End": (
            replenishment_event.end
            if replenishment_event is not None
            else None
        ),
        "Manual Location": (
            replenishment_event.location
            if replenishment_event is not None
            else ""
        ),
        "Address Station Evidence Anywhere": bool(
            (
                (
                    device_gap_long
                    & charge_zone_evidence
                )
                | zero_episode_replenishment
            ).any()
        ),
        "Manual Charge Location Evidence Anywhere": bool(
            manual_charge_location.any()
        ),
        "Charge Zone Evidence Anywhere": bool(
            charge_zone_evidence.any()
        ),
        "Final Charging/Refuelling Rows": int(
            replenishment.sum()
        ),
        "Observed Charging/Refuelling Gap Time (min)": float(
            gap[
                replenishment
            ].sum()
            / 60.0
        ),
        "Recommended Replenishment Time for TAT (min)": (
            replenishment_event.duration_seconds
            / 60.0
            if (
                replenishment_event is not None
                and replenishment_event.duration_seconds
                is not None
                and replenishment.any()
            )
            else float(
                gap[
                    replenishment
                ].sum()
                / 60.0
            )
        ),
        "Recommended Time Source": (
            "MANUAL_DURATION"
            if (
                replenishment_event is not None
                and replenishment_event.duration_seconds
                is not None
                and replenishment.any()
            )
            else (
                "TELEMETRY_GAP"
                if replenishment.any()
                else "NO_REPLENISHMENT_EVENT"
            )
        ),
    }

    # ------------------------------------------------------------------
    # 14. Exhaustiveness validation
    # ------------------------------------------------------------------
    unresolved_positive = (
        positive
        & (
            phase == ""
        )
    )

    if unresolved_positive.any():
        rows = (
            np.flatnonzero(
                unresolved_positive
            )[:20]
            + 2
        ).tolist()

        raise RuntimeError(
            f"{vehicle} still has positive-duration rows without phase "
            f"after exhaustive fallback. Excel rows: {rows}"
        )

    # ------------------------------------------------------------------
    # 15. Metrics
    # ------------------------------------------------------------------
    metrics = {
        "route_category": route_category,
        "applied_stop_threshold_seconds": stop_threshold_seconds,
        "rows": len(
            frame
        ),
        "positive_interval_rows": int(
            positive.sum()
        ),
        "stationary_like_rows": int(
            stationary_like.sum()
        ),
        "gap_stationary_rows": int(
            gap_stationary_like.sum()
        ),
        "recorded_zero_speed_episodes": int(
            len(
                zero_speed_episodes
            )
        ),
        "recorded_zero_speed_episode_rows": int(
            recorded_zero_stationary.sum()
        ),
        "recorded_zero_speed_stop_rows": int(
            mobile_episode_stop.sum()
        ),
        "recorded_zero_speed_idle_rows": int(
            mobile_episode_idle.sum()
        ),
        "manual_stop_events": len(
            stop_events
        ),
        "manual_stop_matches": int(
            manual_stop_support.sum()
        ),
        "manual_stop_start_inferred_events": int(
            manual_stop_start_inferred_events
        ),
        "manual_stop_start_inferred_rows": int(
            manual_stop_start_inferred.sum()
        ),
        "manual_replenishment_matches": int(
            manual_replenishment_support.sum()
        ),
        "charge_zone_rows": int(
            charge_zone_evidence.sum()
        ),
        "manual_charge_location_rows": int(
            manual_charge_location.sum()
        ),
        "high_implied_speed_long_gap_stop_rows": int(
            high_implied_speed_stop.sum()
        ),
        "manual_charging_stop_conflicts_resolved": int(
            manual_source_conflict.sum()
        ),
        "rule_gap_rows": int(
            (
                positive
                & (
                    phase == ""
                )
            ).sum()
        ),
        "low_confidence_fallback_rows": int(
            np.char.startswith(
                confidence.astype(str),
                "LOW"
            ).sum()
            + np.char.startswith(
                confidence.astype(str),
                "VERY_LOW"
            ).sum()
        ),
    }

    result = frame.copy()

    # Important: only this column will be written back to the original
    # telemetry worksheets.
    result[
        PHASE
    ] = phase

    return (
        result,
        pd.DataFrame(
            audit_rows
        ),
        manual_event_frame,
        pd.DataFrame(
            [
                charging_details
            ]
        ),
        metrics,
    )


# ============================================================================
# PHASE SUMMARY
# ============================================================================

def build_phase_summary(
    ev_result: pd.DataFrame,
    ice_result: pd.DataFrame,
    charging_details: pd.DataFrame,
) -> pd.DataFrame:
    rows: list[
        dict[str, Any]
    ] = []

    order = {
        "EV": [
            "CHARGING",
            "STOP",
            "IDLE",
            "INCH",
            "ACC",
            "DEC",
            "CRUISE",
        ],
        "ICE": [
            "REFUELING",
            "STOP",
            "IDLE",
            "INCH",
            "ACC",
            "DEC",
            "CRUISE",
        ],
    }

    for vehicle, frame in (
        ("EV", ev_result),
        ("ICE", ice_result),
    ):
        gap = pd.to_numeric(
            frame[TIME_GAP],
            errors="coerce",
        ).fillna(
            0.0
        )

        distance = pd.to_numeric(
            frame[DISTANCE],
            errors="coerce",
        ).fillna(
            0.0
        )

        total_time = float(
            gap[
                gap > 0
            ].sum()
        )

        for phase_name in order[
            vehicle
        ]:
            mask = (
                frame[
                    PHASE
                ]
                == phase_name
            )

            phase_time = float(
                gap[
                    mask
                ].sum()
            )

            phase_distance = float(
                distance[
                    mask
                ].sum()
            )

            rows.append(
                {
                    "Vehicle": vehicle,
                    "Phase": phase_name,
                    "Telemetry Time (s)": phase_time,
                    "Telemetry Time (min)": phase_time
                    / 60.0,
                    "Distance (m)": phase_distance,
                    "Distance (km)": phase_distance
                    / 1000.0,
                    "TAT Contribution (%)": (
                        phase_time
                        / total_time
                        * 100.0
                        if total_time
                        else 0.0
                    ),
                }
            )

        rule_gap = (
            (
                gap > 0
            )
            & (
                frame[
                    PHASE
                ]
                == ""
            )
        )

        rule_gap_time = float(
            gap[
                rule_gap
            ].sum()
        )

        rule_gap_distance = float(
            distance[
                rule_gap
            ].sum()
        )

        rows.append(
            {
                "Vehicle": vehicle,
                "Phase": "RULE_GAP",
                "Telemetry Time (s)": rule_gap_time,
                "Telemetry Time (min)": rule_gap_time
                / 60.0,
                "Distance (m)": rule_gap_distance,
                "Distance (km)": rule_gap_distance
                / 1000.0,
                "TAT Contribution (%)": (
                    rule_gap_time
                    / total_time
                    * 100.0
                    if total_time
                    else 0.0
                ),
            }
        )

        rows.append(
            {
                "Vehicle": vehicle,
                "Phase": "TOTAL",
                "Telemetry Time (s)": total_time,
                "Telemetry Time (min)": total_time
                / 60.0,
                "Distance (m)": float(
                    distance[
                        gap > 0
                    ].sum()
                ),
                "Distance (km)": float(
                    distance[
                        gap > 0
                    ].sum()
                )
                / 1000.0,
                "TAT Contribution (%)": 100.0
                if total_time
                else 0.0,
            }
        )

    # Add manual charging/refuelling duration as an explicit reference row,
    # not as a silent replacement of telemetry interval time. Build the final
    # DataFrame only once so pandas never concatenates empty/all-NA columns.
    if not charging_details.empty:
        for _, detail_row in charging_details.iterrows():
            rows.append(
                {
                    "Vehicle": detail_row.get("Vehicle"),
                    "Phase": (
                        "MANUAL_CHARGING_TIME_REFERENCE"
                        if detail_row.get("Vehicle") == "EV"
                        else "MANUAL_REFUELING_TIME_REFERENCE"
                    ),
                    "Telemetry Time (s)": None,
                    "Telemetry Time (min)": detail_row.get(
                        "Recommended Replenishment Time for TAT (min)"
                    ),
                    "Distance (m)": None,
                    "Distance (km)": None,
                    "TAT Contribution (%)": None,
                }
            )

    return pd.DataFrame(rows)


# ============================================================================
# EXCEL OUTPUT HELPERS
# ============================================================================

HEADER_FILL = PatternFill(
    "solid",
    fgColor="1F4E78",
)

HEADER_FONT = Font(
    color="FFFFFF",
    bold=True,
)


def write_phase_column(
    worksheet,
    phases: pd.Series,
) -> None:
    phase_column = None

    for cell in worksheet[1]:
        if normalise_text(
            cell.value
        ) == normalise_text(
            PHASE
        ):
            phase_column = cell.column
            break

    if phase_column is None:
        phase_column = (
            worksheet.max_column
            + 1
        )

        header = worksheet.cell(
            row=1,
            column=phase_column,
            value=PHASE,
        )

        header.fill = HEADER_FILL
        header.font = HEADER_FONT

    # Clear old phase values.
    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):
        worksheet.cell(
            row=row_number,
            column=phase_column,
        ).value = None

    for row_number, value in enumerate(
        phases.tolist(),
        start=2,
    ):
        worksheet.cell(
            row=row_number,
            column=phase_column,
            value=(
                clean_text(
                    value
                )
                or None
            ),
        )


def replace_dataframe_sheet(
    workbook,
    sheet_name: str,
    frame: pd.DataFrame,
) -> None:
    if sheet_name in workbook.sheetnames:
        workbook.remove(
            workbook[
                sheet_name
            ]
        )

    worksheet = workbook.create_sheet(
        sheet_name
    )

    if frame.empty:
        worksheet[
            "A1"
        ] = "No records"
        return

    for column_index, column in enumerate(
        frame.columns,
        start=1,
    ):
        cell = worksheet.cell(
            row=1,
            column=column_index,
            value=str(
                column
            ),
        )

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    for row_index, values in enumerate(
        frame.itertuples(
            index=False,
            name=None,
        ),
        start=2,
    ):
        for column_index, value in enumerate(
            values,
            start=1,
        ):
            if isinstance(
                value,
                pd.Timestamp,
            ):
                value = value.to_pydatetime()

            try:
                if pd.isna(value):
                    value = None
            except (TypeError, ValueError):
                pass

            worksheet.cell(
                row=row_index,
                column=column_index,
                value=value,
            )

    worksheet.freeze_panes = "A2"

    # Conservative readable widths.
    for column_index, column in enumerate(
        frame.columns,
        start=1,
    ):
        maximum = len(
            str(
                column
            )
        )

        for value in frame[
            column
        ].head(
            100
        ):
            maximum = max(
                maximum,
                len(
                    clean_text(
                        value
                    )
                ),
            )

        worksheet.column_dimensions[
            get_column_letter(
                column_index
            )
        ].width = min(
            45,
            max(
                12,
                maximum
                + 2,
            ),
        )


def save_pair_workbook(
    workbook_path: Path,
    ev_result: pd.DataFrame,
    ice_result: pd.DataFrame,
    summary: pd.DataFrame,
    classification_audit: pd.DataFrame,
    manual_event_audit: pd.DataFrame,
    charging_details: pd.DataFrame,
) -> None:
    workbook = load_workbook(
        workbook_path,
        data_only=False,
    )

    temporary_path = None

    try:
        if EV_SHEET not in workbook.sheetnames:
            raise ValueError(
                f"{workbook_path.name}: "
                f"missing {EV_SHEET}"
            )

        if ICE_SHEET not in workbook.sheetnames:
            raise ValueError(
                f"{workbook_path.name}: "
                f"missing {ICE_SHEET}"
            )

        # Only phase is changed in original telemetry sheets.
        write_phase_column(
            workbook[
                EV_SHEET
            ],
            ev_result[
                PHASE
            ],
        )

        write_phase_column(
            workbook[
                ICE_SHEET
            ],
            ice_result[
                PHASE
            ],
        )

        replace_dataframe_sheet(
            workbook,
            PHASE_SUMMARY_SHEET,
            summary,
        )

        replace_dataframe_sheet(
            workbook,
            CLASSIFICATION_AUDIT_SHEET,
            classification_audit,
        )

        replace_dataframe_sheet(
            workbook,
            MANUAL_EVENT_AUDIT_SHEET,
            manual_event_audit,
        )

        replace_dataframe_sheet(
            workbook,
            CHARGING_DETAILS_SHEET,
            charging_details,
        )

        with NamedTemporaryFile(
            prefix=(
                f".{workbook_path.stem}"
                "_phase_"
            ),
            suffix=".xlsx",
            dir=workbook_path.parent,
            delete=False,
        ) as handle:
            temporary_path = Path(
                handle.name
            )

        workbook.save(
            temporary_path
        )

        workbook.close()
        workbook = None

        os.replace(
            temporary_path,
            workbook_path,
        )

    finally:
        if workbook is not None:
            workbook.close()

        if (
            temporary_path is not None
            and temporary_path.exists()
        ):
            temporary_path.unlink(
                missing_ok=True
            )


# ============================================================================
# ONE PAIR WORKBOOK
# ============================================================================

def process_pair_workbook(
    workbook_path: Path,
    manual_log: pd.DataFrame,
) -> dict[str, Any]:
    with pd.ExcelFile(
        workbook_path,
        engine="openpyxl",
    ) as excel:
        if EV_SHEET not in excel.sheet_names:
            raise ValueError(
                f"Missing {EV_SHEET}"
            )

        if ICE_SHEET not in excel.sheet_names:
            raise ValueError(
                f"Missing {ICE_SHEET}"
            )

        ev_source = pd.read_excel(
            excel,
            sheet_name=EV_SHEET,
        )

        ice_source = pd.read_excel(
            excel,
            sheet_name=ICE_SHEET,
        )

    pair_info = read_pair_info(
        workbook_path
    )

    ev_context = pair_vehicle_context(
        pair_info,
        "EV",
        ev_source,
        workbook_path,
    )

    ice_context = pair_vehicle_context(
        pair_info,
        "ICE",
        ice_source,
        workbook_path,
    )

    # Match each vehicle to the authoritative manual trip log.
    ev_manual = match_manual_record(
        manual_log,
        ev_context,
    )
    ice_manual = match_manual_record(
        manual_log,
        ice_context,
    )

    (
        ev_result,
        ev_audit,
        ev_manual_audit,
        ev_charging,
        ev_metrics,
    ) = segment_vehicle(
        ev_source,
        "EV",
        ev_context,
        ev_manual,
    )

    (
        ice_result,
        ice_audit,
        ice_manual_audit,
        ice_charging,
        ice_metrics,
    ) = segment_vehicle(
        ice_source,
        "ICE",
        ice_context,
        ice_manual,
    )

    charging_details = concat_frames_without_futurewarning(
        [
            ev_charging,
            ice_charging,
        ]
    )

    summary = build_phase_summary(
        ev_result,
        ice_result,
        charging_details,
    )

    audit_frames = [
        frame
        for frame in (
            ev_audit,
            ice_audit,
        )
        if not frame.empty
    ]

    classification_audit = (
        concat_frames_without_futurewarning(
            audit_frames
        )
        if audit_frames
        else pd.DataFrame(
            columns=[
                "Vehicle",
                "Route Category",
                "Applied Stop Threshold (s)",
                "Excel Row",
                "Timestamp",
                "Interval Start",
                "Time Gap (s)",
                "Distance (m)",
                "Gap Equivalent Speed (km/h)",
                "Device Speed (km/h)",
                "Recorded Zero-Speed Episode",
                "Zero-Speed Episode ID",
                "Zero-Speed Episode Duration (s)",
                "Acceleration (m/s^2)",
                "Phase",
                "Classification Reason",
                "Confidence",
                "Manual Stop Support",
                "Manual Charging/Refuelling Support",
                "Charging-Stop Conflict Resolved",
                "Station Address Evidence",
                "Manual Charge Location Evidence",
                "Charge Zone Evidence",
                "Boundary Gap",
                "Very Long Gap",
                "Scenario",
                "Address",
                "Reverse Geocoded Address",
            ]
        )
    )

    manual_frames = [
        frame
        for frame in (
            ev_manual_audit,
            ice_manual_audit,
        )
        if frame is not None
        and not frame.empty
    ]

    manual_event_audit = (
        concat_frames_without_futurewarning(
            manual_frames
        )
        if manual_frames
        else pd.DataFrame(
            columns=[
                "Vehicle",
                "Event Type",
                "Event Index",
                "Manual Start",
                "Manual End",
                "Manual Duration (s)",
                "Manual Location",
                "Raw Manual Text",
                "Match Status",
                "Match Method",
                "Telemetry Row Index",
                "Excel Row",
                "Telemetry Interval Start",
                "Telemetry Interval End",
                "Telemetry Gap (s)",
                "Telemetry Distance (m)",
                "Gap Equivalent Speed (km/h)",
                "Time Offset (s)",
                "Boundary Gap",
                "Station Address Evidence",
                "Notes",
            ]
        )
    )

    save_pair_workbook(
        workbook_path=workbook_path,
        ev_result=ev_result,
        ice_result=ice_result,
        summary=summary,
        classification_audit=classification_audit,
        manual_event_audit=manual_event_audit,
        charging_details=charging_details,
    )

    return {
        "pair_workbook": workbook_path.name,
        "route_category": ev_metrics[
            "route_category"
        ],
        "applied_stop_threshold_seconds": ev_metrics[
            "applied_stop_threshold_seconds"
        ],
        "ev_manual_match": (
            "MATCHED"
            if ev_manual is not None
            else "NOT MATCHED"
        ),
        "ice_manual_match": (
            "MATCHED"
            if ice_manual is not None
            else "NOT MATCHED"
        ),
        "ev_manual_stop_events": ev_metrics[
            "manual_stop_events"
        ],
        "ev_manual_stop_matches": ev_metrics[
            "manual_stop_matches"
        ],
        "ev_manual_stop_start_inferred_events": ev_metrics[
            "manual_stop_start_inferred_events"
        ],
        "ev_manual_stop_start_inferred_rows": ev_metrics[
            "manual_stop_start_inferred_rows"
        ],
        "ice_manual_stop_events": ice_metrics[
            "manual_stop_events"
        ],
        "ice_manual_stop_matches": ice_metrics[
            "manual_stop_matches"
        ],
        "ice_manual_stop_start_inferred_events": ice_metrics[
            "manual_stop_start_inferred_events"
        ],
        "ice_manual_stop_start_inferred_rows": ice_metrics[
            "manual_stop_start_inferred_rows"
        ],
        "ev_manual_replenishment_rows": ev_metrics[
            "manual_replenishment_matches"
        ],
        "ice_manual_replenishment_rows": ice_metrics[
            "manual_replenishment_matches"
        ],
        "ev_charging_stop_conflicts_resolved": ev_metrics[
            "manual_charging_stop_conflicts_resolved"
        ],
        "ice_charging_stop_conflicts_resolved": ice_metrics[
            "manual_charging_stop_conflicts_resolved"
        ],
        "ev_final_charging_time_min": float(
            pd.to_numeric(
                ev_result.loc[
                    ev_result[PHASE].eq("CHARGING"),
                    TIME_GAP,
                ],
                errors="coerce",
            ).fillna(0.0).sum()
            / 60.0
        ),
        "ice_final_refueling_time_min": float(
            pd.to_numeric(
                ice_result.loc[
                    ice_result[PHASE].eq("REFUELING"),
                    TIME_GAP,
                ],
                errors="coerce",
            ).fillna(0.0).sum()
            / 60.0
        ),
        "ev_rule_gap_rows": ev_metrics[
            "rule_gap_rows"
        ],
        "ice_rule_gap_rows": ice_metrics[
            "rule_gap_rows"
        ],
        "ev_recorded_zero_speed_episodes": ev_metrics[
            "recorded_zero_speed_episodes"
        ],
        "ice_recorded_zero_speed_episodes": ice_metrics[
            "recorded_zero_speed_episodes"
        ],
        "ev_recorded_zero_speed_stop_rows": ev_metrics[
            "recorded_zero_speed_stop_rows"
        ],
        "ice_recorded_zero_speed_stop_rows": ice_metrics[
            "recorded_zero_speed_stop_rows"
        ],
        "ev_recorded_zero_speed_idle_rows": ev_metrics[
            "recorded_zero_speed_idle_rows"
        ],
        "ice_recorded_zero_speed_idle_rows": ice_metrics[
            "recorded_zero_speed_idle_rows"
        ],
        "status": "COMPLETED",
        "error": "",
    }


# ============================================================================
# MAIN
# ============================================================================

def discover_pair_workbooks() -> list[Path]:
    if not PAIRED_ROOT.exists():
        raise FileNotFoundError(
            f"Paired folder not found: {PAIRED_ROOT}"
        )

    files = sorted(
        path
        for path in PAIRED_ROOT.rglob(
            "*.xlsx"
        )
        if not path.name.startswith(
            "~$"
        )
        and not path.name.startswith(
            "."
        )
    )

    if PAIR_NAME_FILTER:
        wanted = PAIR_NAME_FILTER.casefold()

        files = [
            path
            for path in files
            if wanted
            in path.name.casefold()
        ]

    return files


def main() -> None:

    print(
        "=" * 88
    )

    print(
        "TAT STAGE 4 - INTERCITY PHASE SEGMENTATION"
    )

    print(
        f"SCRIPT VERSION: {SCRIPT_VERSION}"
    )

    print(
        "=" * 88
    )

    manual_path = resolve_manual_trip_log_path(
        MANUAL_TRIP_LOG_PATH
    )
    manual_log = load_manual_trip_log(
        manual_path
    )

    print(
        f"Manual trip log: ENABLED | {manual_path}"
    )
    print(
        f"Usable manual trip records loaded: {len(manual_log)}"
    )
    print(
        "Manual-first INTERCITY hierarchy:"
    )
    print(
        "    1. Manual charging/refuelling start+duration or matched event has priority"
    )
    print(
        "    2. Manual stop intervals/start timestamps are matched to telemetry"
    )
    print(
        "    3. Unmatched stationary telemetry uses the 180 s INTERCITY STOP boundary"
    )
    print(
        "    4. CHARGING/REFUELING requires a matched manual charging/refuelling "
        "event; station/address evidence only helps matching"
    )
    print(
        "    5. Remaining moving rows use INCH / ACC / DEC / CRUISE rules"
    )
    print(
        "All positive-duration rows are assigned one of the seven phases; "
        "low-confidence fallbacks remain explicitly audited."
    )
    print()

    pair_workbooks = discover_pair_workbooks()

    print(
        f"Pair workbooks found: "
        f"{len(pair_workbooks)}"
    )

    print()

    SYSTEM_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    started = time.monotonic()

    global_rows: list[
        dict[str, Any]
    ] = []

    completed = 0
    errors = 0

    for number, workbook_path in enumerate(
        pair_workbooks,
        start=1,
    ):
        pair_started = time.monotonic()

        print(
            f"[{number}/{len(pair_workbooks)}] "
            f"{workbook_path.name}"
        )

        try:
            result = process_pair_workbook(
                workbook_path,
                manual_log,
            )

            completed += 1

            print(
                "    COMPLETED | "
                f"Category: {result['route_category']} | "
                f"STOP threshold: "
                f"{result['applied_stop_threshold_seconds']:g} s | "
                f"EV manual: "
                f"{result['ev_manual_match']} | "
                f"ICE manual: "
                f"{result['ice_manual_match']}"
            )

            print(
                "    Manual stop matching | "
                f"EV: "
                f"{result['ev_manual_stop_matches']}/"
                f"{result['ev_manual_stop_events']} | "
                f"ICE: "
                f"{result['ice_manual_stop_matches']}/"
                f"{result['ice_manual_stop_events']}"
            )

            print(
                "    Start-only stop inference | "
                "EV events/rows: "
                f"{result['ev_manual_stop_start_inferred_events']}/"
                f"{result['ev_manual_stop_start_inferred_rows']} | "
                "ICE events/rows: "
                f"{result['ice_manual_stop_start_inferred_events']}/"
                f"{result['ice_manual_stop_start_inferred_rows']}"
            )

            print(
                "    Manual replenishment rows | "
                f"EV: {result['ev_manual_replenishment_rows']} | "
                f"ICE: {result['ice_manual_replenishment_rows']} | "
                "final time (min) | "
                f"EV CHARGING: {result['ev_final_charging_time_min']:.2f} | "
                "ICE REFUELING: "
                f"{result['ice_final_refueling_time_min']:.2f}"
            )

            print(
                "    Rule gaps | "
                f"EV: "
                f"{result['ev_rule_gap_rows']} | "
                f"ICE: "
                f"{result['ice_rule_gap_rows']}"
            )

            print(
                "    Recorded-zero episodes | "
                f"EV: {result['ev_recorded_zero_speed_episodes']} "
                f"(STOP rows {result['ev_recorded_zero_speed_stop_rows']}, "
                f"IDLE rows {result['ev_recorded_zero_speed_idle_rows']}) | "
                f"ICE: {result['ice_recorded_zero_speed_episodes']} "
                f"(STOP rows {result['ice_recorded_zero_speed_stop_rows']}, "
                f"IDLE rows {result['ice_recorded_zero_speed_idle_rows']})"
            )

            global_rows.append(
                result
            )

        except PermissionError as error:
            errors += 1

            message = (
                "Workbook/folder is locked. "
                "Close the pair workbook in Excel/WPS "
                "and rerun. "
                f"{error}"
            )

            print(
                f"    ERROR | {message}"
            )

            global_rows.append(
                {
                    "pair_workbook": workbook_path.name,
                    "status": "ERROR",
                    "error": message,
                }
            )

        except Exception as error:
            errors += 1

            print(
                f"    ERROR | {error}"
            )

            global_rows.append(
                {
                    "pair_workbook": workbook_path.name,
                    "status": "ERROR",
                    "error": str(
                        error
                    ),
                }
            )

        elapsed = (
            time.monotonic()
            - started
        )

        eta = (
            elapsed
            / number
            * (
                len(
                    pair_workbooks
                )
                - number
            )
        )

        print(
            "    Time | "
            f"this pair: "
            f"{format_duration(time.monotonic() - pair_started)} | "
            f"elapsed: "
            f"{format_duration(elapsed)} | "
            f"ETA: "
            f"{format_duration(eta)}"
        )

        print(
            "-" * 88
        )

    log_path = (
        SYSTEM_OUTPUT_DIR
        / GLOBAL_LOG_NAME
    )

    pd.DataFrame(
        global_rows
    ).to_excel(
        log_path,
        index=False,
        engine="openpyxl",
    )

    print()

    print(
        "=" * 88
    )

    print(
        "STAGE 4 SUMMARY"
    )

    print(
        "=" * 88
    )

    print(
        f"Completed: "
        f"{completed}/"
        f"{len(pair_workbooks)}"
    )

    print(
        f"Errors: {errors}"
    )

    print(
        f"Global log: "
        f"{log_path}"
    )

    print()

    print(
        "IMPORTANT: rerun tat_computation.py and results_visualization.py "
        "after this Stage 4. Existing downstream TAT sheets may contain "
        "stale results from an older segmentation."
    )

    print(
        "=" * 88
    )

    if errors:
        raise SystemExit(
            1
        )


if __name__ == "__main__":
    main()