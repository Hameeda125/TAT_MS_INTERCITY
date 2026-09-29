"""Stage 5: add interval features to every EV-ICE pair workbook.



Run this module after ``pair_cropping.py``.  It updates the two cropped-data

worksheets already stored below ``Results/Paired`` and preserves the pair-info

worksheet.  Each cropped-data worksheet is replaced by these columns while preserving the existing input Address column:



1. Timestamp

2. Latitude

3. Longitude

4. Speed (m/s)

5. Speed (km/h)

6. Time Difference (s)

7. Haversine Distance (m)

8. Speed Range (km/h)

9. Acceleration (m/s^2)

10. Address



The interval fields on row i describe the interval from row i-1 to row i.

The first row therefore uses zero for time difference, distance and

acceleration.  No cropped datapoint is removed.  Datapoint rows have no fill

colours.  Pair classification remains STRICT or NONSTRICT only.



Important implementation safeguards

-----------------------------------

* Equivalent headers such as ``Speed(m/s)`` and ``Speed (m/s)`` are coalesced

  rather than allowing one old/blank version to hide another valid version.

* Every workbook listed in ``pair_index.csv`` is read and calculated successfully

  before any workbook is changed.  A genuine preflight problem therefore cannot create a partly

  processed dataset.

* Only the EV and ICE data sheets are replaced.  ``03_Pair_Crop_Info`` is

  updated cell-by-cell as text, avoiding the pandas datetime/string error that

  previously stopped workbook saving.

* Every temporary workbook is reopened and its calculations are verified

  before it atomically replaces the original workbook.

* The operation is idempotent: rerunning it does not duplicate columns.

"""



from __future__ import annotations



import math

import os

import re

import shutil

import stat

import time

import unicodedata

from datetime import datetime

from pathlib import Path

from tempfile import NamedTemporaryFile



import numpy as np

import pandas as pd

from openpyxl.styles import Alignment, Font, PatternFill

from openpyxl.utils import get_column_letter



from config_loader import load_config





SCRIPT_VERSION = "2026-09-29_STAGE5_PRESERVE_CLASS_ISO_QC_V4_REGEX_FIX"



PAIR_INDEX_NAME = "pair_index.csv"

PAIR_INDEX_WORKBOOK_NAME = "Pair_Index.xlsx"

FEATURE_LOG_NAME = "Feature_Engineering_Log.xlsx"



EV_SHEET_NAME = "01_EV_Cropped_Data"

ICE_SHEET_NAME = "02_ICE_Cropped_Data"

INFO_SHEET_NAME = "03_Pair_Crop_Info"



TIMESTAMP = "Timestamp"

LATITUDE = "Latitude"

LONGITUDE = "Longitude"

SPEED_MPS = "Speed (m/s)"

SPEED_KMH = "Speed (km/h)"

TIME_DIFFERENCE = "Time Difference (s)"

HAVERSINE_DISTANCE = "Haversine Distance (m)"

SPEED_RANGE = "Speed Range (km/h)"

ACCELERATION = "Acceleration (m/s^2)"

# ===== ADDRESS PRESERVATION START =====

ADDRESS = "Address"

# ===== ADDRESS PRESERVATION END =====



OUTPUT_COLUMNS = [

    TIMESTAMP,

    LATITUDE,

    LONGITUDE,

    SPEED_MPS,

    SPEED_KMH,

    TIME_DIFFERENCE,

    HAVERSINE_DISTANCE,

    SPEED_RANGE,

    ACCELERATION,

    ADDRESS,

]



ALIASES = {

    TIMESTAMP: [

        "Timestamp",

        "Date Time",

        "Datetime",

        "Packet Time",

    ],

    LATITUDE: ["Latitude", "Lat"],

    LONGITUDE: ["Longitude", "Longitude (deg)", "Lon", "Lng", "Long"],

    SPEED_MPS: [

        "Speed (m/s)",

        "Speed(m/s)",

        "Speed m/s",

        "Speed_mps",

        "Speed_ms",

    ],

    SPEED_KMH: [

        "Speed (km/h)",

        "Speed(km/h)",

        "Speed km/h",

        "Speed_kmh",

        "Speed kmph",

        "Speed(kmph)",

        "Speed",

    ],

    # ===== ADDRESS PRESERVATION START =====

    ADDRESS: [

        "Address",

        "Location Address",

        "Full Address",

        "Formatted Address",

    ],

    # ===== ADDRESS PRESERVATION END =====

}



EARTH_RADIUS_M = 6_371_008.8

HEADER_BLUE = "1F4E78"

STRICT_ROW = "E2F0D9"

NONSTRICT_ROW = "DDEBF7"





class PreflightError(RuntimeError):

    """Raised before writing when one or more workbooks cannot be processed."""





def _clean_text(value) -> str:

    if value is None:

        return ""

    try:

        if pd.isna(value):

            return ""

    except (TypeError, ValueError):

        pass

    return str(value).strip()





def _normalise_key(value) -> str:

    text = unicodedata.normalize("NFKD", _clean_text(value))

    text = "".join(

        character for character in text if not unicodedata.combining(character)

    )

    return re.sub(r"[^a-z0-9]+", "", text.casefold())





def _read_csv(path: Path) -> pd.DataFrame:

    last_error = None

    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):

        try:

            return pd.read_csv(path, encoding=encoding)

        except UnicodeDecodeError as error:

            last_error = error

    raise ValueError(f"Could not decode {path}: {last_error}")





def _matching_columns(dataframe: pd.DataFrame, aliases: list[str]) -> list[str]:

    """Return all matching columns in alias order, without silent overwriting."""



    by_key: dict[str, list[str]] = {}

    for column in dataframe.columns:

        by_key.setdefault(_normalise_key(column), []).append(column)



    matches: list[str] = []

    for alias in aliases:

        for column in by_key.get(_normalise_key(alias), []):

            if column not in matches:

                matches.append(column)

    return matches





def _parse_numeric(series: pd.Series) -> pd.Series:

    if pd.api.types.is_numeric_dtype(series.dtype):

        return pd.to_numeric(series, errors="coerce").astype(float)



    text = series.astype("string").str.strip()

    text = text.str.replace("\u2212", "-", regex=False)

    text = text.str.replace(",", "", regex=False)

    extracted = text.str.extract(

        r"([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?)",

        expand=False,

    )

    values = pd.to_numeric(extracted, errors="coerce").astype(float)

    south_or_west = text.str.contains(r"[SW]\s*$", case=False, na=False)

    values.loc[south_or_west & values.notna()] = -values.loc[

        south_or_west & values.notna()

    ].abs()

    return values





def _parse_datetime_pass(series: pd.Series, dayfirst: bool) -> pd.Series:

    try:

        return pd.to_datetime(

            series,

            errors="coerce",

            dayfirst=dayfirst,

            format="mixed",

        )

    except (TypeError, ValueError):

        # Compatibility with older pandas versions that do not support

        # format="mixed".

        return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)





def _parse_datetimes(series: pd.Series) -> pd.Series:

    """Parse timestamps without swapping month/day in ISO-style dates.



    Rules:

    * Existing datetime values are preserved.

    * Excel serial dates are supported.

    * Strings beginning YYYY-MM-DD or YYYY/MM/DD are parsed year-first.

    * Remaining text is parsed day-first first, then non-day-first as fallback.

    """



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



    text = series.astype("string").str.strip()

    unresolved = result.isna() & series.notna() & text.ne("")



    iso_style = unresolved & text.str.match(

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



    unresolved = result.isna() & series.notna() & text.ne("")

    if bool(unresolved.any()):

        first = _parse_datetime_pass(series.loc[unresolved], dayfirst=True)

        result.loc[unresolved] = first



    unresolved = result.isna() & series.notna() & text.ne("")

    if bool(unresolved.any()):

        second = _parse_datetime_pass(series.loc[unresolved], dayfirst=False)

        result.loc[unresolved] = second



    return result





def _coalesce_numeric(

    dataframe: pd.DataFrame,

    aliases: list[str],

) -> tuple[pd.Series, list[str]]:

    columns = _matching_columns(dataframe, aliases)

    result = pd.Series(np.nan, index=dataframe.index, dtype=float)

    for column in columns:

        values = _parse_numeric(dataframe[column])

        result = result.combine_first(values)

    return result, columns





def _coalesce_datetimes(

    dataframe: pd.DataFrame,

    aliases: list[str],

) -> tuple[pd.Series, list[str]]:

    columns = _matching_columns(dataframe, aliases)

    result = pd.Series(pd.NaT, index=dataframe.index, dtype="datetime64[ns]")

    for column in columns:

        values = _parse_datetimes(dataframe[column])

        result = result.combine_first(values)

    return result, columns







# ===== ADDRESS PRESERVATION START =====

def _coalesce_text(

    dataframe: pd.DataFrame,

    aliases: list[str],

) -> tuple[pd.Series, list[str]]:

    """Coalesce equivalent text columns without discarding existing values."""

    columns = _matching_columns(dataframe, aliases)

    result = pd.Series("", index=dataframe.index, dtype="string")

    for column in columns:

        values = dataframe[column].astype("string").fillna("").str.strip()

        missing = result.eq("")

        result.loc[missing] = values.loc[missing]

    return result.fillna("").astype(str), columns

# ===== ADDRESS PRESERVATION END =====





def _settings(config: dict) -> dict:

    feature = config.get("feature_engineering", {}) or {}

    cleaning = config.get("cleaning", {}) or {}

    width = float(feature.get("speed_band_width_kmh", 5.0))

    top = float(feature.get("speed_band_top_kmh", 80.0))

    maximum_gps_speed = float(

        feature.get(

            "maximum_gps_implied_speed_kmh",

            cleaning.get("maximum_realistic_speed_kmh", 160.0),

        )

    )

    if width <= 0 or top <= 0 or maximum_gps_speed <= 0:

        raise ValueError("Feature-engineering thresholds must be positive")

    return {

        "speed_band_width_kmh": width,

        "speed_band_top_kmh": top,

        "maximum_gps_implied_speed_kmh": maximum_gps_speed,

    }





def _haversine_distance_m(latitude: pd.Series, longitude: pd.Series) -> pd.Series:

    count = len(latitude)

    output = np.zeros(count, dtype=float)

    if count <= 1:

        return pd.Series(output, index=latitude.index, dtype=float)



    lat1 = np.radians(latitude.shift(1).to_numpy(dtype=float))

    lat2 = np.radians(latitude.to_numpy(dtype=float))

    delta_latitude = lat2 - lat1

    delta_longitude = np.radians(

        longitude.to_numpy(dtype=float)

        - longitude.shift(1).to_numpy(dtype=float)

    )

    haversine_value = (

        np.sin(delta_latitude / 2.0) ** 2

        + np.cos(lat1)

        * np.cos(lat2)

        * np.sin(delta_longitude / 2.0) ** 2

    )

    haversine_value = np.clip(haversine_value, 0.0, 1.0)

    output[1:] = (

        EARTH_RADIUS_M

        * 2.0

        * np.arctan2(

            np.sqrt(haversine_value[1:]),

            np.sqrt(1.0 - haversine_value[1:]),

        )

    )

    return pd.Series(output, index=latitude.index, dtype=float)





def _speed_band(speed_kmh: float, rules: dict) -> str:

    speed = max(0.0, float(speed_kmh))

    width = rules["speed_band_width_kmh"]

    top = rules["speed_band_top_kmh"]

    if speed >= top:

        return f"{top:g}+"

    lower = math.floor(speed / width) * width

    return f"{lower:g}-{lower + width:g}"





def _interpolate_internal_speed_gaps(

    speed_kmh: pd.Series,

    timestamps: pd.Series,

    method: pd.Series,

) -> None:

    missing = np.flatnonzero(speed_kmh.isna().to_numpy())

    cursor = 0

    while cursor < len(missing):

        start = int(missing[cursor])

        end = start

        while cursor + 1 < len(missing) and int(missing[cursor + 1]) == end + 1:

            cursor += 1

            end = int(missing[cursor])



        previous = start - 1

        following = end + 1

        if previous >= 0 and following < len(speed_kmh):

            previous_speed = speed_kmh.iloc[previous]

            following_speed = speed_kmh.iloc[following]

            previous_time = timestamps.iloc[previous]

            following_time = timestamps.iloc[following]

            if (

                pd.notna(previous_speed)

                and pd.notna(following_speed)

                and pd.notna(previous_time)

                and pd.notna(following_time)

            ):

                total_seconds = (following_time - previous_time).total_seconds()

                if total_seconds > 0:

                    for position in range(start, end + 1):

                        current_time = timestamps.iloc[position]

                        fraction = (

                            (current_time - previous_time).total_seconds()

                            / total_seconds

                        )

                        speed_kmh.iloc[position] = max(

                            0.0,

                            float(previous_speed)

                            + fraction

                            * (float(following_speed) - float(previous_speed)),

                        )

                        method.iloc[position] = "LINEAR_TIME_INTERPOLATION"

        cursor += 1





def _fill_nearest_speed(

    speed_kmh: pd.Series,

    timestamps: pd.Series,

    method: pd.Series,

) -> None:

    missing_positions = np.flatnonzero(speed_kmh.isna().to_numpy())

    if not len(missing_positions):

        return

    valid_positions = np.flatnonzero(speed_kmh.notna().to_numpy())

    if not len(valid_positions):

        return



    timestamp_ns = timestamps.astype("int64").to_numpy()

    for position in missing_positions:

        nearest_index = int(

            np.argmin(np.abs(timestamp_ns[valid_positions] - timestamp_ns[position]))

        )

        nearest_position = int(valid_positions[nearest_index])

        speed_kmh.iloc[position] = float(speed_kmh.iloc[nearest_position])

        method.iloc[position] = "NEAREST_RECORDED_SPEED"





def _complete_speed(

    speed_mps_source: pd.Series,

    speed_kmh_source: pd.Series,

    timestamps: pd.Series,

    time_difference: pd.Series,

    distance_m: pd.Series,

    rules: dict,

) -> tuple[pd.Series, dict]:

    """Complete the speed series without deleting a retained datapoint.



    This is a defensive compatibility layer.  Cleaned data should already be

    complete, but earlier Stage 3 versions may have rewritten a few paired

    sheets.  Completion order is: recorded km/h, recorded m/s conversion,

    linear interpolation, GPS interval speed, nearest recorded speed and trip

    median.  A stationary no-evidence trip uses zero as the only neutral value.

    """



    kmh = speed_kmh_source.astype(float).copy()

    mps = speed_mps_source.astype(float).copy()

    kmh = kmh.where(np.isfinite(kmh) & (kmh >= 0.0))

    mps = mps.where(np.isfinite(mps) & (mps >= 0.0))



    method = pd.Series("RECORDED_KMH", index=kmh.index, dtype=object)

    from_mps = kmh.isna() & mps.notna()

    kmh.loc[from_mps] = mps.loc[from_mps] * 3.6

    method.loc[from_mps] = "CONVERTED_FROM_RECORDED_MPS"



    # Count rows that need completion only after both recorded units have been

    # considered.  This prevents a header-spacing mismatch from being reported

    # as missing telemetry.

    completion_needed = kmh.isna()

    completion_count = int(completion_needed.sum())

    method.loc[completion_needed] = "PENDING_COMPLETION"



    # When both units exist, km/h is kept as the recorded reference and the

    # final m/s column is regenerated from it.  This guarantees unit identity.

    both = kmh.notna() & mps.notna()

    unit_reconciled = int(

        (

            both

            & ~np.isclose(

                kmh,

                mps * 3.6,

                rtol=1e-6,

                atol=1e-3,

                equal_nan=True,

            )

        ).sum()

    )



    _interpolate_internal_speed_gaps(kmh, timestamps, method)



    gps_speed = pd.Series(np.nan, index=kmh.index, dtype=float)

    computable = (time_difference > 0.0) & distance_m.notna()

    gps_speed.loc[computable] = (

        distance_m.loc[computable] / time_difference.loc[computable] * 3.6

    )

    gps_speed = gps_speed.where(

        np.isfinite(gps_speed)

        & (gps_speed >= 0.0)

        & (gps_speed <= rules["maximum_gps_implied_speed_kmh"])

    )



    current_interval = kmh.isna() & gps_speed.notna()

    kmh.loc[current_interval] = gps_speed.loc[current_interval]

    method.loc[current_interval] = "GPS_INTERVAL_SPEED"



    next_interval_speed = gps_speed.shift(-1)

    following_interval = kmh.isna() & next_interval_speed.notna()

    kmh.loc[following_interval] = next_interval_speed.loc[following_interval]

    method.loc[following_interval] = "NEXT_GPS_INTERVAL_SPEED"



    _fill_nearest_speed(kmh, timestamps, method)



    if bool(kmh.isna().any()):

        trip_median = kmh.median(skipna=True)

        if pd.notna(trip_median):

            median_rows = kmh.isna()

            kmh.loc[median_rows] = float(trip_median)

            method.loc[median_rows] = "TRIP_MEDIAN_SPEED"



    if bool(kmh.isna().any()) and float(distance_m.sum()) <= 1e-6:

        stationary_rows = kmh.isna()

        kmh.loc[stationary_rows] = 0.0

        method.loc[stationary_rows] = "STATIONARY_ZERO_SPEED"



    if bool(kmh.isna().any()):

        raise ValueError(

            "Speed could not be reconstructed from either recorded unit, "

            "neighbouring timestamps or consecutive GPS movement"

        )



    kmh = kmh.astype(float).clip(lower=0.0)

    counts = method.value_counts().to_dict()

    audit = {

        "speed_values_completed": completion_count,

        "speed_from_mps_count": int(

            counts.get("CONVERTED_FROM_RECORDED_MPS", 0)

        ),

        "speed_linear_interpolation_count": int(

            counts.get("LINEAR_TIME_INTERPOLATION", 0)

        ),

        "speed_gps_completion_count": int(

            counts.get("GPS_INTERVAL_SPEED", 0)

            + counts.get("NEXT_GPS_INTERVAL_SPEED", 0)

        ),

        "speed_nearest_completion_count": int(

            counts.get("NEAREST_RECORDED_SPEED", 0)

        ),

        "speed_median_completion_count": int(

            counts.get("TRIP_MEDIAN_SPEED", 0)

        ),

        "speed_stationary_zero_count": int(

            counts.get("STATIONARY_ZERO_SPEED", 0)

        ),

        "speed_unit_reconciled_count": unit_reconciled,

    }

    return kmh, audit





def engineer_features(

    source: pd.DataFrame,

    rules: dict,

) -> tuple[pd.DataFrame, dict]:

    """Return the Stage 3 feature table while preserving Address."""



    if source.empty:

        raise ValueError("Cropped-data sheet contains no datapoints")



    timestamps, timestamp_columns = _coalesce_datetimes(

        source, ALIASES[TIMESTAMP]

    )

    latitude, latitude_columns = _coalesce_numeric(source, ALIASES[LATITUDE])

    longitude, longitude_columns = _coalesce_numeric(source, ALIASES[LONGITUDE])

    speed_mps_source, speed_mps_columns = _coalesce_numeric(

        source, ALIASES[SPEED_MPS]

    )

    speed_kmh_source, speed_kmh_columns = _coalesce_numeric(

        source, ALIASES[SPEED_KMH]

    )

    address, address_columns = _coalesce_text(

        source, ALIASES[ADDRESS]

    )



    missing_headers = []

    if not timestamp_columns:

        missing_headers.append("Timestamp")

    if not latitude_columns:

        missing_headers.append("Latitude")

    if not longitude_columns:

        missing_headers.append("Longitude")

    if not speed_mps_columns and not speed_kmh_columns:

        missing_headers.append("Speed in m/s or km/h")

    if missing_headers:

        raise ValueError("Missing source column(s): " + ", ".join(missing_headers))



    working = pd.DataFrame(

        {

            TIMESTAMP: timestamps,

            LATITUDE: latitude,

            LONGITUDE: longitude,

            "_speed_mps_source": speed_mps_source,

            "_speed_kmh_source": speed_kmh_source,

            ADDRESS: address,

        }

    )



    timestamp_problem_count = int(working[TIMESTAMP].isna().sum())

    coordinate_problem_count = int(

        working[[LATITUDE, LONGITUDE]].isna().any(axis=1).sum()

    )

    coordinate_range_problem_count = int(

        (

            ~working[LATITUDE].between(-90.0, 90.0)

            | ~working[LONGITUDE].between(-180.0, 180.0)

        ).sum()

    )

    problems = []

    if timestamp_problem_count:

        problems.append(f"{timestamp_problem_count} timestamp row(s)")

    if coordinate_problem_count:

        problems.append(f"{coordinate_problem_count} coordinate row(s)")

    if coordinate_range_problem_count:

        problems.append(

            f"{coordinate_range_problem_count} out-of-range coordinate row(s)"

        )

    if problems:

        raise ValueError(

            "Cropped sheet cannot support interval calculations: "

            + ", ".join(problems)

        )



    # Stable sorting retains all rows and gives consecutive interval fields a

    # single, reproducible chronological meaning.

    working = working.sort_values(TIMESTAMP, kind="stable").reset_index(drop=True)



    time_difference = working[TIMESTAMP].diff().dt.total_seconds().astype(float)

    time_difference.iloc[0] = 0.0

    if bool((time_difference < 0.0).any()):

        raise RuntimeError("Negative time difference remained after stable sorting")



    distance_m = _haversine_distance_m(

        working[LATITUDE], working[LONGITUDE]

    )

    speed_kmh, speed_audit = _complete_speed(

        working["_speed_mps_source"],

        working["_speed_kmh_source"],

        working[TIMESTAMP],

        time_difference,

        distance_m,

        rules,

    )

    speed_mps = speed_kmh / 3.6



    # Zero is an explicit non-interval placeholder for the first row and for

    # a duplicate-timestamp row.  Positive-time rows use the physical formula.

    acceleration = pd.Series(0.0, index=working.index, dtype=float)

    speed_change = speed_mps.diff()

    positive_time = time_difference > 0.0

    acceleration.loc[positive_time] = (

        speed_change.loc[positive_time] / time_difference.loc[positive_time]

    )



    output = pd.DataFrame(

        {

            TIMESTAMP: working[TIMESTAMP],

            LATITUDE: working[LATITUDE].astype(float),

            LONGITUDE: working[LONGITUDE].astype(float),

            SPEED_MPS: speed_mps.astype(float),

            SPEED_KMH: speed_kmh.astype(float),

            TIME_DIFFERENCE: time_difference.astype(float),

            HAVERSINE_DISTANCE: distance_m.astype(float),

            SPEED_RANGE: speed_kmh.map(lambda value: _speed_band(value, rules)),

            ACCELERATION: acceleration.astype(float),

            ADDRESS: working[ADDRESS].fillna("").astype(str),

        },

        columns=OUTPUT_COLUMNS,

    )



    numeric_columns = [

        LATITUDE,

        LONGITUDE,

        SPEED_MPS,

        SPEED_KMH,

        TIME_DIFFERENCE,

        HAVERSINE_DISTANCE,

        ACCELERATION,

    ]

    required_nonblank_columns = [

        TIMESTAMP,

        LATITUDE,

        LONGITUDE,

        SPEED_MPS,

        SPEED_KMH,

        TIME_DIFFERENCE,

        HAVERSINE_DISTANCE,

        SPEED_RANGE,

        ACCELERATION,

    ]

    if output[required_nonblank_columns].isna().any().any():

        blank_columns = output[required_nonblank_columns].columns[

            output[required_nonblank_columns].isna().any()

        ].tolist()

        raise RuntimeError(

            "Feature table contains blank required cells in: "

            + ", ".join(blank_columns)

        )

    if not np.isfinite(output[numeric_columns].to_numpy(dtype=float)).all():

        raise RuntimeError("Feature table contains a non-finite numeric value")



    audit = {

        "datapoint_count": len(output),

        "consecutive_interval_count": max(len(output) - 1, 0),

        "positive_time_interval_count": int(positive_time.sum()),

        "zero_time_interval_count": int((time_difference == 0.0).sum()),

        "time_difference_sum_seconds": float(time_difference.sum()),

        "haversine_distance_sum_m": float(distance_m.sum()),

        "acceleration_value_count": int(output[ACCELERATION].notna().sum()),

        "timestamp_source_columns": " | ".join(map(str, timestamp_columns)),

        "latitude_source_columns": " | ".join(map(str, latitude_columns)),

        "longitude_source_columns": " | ".join(map(str, longitude_columns)),

        "speed_mps_source_columns": " | ".join(map(str, speed_mps_columns)),

        "speed_kmh_source_columns": " | ".join(map(str, speed_kmh_columns)),

        "address_source_columns": " | ".join(map(str, address_columns)),

        "address_nonblank_count": int(

            output[ADDRESS].astype(str).str.strip().ne("").sum()

        ),

        **speed_audit,

    }

    return output, audit





def _normalise_pair_class(value) -> str:

    text = re.sub(r"[^A-Z0-9]+", "_", _clean_text(value).upper()).strip("_")

    if text in {"STRICT", "PRIMARY_STRICT", "STRICT_CANDIDATE"}:

        return "STRICT"

    if (

        text in {

            "NONSTRICT",

            "NON_STRICT",

            "SECONDARY_NONSTRICT",

            "SECONDARY_NON_STRICT",

            "NONSTRICT_CANDIDATE",

            "NON_STRICT_CANDIDATE",

        }

        or "NONSTRICT" in text

        or "NON_STRICT" in text

    ):

        return "NONSTRICT"

    raise ValueError(f"Unrecognised pair class: {value!r}")





def _remove_warning_token(value) -> str:

    text = _clean_text(value)

    text = re.sub(

        r"(?i)(?:_WITH_ROUTE_WARNING|_WITH_WARNING|_ROUTE_WARNING|_WARNING)",

        "",

        text,

    )

    text = re.sub(r"(?i)(?:^|[ _|])WARNING(?:$|[ _|])", "_", text)

    text = re.sub(r"_{2,}", "_", text)

    text = re.sub(r"\s{2,}", " ", text)

    text = re.sub(r"\|\s*\|", "|", text)

    return text.strip(" _|")





def _neutralise_warning_text(value):

    if value is None:

        return value

    try:

        if pd.isna(value):

            return value

    except (TypeError, ValueError):

        pass

    text = str(value)

    replacements = {

        "ACCEPTED_WITH_ROUTE_WARNING": "ACCEPTED_USING_AVAILABLE_ROUTE",

        "RETAINED_WITH_ROUTE_WARNING": "RETAINED_USING_AVAILABLE_ROUTE",

        "CROPPED_PAIR_WORKBOOK_CREATED_WITH_ROUTE_WARNING": (

            "CROPPED_PAIR_WORKBOOK_CREATED"

        ),

        "BEST_AVAILABLE_ORDERED_SECTION_WITH_WARNING": (

            "BEST_AVAILABLE_ORDERED_SECTION"

        ),

        "VERIFIED_BOUNDARY_FALLBACK_WITH_WARNING": "VERIFIED_BOUNDARY_FALLBACK",

        "FULL_VERIFIED_TRIP_FALLBACK_WARNING": "FULL_VERIFIED_TRIP_FALLBACK",

        "WARNING_DOES_NOT_BLOCK_PROCESSING": "PROCESSED_USING_AVAILABLE_DATA",

    }

    for old, new in replacements.items():

        text = re.sub(re.escape(old), new, text, flags=re.IGNORECASE)

    return _remove_warning_token(text)





def _safe_route(route_direction: str) -> str:

    route = _clean_text(route_direction) or "Unknown Route"

    route = re.sub(r"(?i)\s+to\s+", "_to_", route)

    route = re.sub(r"[^A-Za-z0-9_-]+", "_", route)

    return re.sub(r"_{2,}", "_", route).strip("_") or "Unknown_Route"





def _desired_workbook_name(row: pd.Series, current_path: Path) -> str:

    pair_id = _clean_text(row.get("pair_id"))

    pair_class = _normalise_pair_class(row.get("final_pair_class"))

    route_direction = _clean_text(row.get("route_direction"))

    cleaned_current = _remove_warning_token(current_path.name)

    match = re.match(

        r"^PAIR_\d{3}_(?:STRICT|NONSTRICT)_(.+)\.xlsx$",

        cleaned_current,

        flags=re.IGNORECASE,

    )

    route_suffix = match.group(1) if match else _safe_route(route_direction)

    route_suffix = _remove_warning_token(route_suffix).strip("_")

    return f"{pair_id}_{pair_class}_{route_suffix}.xlsx"





def _resolve_workbook_path(row: pd.Series, results_root: Path) -> Path:

    paired_root = results_root / "Paired"

    pair_id = _clean_text(row.get("pair_id"))

    configured_text = _clean_text(row.get("pair_workbook_path"))

    configured = Path(configured_text) if configured_text else None

    name = _clean_text(row.get("pair_workbook_name"))

    folder = _clean_text(row.get("output_date_folder"))



    candidates: list[Path] = []

    if name and folder:

        candidates.append(paired_root / folder / _remove_warning_token(name))

    if configured is not None:

        candidates.append(configured.with_name(_remove_warning_token(configured.name)))

    if configured is not None:

        candidates.append(configured)

    if name and folder:

        candidates.append(paired_root / folder / name)



    for candidate in candidates:

        if candidate.exists() and candidate.is_file():

            return candidate



    if pair_id and paired_root.exists():

        matches = [

            path

            for path in paired_root.rglob(f"{pair_id}_*.xlsx")

            if not path.name.startswith("~$") and not path.name.startswith(".")

        ]

        if matches:

            nonwarning = [

                path for path in matches if "WARNING" not in path.name.upper()

            ]

            pool = nonwarning or matches

            return max(pool, key=lambda path: path.stat().st_mtime)



    attempted = "; ".join(map(str, candidates)) or "no usable index path"

    raise FileNotFoundError(

        f"Pair workbook not found for {pair_id or name}. Tried: {attempted}"

    )





def _read_pair_data(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:

    with pd.ExcelFile(path, engine="openpyxl") as excel:

        required = [EV_SHEET_NAME, ICE_SHEET_NAME, INFO_SHEET_NAME]

        missing = [sheet for sheet in required if sheet not in excel.sheet_names]

        if missing:

            raise ValueError("Missing worksheet(s): " + ", ".join(missing))

        ev_source = pd.read_excel(excel, sheet_name=EV_SHEET_NAME)

        ice_source = pd.read_excel(excel, sheet_name=ICE_SHEET_NAME)

    return ev_source, ice_source





def _make_writable(path: Path) -> None:

    if path.exists():

        try:

            os.chmod(path, stat.S_IREAD | stat.S_IWRITE)

        except OSError:

            pass





def _check_not_locked(path: Path, attempts: int = 3) -> None:

    last_error = None

    for attempt in range(attempts):

        try:

            with path.open("r+b"):

                return

        except PermissionError as error:

            last_error = error

            time.sleep(0.35 * (attempt + 1))

    raise PermissionError(

        f"Workbook is open or locked: {path}. Close Excel and the File Explorer "

        f"Preview pane before running Stage 3. Original error: {last_error}"

    )





def _replace_with_retry(source: Path, destination: Path, attempts: int = 12) -> None:

    last_error = None

    for attempt in range(attempts):

        try:

            _make_writable(destination)

            os.replace(source, destination)

            return

        except PermissionError as error:

            last_error = error

            time.sleep(min(0.35 * (attempt + 1), 2.0))

    raise PermissionError(

        f"Could not replace {destination}. Close Excel and the File Explorer "

        f"Preview pane. Original error: {last_error}"

    )





def _unlink_with_retry(path: Path, attempts: int = 12) -> None:

    if not path.exists():

        return

    last_error = None

    for attempt in range(attempts):

        try:

            _make_writable(path)

            path.unlink()

            return

        except PermissionError as error:

            last_error = error

            time.sleep(min(0.35 * (attempt + 1), 2.0))

    raise PermissionError(f"Could not remove old workbook {path}: {last_error}")





def _format_data_sheet(worksheet) -> None:

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = worksheet.dimensions

    worksheet.sheet_view.showGridLines = True

    header_fill = PatternFill(fill_type="solid", fgColor=HEADER_BLUE)

    for cell in worksheet[1]:

        cell.fill = header_fill

        cell.font = Font(color="FFFFFF", bold=True)

        cell.alignment = Alignment(

            horizontal="center", vertical="center", wrap_text=True

        )

    worksheet.row_dimensions[1].height = 34

    widths = [22, 14, 14, 15, 15, 21, 24, 23, 23, 48]

    for number, width in enumerate(widths, start=1):

        worksheet.column_dimensions[get_column_letter(number)].width = width





def _format_table_sheet(worksheet, dataframe: pd.DataFrame) -> None:

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = worksheet.dimensions

    header_fill = PatternFill(fill_type="solid", fgColor=HEADER_BLUE)

    for cell in worksheet[1]:

        cell.fill = header_fill

        cell.font = Font(color="FFFFFF", bold=True)

        cell.alignment = Alignment(

            horizontal="center", vertical="center", wrap_text=True

        )

    worksheet.row_dimensions[1].height = 34

    for column_number, column_name in enumerate(dataframe.columns, start=1):

        values = [str(column_name)] + [

            _clean_text(value) for value in dataframe[column_name].head(200)

        ]

        width = min(max(max(map(len, values), default=0) + 2, 12), 48)

        worksheet.column_dimensions[get_column_letter(column_number)].width = width





def _apply_pair_class_colours(worksheet, dataframe: pd.DataFrame) -> None:

    class_column = None

    for candidate in ("final_pair_class", "pair_class"):

        if candidate in dataframe.columns:

            class_column = dataframe.columns.get_loc(candidate) + 1

            break

    if class_column is None:

        return

    for row_number in range(2, worksheet.max_row + 1):

        pair_class = _clean_text(

            worksheet.cell(row=row_number, column=class_column).value

        ).upper()

        colour = STRICT_ROW if pair_class == "STRICT" else NONSTRICT_ROW

        fill = PatternFill(fill_type="solid", fgColor=colour)

        for cell in worksheet[row_number]:

            cell.fill = fill





def _update_info_sheet(worksheet, updates: dict[str, str]) -> None:

    """Update metadata as strings without converting the whole worksheet."""



    headers = {

        _normalise_key(cell.value): cell.column

        for cell in worksheet[1]

        if _clean_text(cell.value)

    }

    parameter_column = headers.get(_normalise_key("Parameter"), 1)

    value_column = headers.get(_normalise_key("Value"), 2)



    # Preserve every upstream route/distance/sampling QC field exactly as it

    # arrived from pair_cropping.py. Feature engineering may add/update its own

    # metadata, but it must not delete or neutralise upstream review evidence.



    existing = {}

    for row_number in range(2, worksheet.max_row + 1):

        key = _clean_text(

            worksheet.cell(row=row_number, column=parameter_column).value

        )

        if key:

            existing[_normalise_key(key)] = row_number



    for key, value in updates.items():

        normalised = _normalise_key(key)

        row_number = existing.get(normalised)

        if row_number is None:

            row_number = worksheet.max_row + 1

            worksheet.cell(row=row_number, column=parameter_column).value = str(key)

            existing[normalised] = row_number

        worksheet.cell(row=row_number, column=value_column).value = str(value)



    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = worksheet.dimensions

    worksheet.column_dimensions[get_column_letter(parameter_column)].width = 55

    worksheet.column_dimensions[get_column_letter(value_column)].width = 90





def _verify_feature_frame(

    dataframe: pd.DataFrame,

    expected_rows: int,

    label: str,

) -> None:

    if list(dataframe.columns) != OUTPUT_COLUMNS:

        raise RuntimeError(

            f"{label} columns differ from the required feature schema"

        )

    if len(dataframe) != expected_rows:

        raise RuntimeError(

            f"{label} row count changed: {len(dataframe)} != {expected_rows}"

        )

    required_nonblank_columns = [

        TIMESTAMP,

        LATITUDE,

        LONGITUDE,

        SPEED_MPS,

        SPEED_KMH,

        TIME_DIFFERENCE,

        HAVERSINE_DISTANCE,

        SPEED_RANGE,

        ACCELERATION,

    ]

    if dataframe[required_nonblank_columns].isna().any().any():

        blank_columns = dataframe[required_nonblank_columns].columns[

            dataframe[required_nonblank_columns].isna().any()

        ].tolist()

        raise RuntimeError(

            f"{label} contains blank saved values in: "

            f"{', '.join(blank_columns)}"

        )



    timestamp = _parse_datetimes(dataframe[TIMESTAMP])

    latitude = _parse_numeric(dataframe[LATITUDE])

    longitude = _parse_numeric(dataframe[LONGITUDE])

    speed_mps = _parse_numeric(dataframe[SPEED_MPS])

    speed_kmh = _parse_numeric(dataframe[SPEED_KMH])

    time_difference = _parse_numeric(dataframe[TIME_DIFFERENCE])

    distance = _parse_numeric(dataframe[HAVERSINE_DISTANCE])

    acceleration = _parse_numeric(dataframe[ACCELERATION])



    numeric = pd.DataFrame(

        {

            "latitude": latitude,

            "longitude": longitude,

            "speed_mps": speed_mps,

            "speed_kmh": speed_kmh,

            "time_difference": time_difference,

            "distance": distance,

            "acceleration": acceleration,

        }

    )

    if timestamp.isna().any() or not np.isfinite(numeric.to_numpy()).all():

        raise RuntimeError(f"{label} contains an unparseable saved value")

    if (time_difference < 0.0).any() or (distance < 0.0).any():

        raise RuntimeError(f"{label} contains a negative time or distance value")

    if not np.allclose(speed_mps * 3.6, speed_kmh, rtol=1e-7, atol=1e-7):

        raise RuntimeError(f"{label} saved speed units are inconsistent")



    expected_time = timestamp.diff().dt.total_seconds().fillna(0.0)

    if not np.allclose(time_difference, expected_time, rtol=0.0, atol=1e-7):

        raise RuntimeError(f"{label} time differences failed verification")



    expected_distance = _haversine_distance_m(latitude, longitude)

    if not np.allclose(distance, expected_distance, rtol=1e-8, atol=1e-5):

        raise RuntimeError(f"{label} Haversine distances failed verification")



    expected_acceleration = pd.Series(0.0, index=dataframe.index, dtype=float)

    positive = time_difference > 0.0

    expected_acceleration.loc[positive] = (

        speed_mps.diff().loc[positive] / time_difference.loc[positive]

    )

    if not np.allclose(

        acceleration, expected_acceleration, rtol=1e-8, atol=1e-8

    ):

        raise RuntimeError(f"{label} acceleration failed verification")

    if dataframe[SPEED_RANGE].astype(str).str.strip().eq("").any():

        raise RuntimeError(f"{label} contains a blank speed range")





def _verify_workbook(

    path: Path,

    expected_ev_rows: int,

    expected_ice_rows: int,

) -> None:

    with pd.ExcelFile(path, engine="openpyxl") as excel:

        required = [EV_SHEET_NAME, ICE_SHEET_NAME, INFO_SHEET_NAME]

        missing = [sheet for sheet in required if sheet not in excel.sheet_names]

        if missing:

            raise RuntimeError(

                "Saved workbook is missing worksheet(s): " + ", ".join(missing)

            )

        ev = pd.read_excel(excel, sheet_name=EV_SHEET_NAME)

        ice = pd.read_excel(excel, sheet_name=ICE_SHEET_NAME)

    _verify_feature_frame(ev, expected_ev_rows, EV_SHEET_NAME)

    _verify_feature_frame(ice, expected_ice_rows, ICE_SHEET_NAME)





def _write_pair_workbook(record: dict) -> None:

    source_path: Path = record["source_path"]

    destination_path: Path = record["destination_path"]

    destination_path.parent.mkdir(parents=True, exist_ok=True)



    with NamedTemporaryFile(

        prefix=f".{source_path.stem}_stage3_",

        suffix=".xlsx",

        dir=source_path.parent,

        delete=False,

    ) as temporary_file:

        temporary_path = Path(temporary_file.name)



    try:

        shutil.copy2(source_path, temporary_path)

        with pd.ExcelWriter(

            temporary_path,

            engine="openpyxl",

            mode="a",

            if_sheet_exists="replace",

            datetime_format="dd-mm-yyyy hh:mm:ss",

        ) as writer:

            record["ev_features"].to_excel(

                writer, sheet_name=EV_SHEET_NAME, index=False

            )

            record["ice_features"].to_excel(

                writer, sheet_name=ICE_SHEET_NAME, index=False

            )

            _format_data_sheet(writer.sheets[EV_SHEET_NAME])

            _format_data_sheet(writer.sheets[ICE_SHEET_NAME])



            workbook = writer.book

            info_sheet = workbook[INFO_SHEET_NAME]

            _update_info_sheet(info_sheet, record["info_updates"])

            preferred = [EV_SHEET_NAME, ICE_SHEET_NAME, INFO_SHEET_NAME]

            workbook._sheets = [workbook[name] for name in preferred] + [

                sheet for sheet in workbook.worksheets if sheet.title not in preferred

            ]



        _verify_workbook(

            temporary_path,

            expected_ev_rows=len(record["ev_features"]),

            expected_ice_rows=len(record["ice_features"]),

        )

        _replace_with_retry(temporary_path, destination_path)

        if os.path.normcase(os.path.abspath(source_path)) != os.path.normcase(

            os.path.abspath(destination_path)

        ):

            _unlink_with_retry(source_path)

    except Exception:

        temporary_path.unlink(missing_ok=True)

        raise





def _clean_pair_index(pair_index: pd.DataFrame) -> pd.DataFrame:

    """Prepare pair_index without erasing upstream scientific/QC evidence.



    STRICT/NONSTRICT is assigned upstream by candidate_pairing.py and preserved

    by pair_cropping.py. This stage only normalises the stored label spelling;

    it never recalculates the class from dates, endpoints, route metrics,

    distance-review status, or sampling-review status.

    """



    result = pair_index.copy()



    if "final_pair_class" not in result.columns:

        raise ValueError("pair_index.csv is missing final_pair_class")



    result["final_pair_class"] = result["final_pair_class"].map(

        _normalise_pair_class

    )



    # Preserve all upstream warning/review/QC columns and their text. Only an

    # obsolete error field owned by this same stage may be removed after a

    # successful rerun.

    drop_after_success = [

        column

        for column in result.columns

        if _normalise_key(column) == "featureengineeringerror"

    ]

    if drop_after_success:

        result = result.drop(columns=drop_after_success)



    return result





def _preflight(

    pair_index: pd.DataFrame,

    results_root: Path,

    rules: dict,

) -> list[dict]:

    """Calculate every workbook before any workbook is modified."""



    records: list[dict] = []

    problems: list[str] = []

    resolved_paths: set[str] = set()

    destination_paths: set[str] = set()



    for sequence, (index_position, row) in enumerate(pair_index.iterrows(), start=1):

        pair_id = _clean_text(row.get("pair_id")) or f"ROW_{sequence}"

        try:

            pair_class = _normalise_pair_class(row.get("final_pair_class"))

            source_path = _resolve_workbook_path(row, results_root)

            source_key = os.path.normcase(os.path.abspath(source_path))

            if source_key in resolved_paths:

                raise ValueError("The same pair workbook is referenced more than once")

            resolved_paths.add(source_key)

            _check_not_locked(source_path)



            ev_source, ice_source = _read_pair_data(source_path)

            ev_features, ev_audit = engineer_features(ev_source, rules)

            ice_features, ice_audit = engineer_features(ice_source, rules)



            desired_name = _desired_workbook_name(row, source_path)

            destination_path = source_path.with_name(desired_name)

            destination_key = os.path.normcase(os.path.abspath(destination_path))

            if destination_key in destination_paths:

                raise ValueError("Two pair records resolve to the same output workbook")

            destination_paths.add(destination_key)



            route_direction = _clean_text(row.get("route_direction"))

            pair_title = " | ".join(

                [pair_id, pair_class, route_direction or "Unknown Route"]

            )

            completed_text = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            records.append(

                {

                    "index_position": index_position,

                    "pair_id": pair_id,

                    "pair_class": pair_class,

                    "candidate_id": _clean_text(row.get("candidate_id")),

                    "route_direction": route_direction,

                    "source_path": source_path,

                    "destination_path": destination_path,

                    "desired_name": desired_name,

                    "pair_title": pair_title,

                    "ev_features": ev_features,

                    "ice_features": ice_features,

                    "ev_audit": ev_audit,

                    "ice_audit": ice_audit,

                    "info_updates": {

                        "final_pair_class": pair_class,

                        "pair_title": pair_title,

                        "pair_workbook_name": desired_name,

                        "pair_workbook_path": str(destination_path),

                        "feature_engineering_status": "COMPLETED",

                        "feature_engineering_version": SCRIPT_VERSION,

                        "feature_engineering_completed_at": completed_text,

                        "feature_engineering_columns": " | ".join(OUTPUT_COLUMNS),

                        "feature_interval_convention": (

                            "PREVIOUS_DATAPOINT_TO_CURRENT_DATAPOINT"

                        ),

                        "feature_acceleration_formula": (

                            "(CURRENT_SPEED_MPS-PREVIOUS_SPEED_MPS)/"

                            "TIME_DIFFERENCE_SECONDS"

                        ),

                        "first_row_interval_values": (

                            "TIME=0 | DISTANCE=0 | ACCELERATION=0"

                        ),

                        "datapoint_row_colours": "NONE",

                        "phase_segmentation_status": "NOT_STARTED",

                    },

                }

            )

        except Exception as error:

            problems.append(f"{pair_id}: {error}")



        if sequence % 10 == 0 or sequence == len(pair_index):

            print(f"  Preflight checked {sequence}/{len(pair_index)} workbooks")



    if problems:

        detail = "\n".join(f"  - {problem}" for problem in problems[:20])

        if len(problems) > 20:

            detail += f"\n  - ... and {len(problems) - 20} more"

        raise PreflightError(

            "Preflight stopped before changing any workbook:\n" + detail

        )

    if len(records) != len(pair_index):

        raise PreflightError(

            f"Preflight prepared {len(records)} of {len(pair_index)} pair workbooks"

        )

    return records





def _atomic_write_csv(path: Path, dataframe: pd.DataFrame) -> None:

    path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile(

        prefix=f".{path.stem}_",

        suffix=".csv",

        dir=path.parent,

        delete=False,

    ) as temporary_file:

        temporary_path = Path(temporary_file.name)

    try:

        dataframe.to_csv(temporary_path, index=False, encoding="utf-8-sig")

        _replace_with_retry(temporary_path, path)

    except Exception:

        temporary_path.unlink(missing_ok=True)

        raise





def _atomic_write_excel(

    path: Path,

    tables: dict[str, pd.DataFrame],

    colour_pair_classes: bool = False,

) -> None:

    path.parent.mkdir(parents=True, exist_ok=True)

    with NamedTemporaryFile(

        prefix=f".{path.stem}_",

        suffix=".xlsx",

        dir=path.parent,

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

                worksheet = writer.sheets[sheet_name]

                _format_table_sheet(worksheet, dataframe)

                if colour_pair_classes:

                    _apply_pair_class_colours(worksheet, dataframe)

        _replace_with_retry(temporary_path, path)

    except Exception:

        temporary_path.unlink(missing_ok=True)

        raise





def _rules_table(rules: dict) -> pd.DataFrame:

    return pd.DataFrame(

        [

            {"Rule": "Output columns", "Value": " | ".join(OUTPUT_COLUMNS)},

            {

                "Rule": "Interval convention",

                "Value": "Previous datapoint to current datapoint",

            },

            {

                "Rule": "First row",

                "Value": "Time difference = 0; distance = 0; acceleration = 0",

            },

            {

                "Rule": "Distance",

                "Value": f"Haversine; Earth radius = {EARTH_RADIUS_M} m",

            },

            {

                "Rule": "Acceleration",

                "Value": "(current speed m/s - previous speed m/s) / time difference s",

            },

            {

                "Rule": "Speed range",

                "Value": (

                    f"{rules['speed_band_width_kmh']:g} km/h bands; "

                    f"{rules['speed_band_top_kmh']:g}+ upper band"

                ),

            },

            {

                "Rule": "Speed compatibility",

                "Value": (

                    "Coalesce all recognized m/s and km/h columns; complete only "

                    "unavailable cells by linear interpolation, GPS interval, "

                    "nearest recorded value, then trip median"

                ),

            },

            {"Rule": "Rows removed", "Value": "0"},

            {"Rule": "Datapoint row colours", "Value": "NONE"},

            {"Rule": "Normalization/standardization", "Value": "NOT APPLIED"},

            {"Rule": "Pair classes", "Value": "PRESERVED FROM pair_index.csv; STRICT or NONSTRICT only"},

            {"Rule": "Phase segmentation", "Value": "NOT PERFORMED IN STAGE 3"},

        ]

    )





def run_feature_engineering(config: dict, pair_index_path=None) -> dict:

    rules = _settings(config)

    results_root = Path(config["project"]["results_folder"])

    pair_index_path = Path(

        pair_index_path or results_root / "System" / PAIR_INDEX_NAME

    )

    if not pair_index_path.exists():

        raise FileNotFoundError(

            f"Pair index not found: {pair_index_path}. Run pair_cropping.py first."

        )



    pair_index = _clean_pair_index(_read_csv(pair_index_path))

    if pair_index.empty:

        raise ValueError("pair_index.csv is empty")

    required = ["pair_id", "final_pair_class", "pair_workbook_name"]

    missing = [column for column in required if column not in pair_index.columns]

    if missing:

        raise ValueError("pair_index.csv is missing: " + ", ".join(missing))



    pair_classes = pair_index["final_pair_class"]

    pair_class_snapshot = pair_classes.copy(deep=True)

    strict_count = int((pair_classes == "STRICT").sum())

    nonstrict_count = int((pair_classes == "NONSTRICT").sum())

    if strict_count + nonstrict_count != len(pair_index):

        raise ValueError("Every pair must be either STRICT or NONSTRICT")



    print("Preflight: reading and calculating every pair workbook")

    records = _preflight(pair_index, results_root, rules)

    print(

        f"Preflight passed: {len(records)} workbooks; "

        f"{strict_count} STRICT + {nonstrict_count} NONSTRICT\n"

    )



    started_at = datetime.now()

    pair_audit_rows = []

    vehicle_audit_rows = []



    for sequence, record in enumerate(records, start=1):

        _write_pair_workbook(record)

        completed_text = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

        index_position = record["index_position"]

        ev_audit = record["ev_audit"]

        ice_audit = record["ice_audit"]



        pair_index.at[index_position, "pair_workbook_name"] = record["desired_name"]

        pair_index.at[index_position, "pair_workbook_path"] = str(

            record["destination_path"]

        )

        pair_index.at[index_position, "pair_title"] = record["pair_title"]

        pair_index.at[index_position, "feature_engineering_version"] = SCRIPT_VERSION

        pair_index.at[index_position, "feature_engineering_status"] = "COMPLETED"

        pair_index.at[index_position, "feature_engineering_completed_at"] = (

            completed_text

        )

        pair_index.at[index_position, "ev_datapoints_processed"] = ev_audit[

            "datapoint_count"

        ]

        pair_index.at[index_position, "ice_datapoints_processed"] = ice_audit[

            "datapoint_count"

        ]

        pair_index.at[index_position, "ev_time_difference_sum_seconds"] = ev_audit[

            "time_difference_sum_seconds"

        ]

        pair_index.at[index_position, "ice_time_difference_sum_seconds"] = ice_audit[

            "time_difference_sum_seconds"

        ]

        pair_index.at[index_position, "ev_haversine_distance_sum_m"] = ev_audit[

            "haversine_distance_sum_m"

        ]

        pair_index.at[index_position, "ice_haversine_distance_sum_m"] = ice_audit[

            "haversine_distance_sum_m"

        ]



        pair_audit_rows.append(

            {

                "pair_id": record["pair_id"],

                "final_pair_class": record["pair_class"],

                "candidate_id": record["candidate_id"],

                "route_direction": record["route_direction"],

                "pair_workbook_name": record["desired_name"],

                "pair_workbook_path": str(record["destination_path"]),

                "ev_datapoints": ev_audit["datapoint_count"],

                "ice_datapoints": ice_audit["datapoint_count"],

                "ev_time_difference_sum_seconds": ev_audit[

                    "time_difference_sum_seconds"

                ],

                "ice_time_difference_sum_seconds": ice_audit[

                    "time_difference_sum_seconds"

                ],

                "ev_haversine_distance_sum_m": ev_audit[

                    "haversine_distance_sum_m"

                ],

                "ice_haversine_distance_sum_m": ice_audit[

                    "haversine_distance_sum_m"

                ],

                "feature_engineering_status": "COMPLETED",

                "verified_after_save": "YES",

                "processing_completed_at": completed_text,

            }

        )



        for vehicle_type, audit in (("EV", ev_audit), ("ICE", ice_audit)):

            vehicle_audit_rows.append(

                {

                    "pair_id": record["pair_id"],

                    "final_pair_class": record["pair_class"],

                    "vehicle_type": vehicle_type,

                    "pair_workbook_name": record["desired_name"],

                    **audit,

                    "feature_engineering_status": "COMPLETED",

                    "verified_after_save": "YES",

                }

            )



        if sequence % 5 == 0 or sequence == len(records):

            print(f"  Saved and verified {sequence}/{len(records)} workbooks")



    # Feature engineering must never redefine STRICT/NONSTRICT.

    if not pair_index["final_pair_class"].equals(pair_class_snapshot):

        raise RuntimeError(

            "STRICT/NONSTRICT classification changed during feature engineering"

        )



    completed_at = datetime.now()

    pair_audit = pd.DataFrame(pair_audit_rows)

    vehicle_audit = pd.DataFrame(vehicle_audit_rows)

    run_summary = pd.DataFrame(

        [

            {"Metric": "Script version", "Value": SCRIPT_VERSION},

            {

                "Metric": "Run started",

                "Value": started_at.strftime("%d-%m-%Y %H:%M:%S"),

            },

            {

                "Metric": "Run completed",

                "Value": completed_at.strftime("%d-%m-%Y %H:%M:%S"),

            },

            {"Metric": "Pair-index rows received", "Value": len(pair_index)},

            {"Metric": "Pair workbooks completed", "Value": len(records)},

            {"Metric": "Pair workbooks verified", "Value": len(records)},

            {"Metric": "STRICT pairs retained", "Value": strict_count},

            {"Metric": "NONSTRICT pairs retained", "Value": nonstrict_count},

            {

                "Metric": "EV/ICE sheets processed",

                "Value": len(vehicle_audit),

            },

            {

                "Metric": "Cropped datapoints processed",

                "Value": int(vehicle_audit["datapoint_count"].sum()),

            },

            {"Metric": "Datapoints excluded", "Value": 0},

            {"Metric": "Blank cells in required feature columns", "Value": 0},

            {"Metric": "Datapoint row colours", "Value": "NONE"},

            {

                "Metric": "Normalization/standardization",

                "Value": "NOT APPLIED",

            },

            {"Metric": "Next stage", "Value": "phase_segmentation.py"},

        ]

    )



    log_path = results_root / "Logs" / FEATURE_LOG_NAME

    pair_index_workbook_path = results_root / "System" / PAIR_INDEX_WORKBOOK_NAME

    _atomic_write_csv(pair_index_path, pair_index)

    _atomic_write_excel(

        pair_index_workbook_path,

        {"Pair Index": pair_index},

        colour_pair_classes=True,

    )

    _atomic_write_excel(

        log_path,

        {

            "Run Summary": run_summary,

            "Pair Feature Audit": pair_audit,

            "Vehicle Feature Audit": vehicle_audit,

            "Rules Applied": _rules_table(rules),

        },

        colour_pair_classes=True,

    )



    return {

        "pair_count": len(pair_index),

        "completed_count": len(records),

        "strict_count": strict_count,

        "nonstrict_count": nonstrict_count,

        "vehicle_sheet_count": len(vehicle_audit),

        "datapoints_processed": int(vehicle_audit["datapoint_count"].sum()),

        "log_path": str(log_path),

        "pair_index_path": str(pair_index_path),

        "pair_index_workbook_path": str(pair_index_workbook_path),

    }





if __name__ == "__main__":

    try:

        settings = load_config()

        print(f"Script version: {SCRIPT_VERSION}")

        print("Stage 5 feature engineering started")

        print("Target: feature columns plus preserved Address; no datapoint colours\n")

        result = run_feature_engineering(settings)

        print("\nFeature engineering completed and verified")

        print(f"Pair-index rows received: {result['pair_count']}")

        print(f"Pair workbooks completed: {result['completed_count']}")

        print(f"STRICT pairs retained: {result['strict_count']}")

        print(f"NONSTRICT pairs retained: {result['nonstrict_count']}")

        print(f"EV/ICE cropped-data sheets processed: {result['vehicle_sheet_count']}")

        print(f"Cropped datapoints processed: {result['datapoints_processed']}")

        print("Datapoints excluded: 0")

        print("Blank cells in required feature columns: 0")

        print("Datapoint row colours: NONE")

        print("Normalization/standardization applied: NO")

        print(f"Feature engineering audit: {result['log_path']}")

        print(f"Updated pair index: {result['pair_index_path']}")

        print(f"Pair index workbook: {result['pair_index_workbook_path']}")

        print("\nNext stage: phase_segmentation.py")

    except Exception as error:

        print(f"\nFEATURE ENGINEERING STOPPED | {error}")

        raise SystemExit(1)