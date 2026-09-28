"""
STAGE 4 DIAGNOSTIC TEST
=======================

Purpose
-------
Compare manually entered Stop / Charging information from:

    D:\My Desktop\TAT Data Collection Intercity.xlsx

against the actual telemetry contained in the current Stage-3 pair files:

    D:\TAT_MS_INTERCITY\Data\Results\Paired

This is a DIAGNOSTIC ONLY script.
It does NOT modify Phase Segmentation or any existing pipeline result.

Project rules for this diagnostic
---------------------------------
1. Manual charging/refuelling information is authoritative for identifying
   a replenishment event.
2. The manual charging timestamp is an approximate event time. Because the
   telemetry may lag, the script searches around the manual time.
3. Manual Stop Timestamps are approximate START timestamps only.
4. The telemetry is used to find the actual stationary episode and therefore
   determine the actual end and duration.
5. Stationary duration:
       < 3 minutes  -> IDLE
       >= 3 minutes -> STOP
6. Charging/refuelling has priority over STOP/IDLE.
7. The diagnostic does not assume a fixed lag. It MEASURES the lag.
8. The script uses the pair workbook's 03_Pair_Crop_Info metadata and the
   actual cropped telemetry. It does NOT require the vehicle number to be
   present in the PAIR filename.

Output
------
Data\Results\Diagnostics\Manual_Event_Telemetry_Lag_Test.xlsx

Sheets:
    Event_Matching
    Trip_Matching
    Lag_Summary
    Classification_Summary
    Unmatched_Manual_Records
    Method_Notes
"""

from __future__ import annotations

import math
import re
from pathlib import Path
from datetime import datetime, date, time

import numpy as np
import pandas as pd


# =====================================================================
# PATHS
# =====================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

MANUAL_EXCEL = Path(r"D:\My Desktop\TAT Data Collection Intercity.xlsx")

RESULTS_PAIRED = Path(r"D:\TAT_MS_INTERCITY\Data\Results\Paired")

OUTPUT_DIR = PROJECT_ROOT / "Data" / "Results" / "Diagnostics"
OUTPUT_FILE = OUTPUT_DIR / "Manual_Event_Telemetry_Lag_Test.xlsx"

# The intern timestamp may not equal the telemetry timestamp.
# Start with a deliberately broad diagnostic window.
SEARCH_BEFORE_MIN = 90
SEARCH_AFTER_MIN = 90

# Vehicle is treated as stationary at <= 0.5 km/h.
STATIONARY_SPEED_KMH = 0.5

# User's classification rule:
# less than 3 minutes = IDLE
# 3 minutes or more = STOP
IDLE_THRESHOLD_SECONDS = 180.0


# =====================================================================
# GENERAL HELPERS
# =====================================================================

def clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_text(value) -> str:
    text = clean_text(value).upper()
    text = text.replace("&", "AND")
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def normalize_vehicle(value) -> str:
    return normalize_text(value)


def parse_date_value(value):
    if value is None:
        return None

    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    text = clean_text(value)
    if not text:
        return None

    for fmt in (
        "%d-%m-%y",
        "%d-%m-%Y",
        "%d/%m/%y",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def parse_time_value(value):
    if value is None:
        return None

    if isinstance(value, pd.Timestamp):
        return value.time()
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    # Excel time stored as fraction of a day.
    if isinstance(value, (float, int)) and not isinstance(value, bool):
        x = float(value)
        if 0 <= x < 1:
            total_seconds = int(round(x * 86400)) % 86400
            return time(
                total_seconds // 3600,
                (total_seconds % 3600) // 60,
                total_seconds % 60,
            )

    text = clean_text(value)
    if not text:
        return None

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            pass

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.time()


def combine_date_time(d, t):
    if d is None or t is None:
        return pd.NaT
    return pd.Timestamp(datetime.combine(d, t))


def parse_number(value):
    try:
        x = pd.to_numeric(value, errors="coerce")
        if pd.isna(x):
            return np.nan
        return float(x)
    except Exception:
        return np.nan


def haversine_m(lat1, lon1, lat2, lon2):
    """
    Vector-friendly haversine distance in metres.
    """
    r = 6371000.0

    lat1 = np.asarray(lat1, dtype=float)
    lon1 = np.asarray(lon1, dtype=float)
    lat2 = np.asarray(lat2, dtype=float)
    lon2 = np.asarray(lon2, dtype=float)

    p1 = np.radians(lat1)
    p2 = np.radians(lat2)
    dp = np.radians(lat2 - lat1)
    dl = np.radians(lon2 - lon1)

    a = (
        np.sin(dp / 2.0) ** 2
        + np.cos(p1) * np.cos(p2) * np.sin(dl / 2.0) ** 2
    )

    return 2.0 * r * np.arcsin(np.sqrt(np.clip(a, 0, 1)))


# =====================================================================
# ROUTE NORMALIZATION
# =====================================================================

def route_direction_from_manual(row):
    start = clean_text(row.get("Start Location"))
    end = clean_text(row.get("End Location"))

    if not start and not end:
        return ""

    return f"{start} to {end}"


def route_key(route):
    """
    Normalize route names enough to compare common variants such as:
        Chennai -> Pondicherry
        Chennai to Pondy
        Pondicherry -> Chennai
    """
    text = normalize_text(route)

    # Common spelling variants in this project.
    replacements = [
        ("PONDICHERRY", "PONDY"),
        ("PONDICHERY", "PONDY"),
        ("PONDY", "PONDY"),
        ("MADURAVAYOL", "MADURAVOYAL"),
        ("MADURAVOYAL", "MADURAVOYAL"),
        ("AVRROUNDANASALEM", "AVRROUNDANASALEM"),
        ("SILKBOARD", "SILKBOARD"),
        ("CHENNAI", "CHENNAI"),
    ]

    for old, new in replacements:
        text = text.replace(old, new)

    return text


def route_match_score(manual_start, manual_end, pair_route):
    manual_route = route_key(f"{manual_start} to {manual_end}")
    pair_route_key = route_key(pair_route)

    if not manual_route or not pair_route_key:
        return 0

    if manual_route == pair_route_key:
        return 100

    # Direction-specific endpoint token comparison.
    m_start = route_key(manual_start)
    m_end = route_key(manual_end)
    p = pair_route_key

    score = 0
    if m_start and m_start in p:
        score += 35
    if m_end and m_end in p:
        score += 35

    return score


# =====================================================================
# MANUAL STOP PARSING
# =====================================================================

TIME_PATTERN = re.compile(
    r"(?:(\d+)\s*\([^)]*\)\s*)?"
    r"(\d{1,2}:\d{2}(?::\d{2})?)"
    r"\s*(?:[-–—]\s*(\d{1,2}:\d{2}(?::\d{2})?))?"
)


def parse_stop_entries(value, trip_date):
    """
    Parse manual stop strings.

    Examples:
        1 (11:45), 2 (12:02), ...
        5 (03:13–03:29)
        (09:40, 09:59, 10:36)

    IMPORTANT:
    Stop timestamps are interpreted sequentially.
    If the clock time goes backwards, the calendar date is advanced by
    one day. This handles trips crossing midnight.
    """

    text = clean_text(value)
    if not text:
        return []

    raw_events = []

    for match in TIME_PATTERN.finditer(text):
        label = match.group(1) or ""
        start_s = match.group(2)
        end_s = match.group(3)

        try:
            start_t = datetime.strptime(start_s, "%H:%M").time()
        except ValueError:
            try:
                start_t = datetime.strptime(start_s, "%H:%M:%S").time()
            except ValueError:
                continue

        end_t = None
        if end_s:
            try:
                end_t = datetime.strptime(end_s, "%H:%M").time()
            except ValueError:
                try:
                    end_t = datetime.strptime(end_s, "%H:%M:%S").time()
                except ValueError:
                    end_t = None

        raw_events.append((label, start_t, end_t))

    events = []
    current_date = trip_date
    previous_start = None

    for number, (label, start_t, end_t) in enumerate(raw_events, start=1):
        start_dt = combine_date_time(current_date, start_t)

        # If the sequence goes backwards in clock time, it crossed midnight.
        if previous_start is not None and start_dt < previous_start:
            current_date = current_date + pd.Timedelta(days=1)
            current_date = current_date.date() if hasattr(current_date, "date") else current_date
            start_dt = combine_date_time(current_date, start_t)

        end_dt = pd.NaT
        if end_t is not None:
            end_dt = combine_date_time(current_date, end_t)
            if end_dt < start_dt:
                end_dt += pd.Timedelta(days=1)

        events.append({
            "event_no": label or number,
            "manual_start": start_dt,
            "manual_end": end_dt,
            "explicit_interval": not pd.isna(end_dt),
        })

        previous_start = start_dt

    return events


# =====================================================================
# LOAD MANUAL TRIP RECORDS AND EVENTS
# =====================================================================

def load_manual_workbook():
    if not MANUAL_EXCEL.exists():
        raise FileNotFoundError(
            f"Manual Excel not found:\n{MANUAL_EXCEL}"
        )

    df = pd.read_excel(MANUAL_EXCEL)
    df.columns = [clean_text(c) for c in df.columns]

    required = [
        "Date",
        "Vehicle Number",
        "Vehicle Type",
        "Start Location",
        "End Location",
        "Start time",
        "End Time",
        "Charging time (min)",
        "charging timestamp",
        "stop timestamps",
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "The manual Excel is missing required columns:\n"
            + "\n".join(f"- {x}" for x in missing)
        )

    return df


def build_manual_trip_records(df):
    trips = []

    for idx, row in df.iterrows():
        trip_date = parse_date_value(row["Date"])

        if trip_date is None:
            continue

        start_time = parse_time_value(row["Start time"])
        end_time = parse_time_value(row["End Time"])

        trip_start = combine_date_time(trip_date, start_time)
        trip_end = combine_date_time(trip_date, end_time)

        if not pd.isna(trip_start) and not pd.isna(trip_end):
            if trip_end < trip_start:
                trip_end += pd.Timedelta(days=1)

        vehicle_type = clean_text(row["Vehicle Type"]).upper()

        charging_duration = parse_number(row["Charging time (min)"])
        charging_clock = parse_time_value(row["charging timestamp"])

        charging_start = combine_date_time(trip_date, charging_clock)
        if (
            not pd.isna(trip_start)
            and not pd.isna(charging_start)
            and charging_start < trip_start - pd.Timedelta(hours=6)
        ):
            charging_start += pd.Timedelta(days=1)

        charging_end = pd.NaT
        if not pd.isna(charging_start) and not pd.isna(charging_duration):
            charging_end = charging_start + pd.Timedelta(
                minutes=float(charging_duration)
            )

        stop_events = parse_stop_entries(
            row["stop timestamps"],
            trip_date,
        )

        trips.append({
            "manual_row": idx + 2,
            "vehicle": normalize_vehicle(row["Vehicle Number"]),
            "vehicle_raw": clean_text(row["Vehicle Number"]),
            "vehicle_type": vehicle_type,
            "date": trip_date,
            "route_start": clean_text(row["Start Location"]),
            "route_end": clean_text(row["End Location"]),
            "route": route_direction_from_manual(row),
            "manual_trip_start": trip_start,
            "manual_trip_end": trip_end,
            "charging_duration_min": charging_duration,
            "charging_start": charging_start,
            "charging_end": charging_end,
            "charging_timestamp_present": not pd.isna(charging_start),
            "stop_events": stop_events,
            "raw_stop_text": clean_text(row["stop timestamps"]),
        })

    return trips


# =====================================================================
# PAIR WORKBOOK METADATA
# =====================================================================

def read_pair_info(pair_path):
    try:
        info_df = pd.read_excel(
            pair_path,
            sheet_name="03_Pair_Crop_Info",
        )
    except Exception as exc:
        return {
            "error": f"Could not read 03_Pair_Crop_Info: {exc}"
        }

    if info_df.empty:
        return {}

    info_df.columns = [clean_text(c) for c in info_df.columns]

    # Expected format: Parameter | Value
    if "Parameter" in info_df.columns and "Value" in info_df.columns:
        return {
            clean_text(row["Parameter"]): row["Value"]
            for _, row in info_df.iterrows()
        }

    # Fallback for a two-column sheet with different headers.
    if len(info_df.columns) >= 2:
        c1, c2 = info_df.columns[:2]
        return {
            clean_text(row[c1]): row[c2]
            for _, row in info_df.iterrows()
        }

    return {}


def parse_info_timestamp(value):
    if value is None:
        return pd.NaT

    parsed = pd.to_datetime(value, errors="coerce")
    return parsed


def pair_side_metadata(pair_path, info, side):
    prefix = "ev" if side == "EV" else "ice"

    cropped_start_key = f"{prefix}_cropped_start_timestamp"
    cropped_end_key = f"{prefix}_cropped_end_timestamp"

    verified_start_key = f"{prefix}_verified_start"
    verified_end_key = f"{prefix}_verified_end"

    file_key = f"{prefix}_file_name"
    clean_key = f"{prefix}_cleaned_file_path"

    return {
        "pair_path": str(pair_path),
        "pair_name": pair_path.name,
        "side": side,
        "vehicle_type": "EV" if side == "EV" else "DIESEL",
        "pair_id": clean_text(info.get("pair_id")),
        "final_pair_class": clean_text(info.get("final_pair_class")),
        "route_direction": clean_text(info.get("route_direction")),
        "source_file_name": clean_text(info.get(file_key)),
        "cleaned_file_path": clean_text(info.get(clean_key)),
        "verified_start": parse_info_timestamp(info.get(verified_start_key)),
        "verified_end": parse_info_timestamp(info.get(verified_end_key)),
        "cropped_start": parse_info_timestamp(info.get(cropped_start_key)),
        "cropped_end": parse_info_timestamp(info.get(cropped_end_key)),
    }


def discover_pair_sides():
    if not RESULTS_PAIRED.exists():
        raise FileNotFoundError(
            f"Paired results folder not found:\n{RESULTS_PAIRED}"
        )

    pair_files = sorted(RESULTS_PAIRED.rglob("PAIR_*.xlsx"))

    # Exclude temporary Excel files.
    pair_files = [
        p for p in pair_files
        if not p.name.startswith("~$")
    ]

    sides = []

    for pair_path in pair_files:
        info = read_pair_info(pair_path)

        if "error" in info:
            print(f"WARNING: {pair_path.name}: {info['error']}")
            continue

        for side in ("EV", "ICE"):
            meta = pair_side_metadata(pair_path, info, side)

            if pd.isna(meta["cropped_start"]):
                # Fall back to verified start if necessary.
                meta["cropped_start"] = meta["verified_start"]

            if pd.isna(meta["cropped_end"]):
                meta["cropped_end"] = meta["verified_end"]

            sides.append(meta)

    return sides


# =====================================================================
# MANUAL TRIP -> PAIR SIDE MATCHING
# =====================================================================

def trip_match_score(manual_trip, side):
    score = 0

    # Vehicle type is a hard compatibility condition.
    mtype = manual_trip["vehicle_type"]
    stype = side["vehicle_type"]

    if "EV" in mtype and stype == "EV":
        score += 100
    elif "DIESEL" in mtype and stype == "DIESEL":
        score += 100
    elif "ICE" in mtype and stype == "DIESEL":
        score += 100
    elif "PETROL" in mtype and stype == "DIESEL":
        # Not normally expected, but do not make it impossible.
        score += 20
    else:
        return -999

    # Date is very strong evidence.
    mdate = manual_trip["date"]
    side_start = side["cropped_start"]

    if not pd.isna(side_start):
        if side_start.date() == mdate:
            score += 80
        elif side_start.date() == mdate + pd.Timedelta(days=1):
            # Usually only possible if the actual trip begins after midnight.
            score += 20
        else:
            return -999

    # Route/direction.
    rscore = route_match_score(
        manual_trip["route_start"],
        manual_trip["route_end"],
        side["route_direction"],
    )
    score += rscore

    # Manual trip start vs actual cropped start.
    if (
        not pd.isna(manual_trip["manual_trip_start"])
        and not pd.isna(side["cropped_start"])
    ):
        # Compare clock time on the same trip date, allowing midnight wrap.
        a = manual_trip["manual_trip_start"]
        b = side["cropped_start"]

        # If dates are close, calculate absolute datetime difference.
        diff_min = abs((b - a).total_seconds()) / 60.0

        # Penalize very large differences, but don't reject automatically.
        if diff_min <= 30:
            score += 40
        elif diff_min <= 90:
            score += 25
        elif diff_min <= 180:
            score += 10
        else:
            score -= min(30, int(diff_min / 60))

    return score


def match_manual_trips_to_pairs(manual_trips, pair_sides):
    rows = []
    assignments = {}

    for trip in manual_trips:
        candidates = []

        for side in pair_sides:
            score = trip_match_score(trip, side)
            if score > 0:
                candidates.append((score, side))

        candidates.sort(
            key=lambda x: (
                x[0],
                -abs(
                    (
                        x[1]["cropped_start"] - trip["manual_trip_start"]
                    ).total_seconds()
                )
                if not pd.isna(x[1]["cropped_start"])
                and not pd.isna(trip["manual_trip_start"])
                else 0,
            ),
            reverse=True,
        )

        if candidates:
            best_score, best_side = candidates[0]

            # If there is a close tie, flag it rather than pretending
            # the match is certain.
            ambiguous = False
            if len(candidates) > 1:
                second_score = candidates[1][0]
                if best_score - second_score <= 10:
                    ambiguous = True

            key = (trip["manual_row"], best_side["pair_path"], best_side["side"])
            assignments[key] = trip

            rows.append({
                "manual_row": trip["manual_row"],
                "vehicle": trip["vehicle_raw"],
                "vehicle_type": trip["vehicle_type"],
                "manual_date": trip["date"],
                "manual_route": trip["route"],
                "manual_trip_start": trip["manual_trip_start"],
                "manual_trip_end": trip["manual_trip_end"],
                "matched_pair": best_side["pair_name"],
                "matched_side": best_side["side"],
                "pair_route": best_side["route_direction"],
                "pair_cropped_start": best_side["cropped_start"],
                "pair_cropped_end": best_side["cropped_end"],
                "trip_match_score": best_score,
                "trip_match_status": (
                    "AMBIGUOUS_BEST_MATCH" if ambiguous else "MATCHED"
                ),
            })
        else:
            rows.append({
                "manual_row": trip["manual_row"],
                "vehicle": trip["vehicle_raw"],
                "vehicle_type": trip["vehicle_type"],
                "manual_date": trip["date"],
                "manual_route": trip["route"],
                "manual_trip_start": trip["manual_trip_start"],
                "manual_trip_end": trip["manual_trip_end"],
                "matched_pair": "",
                "matched_side": "",
                "pair_route": "",
                "pair_cropped_start": pd.NaT,
                "pair_cropped_end": pd.NaT,
                "trip_match_score": 0,
                "trip_match_status": "NO_PAIR_MATCH",
            })

    return pd.DataFrame(rows), assignments


# =====================================================================
# TELEMETRY
# =====================================================================

def load_telemetry(pair_path, side):
    sheet = (
        "01_EV_Cropped_Data"
        if side == "EV"
        else "02_ICE_Cropped_Data"
    )

    try:
        df = pd.read_excel(pair_path, sheet_name=sheet)
    except Exception as exc:
        return pd.DataFrame(), f"Could not read {sheet}: {exc}"

    if df.empty:
        return df, "Sheet is empty."

    df.columns = [clean_text(c) for c in df.columns]

    rename = {}

    for c in df.columns:
        lc = c.lower().strip()

        if lc == "timestamp":
            rename[c] = "Timestamp"
        elif lc in {"speed (km/h)", "speed_kmh", "speed km/h", "speed"}:
            rename[c] = "Speed_kmh"
        elif lc in {"latitude", "lat"}:
            rename[c] = "Latitude"
        elif lc in {"longitude", "lon", "lng"}:
            rename[c] = "Longitude"
        elif lc in {"time difference (s)", "time_difference_s"}:
            rename[c] = "Time_Difference_s"
        elif lc in {"address", "reverse geocoded address"}:
            rename[c] = "Address"

    df = df.rename(columns=rename)

    if "Timestamp" not in df.columns:
        return pd.DataFrame(), "Timestamp column missing."

    df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
    df = df.dropna(subset=["Timestamp"]).sort_values("Timestamp").reset_index(drop=True)

    if "Speed_kmh" not in df.columns:
        df["Speed_kmh"] = np.nan
    else:
        df["Speed_kmh"] = pd.to_numeric(df["Speed_kmh"], errors="coerce")

    for c in ("Latitude", "Longitude"):
        if c not in df.columns:
            df[c] = np.nan
        else:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    if "Time_Difference_s" in df.columns:
        df["Time_Difference_s"] = pd.to_numeric(
            df["Time_Difference_s"], errors="coerce"
        )
    else:
        df["Time_Difference_s"] = (
            df["Timestamp"].diff().dt.total_seconds()
        )

    df["Derived_Gap_s"] = df["Timestamp"].diff().dt.total_seconds()

    # Derive GPS movement.
    df["Gap_Distance_m"] = np.nan

    if len(df) > 1:
        df.loc[1:, "Gap_Distance_m"] = haversine_m(
            df.loc[1:, "Latitude"].to_numpy(),
            df.loc[1:, "Longitude"].to_numpy(),
            df.loc[:-1, "Latitude"].to_numpy(),
            df.loc[:-1, "Longitude"].to_numpy(),
        )

    df["Implied_Gap_Speed_kmh"] = np.where(
        (df["Derived_Gap_s"] > 0) & df["Gap_Distance_m"].notna(),
        df["Gap_Distance_m"] / df["Derived_Gap_s"] * 3.6,
        np.nan,
    )

    return df, ""


def stationary_mask(df):
    speed = pd.to_numeric(df["Speed_kmh"], errors="coerce")
    implied = pd.to_numeric(
        df["Implied_Gap_Speed_kmh"], errors="coerce"
    )

    return (
        speed.le(STATIONARY_SPEED_KMH)
        | (speed.isna() & implied.le(STATIONARY_SPEED_KMH))
    ).fillna(False)


def find_stationary_episodes(df):
    """
    Build stationary episodes from consecutive stationary telemetry rows.

    The diagnostic uses the actual timestamps. The final Phase Segmentation
    code can later use the project's full interval convention and gap logic.
    """
    if df.empty:
        return []

    mask = stationary_mask(df).to_numpy()

    episodes = []
    start_idx = None

    for i, stationary in enumerate(mask):
        if stationary and start_idx is None:
            start_idx = i

        is_last = i == len(mask) - 1

        if start_idx is not None and (not stationary or is_last):
            end_idx = i if stationary and is_last else i - 1

            if end_idx >= start_idx:
                start_ts = pd.Timestamp(
                    df.iloc[start_idx]["Timestamp"]
                )
                end_ts = pd.Timestamp(
                    df.iloc[end_idx]["Timestamp"]
                )

                duration_sec = max(
                    0.0,
                    (end_ts - start_ts).total_seconds()
                )

                first_row = df.iloc[start_idx]
                last_row = df.iloc[end_idx]

                episodes.append({
                    "start_idx": start_idx,
                    "end_idx": end_idx,
                    "start": start_ts,
                    "end": end_ts,
                    "duration_sec": duration_sec,
                    "duration_min": duration_sec / 60.0,
                    "rows": end_idx - start_idx + 1,
                    "start_lat": first_row.get("Latitude", np.nan),
                    "start_lon": first_row.get("Longitude", np.nan),
                    "end_lat": last_row.get("Latitude", np.nan),
                    "end_lon": last_row.get("Longitude", np.nan),
                    "start_address": clean_text(first_row.get("Address")),
                    "end_address": clean_text(last_row.get("Address")),
                })

            start_idx = None

    return episodes


def event_distance_to_episode_minutes(manual_time, episode):
    if pd.isna(manual_time):
        return np.inf

    # Distance to interval rather than only distance to start.
    if manual_time < episode["start"]:
        return (episode["start"] - manual_time).total_seconds() / 60.0

    if manual_time > episode["end"]:
        return (manual_time - episode["end"]).total_seconds() / 60.0

    return 0.0


def choose_episode(episodes, manual_start):
    if not episodes or pd.isna(manual_start):
        return None, "NO_STATIONARY_EPISODES"

    window_start = manual_start - pd.Timedelta(minutes=SEARCH_BEFORE_MIN)
    window_end = manual_start + pd.Timedelta(minutes=SEARCH_AFTER_MIN)

    candidates = [
        ep for ep in episodes
        if ep["end"] >= window_start
        and ep["start"] <= window_end
    ]

    if not candidates:
        return None, "NO_EPISODE_IN_SEARCH_WINDOW"

    # Primary: distance from manual timestamp to episode.
    # Secondary: longer episode.
    candidates.sort(
        key=lambda ep: (
            event_distance_to_episode_minutes(manual_start, ep),
            -ep["duration_sec"],
        )
    )

    best = candidates[0]

    # Detect an ambiguous case if another episode is almost equally close.
    status = "MATCHED"
    if len(candidates) > 1:
        d1 = event_distance_to_episode_minutes(manual_start, best)
        d2 = event_distance_to_episode_minutes(
            manual_start, candidates[1]
        )
        if abs(d2 - d1) <= 2.0:
            status = "AMBIGUOUS_EPISODE_MATCH"

    return best, status


# =====================================================================
# CHARGING PROXIMITY
# =====================================================================

def episode_overlap(ep, start, end):
    if pd.isna(start):
        return False

    if pd.isna(end):
        return (
            abs((ep["start"] - start).total_seconds()) <= 90 * 60
            or ep["start"] <= start <= ep["end"]
        )

    return ep["end"] >= start and ep["start"] <= end


def distance_between_points_m(lat1, lon1, lat2, lon2):
    if any(pd.isna(x) for x in [lat1, lon1, lat2, lon2]):
        return np.nan

    return float(
        haversine_m(
            np.array([lat1]),
            np.array([lon1]),
            np.array([lat2]),
            np.array([lon2]),
        )[0]
    )


def find_nearby_charging_episode(episodes, charging_start, charging_end):
    if pd.isna(charging_start):
        return None

    candidates = [
        ep for ep in episodes
        if episode_overlap(ep, charging_start, charging_end)
    ]

    if not candidates:
        # Also allow a broad search around the manual timestamp.
        candidates = [
            ep for ep in episodes
            if abs(
                (ep["start"] - charging_start).total_seconds()
            ) <= SEARCH_AFTER_MIN * 60
        ]

    if not candidates:
        return None

    candidates.sort(
        key=lambda ep: (
            event_distance_to_episode_minutes(charging_start, ep),
            -ep["duration_sec"],
        )
    )

    return candidates[0]


# =====================================================================
# EVENT ANALYSIS
# =====================================================================

def analyze_trip_events(trip, side, telemetry):
    """
    Analyze all charging and manual-stop events belonging to one matched
    manual trip and one Stage-3 EV/ICE telemetry side.
    """
    rows = []

    episodes = find_stationary_episodes(telemetry)

    # ---------------------------------------------------------------
    # Charging / refuelling
    # ---------------------------------------------------------------

    charging_duration = trip["charging_duration_min"]
    charging_start = trip["charging_start"]
    charging_end = trip["charging_end"]

    if not pd.isna(charging_duration):
        if pd.isna(charging_start):
            rows.append({
                "manual_row": trip["manual_row"],
                "vehicle": trip["vehicle_raw"],
                "vehicle_type": trip["vehicle_type"],
                "manual_date": trip["date"],
                "route": trip["route"],
                "event_type": (
                    "CHARGING" if side["side"] == "EV"
                    else "REFUELING"
                ),
                "event_no": 1,
                "manual_start": pd.NaT,
                "manual_end": pd.NaT,
                "manual_duration_min": charging_duration,
                "matched_pair": side["pair_name"],
                "matched_side": side["side"],
                "telemetry_start": pd.NaT,
                "telemetry_end": pd.NaT,
                "telemetry_duration_min": np.nan,
                "start_lag_min": np.nan,
                "classification": "MANUAL_DURATION_WITHOUT_TIMESTAMP",
                "detection_status": "NO_CHARGING_TIMESTAMP",
                "confidence": "LOW",
                "charging_proximity": "NOT_TESTABLE",
                "notes": (
                    "Charging duration exists in the manual Excel, "
                    "but charging timestamp is blank. The event cannot "
                    "be temporally matched by this diagnostic."
                ),
            })
        else:
            ep = find_nearby_charging_episode(
                episodes,
                charging_start,
                charging_end,
            )

            if ep is None:
                rows.append({
                    "manual_row": trip["manual_row"],
                    "vehicle": trip["vehicle_raw"],
                    "vehicle_type": trip["vehicle_type"],
                    "manual_date": trip["date"],
                    "route": trip["route"],
                    "event_type": (
                        "CHARGING" if side["side"] == "EV"
                        else "REFUELING"
                    ),
                    "event_no": 1,
                    "manual_start": charging_start,
                    "manual_end": charging_end,
                    "manual_duration_min": charging_duration,
                    "matched_pair": side["pair_name"],
                    "matched_side": side["side"],
                    "telemetry_start": pd.NaT,
                    "telemetry_end": pd.NaT,
                    "telemetry_duration_min": np.nan,
                    "start_lag_min": np.nan,
                    "classification": (
                        "CHARGING" if side["side"] == "EV"
                        else "REFUELING"
                    ),
                    "detection_status": "NO_STATIONARY_EPISODE_FOUND",
                    "confidence": "LOW",
                    "charging_proximity": "NO_TELEMETRY_MATCH",
                    "notes": (
                        "Manual charging/refuelling record exists, "
                        "but no stationary telemetry episode was found "
                        f"within +/- {SEARCH_BEFORE_MIN}/{SEARCH_AFTER_MIN} min."
                    ),
                })
            else:
                lag = (
                    ep["start"] - charging_start
                ).total_seconds() / 60.0

                rows.append({
                    "manual_row": trip["manual_row"],
                    "vehicle": trip["vehicle_raw"],
                    "vehicle_type": trip["vehicle_type"],
                    "manual_date": trip["date"],
                    "route": trip["route"],
                    "event_type": (
                        "CHARGING" if side["side"] == "EV"
                        else "REFUELING"
                    ),
                    "event_no": 1,
                    "manual_start": charging_start,
                    "manual_end": charging_end,
                    "manual_duration_min": charging_duration,
                    "matched_pair": side["pair_name"],
                    "matched_side": side["side"],
                    "telemetry_start": ep["start"],
                    "telemetry_end": ep["end"],
                    "telemetry_duration_min": ep["duration_min"],
                    "start_lag_min": lag,
                    "classification": (
                        "CHARGING" if side["side"] == "EV"
                        else "REFUELING"
                    ),
                    "detection_status": "MATCHED",
                    "confidence": (
                        "HIGH" if abs(lag) <= 15
                        else "MEDIUM" if abs(lag) <= 30
                        else "LOW"
                    ),
                    "charging_proximity": "STATIONARY_EPISODE_MATCHED",
                    "notes": (
                        "Charging/refuelling takes priority over "
                        "STOP/IDLE for the diagnostic."
                    ),
                })

    # ---------------------------------------------------------------
    # Manual stop timestamps
    # ---------------------------------------------------------------

    for stop in trip["stop_events"]:
        manual_start = stop["manual_start"]
        manual_end = stop["manual_end"]

        ep, status = choose_episode(
            episodes,
            manual_start,
        )

        if ep is None:
            rows.append({
                "manual_row": trip["manual_row"],
                "vehicle": trip["vehicle_raw"],
                "vehicle_type": trip["vehicle_type"],
                "manual_date": trip["date"],
                "route": trip["route"],
                "event_type": "MANUAL_STOP",
                "event_no": stop["event_no"],
                "manual_start": manual_start,
                "manual_end": manual_end,
                "manual_duration_min": (
                    (manual_end - manual_start).total_seconds() / 60.0
                    if not pd.isna(manual_end) else np.nan
                ),
                "matched_pair": side["pair_name"],
                "matched_side": side["side"],
                "telemetry_start": pd.NaT,
                "telemetry_end": pd.NaT,
                "telemetry_duration_min": np.nan,
                "start_lag_min": np.nan,
                "classification": "NOT_MATCHED",
                "detection_status": status,
                "confidence": "LOW",
                "charging_proximity": "NOT_TESTABLE",
                "notes": (
                    "Manual stop timestamp was used as the search reference."
                ),
            })
            continue

        lag = (
            ep["start"] - manual_start
        ).total_seconds() / 60.0

        # IMPORTANT:
        # For this diagnostic, a manually entered stop is classified by
        # telemetry duration unless it is actually a charging/refuelling
        # event. Charging is handled separately above.
        if ep["duration_sec"] < IDLE_THRESHOLD_SECONDS:
            classification = "IDLE"
        else:
            classification = "STOP"

        rows.append({
            "manual_row": trip["manual_row"],
            "vehicle": trip["vehicle_raw"],
            "vehicle_type": trip["vehicle_type"],
            "manual_date": trip["date"],
            "route": trip["route"],
            "event_type": "MANUAL_STOP",
            "event_no": stop["event_no"],
            "manual_start": manual_start,
            "manual_end": manual_end,
            "manual_duration_min": (
                (manual_end - manual_start).total_seconds() / 60.0
                if not pd.isna(manual_end) else np.nan
            ),
            "matched_pair": side["pair_name"],
            "matched_side": side["side"],
            "telemetry_start": ep["start"],
            "telemetry_end": ep["end"],
            "telemetry_duration_min": ep["duration_min"],
            "start_lag_min": lag,
            "classification": classification,
            "detection_status": status,
            "confidence": (
                "HIGH" if abs(lag) <= 15
                else "MEDIUM" if abs(lag) <= 30
                else "LOW"
            ),
            "charging_proximity": "NOT_CHARGING_RECORD",
            "notes": (
                "Manual stop timestamp used to locate the event; "
                "telemetry-derived duration determines STOP vs IDLE. "
                "Charging/refuelling records take priority when present."
            ),
        })

    return rows


# =====================================================================
# MAIN
# =====================================================================

def run_test():
    print("=" * 80)
    print("MANUAL STOP / CHARGING vs TELEMETRY LAG DIAGNOSTIC")
    print("=" * 80)

    # ---------------------------------------------------------------
    # 1. Load manual workbook
    # ---------------------------------------------------------------

    manual_df = load_manual_workbook()
    manual_trips = build_manual_trip_records(manual_df)

    print(f"Manual trip records loaded: {len(manual_trips)}")

    total_stop_events = sum(
        len(t["stop_events"])
        for t in manual_trips
    )

    charging_duration_records = sum(
        not pd.isna(t["charging_duration_min"])
        for t in manual_trips
    )

    charging_timestamp_records = sum(
        t["charging_timestamp_present"]
        for t in manual_trips
    )

    print(f"Manual stop timestamps parsed: {total_stop_events}")
    print(
        "Manual charging-duration records: "
        f"{charging_duration_records}"
    )
    print(
        "Manual charging-timestamp records: "
        f"{charging_timestamp_records}"
    )

    # ---------------------------------------------------------------
    # 2. Find current Stage-3 pair sides
    # ---------------------------------------------------------------

    pair_sides = discover_pair_sides()

    unique_pairs = sorted({
        x["pair_path"]
        for x in pair_sides
    })

    print(f"Pair workbooks found: {len(unique_pairs)}")
    print(f"Pair telemetry sides found: {len(pair_sides)}")

    if not pair_sides:
        raise RuntimeError(
            f"No usable PAIR_*.xlsx files found below:\n"
            f"{RESULTS_PAIRED}"
        )

    # ---------------------------------------------------------------
    # 3. Match manual trip records to pair sides
    # ---------------------------------------------------------------

    trip_match_df, assignments = match_manual_trips_to_pairs(
        manual_trips,
        pair_sides,
    )

    matched_count = int(
        (trip_match_df["trip_match_status"] != "NO_PAIR_MATCH").sum()
    )

    print(f"Manual trips matched to pair sides: {matched_count}")
    print(
        "Manual trips without a pair-side match: "
        f"{len(trip_match_df) - matched_count}"
    )

    # ---------------------------------------------------------------
    # 4. Load telemetry only for matched pair sides
    # ---------------------------------------------------------------

    telemetry_cache = {}
    event_rows = []

    unmatched_manual_rows = []

    for trip in manual_trips:
        # Find the assignment for this manual row.
        assignment_candidates = [
            (key, assigned_trip)
            for key, assigned_trip in assignments.items()
            if assigned_trip["manual_row"] == trip["manual_row"]
        ]

        if not assignment_candidates:
            unmatched_manual_rows.append({
                "manual_row": trip["manual_row"],
                "vehicle": trip["vehicle_raw"],
                "vehicle_type": trip["vehicle_type"],
                "date": trip["date"],
                "route": trip["route"],
                "reason": "NO_PAIR_SIDE_MATCH",
            })
            continue

        # A manual trip should normally have one selected pair side.
        key, _ = assignment_candidates[0]

        _, pair_path_str, side_name = key
        pair_path = Path(pair_path_str)

        cache_key = (str(pair_path), side_name)

        if cache_key not in telemetry_cache:
            telemetry, error = load_telemetry(
                pair_path,
                side_name,
            )
            telemetry_cache[cache_key] = (
                telemetry,
                error,
            )

        telemetry, error = telemetry_cache[cache_key]

        if telemetry.empty:
            unmatched_manual_rows.append({
                "manual_row": trip["manual_row"],
                "vehicle": trip["vehicle_raw"],
                "vehicle_type": trip["vehicle_type"],
                "date": trip["date"],
                "route": trip["route"],
                "reason": f"Telemetry unavailable: {error}",
            })
            continue

        side = next(
            x for x in pair_sides
            if x["pair_path"] == str(pair_path)
            and x["side"] == side_name
        )

        rows = analyze_trip_events(
            trip,
            side,
            telemetry,
        )

        event_rows.extend(rows)

    # ---------------------------------------------------------------
    # 5. Build result tables
    # ---------------------------------------------------------------

    event_df = pd.DataFrame(event_rows)

    if event_df.empty:
        event_df = pd.DataFrame(columns=[
            "manual_row",
            "vehicle",
            "vehicle_type",
            "manual_date",
            "route",
            "event_type",
            "event_no",
            "manual_start",
            "manual_end",
            "manual_duration_min",
            "matched_pair",
            "matched_side",
            "telemetry_start",
            "telemetry_end",
            "telemetry_duration_min",
            "start_lag_min",
            "classification",
            "detection_status",
            "confidence",
            "charging_proximity",
            "notes",
        ])

    # ---------------------------------------------------------------
    # 6. Lag summary
    # ---------------------------------------------------------------

    summary_rows = []

    for event_type in [
        "CHARGING",
        "REFUELING",
        "MANUAL_STOP",
    ]:
        sub = event_df[
            event_df["event_type"] == event_type
        ].copy()

        matched = sub["telemetry_start"].notna()

        lag = pd.to_numeric(
            sub.loc[matched, "start_lag_min"],
            errors="coerce",
        ).dropna()

        duration = pd.to_numeric(
            sub.loc[matched, "telemetry_duration_min"],
            errors="coerce",
        ).dropna()

        summary_rows.append({
            "event_type": event_type,
            "manual_events": len(sub),
            "matched_events": int(matched.sum()),
            "unmatched_events": int((~matched).sum()),
            "mean_lag_min": lag.mean() if not lag.empty else np.nan,
            "median_lag_min": lag.median() if not lag.empty else np.nan,
            "min_lag_min": lag.min() if not lag.empty else np.nan,
            "max_lag_min": lag.max() if not lag.empty else np.nan,
            "std_lag_min": (
                lag.std() if len(lag) > 1 else np.nan
            ),
            "mean_telemetry_duration_min": (
                duration.mean() if not duration.empty else np.nan
            ),
            "median_telemetry_duration_min": (
                duration.median() if not duration.empty else np.nan
            ),
        })

    lag_summary_df = pd.DataFrame(summary_rows)

    # ---------------------------------------------------------------
    # 7. Classification summary
    # ---------------------------------------------------------------

    if not event_df.empty:
        classification_df = (
            event_df["classification"]
            .value_counts(dropna=False)
            .rename_axis("classification")
            .reset_index(name="count")
        )
    else:
        classification_df = pd.DataFrame(
            columns=["classification", "count"]
        )

    # ---------------------------------------------------------------
    # 8. Method notes
    # ---------------------------------------------------------------

    method_notes = pd.DataFrame([
        ["Purpose", "Diagnostic only; does not modify Stage 4."],
        ["Manual source", str(MANUAL_EXCEL)],
        ["Telemetry source", str(RESULTS_PAIRED)],
        ["Manual stop meaning", "Approximate START timestamp only."],
        ["Stop end", "Derived from the corresponding telemetry stationary episode."],
        ["IDLE criterion", "< 3 minutes (< 180 seconds)."],
        ["STOP criterion", ">= 3 minutes (>= 180 seconds)."],
        ["Charging priority", "Charging/refuelling takes priority over STOP/IDLE."],
        ["Charging duration", "Read from 'Charging time (min)' in the manual Excel."],
        ["Charging timestamp", "Read from 'charging timestamp' in the manual Excel."],
        ["Lag", "Measured between manual event start and detected telemetry episode start."],
        ["Search window", f"-{SEARCH_BEFORE_MIN} to +{SEARCH_AFTER_MIN} minutes."],
        ["Stationary speed", f"Speed <= {STATIONARY_SPEED_KMH} km/h, with GPS implied-speed fallback when Speed is missing."],
        ["Vehicle matching", "Does not require vehicle number in the PAIR filename; uses date, vehicle type, route and trip timing."],
        ["Midnight stops", "Stop timestamps are parsed sequentially; a backwards clock time advances to the next calendar day."],
        ["Final use", "Use measured lag results to design the actual Stage-4 matching tolerance."],
    ], columns=["Item", "Rule / Description"])

    # ---------------------------------------------------------------
    # 9. Save
    # ---------------------------------------------------------------

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:

        event_df.to_excel(
            writer,
            sheet_name="Event_Matching",
            index=False,
        )

        trip_match_df.to_excel(
            writer,
            sheet_name="Trip_Matching",
            index=False,
        )

        lag_summary_df.to_excel(
            writer,
            sheet_name="Lag_Summary",
            index=False,
        )

        classification_df.to_excel(
            writer,
            sheet_name="Classification_Summary",
            index=False,
        )

        pd.DataFrame(unmatched_manual_rows).to_excel(
            writer,
            sheet_name="Unmatched_Manual_Records",
            index=False,
        )

        method_notes.to_excel(
            writer,
            sheet_name="Method_Notes",
            index=False,
        )

    # ---------------------------------------------------------------
    # 10. Terminal summary
    # ---------------------------------------------------------------

    print()
    print("=" * 80)
    print("DIAGNOSTIC COMPLETED")
    print("=" * 80)
    print(f"Output file:\n{OUTPUT_FILE}")

    if not event_df.empty:
        print()
        print("Classification counts:")
        print(
            classification_df.to_string(index=False)
        )

        print()
        print("Lag summary:")
        print(
            lag_summary_df.to_string(index=False)
        )

    print()
    print("IMPORTANT:")
    print(
        "This report is for measuring the relationship between the "
        "intern-entered times and telemetry."
    )
    print(
        "Do not use the measured lag as a final Stage-4 rule until "
        "the diagnostic results have been reviewed."
    )


if __name__ == "__main__":
    run_test()
