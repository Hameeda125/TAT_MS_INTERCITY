"""
Plot each unique vehicle route once and create an automatically updated route-wise trip summary.

Outputs
-------
1. all_routes_interactive.html  : interactive map; click routes for trip details
2. all_routes_static.png        : one representative travelled path per route
3. route_collection_summary.xlsx: route summary + trip-level details
4. route_collection_summary.csv : simple route summary

Expected input
--------------
A root folder containing .xlsx, .xls, or .csv trip files, including subfolders.
Each file should contain latitude, longitude, and preferably timestamp/speed columns.
A combined location column such as "Lat-Lng" is also supported.

Install once
------------
pip install pandas numpy matplotlib folium openpyxl
"""

from __future__ import annotations

import math
import re
import warnings
from pathlib import Path
from typing import Iterable, Optional

try:
    import folium
    from folium.plugins import Fullscreen, MeasureControl, MiniMap
except ImportError as exc:
    raise SystemExit(
        "Folium is not installed in the active Python environment.\n"
        "Activate tat_pipeline and run:  python -m pip install folium\n"
        "Then run this script again with: python src\\plot_unique_routes_and_update_summary.py"
    ) from exc

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

# ============================================================================
# USER SETTINGS: EDIT ONLY THIS SECTION
# ============================================================================
# Your current project structure:
#   D:\TAT_MS\TAT_Pipeline\src\this_file.py
#   D:\TAT_MS\TAT_Data\...trip files...
#
# Change only these two lines if your folders are different.
DATA_ROOT = Path(r"D:\TAT_MS\TAT_Data")
OUTPUT_DIR = Path(r"D:\TAT_MS\TAT_Data\Results\Route_Map")

# Set True when the route is explicitly available as a column in each file.
# The script still works when it is False by extracting the route from filename.
PREFER_ROUTE_COLUMN = True

# Keep every Nth GPS point when drawing, to reduce HTML size. 1 = keep all points.
MAP_POINT_STEP = 2

# Plot only one representative GPS path for each route.
# The representative trip is selected as the trip with the largest number of valid GPS points.
PLOT_ONE_PATH_PER_ROUTE = True

# Ignore files with fewer than this many valid coordinates.
MIN_VALID_POINTS = 2

# Optional: change to True to join reverse directions as one corridor.
# Example: "A to B" and "B to A" become one route group.
GROUP_REVERSE_DIRECTIONS = False
# ============================================================================


LAT_CANDIDATES = [
    "latitude", "lat", "gps latitude", "gps_latitude", "y"
]
LON_CANDIDATES = [
    "longitude", "lon", "lng", "long", "gps longitude", "gps_longitude", "x"
]
COMBINED_COORD_CANDIDATES = [
    "lat-lng", "lat lng", "lat_lon", "latlon", "location", "coordinates",
    "gps coordinates", "gps location"
]
TIME_CANDIDATES = [
    "timestamp", "time stamp", "datetime", "date time", "packet time",
    "actual received time", "time", "date"
]
ROUTE_CANDIDATES = [
    "route", "route name", "route_name", "trip route", "corridor", "sector"
]
VEHICLE_CANDIDATES = [
    "vehicle type", "vehicle_type", "fuel type", "fuel_type", "powertrain"
]
TRIP_ID_CANDIDATES = [
    "trip id", "trip_id", "trip no", "trip number", "trip_number"
]


def clean_col_name(value: object) -> str:
    """Normalize a column name for robust matching."""
    text = str(value).strip().lower()
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def find_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    """Return the original dataframe column matching one of the candidates."""
    normalized = {clean_col_name(col): col for col in df.columns}
    for candidate in candidates:
        key = clean_col_name(candidate)
        if key in normalized:
            return normalized[key]

    # Partial matching is useful for names such as "Vehicle Latitude (deg)".
    for candidate in candidates:
        key = clean_col_name(candidate)
        for normalized_name, original_name in normalized.items():
            if key in normalized_name:
                return original_name
    return None


def read_trip_file(path: Path) -> pd.DataFrame:
    """Read CSV/Excel, selecting the first non-empty sheet when needed."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        try:
            return pd.read_csv(path)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin-1")

    if suffix in {".xlsx", ".xls"}:
        sheets = pd.read_excel(path, sheet_name=None)
        for _, sheet in sheets.items():
            if not sheet.empty:
                return sheet
        return pd.DataFrame()

    raise ValueError(f"Unsupported file type: {suffix}")


def extract_coordinates(df: pd.DataFrame) -> pd.DataFrame:
    """Return a dataframe with numeric latitude/longitude and optional timestamp."""
    lat_col = find_column(df, LAT_CANDIDATES)
    lon_col = find_column(df, LON_CANDIDATES)

    out = pd.DataFrame(index=df.index)

    if lat_col and lon_col:
        out["latitude"] = pd.to_numeric(df[lat_col], errors="coerce")
        out["longitude"] = pd.to_numeric(df[lon_col], errors="coerce")
    else:
        combined_col = find_column(df, COMBINED_COORD_CANDIDATES)
        if not combined_col:
            raise ValueError("Latitude/longitude columns or a combined coordinate column were not found")

        # Supports values like "13.0827, 80.2707", "13.0827 80.2707",
        # and strings containing brackets.
        extracted = (
            df[combined_col]
            .astype(str)
            .str.extract(r"([-+]?\d{1,3}(?:\.\d+)?)\s*[,;| ]\s*([-+]?\d{1,3}(?:\.\d+)?)")
        )
        out["latitude"] = pd.to_numeric(extracted[0], errors="coerce")
        out["longitude"] = pd.to_numeric(extracted[1], errors="coerce")

    time_col = find_column(df, TIME_CANDIDATES)
    if time_col:
        out["timestamp"] = pd.to_datetime(df[time_col], errors="coerce", dayfirst=True)
    else:
        out["timestamp"] = pd.NaT

    # Valid geographic ranges; removes headers repeated inside data and bad values.
    out = out[
        out["latitude"].between(-90, 90)
        & out["longitude"].between(-180, 180)
    ].copy()

    # Preserve travel order using timestamp whenever timestamps are available.
    if out["timestamp"].notna().sum() >= 2:
        out = out.sort_values("timestamp")

    # Remove exact consecutive duplicate points.
    duplicate_previous = (
        out["latitude"].eq(out["latitude"].shift())
        & out["longitude"].eq(out["longitude"].shift())
    )
    return out.loc[~duplicate_previous].reset_index(drop=True)


def filename_route_name(path: Path) -> str:
    """
    Infer route from filename.

    This handles names such as:
      T_Nagar_to_Thiruporur_trip_01_2026-08-01_EV.xlsx
      chennai-thiruvallur-ice-trip2.xlsx

    Adjust this function if your file naming convention is different.
    """
    text = path.stem

    # Remove common non-route parts.
    text = re.sub(r"(?i)\b(ev|ice|electric|diesel|cng|bus|truck|vehicle)\b", " ", text)
    text = re.sub(r"(?i)\btrip\s*[-_ ]*\d+\b", " ", text)
    text = re.sub(r"\b\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}\b", " ", text)
    text = re.sub(r"\b\d{4}[-_.]\d{1,2}[-_.]\d{1,2}\b", " ", text)
    text = re.sub(r"[_]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" -_")

    # Make route direction readable.
    text = re.sub(r"(?i)\s+(to|2)\s+", " → ", text)
    return text.title() if text else path.stem


def get_route_name(df: pd.DataFrame, path: Path) -> str:
    route_col = find_column(df, ROUTE_CANDIDATES) if PREFER_ROUTE_COLUMN else None
    if route_col:
        values = df[route_col].dropna().astype(str).str.strip()
        values = values[values.ne("")]
        if not values.empty:
            return values.mode().iloc[0]
    return filename_route_name(path)


def get_vehicle_type(df: pd.DataFrame, path: Path) -> str:
    col = find_column(df, VEHICLE_CANDIDATES)
    if col:
        values = df[col].dropna().astype(str).str.upper()
        if not values.empty:
            value = values.mode().iloc[0]
            if "EV" in value or "ELECTRIC" in value:
                return "EV"
            if "ICE" in value or any(x in value for x in ["DIESEL", "PETROL", "CNG"]):
                return "ICE"

    name = path.stem.upper()
    if re.search(r"(^|[^A-Z])EV([^A-Z]|$)", name) or "ELECTRIC" in name:
        return "EV"
    if re.search(r"(^|[^A-Z])ICE([^A-Z]|$)", name) or any(x in name for x in ["DIESEL", "PETROL", "CNG"]):
        return "ICE"
    return "Unknown"


def get_trip_id(df: pd.DataFrame, path: Path) -> str:
    col = find_column(df, TRIP_ID_CANDIDATES)
    if col:
        values = df[col].dropna().astype(str).str.strip()
        if not values.empty:
            return values.iloc[0]
    return path.stem


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlambda / 2) ** 2
    return 2 * radius * math.asin(math.sqrt(a))


def route_distance_km(points: pd.DataFrame) -> float:
    if len(points) < 2:
        return 0.0
    lat1 = points["latitude"].to_numpy()[:-1]
    lon1 = points["longitude"].to_numpy()[:-1]
    lat2 = points["latitude"].to_numpy()[1:]
    lon2 = points["longitude"].to_numpy()[1:]

    # Vectorized haversine calculation.
    r = 6371.0088
    phi1, phi2 = np.radians(lat1), np.radians(lat2)
    dphi = np.radians(lat2 - lat1)
    dlambda = np.radians(lon2 - lon1)
    a = np.sin(dphi / 2) ** 2 + np.cos(phi1) * np.cos(phi2) * np.sin(dlambda / 2) ** 2
    return float(np.sum(2 * r * np.arcsin(np.sqrt(a))))


def canonical_route(route: str) -> str:
    """Optionally group A→B and B→A into one route corridor."""
    route = re.sub(r"\s+", " ", str(route)).strip()
    if not GROUP_REVERSE_DIRECTIONS:
        return route

    parts = re.split(r"\s*(?:→|->|\bto\b)\s*", route, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) == 2:
        return " ↔ ".join(sorted([parts[0].strip(), parts[1].strip()], key=str.casefold))
    return route


def discover_files(root: Path) -> list[Path]:
    """Find input trip files while ignoring generated pipeline outputs."""
    files: list[Path] = []
    excluded_folders = {"results", "result", "outputs", "output", "system", "route_map"}

    for pattern in ("*.xlsx", "*.xls", "*.csv"):
        for path in root.rglob(pattern):
            relative_parts = [part.lower() for part in path.relative_to(root).parts[:-1]]
            if path.name.startswith("~$"):
                continue
            if any(part in excluded_folders for part in relative_parts):
                continue
            files.append(path)

    return sorted(set(files))


def create_excel_output(summary: pd.DataFrame, trips: pd.DataFrame, errors: pd.DataFrame, output: Path) -> None:
    """Save a formatted multi-sheet Excel report using pandas/XlsxWriter."""
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Route Summary", index=False)
        trips.to_excel(writer, sheet_name="Trip Details", index=False)
        errors.to_excel(writer, sheet_name="Skipped Files", index=False)

        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": "#1F4E78",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        number_fmt = workbook.add_format({"num_format": "0.00"})

        for sheet_name, frame in {
            "Route Summary": summary,
            "Trip Details": trips,
            "Skipped Files": errors,
        }.items():
            sheet = writer.sheets[sheet_name]
            sheet.freeze_panes(1, 0)
            sheet.autofilter(0, 0, max(len(frame), 1), max(len(frame.columns) - 1, 0))
            for col_idx, col_name in enumerate(frame.columns):
                sheet.write(0, col_idx, col_name, header_fmt)
                sample_width = frame[col_name].astype(str).map(len).max() if not frame.empty else len(col_name)
                width = min(max(len(col_name) + 2, int(sample_width) + 2), 45)
                fmt = number_fmt if "km" in col_name.lower() else None
                sheet.set_column(col_idx, col_idx, width, fmt)


def main() -> None:
    if not DATA_ROOT.exists():
        raise FileNotFoundError(
            f"Input folder does not exist: {DATA_ROOT}\n"
            "Open this script and correct DATA_ROOT near the top of the file."
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    files = discover_files(DATA_ROOT)
    if not files:
        raise FileNotFoundError(f"No CSV/Excel trip files found under {DATA_ROOT}")

    print(f"Found {len(files)} candidate files")

    trip_records: list[dict] = []
    route_geometries: list[dict] = []
    skipped: list[dict] = []

    for index, path in enumerate(files, start=1):
        print(f"[{index}/{len(files)}] Reading {path.name}")
        try:
            raw = read_trip_file(path)
            if raw.empty:
                raise ValueError("File/sheet is empty")

            points = extract_coordinates(raw)
            if len(points) < MIN_VALID_POINTS:
                raise ValueError(f"Only {len(points)} valid coordinate points")

            raw_route = get_route_name(raw, path)
            route = canonical_route(raw_route)
            vehicle_type = get_vehicle_type(raw, path)
            trip_id = get_trip_id(raw, path)
            distance = route_distance_km(points)

            start_time = points["timestamp"].min()
            end_time = points["timestamp"].max()
            duration_minutes = (
                (end_time - start_time).total_seconds() / 60
                if pd.notna(start_time) and pd.notna(end_time)
                else np.nan
            )

            trip_records.append({
                "Route": route,
                "Original Route Label": raw_route,
                "Trip ID": trip_id,
                "Vehicle Type": vehicle_type,
                "File Name": path.name,
                "Relative Folder": str(path.parent.relative_to(DATA_ROOT)),
                "Start Time": start_time,
                "End Time": end_time,
                "Duration (min)": duration_minutes,
                "GPS Points": len(points),
                "Travelled Distance (km)": distance,
                "Start Latitude": points.iloc[0]["latitude"],
                "Start Longitude": points.iloc[0]["longitude"],
                "End Latitude": points.iloc[-1]["latitude"],
                "End Longitude": points.iloc[-1]["longitude"],
            })

            route_geometries.append({
                "route": route,
                "trip_id": trip_id,
                "vehicle_type": vehicle_type,
                "file_name": path.name,
                "distance_km": distance,
                "points": points.iloc[::max(MAP_POINT_STEP, 1)].copy(),
            })

        except Exception as exc:
            skipped.append({"File Name": path.name, "Path": str(path), "Reason": str(exc)})

    if not trip_records:
        raise RuntimeError("No valid trip files could be processed. Check column names and skipped-file reasons.")

    trips = pd.DataFrame(trip_records).sort_values(["Route", "Start Time", "File Name"])

    # Route-level summary requested by the user.
    summary = (
        trips.groupby("Route", dropna=False)
        .agg(
            Total_Trips=("Trip ID", "count"),
            EV_Trips=("Vehicle Type", lambda s: int((s == "EV").sum())),
            ICE_Trips=("Vehicle Type", lambda s: int((s == "ICE").sum())),
            Unknown_Trips=("Vehicle Type", lambda s: int((s == "Unknown").sum())),
            Total_Travelled_Distance_km=("Travelled Distance (km)", "sum"),
            Average_Trip_Distance_km=("Travelled Distance (km)", "mean"),
            Minimum_Trip_Distance_km=("Travelled Distance (km)", "min"),
            Maximum_Trip_Distance_km=("Travelled Distance (km)", "max"),
            First_Data_Date=("Start Time", "min"),
            Last_Data_Date=("Start Time", "max"),
        )
        .reset_index()
        .rename(columns={
            "Total_Trips": "Total Trips",
            "EV_Trips": "EV Trips",
            "ICE_Trips": "ICE Trips",
            "Unknown_Trips": "Unknown Trips",
            "Total_Travelled_Distance_km": "Total Travelled Distance (km)",
            "Average_Trip_Distance_km": "Average Trip Distance (km)",
            "Minimum_Trip_Distance_km": "Minimum Trip Distance (km)",
            "Maximum_Trip_Distance_km": "Maximum Trip Distance (km)",
            "First_Data_Date": "First Data Date",
            "Last_Data_Date": "Last Data Date",
        })
        .sort_values(["Total Trips", "Route"], ascending=[False, True])
    )

    for col in [
        "Total Travelled Distance (km)", "Average Trip Distance (km)",
        "Minimum Trip Distance (km)", "Maximum Trip Distance (km)",
    ]:
        summary[col] = summary[col].round(2)
    trips["Travelled Distance (km)"] = trips["Travelled Distance (km)"].round(2)
    trips["Duration (min)"] = trips["Duration (min)"].round(2)

    skipped_df = pd.DataFrame(skipped, columns=["File Name", "Path", "Reason"])

    # Select only one representative travelled path for each route.
    # This prevents five Chennai-to-Bangalore trips from being drawn five times.
    # Trip counting still uses every input file, so the summary table remains correct.
    if PLOT_ONE_PATH_PER_ROUTE:
        representative_geometries: list[dict] = []
        for route in summary["Route"]:
            candidates = [g for g in route_geometries if g["route"] == route]
            representative = max(
                candidates,
                key=lambda g: (len(g["points"]), g["distance_km"]),
            )
            representative_geometries.append(representative)
    else:
        representative_geometries = route_geometries

    # ---------------------------- Interactive map ----------------------------
    all_lats = np.concatenate([g["points"]["latitude"].to_numpy() for g in representative_geometries])
    all_lons = np.concatenate([g["points"]["longitude"].to_numpy() for g in representative_geometries])
    centre = [float(np.nanmean(all_lats)), float(np.nanmean(all_lons))]

    fmap = folium.Map(location=centre, zoom_start=10, tiles="OpenStreetMap", control_scale=True)
    Fullscreen(position="topright").add_to(fmap)
    MeasureControl(position="topleft", primary_length_unit="kilometers").add_to(fmap)
    MiniMap(toggle_display=True).add_to(fmap)

    palette = [
        "red", "blue", "green", "purple", "orange", "darkred", "cadetblue",
        "darkgreen", "darkblue", "pink", "gray", "black", "lightblue"
    ]
    routes = list(summary["Route"])
    route_colours = {route: palette[i % len(palette)] for i, route in enumerate(routes)}

    feature_groups: dict[str, folium.FeatureGroup] = {}
    for route in routes:
        count = int(summary.loc[summary["Route"] == route, "Total Trips"].iloc[0])
        group = folium.FeatureGroup(name=f"{route} ({count} trips)", show=True)
        group.add_to(fmap)
        feature_groups[route] = group

    for geo in representative_geometries:
        points = geo["points"]
        coordinates = points[["latitude", "longitude"]].values.tolist()
        colour = route_colours[geo["route"]]
        route_row = summary.loc[summary["Route"] == geo["route"]].iloc[0]
        total_trips = int(route_row["Total Trips"])
        ev_trips = int(route_row["EV Trips"])
        ice_trips = int(route_row["ICE Trips"])
        tooltip = f"{geo['route']} | {total_trips} trips"
        popup = folium.Popup(
            f"<b>Route:</b> {geo['route']}<br>"
            f"<b>Total trips:</b> {total_trips}<br>"
            f"<b>EV trips:</b> {ev_trips}<br>"
            f"<b>ICE trips:</b> {ice_trips}<br>"
            f"<b>Representative file:</b> {geo['file_name']}<br>"
            f"<b>Representative distance:</b> {geo['distance_km']:.2f} km",
            max_width=420,
        )
        folium.PolyLine(
            coordinates,
            color=colour,
            weight=4,
            opacity=0.70,
            tooltip=tooltip,
            popup=popup,
        ).add_to(feature_groups[geo["route"]])

    # One start/end marker per plotted route.
    for geo in representative_geometries:
        points = geo["points"]
        colour = route_colours[geo["route"]]
        folium.CircleMarker(
            [points.iloc[0]["latitude"], points.iloc[0]["longitude"]],
            radius=3, color=colour, fill=True, fill_opacity=1,
            tooltip=f"Start: {geo['route']} | {geo['trip_id']}",
        ).add_to(feature_groups[geo["route"]])
        folium.CircleMarker(
            [points.iloc[-1]["latitude"], points.iloc[-1]["longitude"]],
            radius=3, color=colour, fill=True, fill_opacity=1,
            tooltip=f"End: {geo['route']} | {geo['trip_id']}",
        ).add_to(feature_groups[geo["route"]])

    fmap.fit_bounds([[float(all_lats.min()), float(all_lons.min())],
                     [float(all_lats.max()), float(all_lons.max())]])
    folium.LayerControl(collapsed=False).add_to(fmap)
    fmap.save(OUTPUT_DIR / "all_routes_interactive.html")

    # ------------------------------ Static map -------------------------------
    fig, ax = plt.subplots(figsize=(14, 11))
    for route in routes:
        geo = next(g for g in representative_geometries if g["route"] == route)
        colour = route_colours[route]
        points = geo["points"]
        trip_count = int(summary.loc[summary["Route"] == route, "Total Trips"].iloc[0])
        ax.plot(
            points["longitude"], points["latitude"],
            linewidth=2.4, alpha=0.80, color=colour,
            label=f"{route} ({trip_count} trips)",
        )
        # Label route near the representative travelled path midpoint.
        mid = points.iloc[len(points) // 2]
        ax.annotate(
            route,
            (mid["longitude"], mid["latitude"]),
            xytext=(4, 4), textcoords="offset points", fontsize=8,
            bbox={"boxstyle": "round,pad=0.2", "facecolor": "white", "alpha": 0.65, "edgecolor": "none"},
        )

    margin_lon = max((all_lons.max() - all_lons.min()) * 0.06, 0.01)
    margin_lat = max((all_lats.max() - all_lats.min()) * 0.06, 0.01)
    ax.set_xlim(all_lons.min() - margin_lon, all_lons.max() + margin_lon)
    ax.set_ylim(all_lats.min() - margin_lat, all_lats.max() + margin_lat)
    ax.set_title(
        f"Routes Covered by Collected Vehicle Data\n"
        f"{len(trips)} trips across {len(summary)} routes",
        fontsize=18, fontweight="bold", pad=16,
    )
    ax.set_xlabel("Longitude")
    ax.set_ylabel("Latitude")
    ax.grid(True, linestyle="--", alpha=0.30)
    ax.set_aspect("equal", adjustable="datalim")
    ax.legend(loc="upper left", bbox_to_anchor=(1.01, 1), fontsize=8, frameon=True)
    fig.tight_layout()
    fig.savefig(OUTPUT_DIR / "all_routes_static.png", dpi=300, bbox_inches="tight")
    plt.close(fig)

    # ------------------------------- Tables ----------------------------------
    summary.to_csv(OUTPUT_DIR / "route_collection_summary.csv", index=False)
    create_excel_output(
        summary,
        trips,
        skipped_df,
        OUTPUT_DIR / "route_collection_summary.xlsx",
    )

    print("\nCompleted successfully")
    print(f"Valid trips : {len(trips)}")
    print(f"Routes      : {len(summary)} (one path plotted per route)")
    print(f"Skipped     : {len(skipped_df)}")
    print(f"Outputs     : {OUTPUT_DIR}")


if __name__ == "__main__":
    main()