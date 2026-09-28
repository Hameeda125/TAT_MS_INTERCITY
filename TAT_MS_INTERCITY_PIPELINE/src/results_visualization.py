"""
Stage 6 - INTERCITY EV-ICE Results Visualization
================================================

Run after:
    1. phase_segmentation.py
    2. tat_computation.py

Purpose
-------
Create analytical visualizations for the INTERCITY-only pipeline without
duplicating the basic plots already produced by Stage 9 Master Summary.

Stage 9 Summary owns these basic comparison plots:
    - EV vs ICE TAT summary
    - Average moving speed
    - Average phase time
    - Average phase distance

Therefore this Stage 6 script deliberately DOES NOT generate those plots.

Stage 6 focuses on analytical visualizations that answer different questions:
    - pair-level TAT differences
    - TAT difference distribution
    - EV/ICE distribution agreement
    - Bland-Altman agreement
    - phase contribution differences
    - phase difference heatmap
    - correlation of pair-level differences
    - speed-range behaviour
    - speed-distance relationship
    - route-level uncertainty
    - STRICT vs NONSTRICT behaviour
    - PCA / clustering of EV-ICE operational differences

The script reads the Stage-5 outputs and does not recalculate TAT,
phase segmentation, pairing, or cropping methodology.

Output
------
<results_folder>/Visualizations/INTERCITY/
    Per_Pair/<PAIR_ID>/
    Average/
        Master_Results.xlsx
        Master_Visualizations.pdf
        PNG/
        Master_Plot_Audit.xlsx
    Route_Wise/
        <route>/
            Average/
                Master_Results.xlsx
                Master_Visualizations.pdf
                PNG/
        Route_Wise_Index.xlsx

At pair level, the script may add:
    10_Visual_Dashboard
    11_Visual_Charts

Only INTERCITY pairs are processed.
"""

from __future__ import annotations

import math
import os
import re
import shutil
import time
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import pandas as pd
import seaborn as sns

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config_loader import load_config


# ============================================================================
# CONFIGURATION
# ============================================================================

SCRIPT_VERSION = (
    f"{datetime.now():%Y-%m-%d}_INTERCITY_ONLY_VISUALIZATION_V1"
)

PAIR_INDEX_NAME = "pair_index.csv"

EV_SOURCE_SHEET = "01_EV_Cropped_Data"
ICE_SOURCE_SHEET = "02_ICE_Cropped_Data"
PAIR_INFO_SHEET = "03_Pair_Crop_Info"
TAT_RESULT_SHEET = "08_TAT_Phase_Comparison"
SPEED_RESULT_SHEET = "09_SpeedRange_Comparison"

DASHBOARD_SHEET = "10_Visual_Dashboard"
CHARTS_SHEET = "11_Visual_Charts"

TIMESTAMP = "Timestamp"
SPEED_KMH = "Speed (km/h)"
TIME_DIFFERENCE = "Time Difference (s)"
DISTANCE_M = "Haversine Distance (m)"
PHASE = "phase"

PHASES = [
    "CHARGING / REFUELING",
    "STOP",
    "IDLE",
    "INCH",
    "ACC",
    "DEC",
    "CRUISE",
]

EV_COLOR = "#2E86DE"
ICE_COLOR = "#E67E22"
POSITIVE_COLOR = "#C0392B"
NEGATIVE_COLOR = "#1F618D"

# Lower DPI than the previous visualization script to reduce runtime/file size.
DPI = 150

# Set to "PAIR_002" for a single-pair test.
# Keep None for all 18 INTERCITY pairs.
PAIR_NAME_FILTER: str | None = None

# Preserve the existing behaviour of embedding visualizations in each
# paired workbook. Set False only when a faster, file-only run is desired.
EMBED_VISUALS_IN_PAIR_WORKBOOKS = True

# The summary workbook already owns the following basic comparisons.
SUMMARY_OWNED_PLOTS = {
    "EV vs ICE TAT summary",
    "Average moving speed",
    "Average phase time",
    "Average phase distance",
}

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)

sns.set_theme(
    style="whitegrid",
    context="notebook",
    font_scale=0.9,
)


# ============================================================================
# DATA CLASS
# ============================================================================

@dataclass
class PairData:
    pair_id: str
    workbook_path: Path
    metadata: dict[str, Any]
    overall: pd.DataFrame
    phase: pd.DataFrame
    speed: pd.DataFrame
    ev_source: pd.DataFrame
    ice_source: pd.DataFrame


# ============================================================================
# BASIC HELPERS
# ============================================================================

def clean(value: Any) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    return str(value).strip()


def slug(text: str) -> str:
    return (
        re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            clean(text),
        )
        .strip("_")
        or "unknown"
    )


def parse_num(
    value: Any,
    default: float = np.nan,
) -> float:
    if value is None:
        return default

    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    try:
        return float(value)
    except (TypeError, ValueError):
        pass

    match = re.search(
        r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?",
        clean(value).replace(",", "").replace("−", "-"),
    )

    if match is None:
        return default

    try:
        return float(match.group(0))
    except ValueError:
        return default


def pair_number(pair_id: str) -> int:
    match = re.search(r"(\d+)", pair_id)
    return int(match.group(1)) if match else 0


def read_csv_flexible(path: Path) -> pd.DataFrame:
    last_error = None

    for encoding in (
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ):
        try:
            return pd.read_csv(path, encoding=encoding)
        except UnicodeDecodeError as error:
            last_error = error

    raise ValueError(
        f"Could not decode {path}: {last_error}"
    )


def savefig(
    fig,
    path: Path,
    pdf: PdfPages | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    fig.savefig(
        path,
        dpi=DPI,
        bbox_inches="tight",
        facecolor="white",
    )

    if pdf is not None:
        pdf.savefig(
            fig,
            bbox_inches="tight",
            facecolor="white",
        )

    plt.close(fig)


def finite_numeric(
    frame: pd.DataFrame,
    columns: list[str],
) -> pd.DataFrame:
    available = [
        column
        for column in columns
        if column in frame.columns
    ]

    if not available:
        return pd.DataFrame()

    numeric = frame[available].apply(
        pd.to_numeric,
        errors="coerce",
    )

    return numeric.replace(
        [np.inf, -np.inf],
        np.nan,
    )


def label_bars(
    ax,
    bars,
    decimals: int = 2,
) -> None:
    labels = []

    for bar in bars:
        height = bar.get_height()

        if not math.isfinite(height):
            labels.append("")
        else:
            labels.append(
                f"{height:.{decimals}f}"
            )

    ax.bar_label(
        bars,
        labels=labels,
        padding=3,
        fontsize=8,
    )


# ============================================================================
# PAIR WORKBOOK DISCOVERY
# ============================================================================

def resolve_workbook(
    row: pd.Series,
    results_root: Path,
) -> Path:
    configured = clean(
        row.get("pair_workbook_path")
    )

    if configured:
        candidate = Path(configured)
        if candidate.exists() and candidate.is_file():
            return candidate

    workbook_name = clean(
        row.get("pair_workbook_name")
    )

    output_folder = clean(
        row.get("output_date_folder")
    )

    if workbook_name and output_folder:
        candidate = (
            results_root
            / "Paired"
            / output_folder
            / workbook_name
        )

        if candidate.exists():
            return candidate

    pair_id = clean(
        row.get("pair_id")
    )

    if pair_id:
        matches = [
            path
            for path in (
                results_root / "Paired"
            ).rglob(
                f"{pair_id}_*.xlsx"
            )
            if path.is_file()
            and not path.name.startswith(
                ("~$", ".")
            )
        ]

        if matches:
            return max(
                matches,
                key=lambda path: path.stat().st_mtime,
            )

    raise FileNotFoundError(
        f"Workbook not found for {pair_id or workbook_name}"
    )


def discover_intercity_pairs(
    results_root: Path,
) -> list[tuple[str, dict[str, Any], Path]]:
    pair_index_path = (
        results_root
        / "System"
        / PAIR_INDEX_NAME
    )

    if not pair_index_path.exists():
        raise FileNotFoundError(
            f"Pair index not found: {pair_index_path}"
        )

    index = read_csv_flexible(
        pair_index_path
    )

    pairs = []

    for _, row in index.iterrows():
        pair_id = clean(
            row.get("pair_id")
        )

        if not pair_id:
            continue

        if (
            PAIR_NAME_FILTER
            and PAIR_NAME_FILTER.casefold()
            not in pair_id.casefold()
        ):
            continue

        # This is an INTERCITY-only project.  Prefer explicit metadata/name
        # evidence and do not run CITY/INTERCITY route detection.
        evidence = " ".join(
            [
                clean(row.get(column))
                for column in (
                    "pair_class",
                    "route",
                    "route_name",
                    "route_direction",
                    "source_file_name",
                    "ev_file_name",
                    "ice_file_name",
                    "pair_workbook_name",
                )
            ]
        ).casefold()

        workbook_path = resolve_workbook(
            row,
            results_root,
        )

        evidence = (
            evidence
            + " "
            + workbook_path.name.casefold()
        )

        if "intercity" not in evidence:
            # Current project is intercity-only. If a future pair index
            # contains no token, it is still accepted because this stage is
            # explicitly the intercity pipeline.
            pass

        pairs.append(
            (
                pair_id,
                row.to_dict(),
                workbook_path,
            )
        )

    pairs.sort(
        key=lambda item: (
            pair_number(item[0]),
            item[0],
        )
    )

    return pairs


# ============================================================================
# STAGE-5 EXTRACTION
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
    except (ValueError, KeyError):
        return {}

    if frame.empty or len(frame.columns) < 2:
        return {}

    first = frame.columns[0]
    second = frame.columns[1]

    return {
        clean(key): value
        for key, value in zip(
            frame[first],
            frame[second],
        )
        if clean(key)
    }


def clean_extracted_table(
    frame: pd.DataFrame,
    key_column: str,
) -> pd.DataFrame:
    cleaned = frame.copy()

    valid_positions = []
    seen_names: set[str] = set()

    for position, raw_name in enumerate(
        cleaned.columns
    ):
        name = clean(raw_name)

        if not name:
            continue

        if name.lower().startswith("unnamed:"):
            continue

        normalized = name.casefold()

        if normalized in seen_names:
            continue

        seen_names.add(normalized)
        valid_positions.append(position)

    cleaned = cleaned.iloc[
        :,
        valid_positions,
    ].copy()

    cleaned.columns = [
        clean(column)
        for column in cleaned.columns
    ]

    if key_column not in cleaned.columns:
        raise ValueError(
            f"Expected column {key_column!r} not found. "
            f"Available: {list(cleaned.columns)}"
        )

    cleaned = cleaned.dropna(how="all")

    cleaned = cleaned.loc[
        cleaned[key_column]
        .astype("string")
        .fillna("")
        .str.strip()
        .ne("")
    ].reset_index(drop=True)

    return cleaned


def extract_tables(
    path: Path,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    with pd.ExcelFile(
        path,
        engine="openpyxl",
    ) as excel:
        required = {
            EV_SOURCE_SHEET,
            ICE_SOURCE_SHEET,
            TAT_RESULT_SHEET,
            SPEED_RESULT_SHEET,
        }

        missing = (
            required
            - set(excel.sheet_names)
        )

        if missing:
            raise ValueError(
                "Missing required Stage-5 sheet(s): "
                + ", ".join(sorted(missing))
            )

        tat = pd.read_excel(
            excel,
            sheet_name=TAT_RESULT_SHEET,
            header=None,
        )

        speed = pd.read_excel(
            excel,
            sheet_name=SPEED_RESULT_SHEET,
            header=None,
        )

        ev = pd.read_excel(
            excel,
            sheet_name=EV_SOURCE_SHEET,
        )

        ice = pd.read_excel(
            excel,
            sheet_name=ICE_SOURCE_SHEET,
        )

    overall_header = None
    phase_header = None

    for index in range(len(tat)):
        values = [
            clean(value)
            for value in tat.iloc[index]
        ]

        if (
            "Metric" in values
            and "EV" in values
            and "ICE" in values
        ):
            overall_header = index

        if (
            "Phase" in values
            and "EV Time (min)" in values
        ):
            phase_header = index

    if (
        overall_header is None
        or phase_header is None
    ):
        raise ValueError(
            "Could not identify Stage-5 TAT overall/phase tables."
        )

    overall = tat.iloc[
        overall_header + 1:
        phase_header - 2
    ].copy()

    overall.columns = tat.iloc[
        overall_header
    ]

    overall = clean_extracted_table(
        overall,
        "Metric",
    )

    phase = tat.iloc[
        phase_header + 1:
    ].copy()

    phase.columns = tat.iloc[
        phase_header
    ]

    phase = clean_extracted_table(
        phase,
        "Phase",
    )

    speed_header = None

    for index in range(len(speed)):
        values = [
            clean(value)
            for value in speed.iloc[index]
        ]

        if (
            "Speed Range" in values
            and "EV Time (min)" in values
        ):
            speed_header = index
            break

    if speed_header is None:
        raise ValueError(
            "Could not identify Stage-5 speed-range table."
        )

    speed_frame = speed.iloc[
        speed_header + 1:
    ].copy()

    speed_frame.columns = speed.iloc[
        speed_header
    ]

    speed_frame = clean_extracted_table(
        speed_frame,
        "Speed Range",
    )

    speed_frame = speed_frame.loc[
        speed_frame["Speed Range"]
        .astype("string")
        .str.strip()
        .str.upper()
        .ne("TOTAL")
    ].reset_index(drop=True)

    for frame in (
        overall,
        phase,
        speed_frame,
    ):
        identifier_columns = {
            "Metric",
            "Phase",
            "Speed Range",
        }

        for position, column in enumerate(
            frame.columns
        ):
            if column in identifier_columns:
                continue

            frame.iloc[
                :,
                position,
            ] = frame.iloc[
                :,
                position,
            ].map(parse_num)

    return (
        overall,
        phase,
        speed_frame,
        ev,
        ice,
    )


def metric(
    overall: pd.DataFrame,
    name: str,
    vehicle: str,
) -> float:
    if vehicle not in overall.columns:
        return np.nan

    matches = overall.loc[
        overall["Metric"]
        .astype(str)
        .str.strip()
        .eq(name),
        vehicle,
    ]

    if matches.empty:
        return np.nan

    return parse_num(
        matches.iloc[0]
    )


def phase_value(
    phase: pd.DataFrame,
    phase_name: str,
    column: str,
) -> float:
    aliases = {
        "CHARGING / REFUELING": {
            "charging / refueling",
            "charging",
            "refueling",
            "refuelling",
        },
        "INCH": {
            "inch",
            "inching",
        },
        "ACC": {
            "acc",
            "acceleration",
            "accelerating",
        },
        "DEC": {
            "dec",
            "deceleration",
            "decelerating",
        },
        "CRUISE": {
            "cruise",
            "cruising",
        },
        "STOP": {"stop"},
        "IDLE": {"idle"},
    }

    accepted = aliases.get(
        phase_name,
        {phase_name.casefold()},
    )

    names = (
        phase["Phase"]
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    mask = names.isin(
        accepted
    )

    if column not in phase.columns:
        return 0.0

    values = phase.loc[
        mask,
        column,
    ]

    if values.empty:
        return 0.0

    return float(
        np.nansum(
            [
                parse_num(value, 0.0)
                for value in values
            ]
        )
    )


# ============================================================================
# METADATA / MASTER ROW
# ============================================================================

def route_label(
    metadata: dict[str, Any],
    pair_info: dict[str, Any],
    workbook_path: Path,
) -> str:
    for source in (
        pair_info,
        metadata,
    ):
        for key in (
            "route_direction",
            "route",
            "route_name",
            "Route",
            "Route Name",
        ):
            value = clean(
                source.get(key)
            )

            if value:
                return value

    name = workbook_path.stem

    name = re.sub(
        r"^PAIR_\d+_",
        "",
        name,
        flags=re.I,
    )

    name = re.sub(
        r"^(STRICT|NONSTRICT)_",
        "",
        name,
        flags=re.I,
    )

    return name.replace("_", " ")


def pair_class(
    metadata: dict[str, Any],
    workbook_path: Path,
) -> str:
    value = clean(
        metadata.get("pair_class")
    )

    if value:
        return value.upper()

    name = workbook_path.name.upper()

    if "NONSTRICT" in name:
        return "NONSTRICT"

    if "STRICT" in name:
        return "STRICT"

    return "UNKNOWN"


def pair_to_master(
    pair: PairData,
) -> dict[str, Any]:
    info = read_pair_info(
        pair.workbook_path
    )

    row: dict[str, Any] = {
        "pair_id": pair.pair_id,
        "pair_number": pair_number(pair.pair_id),
        "pair_class": pair_class(
            pair.metadata,
            pair.workbook_path,
        ),
        "operation_type": "INTERCITY",
        "route": route_label(
            pair.metadata,
            info,
            pair.workbook_path,
        ),
        "workbook_path": str(
            pair.workbook_path
        ),
        "ev_tat_min": metric(
            pair.overall,
            "Total TAT (min)",
            "EV",
        ),
        "ice_tat_min": metric(
            pair.overall,
            "Total TAT (min)",
            "ICE",
        ),
        "ev_distance_km": metric(
            pair.overall,
            "Total Distance (km)",
            "EV",
        ),
        "ice_distance_km": metric(
            pair.overall,
            "Total Distance (km)",
            "ICE",
        ),
        "ev_moving_speed_kmh": metric(
            pair.overall,
            "Average Moving Speed (km/h)",
            "EV",
        ),
        "ice_moving_speed_kmh": metric(
            pair.overall,
            "Average Moving Speed (km/h)",
            "ICE",
        ),
    }

    row["tat_min_diff"] = (
        row["ev_tat_min"]
        - row["ice_tat_min"]
    )

    row["distance_km_diff"] = (
        row["ev_distance_km"]
        - row["ice_distance_km"]
    )

    row["moving_speed_kmh_diff"] = (
        row["ev_moving_speed_kmh"]
        - row["ice_moving_speed_kmh"]
    )

    stationary_ev = (
        phase_value(
            pair.phase,
            "STOP",
            "EV Time (min)",
        )
        + phase_value(
            pair.phase,
            "IDLE",
            "EV Time (min)",
        )
    )

    stationary_ice = (
        phase_value(
            pair.phase,
            "STOP",
            "ICE Time (min)",
        )
        + phase_value(
            pair.phase,
            "IDLE",
            "ICE Time (min)",
        )
    )

    row["stationary_time_min_diff"] = (
        stationary_ev
        - stationary_ice
    )

    for phase_name in PHASES:
        key = slug(
            phase_name
        ).lower()

        ev_time = phase_value(
            pair.phase,
            phase_name,
            "EV Time (min)",
        )

        ice_time = phase_value(
            pair.phase,
            phase_name,
            "ICE Time (min)",
        )

        ev_distance = phase_value(
            pair.phase,
            phase_name,
            "EV Distance (km)",
        )

        ice_distance = phase_value(
            pair.phase,
            phase_name,
            "ICE Distance (km)",
        )

        row[
            f"ev_{key}_time_min"
        ] = ev_time

        row[
            f"ice_{key}_time_min"
        ] = ice_time

        row[
            f"{key}_time_diff_min"
        ] = ev_time - ice_time

        row[
            f"ev_{key}_distance_km"
        ] = ev_distance

        row[
            f"ice_{key}_distance_km"
        ] = ice_distance

        row[
            f"{key}_distance_diff_km"
        ] = ev_distance - ice_distance

    return row


# ============================================================================
# PAIR-LEVEL ANALYTICAL PLOTS
# ============================================================================

def pair_tat_difference_chart(
    pair: PairData,
    out: Path,
) -> Path:
    ev_tat = metric(
        pair.overall,
        "Total TAT (min)",
        "EV",
    )

    ice_tat = metric(
        pair.overall,
        "Total TAT (min)",
        "ICE",
    )

    fig, ax = plt.subplots(
        figsize=(8.5, 5.5)
    )

    bars = ax.bar(
        ["ICE", "EV"],
        [ice_tat, ev_tat],
        color=[ICE_COLOR, EV_COLOR],
        width=0.55,
    )

    label_bars(
        ax,
        bars,
        2,
    )

    if np.isfinite(ev_tat) and np.isfinite(ice_tat):
        difference = ev_tat - ice_tat
        ax.text(
            0.5,
            0.92,
            f"EV − ICE = {difference:+.2f} min",
            transform=ax.transAxes,
            ha="center",
            fontweight="bold",
        )

    ax.set_ylabel(
        "Total TAT (min)"
    )

    ax.set_title(
        f"{pair.pair_id}: Pair-level TAT difference",
        fontweight="bold",
    )

    ax.grid(
        axis="y",
        alpha=0.25,
    )

    path = out / "01_pair_tat_difference.png"
    savefig(fig, path)
    return path


def pair_phase_difference_chart(
    pair: PairData,
    out: Path,
) -> Path:
    phase = pair.phase.copy()

    if (
        "Phase" not in phase.columns
        or "Time Difference EV-ICE (min)" not in phase.columns
    ):
        raise ValueError(
            "Stage-5 phase table does not contain the required "
            "phase difference columns."
        )

    phase["diff"] = pd.to_numeric(
        phase["Time Difference EV-ICE (min)"],
        errors="coerce",
    )

    phase = phase.dropna(
        subset=["diff"]
    )

    phase = phase[
        phase["Phase"]
        .astype(str)
        .str.strip()
        .ne("")
    ]

    phase = phase.sort_values(
        "diff"
    )

    fig, ax = plt.subplots(
        figsize=(10, 5.5)
    )

    values = phase["diff"].to_numpy(
        dtype=float
    )

    colors = np.where(
        values >= 0,
        POSITIVE_COLOR,
        NEGATIVE_COLOR,
    )

    ax.barh(
        phase["Phase"],
        values,
        color=colors,
    )

    ax.axvline(
        0,
        color="black",
        linewidth=1,
    )

    ax.set_xlabel(
        "EV time − ICE time (min)"
    )

    ax.set_title(
        f"{pair.pair_id}: Phase contribution to TAT difference",
        fontweight="bold",
    )

    ax.grid(
        axis="x",
        alpha=0.25,
    )

    path = out / "02_phase_time_difference.png"
    savefig(fig, path)
    return path


def pair_speed_range_chart(
    pair: PairData,
    out: Path,
) -> Path:
    frame = pair.speed.copy()

    required = {
        "Speed Range",
        "EV Time (min)",
        "ICE Time (min)",
    }

    if not required.issubset(
        frame.columns
    ):
        raise ValueError(
            "Stage-5 speed-range table is missing required columns."
        )

    x = np.arange(
        len(frame)
    )

    width = 0.38

    fig, ax = plt.subplots(
        figsize=(10, 5.5)
    )

    ev = pd.to_numeric(
        frame["EV Time (min)"],
        errors="coerce",
    )

    ice = pd.to_numeric(
        frame["ICE Time (min)"],
        errors="coerce",
    )

    ev_bars = ax.bar(
        x - width / 2,
        ev,
        width,
        color=EV_COLOR,
        label="EV",
    )

    ice_bars = ax.bar(
        x + width / 2,
        ice,
        width,
        color=ICE_COLOR,
        label="ICE",
    )

    label_bars(
        ax,
        ev_bars,
        2,
    )

    label_bars(
        ax,
        ice_bars,
        2,
    )

    ax.set_xticks(x)

    ax.set_xticklabels(
        frame["Speed Range"].astype(str),
        rotation=45,
        ha="right",
    )

    ax.set_xlabel(
        "Speed range (km/h)"
    )

    ax.set_ylabel(
        "Time spent (min)"
    )

    ax.set_title(
        f"{pair.pair_id}: Speed-range behaviour",
        fontweight="bold",
    )

    ax.legend()
    ax.grid(
        axis="y",
        alpha=0.25,
    )

    path = out / "03_speed_range_behaviour.png"
    savefig(fig, path)
    return path


def pair_speed_distance_chart(
    pair: PairData,
    out: Path,
) -> Path:
    def cumulative_distance(
        frame: pd.DataFrame,
    ) -> pd.Series:
        if DISTANCE_M not in frame.columns:
            return pd.Series(
                np.arange(len(frame)),
                index=frame.index,
            )

        distance = pd.to_numeric(
            frame[DISTANCE_M],
            errors="coerce",
        ).fillna(0.0)

        return distance.cumsum() / 1000.0

    if (
        SPEED_KMH not in pair.ev_source.columns
        or SPEED_KMH not in pair.ice_source.columns
    ):
        raise ValueError(
            "Speed column not available in cropped telemetry."
        )

    fig, ax = plt.subplots(
        figsize=(10, 5.5)
    )

    ev_distance = cumulative_distance(
        pair.ev_source
    )

    ice_distance = cumulative_distance(
        pair.ice_source
    )

    ev_speed = pd.to_numeric(
        pair.ev_source[SPEED_KMH],
        errors="coerce",
    )

    ice_speed = pd.to_numeric(
        pair.ice_source[SPEED_KMH],
        errors="coerce",
    )

    ax.scatter(
        ev_distance,
        ev_speed,
        s=8,
        alpha=0.45,
        label="EV",
    )

    ax.scatter(
        ice_distance,
        ice_speed,
        s=8,
        alpha=0.45,
        label="ICE",
    )

    ax.set_xlabel(
        "Cumulative distance (km)"
    )

    ax.set_ylabel(
        "Speed (km/h)"
    )

    ax.set_title(
        f"{pair.pair_id}: Speed versus cumulative distance",
        fontweight="bold",
    )

    ax.legend()
    ax.grid(
        alpha=0.25,
    )

    path = out / "04_speed_vs_distance.png"
    savefig(fig, path)
    return path


def pair_charts(
    pair: PairData,
    out: Path,
) -> list[Path]:
    out.mkdir(
        parents=True,
        exist_ok=True,
    )

    paths = []

    builders = [
        pair_tat_difference_chart,
        pair_phase_difference_chart,
        pair_speed_range_chart,
        pair_speed_distance_chart,
    ]

    for builder in builders:
        try:
            paths.append(
                builder(pair, out)
            )
        except Exception as error:
            print(
                f"    Pair plot skipped: "
                f"{builder.__name__}: {error}"
            )

    return paths


# ============================================================================
# MASTER DATA
# ============================================================================

def phase_long(
    master: pd.DataFrame,
) -> pd.DataFrame:
    rows = []

    for _, row in master.iterrows():
        for phase_name in PHASES:
            key = slug(
                phase_name
            ).lower()

            rows.append(
                {
                    "pair_id": row["pair_id"],
                    "phase": phase_name,
                    "ev_time_min": row.get(
                        f"ev_{key}_time_min",
                        np.nan,
                    ),
                    "ice_time_min": row.get(
                        f"ice_{key}_time_min",
                        np.nan,
                    ),
                    "time_diff_min": row.get(
                        f"{key}_time_diff_min",
                        np.nan,
                    ),
                    "ev_distance_km": row.get(
                        f"ev_{key}_distance_km",
                        np.nan,
                    ),
                    "ice_distance_km": row.get(
                        f"ice_{key}_distance_km",
                        np.nan,
                    ),
                    "distance_diff_km": row.get(
                        f"{key}_distance_diff_km",
                        np.nan,
                    ),
                }
            )

    return pd.DataFrame(rows)


# ============================================================================
# MASTER STATISTICAL / ANALYTICAL PLOTS
# ============================================================================

def bootstrap_ci(
    values: np.ndarray,
    iterations: int = 2500,
    seed: int = 42,
) -> tuple[float, float]:
    values = values[
        np.isfinite(values)
    ]

    if len(values) < 2:
        return np.nan, np.nan

    rng = np.random.default_rng(seed)

    samples = rng.choice(
        values,
        size=(iterations, len(values)),
        replace=True,
    )

    means = samples.mean(
        axis=1
    )

    return (
        float(
            np.percentile(
                means,
                2.5,
            )
        ),
        float(
            np.percentile(
                means,
                97.5,
            )
        ),
    )


def _prepare_cluster_matrix(
    master: pd.DataFrame,
    phase_columns: list[str],
) -> tuple[
    pd.DataFrame,
    list[str],
]:
    matrix = finite_numeric(
        master,
        phase_columns,
    )

    if matrix.empty:
        return matrix, []

    matrix.index = master[
        "pair_id"
    ].astype(str)

    constant_columns = []

    variable_columns = []

    for column in matrix.columns:
        finite = matrix[
            column
        ].dropna()

        if finite.nunique() >= 2:
            variable_columns.append(
                column
            )
        else:
            constant_columns.append(
                column
            )

    matrix = matrix[
        variable_columns
    ].copy()

    for column in matrix.columns:
        median = matrix[
            column
        ].median()

        if pd.notna(median):
            matrix[
                column
            ] = matrix[
                column
            ].fillna(median)

    if not matrix.empty:
        finite_rows = np.isfinite(
            matrix.to_numpy(
                dtype=float
            )
        ).all(
            axis=1
        )

        matrix = matrix.loc[
            finite_rows
        ]

    return (
        matrix,
        constant_columns,
    )


def master_figures(
    master: pd.DataFrame,
    out: Path,
    scope_name: str = "INTERCITY",
) -> None:
    out.mkdir(
        parents=True,
        exist_ok=True,
    )

    png = out / "PNG"
    png.mkdir(
        parents=True,
        exist_ok=True,
    )

    master = (
        master.copy()
        .reset_index(drop=True)
    )

    plot_audit: list[
        dict[str, Any]
    ] = []

    phase_frame = phase_long(
        master
    )

    pdf_path = (
        out
        / "Master_Visualizations.pdf"
    )

    with PdfPages(
        pdf_path
    ) as pdf:

        # ------------------------------------------------------------
        # 1. Pair-level TAT difference.
        # This is NOT the Summary's EV-vs-ICE TAT comparison; it shows
        # the signed difference for every pair.
        # ------------------------------------------------------------
        data = master[
            [
                "pair_id",
                "tat_min_diff",
            ]
        ].dropna()

        if not data.empty:
            data = data.sort_values(
                "tat_min_diff"
            )

            fig, ax = plt.subplots(
                figsize=(
                    11,
                    max(
                        6,
                        len(data) * 0.25,
                    ),
                )
            )

            colors = np.where(
                data["tat_min_diff"] >= 0,
                POSITIVE_COLOR,
                NEGATIVE_COLOR,
            )

            ax.barh(
                data["pair_id"],
                data["tat_min_diff"],
                color=colors,
            )

            ax.axvline(
                0,
                color="black",
                linewidth=1,
            )

            ax.set_xlabel(
                "EV TAT − ICE TAT (min)"
            )

            ax.set_title(
                f"{scope_name}: Ranked pair-level TAT difference",
                fontweight="bold",
            )

            savefig(
                fig,
                png / "01_tat_difference_ranked.png",
                pdf,
            )

        # ------------------------------------------------------------
        # 2. TAT difference histogram.
        # ------------------------------------------------------------
        values = pd.to_numeric(
            master["tat_min_diff"],
            errors="coerce",
        ).replace(
            [np.inf, -np.inf],
            np.nan,
        ).dropna()

        if len(values) >= 1:
            fig, ax = plt.subplots(
                figsize=(9, 5.5)
            )

            sns.histplot(
                values,
                bins="auto",
                kde=(
                    len(values) >= 2
                    and values.nunique() >= 2
                ),
                ax=ax,
            )

            ax.axvline(
                0,
                color="black",
                linestyle="--",
                linewidth=1,
            )

            ax.set_xlabel(
                "EV TAT − ICE TAT (min)"
            )

            ax.set_title(
                f"{scope_name}: Distribution of TAT differences",
                fontweight="bold",
            )

            savefig(
                fig,
                png / "02_tat_difference_distribution.png",
                pdf,
            )

        # ------------------------------------------------------------
        # 3. TAT distribution: EV and ICE.
        # ------------------------------------------------------------
        dist = finite_numeric(
            master,
            [
                "ev_tat_min",
                "ice_tat_min",
            ],
        ).dropna()

        if not dist.empty:
            fig, ax = plt.subplots(
                figsize=(8.5, 5.5)
            )

            sns.boxplot(
                data=dist.rename(
                    columns={
                        "ev_tat_min": "EV",
                        "ice_tat_min": "ICE",
                    }
                ),
                ax=ax,
            )

            sns.stripplot(
                data=dist.rename(
                    columns={
                        "ev_tat_min": "EV",
                        "ice_tat_min": "ICE",
                    }
                ),
                ax=ax,
                size=4,
                alpha=0.55,
            )

            ax.set_ylabel(
                "Total TAT (min)"
            )

            ax.set_title(
                f"{scope_name}: EV and ICE TAT distribution",
                fontweight="bold",
            )

            savefig(
                fig,
                png / "03_tat_distribution.png",
                pdf,
            )

        # ------------------------------------------------------------
        # 4. Phase-time difference heatmap.
        # ------------------------------------------------------------
        phase_columns = [
            f"{slug(p).lower()}_time_diff_min"
            for p in PHASES
        ]

        available_phase_columns = [
            column
            for column in phase_columns
            if column in master.columns
        ]

        if available_phase_columns:
            heat = master.set_index(
                "pair_id"
            )[
                available_phase_columns
            ].copy()

            readable = {
                f"{slug(p).lower()}_time_diff_min": p
                for p in PHASES
            }

            heat = heat.rename(
                columns=readable
            )

            heat = heat.loc[
                master.sort_values(
                    "tat_min_diff"
                )["pair_id"]
            ]

            finite_values = heat.to_numpy(
                dtype=float
            )

            if np.isfinite(
                finite_values
            ).any():
                limit = np.nanpercentile(
                    np.abs(
                        finite_values
                    ),
                    95,
                )

                if (
                    not np.isfinite(limit)
                    or limit <= 0
                ):
                    limit = 1.0

                fig, ax = plt.subplots(
                    figsize=(
                        10,
                        max(
                            7,
                            len(heat) * 0.22,
                        ),
                    )
                )

                sns.heatmap(
                    heat,
                    cmap="RdBu_r",
                    center=0,
                    vmin=-limit,
                    vmax=limit,
                    ax=ax,
                    cbar_kws={
                        "label":
                        "EV − ICE phase time (min)"
                    },
                )

                ax.set_title(
                    f"{scope_name}: Pair-by-phase time difference",
                    fontweight="bold",
                )

                savefig(
                    fig,
                    png / "04_phase_difference_heatmap.png",
                    pdf,
                )

        # ------------------------------------------------------------
        # 5. Mean phase contribution to TAT difference.
        # ------------------------------------------------------------
        if not phase_frame.empty:
            phase_mean = (
                phase_frame.groupby(
                    "phase"
                )[
                    "time_diff_min"
                ]
                .mean()
                .reindex(PHASES)
            )

            phase_mean = phase_mean.dropna()

            if not phase_mean.empty:
                fig, ax = plt.subplots(
                    figsize=(10, 5.5)
                )

                colors = np.where(
                    phase_mean.to_numpy()
                    >= 0,
                    POSITIVE_COLOR,
                    NEGATIVE_COLOR,
                )

                ax.bar(
                    phase_mean.index,
                    phase_mean.values,
                    color=colors,
                )

                ax.axhline(
                    0,
                    color="black",
                    linewidth=1,
                )

                ax.set_ylabel(
                    "Mean EV − ICE phase time (min)"
                )

                ax.set_title(
                    f"{scope_name}: Mean phase contribution to TAT difference",
                    fontweight="bold",
                )

                ax.tick_params(
                    axis="x",
                    rotation=35,
                )

                savefig(
                    fig,
                    png / "05_mean_phase_difference.png",
                    pdf,
                )

        # ------------------------------------------------------------
        # 6. Pearson and Spearman correlation of pair-level differences.
        # ------------------------------------------------------------
        corr_columns = [
            "tat_min_diff",
            "distance_km_diff",
            "moving_speed_kmh_diff",
            "stationary_time_min_diff",
            *available_phase_columns,
        ]

        corr_columns = [
            column
            for column in corr_columns
            if column in master.columns
        ]

        if len(corr_columns) >= 2:
            for method in (
                "pearson",
                "spearman",
            ):
                corr = master[
                    corr_columns
                ].corr(
                    method=method
                )

                fig, ax = plt.subplots(
                    figsize=(12, 10)
                )

                sns.heatmap(
                    corr,
                    cmap="vlag",
                    center=0,
                    vmin=-1,
                    vmax=1,
                    annot=True,
                    fmt=".2f",
                    square=True,
                    ax=ax,
                )

                ax.set_title(
                    f"{scope_name}: {method.title()} correlation of pair-level differences",
                    fontweight="bold",
                )

                savefig(
                    fig,
                    png / (
                        f"06_{method}_correlation_heatmap.png"
                    ),
                    pdf,
                )

        # ------------------------------------------------------------
        # 7. EV vs ICE equality scatters.
        # These assess agreement rather than repeating Summary bars.
        # ------------------------------------------------------------
        equality_specs = [
            (
                "ev_tat_min",
                "ice_tat_min",
                "TAT (min)",
                "EV versus ICE total TAT",
                "08_tat_scatter_equality.png",
            ),
            (
                "ev_moving_speed_kmh",
                "ice_moving_speed_kmh",
                "Moving speed (km/h)",
                "EV versus ICE moving speed",
                "09_moving_speed_scatter_equality.png",
            ),
            (
                "ev_distance_km",
                "ice_distance_km",
                "Distance (km)",
                "EV versus ICE distance",
                "10_distance_scatter_equality.png",
            ),
        ]

        for (
            ev_column,
            ice_column,
            label,
            title,
            filename,
        ) in equality_specs:

            data = finite_numeric(
                master,
                [
                    ev_column,
                    ice_column,
                ],
            ).dropna()

            if data.empty:
                plot_audit.append(
                    {
                        "Plot": filename,
                        "Status": "SKIPPED",
                        "Reason": "No complete finite observations.",
                    }
                )
                continue

            fig, ax = plt.subplots(
                figsize=(7.5, 6.5)
            )

            ax.scatter(
                data[ice_column],
                data[ev_column],
                alpha=0.75,
            )

            minimum = float(
                data.min().min()
            )

            maximum = float(
                data.max().max()
            )

            if (
                np.isfinite(minimum)
                and np.isfinite(maximum)
            ):
                if minimum == maximum:
                    pad = max(
                        abs(minimum) * 0.05,
                        1.0,
                    )
                    minimum -= pad
                    maximum += pad

                ax.plot(
                    [minimum, maximum],
                    [minimum, maximum],
                    linestyle="--",
                    color="black",
                    label="Equality line",
                )

            ax.set_xlabel(
                f"ICE {label}"
            )

            ax.set_ylabel(
                f"EV {label}"
            )

            ax.set_title(
                f"{scope_name}: {title}",
                fontweight="bold",
            )

            ax.legend()

            savefig(
                fig,
                png / filename,
                pdf,
            )

        # ------------------------------------------------------------
        # 8. Bland-Altman TAT agreement.
        # ------------------------------------------------------------
        bland = finite_numeric(
            master,
            [
                "ev_tat_min",
                "ice_tat_min",
            ],
        ).dropna()

        if len(bland) >= 1:
            ev = bland[
                "ev_tat_min"
            ].to_numpy(float)

            ice = bland[
                "ice_tat_min"
            ].to_numpy(float)

            mean_values = (
                ev + ice
            ) / 2.0

            differences = (
                ev - ice
            )

            mean_difference = float(
                np.mean(
                    differences
                )
            )

            fig, ax = plt.subplots(
                figsize=(8.5, 6)
            )

            ax.scatter(
                mean_values,
                differences,
                alpha=0.75,
            )

            ax.axhline(
                mean_difference,
                color="black",
                label=(
                    f"Mean = "
                    f"{mean_difference:.2f}"
                ),
            )

            if len(differences) >= 2:
                standard_deviation = float(
                    np.std(
                        differences,
                        ddof=1,
                    )
                )

                if np.isfinite(
                    standard_deviation
                ):
                    ax.axhline(
                        mean_difference
                        + 1.96
                        * standard_deviation,
                        linestyle="--",
                        label="+1.96 SD",
                    )

                    ax.axhline(
                        mean_difference
                        - 1.96
                        * standard_deviation,
                        linestyle="--",
                        label="−1.96 SD",
                    )

            ax.set_xlabel(
                "Mean of EV and ICE TAT (min)"
            )

            ax.set_ylabel(
                "EV − ICE TAT (min)"
            )

            ax.set_title(
                f"{scope_name}: Bland–Altman TAT agreement",
                fontweight="bold",
            )

            ax.legend()

            savefig(
                fig,
                png / "11_bland_altman_tat.png",
                pdf,
            )

        # ------------------------------------------------------------
        # 9. Pareto of phase effects.
        # ------------------------------------------------------------
        if not phase_frame.empty:
            phase_effect = (
                phase_frame.groupby(
                    "phase"
                )[
                    "time_diff_min"
                ]
                .apply(
                    lambda series:
                    series.abs().mean()
                )
                .reindex(PHASES)
                .dropna()
            )

            if (
                not phase_effect.empty
                and phase_effect.sum() > 0
            ):
                phase_effect = phase_effect.sort_values(
                    ascending=False
                )

                cumulative = (
                    phase_effect.cumsum()
                    / phase_effect.sum()
                    * 100.0
                )

                fig, ax = plt.subplots(
                    figsize=(11, 6)
                )

                ax2 = ax.twinx()

                ax.bar(
                    phase_effect.index,
                    phase_effect.values,
                )

                ax2.plot(
                    phase_effect.index,
                    cumulative.values,
                    marker="o",
                )

                ax2.axhline(
                    80,
                    linestyle="--",
                    color="black",
                )

                ax.set_ylabel(
                    "Mean absolute phase difference (min)"
                )

                ax2.set_ylabel(
                    "Cumulative contribution (%)"
                )

                ax.set_title(
                    f"{scope_name}: Pareto analysis of phase effects",
                    fontweight="bold",
                )

                ax.tick_params(
                    axis="x",
                    rotation=35,
                )

                savefig(
                    fig,
                    png / "12_phase_pareto.png",
                    pdf,
                )

        # ------------------------------------------------------------
        # 10. Route-level forest plot.
        # ------------------------------------------------------------
        if "route" in master.columns:
            work = master.loc[
                master["route"]
                .astype(str)
                .str.strip()
                .ne("")
            ].copy()

            if not work.empty:
                route_counts = (
                    work["route"]
                    .value_counts()
                )

                work = work.loc[
                    work["route"].isin(
                        route_counts[
                            route_counts >= 2
                        ].index
                    )
                ]

                if not work.empty:
                    rows = []

                    for route, group in work.groupby(
                        "route"
                    ):
                        values = (
                            pd.to_numeric(
                                group["tat_min_diff"],
                                errors="coerce",
                            )
                            .dropna()
                            .to_numpy(float)
                        )

                        if len(values) < 2:
                            continue

                        lower, upper = bootstrap_ci(
                            values
                        )

                        rows.append(
                            (
                                route,
                                float(values.mean()),
                                lower,
                                upper,
                                len(values),
                            )
                        )

                    if rows:
                        forest = pd.DataFrame(
                            rows,
                            columns=[
                                "route",
                                "mean",
                                "lower",
                                "upper",
                                "n",
                            ],
                        ).sort_values(
                            "mean"
                        )

                        y = np.arange(
                            len(forest)
                        )

                        fig, ax = plt.subplots(
                            figsize=(
                                11,
                                max(
                                    5,
                                    len(forest) * 0.65,
                                ),
                            )
                        )

                        ax.errorbar(
                            forest["mean"],
                            y,
                            xerr=[
                                forest["mean"]
                                - forest["lower"],
                                forest["upper"]
                                - forest["mean"],
                            ],
                            fmt="o",
                            capsize=4,
                        )

                        ax.axvline(
                            0,
                            color="black",
                            linewidth=1,
                        )

                        ax.set_yticks(y)

                        ax.set_yticklabels(
                            [
                                f"{route} (n={n})"
                                for route, n in zip(
                                    forest["route"],
                                    forest["n"],
                                )
                            ]
                        )

                        ax.set_xlabel(
                            "Mean EV − ICE TAT difference (min), "
                            "95% bootstrap CI"
                        )

                        ax.set_title(
                            f"{scope_name}: Route-level TAT difference",
                            fontweight="bold",
                        )

                        savefig(
                            fig,
                            png / "13_route_forest.png",
                            pdf,
                        )

        # ------------------------------------------------------------
        # 11. STRICT vs NONSTRICT.
        # ------------------------------------------------------------
        if (
            "pair_class" in master.columns
            and master["pair_class"]
            .astype(str)
            .str.strip()
            .ne("")
            .sum() > 0
            and master["pair_class"]
            .nunique() > 1
        ):
            fig, ax = plt.subplots(
                figsize=(9, 6)
            )

            sns.violinplot(
                data=master,
                x="pair_class",
                y="tat_min_diff",
                inner=None,
                ax=ax,
            )

            sns.boxplot(
                data=master,
                x="pair_class",
                y="tat_min_diff",
                width=0.22,
                showfliers=False,
                ax=ax,
            )

            sns.stripplot(
                data=master,
                x="pair_class",
                y="tat_min_diff",
                alpha=0.5,
                size=4,
                ax=ax,
            )

            ax.axhline(
                0,
                color="black",
                linewidth=1,
            )

            ax.set_ylabel(
                "EV − ICE TAT (min)"
            )

            ax.set_title(
                f"{scope_name}: STRICT vs NONSTRICT TAT difference",
                fontweight="bold",
            )

            savefig(
                fig,
                png / "14_strict_nonstrict_tat.png",
                pdf,
            )

        # ------------------------------------------------------------
        # 12. PCA of EV-ICE operational differences.
        # ------------------------------------------------------------
        try:
            from sklearn.decomposition import PCA
            from sklearn.preprocessing import StandardScaler

            pca_columns = [
                *available_phase_columns,
                "moving_speed_kmh_diff",
                "distance_km_diff",
            ]

            data = finite_numeric(
                master,
                pca_columns,
            )

            variable_columns = [
                column
                for column in data.columns
                if data[column]
                .dropna()
                .nunique() >= 2
            ]

            data = data[
                variable_columns
            ].copy()

            for column in data.columns:
                median = data[
                    column
                ].median()

                if pd.notna(median):
                    data[
                        column
                    ] = data[
                        column
                    ].fillna(median)

            if not data.empty:
                finite_rows = np.isfinite(
                    data.to_numpy(
                        dtype=float
                    )
                ).all(
                    axis=1
                )

                data = data.loc[
                    finite_rows
                ]

            if (
                len(data) >= 2
                and data.shape[1] >= 2
            ):
                scaled = StandardScaler().fit_transform(
                    data
                )

                pca = PCA(
                    n_components=2
                )

                components = pca.fit_transform(
                    scaled
                )

                fig, ax = plt.subplots(
                    figsize=(8.5, 6.5)
                )

                ax.scatter(
                    components[:, 0],
                    components[:, 1],
                    alpha=0.75,
                )

                ax.set_xlabel(
                    "PC1 "
                    f"({pca.explained_variance_ratio_[0] * 100:.1f}%)"
                )

                ax.set_ylabel(
                    "PC2 "
                    f"({pca.explained_variance_ratio_[1] * 100:.1f}%)"
                )

                ax.set_title(
                    f"{scope_name}: PCA of EV−ICE operational differences",
                    fontweight="bold",
                )

                savefig(
                    fig,
                    png / "15_pca_operational_difference.png",
                    pdf,
                )

            else:
                plot_audit.append(
                    {
                        "Plot": "15_pca_operational_difference.png",
                        "Status": "SKIPPED",
                        "Reason": (
                            "PCA requires at least 2 valid pairs "
                            "and 2 varying numerical features."
                        ),
                    }
                )

        except ImportError:
            plot_audit.append(
                {
                    "Plot": "15_pca_operational_difference.png",
                    "Status": "SKIPPED",
                    "Reason": "scikit-learn is not installed.",
                }
            )

        except Exception as error:
            plot_audit.append(
                {
                    "Plot": "15_pca_operational_difference.png",
                    "Status": "SKIPPED",
                    "Reason": str(error),
                }
            )

        # ------------------------------------------------------------
        # 13. Hierarchical clustering.
        # ------------------------------------------------------------
        try:
            phase_columns = [
                f"{slug(p).lower()}_time_diff_min"
                for p in PHASES
            ]

            matrix, constant_columns = (
                _prepare_cluster_matrix(
                    master,
                    phase_columns,
                )
            )

            if constant_columns:
                plot_audit.append(
                    {
                        "Plot": "16_clustered_phase_heatmap.png",
                        "Status": "INFO",
                        "Reason": (
                            "Constant phase dimensions omitted: "
                            + ", ".join(
                                constant_columns
                            )
                        ),
                    }
                )

            if (
                len(matrix) >= 2
                and matrix.shape[1] >= 1
            ):
                readable = {
                    f"{slug(p).lower()}_time_diff_min": p
                    for p in PHASES
                }

                matrix = matrix.rename(
                    columns=readable
                )

                if np.isfinite(
                    matrix.to_numpy(
                        dtype=float
                    )
                ).all():
                    cluster = sns.clustermap(
                        matrix,
                        cmap="RdBu_r",
                        center=0,
                        z_score=1
                        if matrix.shape[0] >= 2
                        else None,
                        figsize=(
                            10,
                            max(
                                8,
                                len(matrix) * 0.22,
                            ),
                        ),
                    )

                    cluster.fig.suptitle(
                        f"{scope_name}: Hierarchical clustering of "
                        "phase-difference signatures",
                        y=1.02,
                    )

                    cluster.fig.savefig(
                        png / "16_clustered_phase_heatmap.png",
                        dpi=DPI,
                        bbox_inches="tight",
                        facecolor="white",
                    )

                    pdf.savefig(
                        cluster.fig,
                        bbox_inches="tight",
                        facecolor="white",
                    )

                    plt.close(
                        cluster.fig
                    )

            else:
                plot_audit.append(
                    {
                        "Plot": "16_clustered_phase_heatmap.png",
                        "Status": "SKIPPED",
                        "Reason": (
                            "Clustering requires at least 2 valid pairs "
                            "and at least 1 varying phase feature."
                        ),
                    }
                )

        except ImportError:
            plot_audit.append(
                {
                    "Plot": "16_clustered_phase_heatmap.png",
                    "Status": "SKIPPED",
                    "Reason": "SciPy/seaborn clustering dependency unavailable.",
                }
            )

        except Exception as error:
            plot_audit.append(
                {
                    "Plot": "16_clustered_phase_heatmap.png",
                    "Status": "SKIPPED",
                    "Reason": str(error),
                }
            )

    # ------------------------------------------------------------
    # Master Excel workbook.
    # ------------------------------------------------------------
    descriptive = (
        master.select_dtypes(
            include=[np.number]
        )
        .describe()
        .T
        .reset_index()
        .rename(
            columns={
                "index": "Metric"
            }
        )
    )

    if not phase_frame.empty:
        phase_average = (
            phase_frame.groupby(
                "phase"
            )
            .agg(
                EV_Mean_Time_min=(
                    "ev_time_min",
                    "mean",
                ),
                ICE_Mean_Time_min=(
                    "ice_time_min",
                    "mean",
                ),
                Mean_Time_Diff_min=(
                    "time_diff_min",
                    "mean",
                ),
                Median_Time_Diff_min=(
                    "time_diff_min",
                    "median",
                ),
                EV_Mean_Distance_km=(
                    "ev_distance_km",
                    "mean",
                ),
                ICE_Mean_Distance_km=(
                    "ice_distance_km",
                    "mean",
                ),
            )
            .reindex(PHASES)
            .reset_index()
        )
    else:
        phase_average = pd.DataFrame()

    route_summary = pd.DataFrame()

    if "route" in master.columns:
        route_summary = (
            master.loc[
                master["route"]
                .astype(str)
                .str.strip()
                .ne("")
            ]
            .groupby("route")
            .agg(
                Pair_Count=(
                    "pair_id",
                    "size",
                ),
                Mean_EV_TAT_min=(
                    "ev_tat_min",
                    "mean",
                ),
                Mean_ICE_TAT_min=(
                    "ice_tat_min",
                    "mean",
                ),
                Mean_TAT_Difference_min=(
                    "tat_min_diff",
                    "mean",
                ),
                Median_TAT_Difference_min=(
                    "tat_min_diff",
                    "median",
                ),
            )
            .reset_index()
        )

    with pd.ExcelWriter(
        out / "Master_Results.xlsx",
        engine="openpyxl",
    ) as writer:
        master.to_excel(
            writer,
            sheet_name="All_Pairs",
            index=False,
        )

        phase_frame.to_excel(
            writer,
            sheet_name="Pair_Phase_Long",
            index=False,
        )

        phase_average.to_excel(
            writer,
            sheet_name="Phase_Difference_Summary",
            index=False,
        )

        descriptive.to_excel(
            writer,
            sheet_name="Descriptive_Statistics",
            index=False,
        )

        route_summary.to_excel(
            writer,
            sheet_name="Route_Summary",
            index=False,
        )

        if plot_audit:
            pd.DataFrame(
                plot_audit
            ).to_excel(
                writer,
                sheet_name="Plot_Audit",
                index=False,
            )

    if plot_audit:
        pd.DataFrame(
            plot_audit
        ).to_excel(
            out / "Master_Plot_Audit.xlsx",
            index=False,
        )


# ============================================================================
# PAIR WORKBOOK EMBEDDING
# ============================================================================

def _style_excel_title(
    worksheet,
    title: str,
    end_column: int = 14,
) -> None:
    worksheet.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=end_column,
    )

    cell = worksheet.cell(
        1,
        1,
        title,
    )

    cell.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    cell.font = Font(
        color="FFFFFF",
        bold=True,
        size=16,
    )

    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    worksheet.row_dimensions[
        1
    ].height = 28


def _set_dashboard_dimensions(
    worksheet,
) -> None:
    for column in range(
        1,
        15,
    ):
        worksheet.column_dimensions[
            get_column_letter(column)
        ].width = 14

    for row in range(
        1,
        160,
    ):
        worksheet.row_dimensions[
            row
        ].height = 18


def _insert_excel_image(
    worksheet,
    path: Path,
    anchor: str,
    width: int = 700,
    height: int = 350,
) -> None:
    if not path.exists():
        return

    image = ExcelImage(
        str(path)
    )

    image.width = width
    image.height = height

    worksheet.add_image(
        image,
        anchor,
    )


def _write_pair_kpis(
    worksheet,
    pair: PairData,
) -> None:
    info = read_pair_info(
        pair.workbook_path
    )

    route = route_label(
        pair.metadata,
        info,
        pair.workbook_path,
    )

    values = [
        (
            "Pair ID",
            pair.pair_id,
        ),
        (
            "Operation",
            "INTERCITY",
        ),
        (
            "Pair Class",
            pair_class(
                pair.metadata,
                pair.workbook_path,
            ),
        ),
        (
            "Route",
            route,
        ),
        (
            "EV TAT (min)",
            metric(
                pair.overall,
                "Total TAT (min)",
                "EV",
            ),
        ),
        (
            "ICE TAT (min)",
            metric(
                pair.overall,
                "Total TAT (min)",
                "ICE",
            ),
        ),
        (
            "TAT Difference (EV-ICE)",
            (
                metric(
                    pair.overall,
                    "Total TAT (min)",
                    "EV",
                )
                - metric(
                    pair.overall,
                    "Total TAT (min)",
                    "ICE",
                )
            ),
        ),
    ]

    row = 3

    for label, value in values:
        worksheet.cell(
            row=row,
            column=1,
            value=label,
        ).font = Font(
            bold=True
        )

        worksheet.cell(
            row=row,
            column=2,
            value=value,
        )

        row += 1


def embed_pair_visualizations(
    pair: PairData,
    image_dir: Path,
) -> None:
    """
    Embed only Stage-6 analytical plots.

    The Summary-owned basic phase-time, phase-distance, moving-speed and
    EV-vs-ICE summary plots are intentionally not embedded here.
    """
    workbook = load_workbook(
        pair.workbook_path,
        data_only=False,
    )

    try:
        for sheet_name in (
            DASHBOARD_SHEET,
            CHARTS_SHEET,
        ):
            if sheet_name in workbook.sheetnames:
                workbook.remove(
                    workbook[sheet_name]
                )

        dashboard = workbook.create_sheet(
            DASHBOARD_SHEET
        )

        _style_excel_title(
            dashboard,
            (
                f"{pair.pair_id}: "
                "INTERCITY Analytical Visualization"
            ),
        )

        _set_dashboard_dimensions(
            dashboard
        )

        _write_pair_kpis(
            dashboard,
            pair,
        )

        dashboard["A13"] = (
            "Stage 6 contains analytical plots only. "
            "Basic EV/ICE summary comparisons are maintained in "
            "Stage 9 Master Summary."
        )

        dashboard["A13"].alignment = Alignment(
            wrap_text=True
        )

        charts = workbook.create_sheet(
            CHARTS_SHEET
        )

        _style_excel_title(
            charts,
            (
                f"{pair.pair_id}: "
                "Detailed Analytical Visualizations"
            ),
        )

        _set_dashboard_dimensions(
            charts
        )

        layout = [
            (
                "01_pair_tat_difference.png",
                "A3",
            ),
            (
                "02_phase_time_difference.png",
                "H3",
            ),
            (
                "03_speed_range_behaviour.png",
                "A28",
            ),
            (
                "04_speed_vs_distance.png",
                "H28",
            ),
        ]

        for filename, anchor in layout:
            _insert_excel_image(
                charts,
                image_dir / filename,
                anchor,
                width=680,
                height=350,
            )

        temporary_path = (
            pair.workbook_path.with_suffix(
                ".stage6tmp.xlsx"
            )
        )

        workbook.save(
            temporary_path
        )

        workbook.close()

        os.replace(
            temporary_path,
            pair.workbook_path,
        )

    finally:
        try:
            workbook.close()
        except Exception:
            pass


# ============================================================================
# OUTPUT DIRECTORY
# ============================================================================

def prepare_visualization_root(
    root: Path,
) -> None:
    root.mkdir(
        parents=True,
        exist_ok=True,
    )

    # This is now an INTERCITY-only visualization stage.
    for old_name in (
        "CITY",
        "INTERCITY",
        "Per_Pair",
        "Master",
    ):
        old_path = root / old_name

        if old_path.exists():
            shutil.rmtree(
                old_path
            )

    (
        root
        / "INTERCITY"
        / "Per_Pair"
    ).mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        root
        / "INTERCITY"
        / "Average"
    ).mkdir(
        parents=True,
        exist_ok=True,
    )

    (
        root
        / "INTERCITY"
        / "Route_Wise"
    ).mkdir(
        parents=True,
        exist_ok=True,
    )


def generate_route_wise_results(
    master: pd.DataFrame,
    category_root: Path,
) -> None:
    if (
        "route" not in master.columns
        or master.empty
    ):
        return

    route_root = (
        category_root
        / "Route_Wise"
    )

    rows = []

    work = master.loc[
        master["route"]
        .astype(str)
        .str.strip()
        .ne("")
    ].copy()

    for route, route_data in work.groupby(
        "route",
        dropna=False,
    ):
        route_name = clean(route)

        if not route_name:
            continue

        route_folder = (
            route_root
            / slug(route_name)
            / "Average"
        )

        route_folder.mkdir(
            parents=True,
            exist_ok=True,
        )

        master_figures(
            route_data.reset_index(
                drop=True
            ),
            route_folder,
            scope_name=f"Route: {route_name}",
        )

        rows.append(
            {
                "Route": route_name,
                "Paired Workbooks": len(route_data),
                "Average Results Folder": str(
                    route_folder
                ),
            }
        )

    if rows:
        pd.DataFrame(rows).to_excel(
            route_root
            / "Route_Wise_Index.xlsx",
            index=False,
        )


# ============================================================================
# MAIN RUN
# ============================================================================

def run(
    config: dict,
) -> dict[str, Any]:
    results_root = Path(
        config["project"]["results_folder"]
    )

    pair_index_path = (
        results_root
        / "System"
        / PAIR_INDEX_NAME
    )

    if not pair_index_path.exists():
        raise FileNotFoundError(
            pair_index_path
        )

    root = (
        results_root
        / "Visualizations"
    )

    prepare_visualization_root(
        root
    )

    intercity_root = (
        root
        / "INTERCITY"
    )

    pairs = discover_intercity_pairs(
        results_root
    )

    print("=" * 88)
    print(
        "STAGE 6: INTERCITY ANALYTICAL RESULTS VISUALIZATION"
    )
    print("=" * 88)
    print(
        f"Script version: {SCRIPT_VERSION}"
    )
    print(
        f"Pair workbooks: {len(pairs)}"
    )
    print(
        f"Output: {intercity_root}"
    )
    print()
    print(
        "Summary-owned plots NOT generated by this stage:"
    )

    for item in sorted(
        SUMMARY_OWNED_PLOTS
    ):
        print(
            f"  - {item}"
        )

    print()
    print(
        "Stage-6 analytical plots:"
    )
    print(
        "  - pair-level TAT difference"
    )
    print(
        "  - TAT difference distribution"
    )
    print(
        "  - EV/ICE TAT distribution"
    )
    print(
        "  - phase difference heatmap"
    )
    print(
        "  - phase contribution analysis"
    )
    print(
        "  - Pearson/Spearman correlations"
    )
    print(
        "  - EV/ICE equality scatter plots"
    )
    print(
        "  - Bland-Altman"
    )
    print(
        "  - phase Pareto"
    )
    print(
        "  - route forest"
    )
    print(
        "  - STRICT/NONSTRICT"
    )
    print(
        "  - PCA"
    )
    print(
        "  - hierarchical clustering"
    )
    print()

    started = time.monotonic()

    master_rows = []
    audit_rows = []

    for sequence, (
        pair_id,
        metadata,
        workbook_path,
    ) in enumerate(
        pairs,
        start=1,
    ):
        pair_start = time.monotonic()

        print(
            f"[{sequence}/{len(pairs)}] "
            f"{pair_id} | "
            f"{workbook_path.name}",
            flush=True,
        )

        status = "COMPLETED"
        error = "NONE"

        try:
            (
                overall,
                phase,
                speed,
                ev_source,
                ice_source,
            ) = extract_tables(
                workbook_path
            )

            pair = PairData(
                pair_id=pair_id,
                workbook_path=workbook_path,
                metadata=metadata,
                overall=overall,
                phase=phase,
                speed=speed,
                ev_source=ev_source,
                ice_source=ice_source,
            )

            pair_output_dir = (
                intercity_root
                / "Per_Pair"
                / slug(pair_id)
            )

            pair_charts(
                pair,
                pair_output_dir,
            )

            if EMBED_VISUALS_IN_PAIR_WORKBOOKS:
                embed_pair_visualizations(
                    pair,
                    pair_output_dir,
                )

            master_rows.append(
                pair_to_master(
                    pair
                )
            )

            print(
                "    COMPLETED | "
                "INTERCITY",
                flush=True,
            )

        except Exception as exc:
            status = "ERROR"
            error = str(exc)

            print(
                f"    ERROR | {error}",
                flush=True,
            )

        elapsed = time.monotonic() - pair_start

        audit_rows.append(
            {
                "pair_id": pair_id,
                "operation_type": "INTERCITY",
                "workbook_path": str(
                    workbook_path
                ),
                "status": status,
                "error": error,
                "elapsed_s": elapsed,
            }
        )

        print(
            f"    Time: {elapsed:.1f}s",
            flush=True,
        )
        print("-" * 88)

    master = pd.DataFrame(
        master_rows
    )

    if master.empty:
        raise RuntimeError(
            "No INTERCITY pair visualizations were completed."
        )

    master = master.sort_values(
        [
            "pair_number",
            "pair_id",
        ]
    ).reset_index(
        drop=True
    )

    print()
    print("=" * 88)
    print("INTERCITY MASTER ANALYSIS")
    print("=" * 88)
    print(
        f"Completed pairs: {len(master)}"
    )

    master_figures(
        master,
        intercity_root / "Average",
        scope_name="INTERCITY",
    )

    generate_route_wise_results(
        master,
        intercity_root,
    )

    category_index = pd.DataFrame(
        [
            {
                "Category": "INTERCITY",
                "Paired Workbooks": len(master),
            }
        ]
    )

    with pd.ExcelWriter(
        root
        / "Visualization_Category_Index.xlsx",
        engine="openpyxl",
    ) as writer:
        category_index.to_excel(
            writer,
            sheet_name="Category_Summary",
            index=False,
        )

        master.to_excel(
            writer,
            sheet_name="All_Intercity_Pairs",
            index=False,
        )

    pd.DataFrame(
        audit_rows
    ).to_csv(
        root
        / "Visualization_Audit.csv",
        index=False,
        encoding="utf-8-sig",
    )

    elapsed_total = (
        time.monotonic()
        - started
    )

    completed = sum(
        row["status"] == "COMPLETED"
        for row in audit_rows
    )

    errors = sum(
        row["status"] == "ERROR"
        for row in audit_rows
    )

    print()
    print("=" * 88)
    print("STAGE 6 COMPLETE")
    print("=" * 88)
    print(
        f"INTERCITY pairs processed : {len(master)}"
    )
    print(
        f"Completed                  : {completed}"
    )
    print(
        f"Errors                     : {errors}"
    )
    print(
        f"Total elapsed              : {elapsed_total:.1f}s"
    )
    print(
        f"INTERCITY visualizations  : {intercity_root}"
    )
    print(
        f"Average analysis          : "
        f"{intercity_root / 'Average'}"
    )
    print(
        f"Route-wise analysis       : "
        f"{intercity_root / 'Route_Wise'}"
    )
    print("=" * 88)

    if errors:
        print(
            "\nWARNING: Some pair workbooks had errors. "
            "Review Visualization_Audit.csv."
        )

    return {
        "pair_count": len(master),
        "completed_count": completed,
        "error_count": errors,
        "output_root": str(root),
        "intercity_root": str(intercity_root),
        "audit_path": str(
            root / "Visualization_Audit.csv"
        ),
    }


def main() -> None:
    try:
        result = run(
            load_config()
        )

        print()
        print(
            "INTERCITY VISUALIZATION "
            "COMPLETED SUCCESSFULLY"
        )
        print(
            f"Pairs processed: "
            f"{result['completed_count']}"
        )

    except Exception as error:
        print(
            f"\nTAT STAGE 6 FAILED | {error}"
        )
        raise SystemExit(1)


if __name__ == "__main__":
    main()
