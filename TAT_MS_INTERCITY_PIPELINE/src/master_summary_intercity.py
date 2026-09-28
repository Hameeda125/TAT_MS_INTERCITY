"""
Stage 9/9 - Master EV vs ICE Summary Workbook
==============================================

Run AFTER:
    1. phase_segmentation.py
    2. tat_computation.py

Purpose
-------
Create ONE master Excel workbook containing the summary tables requested:

1. ALL PAIRS
   - all pairs in this INTERCITY-only pipeline
   - Table 1: EV vs ICE Trip Summary
   - Table 2: Phase Time Comparison (min)
   - Table 3: Phase Distance Comparison (km)
   - Summary charts with values displayed on the bars

2. INTERCITY
   - same three tables
   - charts

3. SAME-DAY EV-ICE PAIRS
   - all same-calendar-day INTERCITY pairs
   - same three tables
   - pair-wise overall TAT chart
   - phase-time and phase-distance charts

4. DIFFERENT-DAY, SIMILAR-START-TIME EV-ICE PAIRS
   - EV and ICE trip dates must be different
   - circular clock-time difference between trip starts must be <= 2 hours
   - all qualifying INTERCITY pairs
   - same three tables and charts

5. ROUTE-WISE
   - one sheet per INTERCITY route group
   - same three tables
   - charts
   - both route directions are kept inside the same route-group sheet

6. ROUTE-DIRECTION-WISE
   - one additional sheet for every INTERCITY route direction
   - e.g. Chennai → Pondicherry and Pondicherry → Chennai are separate sheets
   - each sheet contains only that direction's results and charts

The script DOES NOT modify any pair workbook.
It only reads the pair workbooks and creates a master summary workbook.

Input Stage-5 sheet expected in every pair workbook:
    08_TAT_Phase_Comparison

Input source sheets:
    01_EV_Cropped_Data
    02_ICE_Cropped_Data

Optional metadata sheet:
    03_Pair_Crop_Info

Primary results location:
    D:/TAT_MS/TAT_Data/Results

Output:
    D:/TAT_MS_INTERCITY/Data/Results/Master_Summary/
        TAT_All_Data_Summary.xlsx
        Charts/*.png
        Training_Data/
            EV_ICE_Training_STRICT.xlsx
            EV_ICE_Training_NONSTRICT.xlsx

Training-data workbooks
-----------------------
Two separate ML-ready Excel files are also generated:
    EV_ICE_Training_STRICT.xlsx
    EV_ICE_Training_NONSTRICT.xlsx

Each row represents one EV-ICE pair and contains:
    - pair metadata and route information
    - EV and ICE overall distance
    - EV and ICE total TAT
    - EV and ICE average moving speed
    - EV and ICE time for every phase
    - EV and ICE distance for every phase
    - EV and ICE average speed for every phase

Phase average speed is calculated as:
    phase distance / phase time * 60

for phases with positive phase time. For stationary phases, the
distance/time result is naturally zero when the phase distance is zero.

Important
---------
The code uses the Stage-5 TAT results as the source of truth for:
    Total TAT
    Total Distance
    Average Moving Speed
    phase-wise time
    phase-wise distance

This keeps the master summary consistent with tat_computation.py.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
import math
import re
import shutil
import warnings

SCRIPT_VERSION = f"{datetime.now():%Y-%m-%d}_INTERCITY_ONLY_SUMMARY_V1"

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter


# ============================================================================
# CONFIGURATION
# ============================================================================

RESULTS_ROOT = Path(
    r"D:\TAT_MS_INTERCITY\Data\Results"
)

PAIRED_ROOT = (
    RESULTS_ROOT
    / "Paired"
)

SYSTEM_ROOT = (
    RESULTS_ROOT
    / "System"
)

PAIR_INDEX_PATH = (
    SYSTEM_ROOT
    / "pair_index.csv"
)

OUTPUT_ROOT = (
    RESULTS_ROOT
    / "Master_Summary"
)

OUTPUT_WORKBOOK = (
    OUTPUT_ROOT
    / "TAT_All_Data_Summary.xlsx"
)

CHART_ROOT = (
    OUTPUT_ROOT
    / "Charts"
)

# ---------------------------------------------------------------------------
# ML TRAINING DATA OUTPUT
# ---------------------------------------------------------------------------
# Two separate workbooks are created so STRICT and NONSTRICT pairs can be
# modelled independently without mixing the two pairing populations.
TRAINING_ROOT = (
    OUTPUT_ROOT
    / "Training_Data"
)

STRICT_TRAINING_WORKBOOK = (
    TRAINING_ROOT
    / "EV_ICE_Training_STRICT.xlsx"
)

NONSTRICT_TRAINING_WORKBOOK = (
    TRAINING_ROOT
    / "EV_ICE_Training_NONSTRICT.xlsx"
)

# IMPORTANT: this script belongs only to the INTERCITY pipeline.
# It must read pair workbooks from the current intercity Results folder
# and must never fall back to the old TAT_MS project folder.
INTERCITY_RESULTS_ROOT = RESULTS_ROOT

EV_SOURCE_SHEET = "01_EV_Cropped_Data"
ICE_SOURCE_SHEET = "02_ICE_Cropped_Data"
PAIR_INFO_SHEET = "03_Pair_Crop_Info"
TAT_RESULT_SHEET = "08_TAT_Phase_Comparison"

TIMESTAMP = "Timestamp"

PHASES = [
    "CHARGING / REFUELING",
    "STOP",
    "IDLE",
    "INCH",
    "ACC",
    "DEC",
    "CRUISE",
]

MOVING_PHASES = {
    "INCH",
    "ACC",
    "DEC",
    "CRUISE",
}

# Set to a pair id such as "PAIR_124" for a quick test.
# Keep None to process all pairs.
PAIR_NAME_FILTER: str | None = None

# Different-day comparison rule: EV and ICE must start within this many
# minutes of the same clock time.  A circular clock difference is used so
# 23:30 and 00:30 are treated as one hour apart, not 23 hours apart.
DIFFERENT_DAY_TIME_TOLERANCE_MINUTES = 120.0

# ============================================================================
# KNOWN ROUTES
# ============================================================================

# Route direction is ignored for route grouping.
# For example:
#   CMBT -> Navalur
#   Navalur -> CMBT
# belong to the same route group.

ROUTE_DEFINITIONS = [
    {
        "label": "T Nagar ↔ Thiruporur",
        "category": "CITY",
        "a": [
            "t nagar",
            "tnagar",
            "thyagaraya nagar",
        ],
        "b": [
            "thiruporur",
            "thirupporur",
            "tiruporur",
        ],
    },
    {
        "label": "CMBT/Koyambedu ↔ Navalur",
        "category": "CITY",
        "a": [
            "cmbt",
            "cbmt",
            "koyambedu",
        ],
        "b": [
            "navalur",
        ],
    },
    {
        "label": "KCBT ↔ Sholinganallur",
        "category": "CITY",
        "a": [
            "kcbt",
        ],
        "b": [
            "sholinganallur",
            "shollinganallur",
        ],
    },
    {
        "label": "Hebbal ↔ Silk Board",
        "category": "CITY",
        "a": [
            "hebbal",
            "hebal",
        ],
        "b": [
            "silk board",
            "silkboard",
        ],
    },
    {
        "label": "Bellandur ↔ KIA",
        "category": "CITY",
        "a": [
            "bellandur",
        ],
        "b": [
            "kia",
            "kempegowda international airport",
            "kempegowda airport",
        ],
    },
    {
        "label": "Kadugodi ↔ Silk Board",
        "category": "CITY",
        "a": [
            "kadugodi",
            "kadugodi bus stand",
        ],
        "b": [
            "silk board",
            "silkboard",
        ],
    },
    {
        "label": "Whitefield ↔ Silk Board",
        "category": "CITY",
        "a": [
            "whitefield",
        ],
        "b": [
            "silk board",
            "silkboard",
        ],
    },
    {
        "label": "Chennai ↔ Pondicherry",
        "category": "INTERCITY",
        "a": [
            "chennai",
        ],
        "b": [
            "pondicherry",
            "pondi",
            "pondy",
            "puducherry",
        ],
    },
    {
        "label": "Maduravoyal ↔ Silk Board",
        "category": "INTERCITY",
        "a": [
            "maduravoyal",
            "madhuravoyal",
            "maduravayol",
            "maduravoyil",
        ],
        "b": [
            "silk board",
            "silkboard",
        ],
    },
    {
        "label": "AVR Roundana (Salem) ↔ Silk Board",
        "category": "INTERCITY",
        "a": [
            "avr roundana",
            "avr roundana salem",
            "salem",
        ],
        "b": [
            "silk board",
            "silkboard",
        ],
    },
]


# ============================================================================
# STYLE
# ============================================================================

TITLE_BLUE = "1F4E78"
HEADER_BLUE = "4472C4"
LIGHT_BLUE = "D9E2F3"
ALT_BLUE = "E9EFF9"
LIGHT_GREEN = "E2F0D9"
LIGHT_GREY = "F2F2F2"
WHITE = "FFFFFF"
RED = "C00000"

THIN_GREY = Side(
    style="thin",
    color="B7B7B7",
)

TABLE_BORDER = Border(
    left=THIN_GREY,
    right=THIN_GREY,
    top=THIN_GREY,
    bottom=THIN_GREY,
)

EV_COLOR = "#2E86DE"
ICE_COLOR = "#E67E22"

DPI = 180

warnings.filterwarnings(
    "ignore",
    category=UserWarning,
)

plt.rcParams.update(
    {
        "font.size": 9,
        "axes.titlesize": 12,
        "axes.labelsize": 9,
        "xtick.labelsize": 8,
        "ytick.labelsize": 8,
    }
)


# ============================================================================
# DATA CLASS
# ============================================================================

@dataclass
class PairRecord:
    pair_id: str
    pair_number: int
    pair_class: str
    workbook_path: Path

    route_direction: str
    route_group: str
    category: str

    ev_date: str
    ice_date: str

    ev_start_time: str
    ev_end_time: str
    ice_start_time: str
    ice_end_time: str

    ev_distance_km: float
    ice_distance_km: float

    ev_tat_min: float
    ice_tat_min: float

    ev_moving_speed_kmh: float
    ice_moving_speed_kmh: float

    phase_time_ev: dict[str, float]
    phase_time_ice: dict[str, float]

    phase_distance_ev: dict[str, float]
    phase_distance_ice: dict[str, float]


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def clean(
    value: Any,
) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(
            value
        ):
            return ""
    except (
        TypeError,
        ValueError,
    ):
        pass

    return str(
        value
    ).strip()


def normalise(
    value: Any,
) -> str:
    value = clean(
        value
    ).casefold()

    value = re.sub(
        r"[_\-/()]+",
        " ",
        value,
    )

    value = re.sub(
        r"[^a-z0-9]+",
        " ",
        value,
    )

    return re.sub(
        r"\s+",
        " ",
        value,
    ).strip()


def parse_num(
    value: Any,
    default: float = np.nan,
) -> float:
    if value is None:
        return default

    try:
        if pd.isna(
            value
        ):
            return default
    except (
        TypeError,
        ValueError,
    ):
        pass

    try:
        return float(
            value
        )
    except (
        TypeError,
        ValueError,
    ):
        pass

    match = re.search(
        r"[-+]?"
        r"(?:\d+(?:\.\d*)?|\.\d+)"
        r"(?:[eE][-+]?\d+)?",
        clean(
            value
        ).replace(
            ",",
            "",
        ).replace(
            "−",
            "-",
        ),
    )

    if match is None:
        return default

    try:
        return float(
            match.group(0)
        )
    except ValueError:
        return default


def parse_dt_series(
    series: pd.Series,
) -> pd.Series:
    try:
        return pd.to_datetime(
            series,
            errors="coerce",
            dayfirst=True,
            format="mixed",
        )
    except TypeError:
        return pd.to_datetime(
            series,
            errors="coerce",
            dayfirst=True,
        )


def pair_number(
    pair_id: str,
) -> int:
    match = re.search(
        r"(\d+)",
        pair_id,
    )

    return (
        int(
            match.group(1)
        )
        if match
        else 0
    )


def safe_sheet_name(
    value: str,
    used: set[str],
) -> str:
    value = re.sub(
        r"[\[\]:*?/\\]",
        "_",
        value,
    )

    value = value.strip() or "Route"

    value = value[
        :31
    ]

    candidate = value
    counter = 2

    while candidate in used:
        suffix = (
            f"_{counter}"
        )

        candidate = (
            value[
                : 31
                - len(
                    suffix
                )
            ]
            + suffix
        )

        counter += 1

    used.add(
        candidate
    )

    return candidate


def read_csv_flexible(
    path: Path,
) -> pd.DataFrame:
    last_error = None

    for encoding in (
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ):
        try:
            return pd.read_csv(
                path,
                encoding=encoding,
            )
        except UnicodeDecodeError as error:
            last_error = error

    raise ValueError(
        f"Could not decode {path}: "
        f"{last_error}"
    )


# ============================================================================
# ROUTE / CATEGORY DETECTION
# ============================================================================

def contains_alias(
    text: str,
    aliases: list[str],
) -> bool:
    text_normal = (
        " "
        + normalise(
            text
        )
        + " "
    )

    for alias in aliases:
        alias_normal = normalise(
            alias
        )

        if (
            f" {alias_normal} "
            in text_normal
            or alias_normal
            in normalise(
                text
            )
        ):
            return True

    return False


def route_group_and_category(
    route_direction: str,
    all_evidence: str,
) -> tuple[
    str,
    str,
]:
    evidence = (
        f"{route_direction} "
        f"{all_evidence}"
    )

    for definition in ROUTE_DEFINITIONS:
        has_a = contains_alias(
            evidence,
            definition[
                "a"
            ],
        )

        has_b = contains_alias(
            evidence,
            definition[
                "b"
            ],
        )

        if (
            has_a
            and has_b
        ):
            return (
                definition[
                    "label"
                ],
                definition[
                    "category"
                ],
            )

    normalized = normalise(
        all_evidence
    )

    # Explicit filename/metadata category takes priority for unknown routes.
    if re.search(
        r"(?:^| )intercity(?: |$)",
        normalized,
    ):
        category = (
            "INTERCITY"
        )

    elif re.search(
        r"(?:^| )city(?: |$)",
        normalized,
    ):
        category = (
            "CITY"
        )

    else:
        category = (
            "UNCLASSIFIED"
        )

    route_label = (
        route_direction
        or "Unknown Route"
    )

    return (
        route_label,
        category,
    )



# ============================================================================
# ROUTE-DIRECTION HELPERS
# ============================================================================

def alias_position(
    text: str,
    aliases: list[str],
) -> int | None:
    """Return the earliest position of any alias in normalized text."""
    normalized_text = normalise(text)
    best_position = None

    for alias in aliases:
        normalized_alias = normalise(alias)
        if not normalized_alias:
            continue

        position = normalized_text.find(normalized_alias)
        if position >= 0 and (
            best_position is None
            or position < best_position
        ):
            best_position = position

    return best_position


def canonical_intercity_direction(
    route_direction: str,
    route_group: str,
) -> str:
    """
    Convert a known INTERCITY route into a stable directional label.

    Examples:
        Chennai to Pondicherry -> Chennai → Pondicherry
        Pondicherry to Chennai -> Pondicherry → Chennai

    Direction is determined from the route-direction text, not from pair
    class or vehicle type. If the direction cannot be resolved, the original
    route-direction text is retained.
    """
    for definition in ROUTE_DEFINITIONS:
        if (
            definition["category"] != "INTERCITY"
            or definition["label"] != route_group
        ):
            continue

        a_position = alias_position(
            route_direction,
            definition["a"],
        )
        b_position = alias_position(
            route_direction,
            definition["b"],
        )

        if (
            a_position is not None
            and b_position is not None
        ):
            if a_position < b_position:
                return (
                    f"{definition['label'].split(' ↔ ')[0]} "
                    f"→ "
                    f"{definition['label'].split(' ↔ ')[1]}"
                )

            if b_position < a_position:
                return (
                    f"{definition['label'].split(' ↔ ')[1]} "
                    f"→ "
                    f"{definition['label'].split(' ↔ ')[0]}"
                )

    # Fallback for an unknown/new INTERCITY route:
    # preserve the route direction supplied by the pair metadata.
    return (
        route_direction.strip()
        if clean(route_direction)
        else route_group
    )


def direction_slug(
    direction: str,
) -> str:
    """Create a compact Excel-safe slug for a directional route sheet."""
    value = direction.replace(
        "→",
        "to",
    )

    value = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        value,
    ).strip("_")

    return value[:70] or "Unknown_Direction"


# ============================================================================
# PAIR INDEX / WORKBOOK RESOLUTION
# ============================================================================

def resolve_workbook(
    row: pd.Series,
) -> Path:
    configured = clean(
        row.get(
            "pair_workbook_path"
        )
    )

    if configured:
        candidate = Path(
            configured
        )

        if (
            candidate.exists()
            and candidate.is_file()
        ):
            return candidate

    workbook_name = clean(
        row.get(
            "pair_workbook_name"
        )
    )

    output_folder = clean(
        row.get(
            "output_date_folder"
        )
    )

    if (
        workbook_name
        and output_folder
    ):
        candidate = (
            PAIRED_ROOT
            / output_folder
            / workbook_name
        )

        if candidate.exists():
            return candidate

    pair_id = clean(
        row.get(
            "pair_id"
        )
    )

    if pair_id:
        matches = [
            path
            for path in PAIRED_ROOT.rglob(
                f"{pair_id}_*.xlsx"
            )
            if path.is_file()
            and not path.name.startswith(
                (
                    "~$",
                    ".",
                )
            )
        ]

        if matches:
            return max(
                matches,
                key=lambda path: (
                    path.stat().st_mtime
                ),
            )

    raise FileNotFoundError(
        "Pair workbook not found for "
        f"{pair_id or workbook_name}"
    )


def discover_pairs() -> list[
    tuple[
        str,
        dict[str, Any],
        Path,
    ]
]:
    pairs = []

    if PAIR_INDEX_PATH.exists():
        index = read_csv_flexible(
            PAIR_INDEX_PATH
        )

        for _, row in index.iterrows():
            pair_id = clean(
                row.get(
                    "pair_id"
                )
            )

            if not pair_id:
                continue

            if (
                PAIR_NAME_FILTER
                and PAIR_NAME_FILTER.casefold()
                not in pair_id.casefold()
            ):
                continue

            try:
                workbook_path = resolve_workbook(
                    row
                )
            except FileNotFoundError:
                continue

            pairs.append(
                (
                    pair_id,
                    row.to_dict(),
                    workbook_path,
                )
            )

    else:
        for workbook_path in sorted(
            PAIRED_ROOT.rglob(
                "PAIR_*.xlsx"
            )
        ):
            if workbook_path.name.startswith(
                (
                    "~$",
                    ".",
                )
            ):
                continue

            pair_id_match = re.match(
                r"(PAIR_\d+)",
                workbook_path.stem,
                flags=re.I,
            )

            pair_id = (
                pair_id_match.group(1).upper()
                if pair_id_match
                else workbook_path.stem
            )

            if (
                PAIR_NAME_FILTER
                and PAIR_NAME_FILTER.casefold()
                not in pair_id.casefold()
            ):
                continue

            pairs.append(
                (
                    pair_id,
                    {},
                    workbook_path,
                )
            )

    pairs.sort(
        key=lambda item: (
            pair_number(
                item[0]
            ),
            item[0],
        )
    )

    return pairs


# ============================================================================
# METADATA / STAGE-5 EXTRACTION
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
    except (
        ValueError,
        KeyError,
    ):
        return {}

    if (
        frame.empty
        or len(
            frame.columns
        )
        < 2
    ):
        return {}

    parameter = frame.columns[
        0
    ]

    value = frame.columns[
        1
    ]

    return {
        clean(
            key
        ): val
        for key, val in zip(
            frame[
                parameter
            ],
            frame[
                value
            ],
        )
        if clean(
            key
        )
    }


def clean_extracted_table(
    frame: pd.DataFrame,
    key_column: str,
) -> pd.DataFrame:
    frame = frame.copy()

    valid_positions = []
    seen = set()

    for index, raw_name in enumerate(
        frame.columns
    ):
        name = clean(
            raw_name
        )

        if (
            not name
            or name.lower().startswith(
                "unnamed:"
            )
        ):
            continue

        key = name.casefold()

        if key in seen:
            continue

        seen.add(
            key
        )

        valid_positions.append(
            index
        )

    frame = frame.iloc[
        :,
        valid_positions,
    ].copy()

    frame.columns = [
        clean(
            column
        )
        for column in frame.columns
    ]

    if key_column not in frame.columns:
        raise ValueError(
            f"Expected column "
            f"{key_column!r} not found. "
            f"Available: "
            f"{list(frame.columns)}"
        )

    frame = frame.dropna(
        how="all"
    )

    frame = frame.loc[
        frame[
            key_column
        ]
        .astype(
            "string"
        )
        .fillna(
            ""
        )
        .str.strip()
        .ne(
            ""
        )
    ].reset_index(
        drop=True
    )

    return frame


def extract_stage5_tables(
    workbook_path: Path,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
]:
    with pd.ExcelFile(
        workbook_path,
        engine="openpyxl",
    ) as excel:
        required = {
            EV_SOURCE_SHEET,
            ICE_SOURCE_SHEET,
            TAT_RESULT_SHEET,
        }

        missing = (
            required
            - set(
                excel.sheet_names
            )
        )

        if missing:
            raise ValueError(
                "Missing sheets: "
                + ", ".join(
                    sorted(
                        missing
                    )
                )
            )

        tat = pd.read_excel(
            excel,
            sheet_name=TAT_RESULT_SHEET,
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

    for row_index in range(
        len(
            tat
        )
    ):
        values = [
            clean(
                value
            )
            for value in tat.iloc[
                row_index
            ]
        ]

        if (
            "Metric"
            in values
            and "EV"
            in values
            and "ICE"
            in values
        ):
            overall_header = (
                row_index
            )

        if (
            "Phase"
            in values
            and "EV Time (min)"
            in values
        ):
            phase_header = (
                row_index
            )

    if (
        overall_header
        is None
        or phase_header
        is None
    ):
        raise ValueError(
            "Could not identify Stage-5 "
            "overall/phase tables"
        )

    overall = tat.iloc[
        overall_header
        + 1:
        phase_header
        - 2
    ].copy()

    overall.columns = tat.iloc[
        overall_header
    ]

    overall = clean_extracted_table(
        overall,
        "Metric",
    )

    phase = tat.iloc[
        phase_header
        + 1:
    ].copy()

    phase.columns = tat.iloc[
        phase_header
    ]

    phase = clean_extracted_table(
        phase,
        "Phase",
    )

    return (
        overall,
        phase,
        ev,
        ice,
    )


def metric(
    overall: pd.DataFrame,
    metric_name: str,
    vehicle: str,
) -> float:
    values = overall.loc[
        overall[
            "Metric"
        ]
        .astype(
            str
        )
        .str.strip()
        .eq(
            metric_name
        ),
        vehicle,
    ]

    if values.empty:
        return np.nan

    return parse_num(
        values.iloc[
            0
        ]
    )


def phase_value(
    phase: pd.DataFrame,
    phase_name: str,
    column: str,
) -> float:
    aliases = {
        "CHARGING / REFUELING": {
            "CHARGING / REFUELING",
            "CHARGING",
            "REFUELING",
            "REFUELLING",
        },
        "INCH": {
            "INCH",
            "INCHING",
        },
        "ACC": {
            "ACC",
            "ACCELERATION",
            "ACCELERATING",
        },
        "DEC": {
            "DEC",
            "DECELERATION",
            "DECELERATING",
        },
        "CRUISE": {
            "CRUISE",
            "CRUISING",
        },
        "STOP": {
            "STOP",
        },
        "IDLE": {
            "IDLE",
        },
    }

    accepted = {
        normalise(
            value
        )
        for value in aliases.get(
            phase_name,
            {
                phase_name
            },
        )
    }

    mask = phase[
        "Phase"
    ].map(
        normalise
    ).isin(
        accepted
    )

    matches = phase.loc[
        mask,
        column,
    ]

    if matches.empty:
        return 0.0

    values = [
        parse_num(
            value,
            default=0.0,
        )
        for value in matches
    ]

    return float(
        np.nansum(
            values
        )
    )


def timestamp_summary(
    source: pd.DataFrame,
) -> tuple[
    str,
    str,
    str,
]:
    if TIMESTAMP not in source.columns:
        return (
            "",
            "",
            "",
        )

    timestamps = parse_dt_series(
        source[
            TIMESTAMP
        ]
    ).dropna()

    if timestamps.empty:
        return (
            "",
            "",
            "",
        )

    start = pd.Timestamp(
        timestamps.iloc[
            0
        ]
    )

    end = pd.Timestamp(
        timestamps.iloc[
            -1
        ]
    )

    date_text = start.strftime(
        "%d-%m-%Y"
    )

    return (
        date_text,
        start.strftime(
            "%H:%M:%S"
        ),
        end.strftime(
            "%H:%M:%S"
        ),
    )


def is_same_calendar_day(
    record: PairRecord,
) -> bool:
    """Return True only when both valid trip dates are identical."""
    return bool(
        record.ev_date
        and record.ice_date
        and record.ev_date
        == record.ice_date
    )


def common_trip_date(
    record: PairRecord,
) -> str:
    """Return the shared date for a same-day pair, otherwise blank."""
    return (
        record.ev_date
        if is_same_calendar_day(
            record
        )
        else ""
    )


def is_different_calendar_day(
    record: PairRecord,
) -> bool:
    """Return True only when both valid trip dates are different."""
    return bool(
        record.ev_date
        and record.ice_date
        and record.ev_date
        != record.ice_date
    )


def clock_time_minutes(
    time_text: str,
) -> float:
    """Convert HH:MM[:SS] text to minutes after midnight."""
    value = clean(
        time_text
    )

    if not value:
        return np.nan

    for time_format in (
        "%H:%M:%S",
        "%H:%M",
    ):
        try:
            parsed = datetime.strptime(
                value,
                time_format,
            )

            return (
                parsed.hour
                * 60.0
                + parsed.minute
                + parsed.second
                / 60.0
            )

        except ValueError:
            continue

    return np.nan


def start_time_difference_minutes(
    record: PairRecord,
) -> float:
    """Return the shortest circular clock-time difference between starts."""
    ev_minutes = clock_time_minutes(
        record.ev_start_time
    )

    ice_minutes = clock_time_minutes(
        record.ice_start_time
    )

    if not (
        math.isfinite(
            ev_minutes
        )
        and math.isfinite(
            ice_minutes
        )
    ):
        return np.nan

    direct_difference = abs(
        ev_minutes
        - ice_minutes
    )

    return min(
        direct_difference,
        24.0
        * 60.0
        - direct_difference,
    )


def is_different_day_similar_time(
    record: PairRecord,
) -> bool:
    """Check different dates and start-time difference within two hours."""
    if not is_different_calendar_day(
        record
    ):
        return False

    difference = start_time_difference_minutes(
        record
    )

    return bool(
        math.isfinite(
            difference
        )
        and difference
        <= DIFFERENT_DAY_TIME_TOLERANCE_MINUTES
    )


def best_route_direction(
    metadata: dict[str, Any],
    pair_info: dict[str, Any],
    workbook_path: Path,
) -> str:
    keys = [
        "route_direction",
        "route",
        "route_name",
        "route_group",
    ]

    for source in (
        pair_info,
        metadata,
    ):
        for key in keys:
            value = clean(
                source.get(
                    key
                )
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

    name = re.sub(
        r"_",
        " ",
        name,
    )

    name = re.sub(
        r"\bto\b",
        "to",
        name,
        flags=re.I,
    )

    return name


def pair_class(
    metadata: dict[str, Any],
    workbook_path: Path,
) -> str:
    value = clean(
        metadata.get(
            "pair_class"
        )
    )

    if value:
        return value.upper()

    name = workbook_path.name.upper()

    if "NONSTRICT" in name:
        return "NONSTRICT"

    if "STRICT" in name:
        return "STRICT"

    return ""


def build_pair_record(
    pair_id: str,
    metadata: dict[str, Any],
    workbook_path: Path,
) -> PairRecord:
    (
        overall,
        phase,
        ev,
        ice,
    ) = extract_stage5_tables(
        workbook_path
    )

    pair_info = read_pair_info(
        workbook_path
    )

    route_direction = best_route_direction(
        metadata,
        pair_info,
        workbook_path,
    )

    all_evidence = " ".join(
        [
            workbook_path.name,
            str(
                workbook_path
            ),
            route_direction,
            *[
                clean(
                    value
                )
                for value in metadata.values()
            ],
            *[
                clean(
                    value
                )
                for value in pair_info.values()
            ],
        ]
    )

    (
        route_group,
        _detected_category,
    ) = route_group_and_category(
        route_direction,
        all_evidence,
    )

    # This Summary script belongs to the INTERCITY-only pipeline.
    # Category filtering therefore cannot accidentally reintroduce CITY data.
    category = "INTERCITY"

    (
        ev_date,
        ev_start,
        ev_end,
    ) = timestamp_summary(
        ev
    )

    (
        ice_date,
        ice_start,
        ice_end,
    ) = timestamp_summary(
        ice
    )

    phase_time_ev = {}
    phase_time_ice = {}

    phase_distance_ev = {}
    phase_distance_ice = {}

    for phase_name in PHASES:
        phase_time_ev[
            phase_name
        ] = phase_value(
            phase,
            phase_name,
            "EV Time (min)",
        )

        phase_time_ice[
            phase_name
        ] = phase_value(
            phase,
            phase_name,
            "ICE Time (min)",
        )

        phase_distance_ev[
            phase_name
        ] = phase_value(
            phase,
            phase_name,
            "EV Distance (km)",
        )

        phase_distance_ice[
            phase_name
        ] = phase_value(
            phase,
            phase_name,
            "ICE Distance (km)",
        )

    return PairRecord(
        pair_id=pair_id,
        pair_number=pair_number(
            pair_id
        ),
        pair_class=pair_class(
            metadata,
            workbook_path,
        ),
        workbook_path=workbook_path,
        route_direction=route_direction,
        route_group=route_group,
        category=category,
        ev_date=ev_date,
        ice_date=ice_date,
        ev_start_time=ev_start,
        ev_end_time=ev_end,
        ice_start_time=ice_start,
        ice_end_time=ice_end,
        ev_distance_km=metric(
            overall,
            "Total Distance (km)",
            "EV",
        ),
        ice_distance_km=metric(
            overall,
            "Total Distance (km)",
            "ICE",
        ),
        ev_tat_min=metric(
            overall,
            "Total TAT (min)",
            "EV",
        ),
        ice_tat_min=metric(
            overall,
            "Total TAT (min)",
            "ICE",
        ),
        ev_moving_speed_kmh=metric(
            overall,
            "Average Moving Speed (km/h)",
            "EV",
        ),
        ice_moving_speed_kmh=metric(
            overall,
            "Average Moving Speed (km/h)",
            "ICE",
        ),
        phase_time_ev=phase_time_ev,
        phase_time_ice=phase_time_ice,
        phase_distance_ev=phase_distance_ev,
        phase_distance_ice=phase_distance_ice,
    )


# ============================================================================
# DATAFRAME BUILDERS
# ============================================================================

def trip_summary_frame(
    records: list[PairRecord],
    include_same_day_columns: bool = False,
    include_different_day_time_columns: bool = False,
) -> pd.DataFrame:
    rows = []

    for record in records:
        row = {
            "Pair": record.pair_number,
            "Pair ID": record.pair_id,
            "Pair Class": record.pair_class,
            "Category": record.category,
            "Route Group": record.route_group,
            "Trips": record.route_direction,
            "EV Date": record.ev_date,
            "ICE Date": record.ice_date,
        }

        # These two audit columns are intentionally restricted to the three
        # new same-day sheets.  The original ALL, CITY, INTERCITY and route
        # sheets therefore keep their previous table structure unchanged.
        if include_same_day_columns:
            row.update(
                {
                    "Same-Day Pair": (
                        "YES"
                        if is_same_calendar_day(
                            record
                        )
                        else "NO"
                    ),
                    "Common Date": common_trip_date(
                        record
                    ),
                }
            )

        # These audit columns appear only in the three different-day,
        # similar-start-time sheets.  All previously existing sheets retain
        # their exact earlier table structure.
        if include_different_day_time_columns:
            time_difference = (
                start_time_difference_minutes(
                    record
                )
            )

            row.update(
                {
                    "Different-Day Pair": (
                        "YES"
                        if is_different_calendar_day(
                            record
                        )
                        else "NO"
                    ),
                    "Start-Time Difference (min)": time_difference,
                    "Allowed Difference (min)": (
                        DIFFERENT_DAY_TIME_TOLERANCE_MINUTES
                    ),
                    "Within 2 Hours": (
                        "YES"
                        if math.isfinite(
                            time_difference
                        )
                        and time_difference
                        <= DIFFERENT_DAY_TIME_TOLERANCE_MINUTES
                        else "NO"
                    ),
                }
            )

        row.update(
            {
                "EV Distance (km)": record.ev_distance_km,
                "ICE Distance (km)": record.ice_distance_km,
                "EV TAT (min)": record.ev_tat_min,
                "ICE TAT (min)": record.ice_tat_min,
                "EV Start Time": record.ev_start_time,
                "EV End Time": record.ev_end_time,
                "ICE Start Time": record.ice_start_time,
                "ICE End Time": record.ice_end_time,
                "EV Moving Avg Speed (km/h)": record.ev_moving_speed_kmh,
                "ICE Moving Avg Speed (km/h)": record.ice_moving_speed_kmh,
            }
        )

        rows.append(
            row
        )

    return pd.DataFrame(
        rows
    )


def phase_time_frame(
    records: list[PairRecord],
) -> pd.DataFrame:
    rows = []

    for record in records:
        for vehicle in (
            "EV",
            "ICE",
        ):
            source = (
                record.phase_time_ev
                if vehicle
                == "EV"
                else record.phase_time_ice
            )

            row = {
                "Pair": record.pair_number,
                "Pair ID": record.pair_id,
                "Vehicle": vehicle,
                "Category": record.category,
                "Route Group": record.route_group,
                "Trips": record.route_direction,
            }

            for phase_name in PHASES:
                display = (
                    "INCHING"
                    if phase_name
                    == "INCH"
                    else phase_name
                )

                row[
                    display
                ] = source.get(
                    phase_name,
                    0.0,
                )

            rows.append(
                row
            )

    return pd.DataFrame(
        rows
    )


def phase_distance_frame(
    records: list[PairRecord],
) -> pd.DataFrame:
    rows = []

    for record in records:
        for vehicle in (
            "EV",
            "ICE",
        ):
            source = (
                record.phase_distance_ev
                if vehicle
                == "EV"
                else record.phase_distance_ice
            )

            row = {
                "Pair": record.pair_number,
                "Pair ID": record.pair_id,
                "Vehicle": vehicle,
                "Category": record.category,
                "Route Group": record.route_group,
                "Trips": record.route_direction,
            }

            for phase_name in PHASES:
                display = (
                    "INCHING"
                    if phase_name
                    == "INCH"
                    else phase_name
                )

                row[
                    display
                ] = source.get(
                    phase_name,
                    0.0,
                )

            rows.append(
                row
            )

    return pd.DataFrame(
        rows
    )


def route_index_frame(
    records: list[PairRecord],
) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()

    frame = pd.DataFrame(
        {
            "Route": [
                record.route_group
                for record in records
            ],
            "Category": [
                record.category
                for record in records
            ],
            "Pair ID": [
                record.pair_id
                for record in records
            ],
        }
    )

    summary = (
        frame.groupby(
            [
                "Category",
                "Route",
            ],
            dropna=False,
        )
        .agg(
            Pair_Count=(
                "Pair ID",
                "nunique",
            )
        )
        .reset_index()
        .sort_values(
            [
                "Category",
                "Route",
            ]
        )
    )

    return summary


# ============================================================================
# CHARTS
# ============================================================================

def label_bars(
    ax,
    bars,
    decimals: int = 2,
) -> None:
    labels = []

    for bar in bars:
        height = (
            bar.get_height()
        )

        if not math.isfinite(
            height
        ):
            labels.append(
                ""
            )
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


def save_chart(
    fig,
    path: Path,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    fig.savefig(
        path,
        dpi=DPI,
        bbox_inches="tight",
        facecolor="white",
    )

    plt.close(
        fig
    )


def scope_slug(
    scope_name: str,
) -> str:
    return re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        scope_name,
    ).strip(
        "_"
    )[:80] or "scope"


def make_overall_tat_chart(
    records: list[PairRecord],
    scope_name: str,
    path: Path,
    force_pairwise: bool = False,
) -> None:
    if not records:
        return

    # Route-specific scopes get pair-wise TAT.
    # Larger scopes get route-wise average TAT for readability.
    routes = sorted(
        {
            record.route_group
            for record in records
        }
    )

    if (
        force_pairwise
        or (
            len(
                routes
            )
            == 1
            and len(
                records
            )
            <= 30
        )
    ):
        labels = [
            str(
                record.pair_number
            )
            for record in records
        ]

        ev = np.array(
            [
                record.ev_tat_min
                for record in records
            ],
            dtype=float,
        )

        ice = np.array(
            [
                record.ice_tat_min
                for record in records
            ],
            dtype=float,
        )

        title = (
            f"{scope_name}: "
            "Pair-wise EV vs ICE TAT"
        )

        xlabel = "Pair"

    else:
        rows = []

        for route in routes:
            group = [
                record
                for record in records
                if record.route_group
                == route
            ]

            rows.append(
                {
                    "Route": route,
                    "EV": np.nanmean(
                        [
                            record.ev_tat_min
                            for record in group
                        ]
                    ),
                    "ICE": np.nanmean(
                        [
                            record.ice_tat_min
                            for record in group
                        ]
                    ),
                }
            )

        labels = [
            row[
                "Route"
            ]
            for row in rows
        ]

        ev = np.array(
            [
                row[
                    "EV"
                ]
                for row in rows
            ],
            dtype=float,
        )

        ice = np.array(
            [
                row[
                    "ICE"
                ]
                for row in rows
            ],
            dtype=float,
        )

        title = (
            f"{scope_name}: "
            "Route-wise Average EV vs ICE TAT"
        )

        xlabel = "Route"

    x = np.arange(
        len(
            labels
        )
    )

    width = 0.36

    fig, ax = plt.subplots(
        figsize=(
            max(
                9,
                len(
                    labels
                )
                * 1.2,
            ),
            5.5,
        )
    )

    ev_bars = ax.bar(
        x
        - width
        / 2,
        ev,
        width,
        label="EV",
        color=EV_COLOR,
    )

    ice_bars = ax.bar(
        x
        + width
        / 2,
        ice,
        width,
        label="ICE",
        color=ICE_COLOR,
    )

    label_bars(
        ax,
        ev_bars,
        1,
    )

    label_bars(
        ax,
        ice_bars,
        1,
    )

    ax.set_ylabel(
        "TAT (min)"
    )

    ax.set_xlabel(
        xlabel
    )

    ax.set_title(
        title,
        fontweight="bold",
    )

    ax.set_xticks(
        x
    )

    ax.set_xticklabels(
        labels,
        rotation=(
            40
            if len(
                labels
            )
            > 5
            else 0
        ),
        ha="right"
        if len(
            labels
        )
        > 5
        else "center",
    )

    ax.legend()

    ax.grid(
        axis="y",
        alpha=0.25,
    )

    fig.tight_layout()

    save_chart(
        fig,
        path,
    )


def make_moving_speed_chart(
    records: list[PairRecord],
    scope_name: str,
    path: Path,
) -> None:
    if not records:
        return

    ev = np.nanmean(
        [
            record.ev_moving_speed_kmh
            for record in records
        ]
    )

    ice = np.nanmean(
        [
            record.ice_moving_speed_kmh
            for record in records
        ]
    )

    fig, ax = plt.subplots(
        figsize=(
            7,
            5,
        )
    )

    bars = ax.bar(
        [
            "EV",
            "ICE",
        ],
        [
            ev,
            ice,
        ],
        color=[
            EV_COLOR,
            ICE_COLOR,
        ],
        width=0.55,
    )

    label_bars(
        ax,
        bars,
        2,
    )

    ax.set_ylabel(
        "Average Moving Speed (km/h)"
    )

    ax.set_title(
        f"{scope_name}: "
        "Average Moving Speed",
        fontweight="bold",
    )

    ax.grid(
        axis="y",
        alpha=0.25,
    )

    fig.tight_layout()

    save_chart(
        fig,
        path,
    )


def average_phase_values(
    records: list[PairRecord],
    attribute_ev: str,
    attribute_ice: str,
) -> tuple[
    list[str],
    np.ndarray,
    np.ndarray,
]:
    labels = [
        (
            "CHG/REFUEL"
            if phase
            == "CHARGING / REFUELING"
            else (
                "INCH"
                if phase
                == "INCH"
                else phase
            )
        )
        for phase in PHASES
    ]

    ev_values = []

    ice_values = []

    for phase_name in PHASES:
        ev_values.append(
            np.nanmean(
                [
                    getattr(
                        record,
                        attribute_ev,
                    ).get(
                        phase_name,
                        0.0,
                    )
                    for record in records
                ]
            )
        )

        ice_values.append(
            np.nanmean(
                [
                    getattr(
                        record,
                        attribute_ice,
                    ).get(
                        phase_name,
                        0.0,
                    )
                    for record in records
                ]
            )
        )

    return (
        labels,
        np.array(
            ev_values,
            dtype=float,
        ),
        np.array(
            ice_values,
            dtype=float,
        ),
    )


def make_phase_time_chart(
    records: list[PairRecord],
    scope_name: str,
    path: Path,
) -> None:
    if not records:
        return

    (
        labels,
        ev,
        ice,
    ) = average_phase_values(
        records,
        "phase_time_ev",
        "phase_time_ice",
    )

    x = np.arange(
        len(
            labels
        )
    )

    width = 0.36

    fig, ax = plt.subplots(
        figsize=(
            10,
            5.5,
        )
    )

    ev_bars = ax.bar(
        x
        - width
        / 2,
        ev,
        width,
        label="EV",
        color=EV_COLOR,
    )

    ice_bars = ax.bar(
        x
        + width
        / 2,
        ice,
        width,
        label="ICE",
        color=ICE_COLOR,
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

    ax.set_xticks(
        x
    )

    ax.set_xticklabels(
        labels,
    )

    ax.set_ylabel(
        "Average Phase Time (min)"
    )

    ax.set_title(
        f"{scope_name}: "
        "Average Phase Time",
        fontweight="bold",
    )

    ax.legend()

    ax.grid(
        axis="y",
        alpha=0.25,
    )

    fig.tight_layout()

    save_chart(
        fig,
        path,
    )


def make_phase_distance_chart(
    records: list[PairRecord],
    scope_name: str,
    path: Path,
) -> None:
    if not records:
        return

    (
        labels,
        ev,
        ice,
    ) = average_phase_values(
        records,
        "phase_distance_ev",
        "phase_distance_ice",
    )

    x = np.arange(
        len(
            labels
        )
    )

    width = 0.36

    fig, ax = plt.subplots(
        figsize=(
            10,
            5.5,
        )
    )

    ev_bars = ax.bar(
        x
        - width
        / 2,
        ev,
        width,
        label="EV",
        color=EV_COLOR,
    )

    ice_bars = ax.bar(
        x
        + width
        / 2,
        ice,
        width,
        label="ICE",
        color=ICE_COLOR,
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

    ax.set_xticks(
        x
    )

    ax.set_xticklabels(
        labels,
    )

    ax.set_ylabel(
        "Average Phase Distance (km)"
    )

    ax.set_title(
        f"{scope_name}: "
        "Average Phase Distance",
        fontweight="bold",
    )

    ax.legend()

    ax.grid(
        axis="y",
        alpha=0.25,
    )

    fig.tight_layout()

    save_chart(
        fig,
        path,
    )


def create_scope_charts(
    records: list[PairRecord],
    scope_name: str,
    force_pairwise_tat: bool = False,
) -> list[
    tuple[
        str,
        Path,
    ]
]:
    slug_name = scope_slug(
        scope_name
    )

    scope_folder = (
        CHART_ROOT
        / slug_name
    )

    scope_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    charts = [
        (
            "TAT",
            scope_folder
            / "01_tat.png",
        ),
        (
            "Moving Speed",
            scope_folder
            / "02_moving_speed.png",
        ),
        (
            "Phase Time",
            scope_folder
            / "03_phase_time.png",
        ),
        (
            "Phase Distance",
            scope_folder
            / "04_phase_distance.png",
        ),
    ]

    make_overall_tat_chart(
        records,
        scope_name,
        charts[
            0
        ][
            1
        ],
        force_pairwise=force_pairwise_tat,
    )

    make_moving_speed_chart(
        records,
        scope_name,
        charts[
            1
        ][
            1
        ],
    )

    make_phase_time_chart(
        records,
        scope_name,
        charts[
            2
        ][
            1
        ],
    )

    make_phase_distance_chart(
        records,
        scope_name,
        charts[
            3
        ][
            1
        ],
    )

    return charts


# ============================================================================
# EXCEL TABLE WRITING
# ============================================================================

def style_title(
    worksheet,
    row: int,
    title: str,
    end_column: int,
) -> None:
    worksheet.merge_cells(
        start_row=row,
        start_column=1,
        end_row=row,
        end_column=end_column,
    )

    cell = worksheet.cell(
        row=row,
        column=1,
        value=title,
    )

    cell.font = Font(
        size=18,
        bold=True,
        color="000000",
    )

    cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
    )

    worksheet.row_dimensions[
        row
    ].height = 28

    # Red line below title, similar to the example.
    for column in range(
        1,
        end_column
        + 1,
    ):
        worksheet.cell(
            row=row
            + 1,
            column=column,
        ).fill = PatternFill(
            fill_type="solid",
            fgColor=RED,
        )

    worksheet.row_dimensions[
        row
        + 1
    ].height = 4


def write_dataframe_table(
    worksheet,
    start_row: int,
    frame: pd.DataFrame,
    title: str,
    decimals: int = 3,
) -> int:
    if frame.empty:
        style_title(
            worksheet,
            start_row,
            title,
            8,
        )

        worksheet.cell(
            row=start_row
            + 3,
            column=1,
            value="No data",
        )

        return start_row + 6

    end_column = max(
        8,
        len(
            frame.columns
        ),
    )

    style_title(
        worksheet,
        start_row,
        title,
        end_column,
    )

    header_row = (
        start_row
        + 3
    )

    for column_index, column in enumerate(
        frame.columns,
        start=1,
    ):
        cell = worksheet.cell(
            row=header_row,
            column=column_index,
            value=str(
                column
            ),
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_BLUE,
        )

        cell.font = Font(
            color=WHITE,
            bold=True,
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        cell.border = TABLE_BORDER

    for row_offset, row_values in enumerate(
        frame.itertuples(
            index=False,
            name=None,
        ),
        start=1,
    ):
        excel_row = (
            header_row
            + row_offset
        )

        fill = (
            LIGHT_BLUE
            if row_offset
            % 2
            == 1
            else ALT_BLUE
        )

        for column_index, value in enumerate(
            row_values,
            start=1,
        ):
            cell = worksheet.cell(
                row=excel_row,
                column=column_index,
            )

            if isinstance(
                value,
                (
                    np.floating,
                    float,
                ),
            ):
                if math.isnan(
                    float(
                        value
                    )
                ):
                    cell.value = None
                else:
                    cell.value = float(
                        value
                    )

                    cell.number_format = (
                        "0."
                        + "0"
                        * decimals
                    )

            elif isinstance(
                value,
                (
                    np.integer,
                    int,
                ),
            ):
                cell.value = int(
                    value
                )

            elif isinstance(
                value,
                pd.Timestamp,
            ):
                cell.value = value.to_pydatetime()

            else:
                try:
                    if pd.isna(
                        value
                    ):
                        value = None
                except (
                    TypeError,
                    ValueError,
                ):
                    pass

                cell.value = value

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor=fill,
            )

            cell.border = TABLE_BORDER

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

    worksheet.auto_filter.ref = (
        f"A{header_row}:"
        f"{get_column_letter(len(frame.columns))}"
        f"{header_row + len(frame)}"
    )

    return (
        header_row
        + len(
            frame
        )
        + 3
    )


def set_reasonable_widths(
    worksheet,
    max_columns: int,
) -> None:
    for column_index in range(
        1,
        max_columns
        + 1,
    ):
        max_length = 0

        for row in range(
            1,
            min(
                worksheet.max_row,
                250,
            )
            + 1,
        ):
            value = worksheet.cell(
                row=row,
                column=column_index,
            ).value

            max_length = max(
                max_length,
                len(
                    clean(
                        value
                    )
                ),
            )

        width = min(
            28,
            max(
                10,
                max_length
                + 2,
            ),
        )

        worksheet.column_dimensions[
            get_column_letter(
                column_index
            )
        ].width = width


def add_chart_images(
    worksheet,
    charts: list[
        tuple[
            str,
            Path,
        ]
    ],
) -> None:
    anchors = [
        "U2",
        "U25",
        "U48",
        "U71",
    ]

    for (
        (
            _,
            path,
        ),
        anchor,
    ) in zip(
        charts,
        anchors,
    ):
        if not path.exists():
            continue

        image = ExcelImage(
            str(
                path
            )
        )

        image.width = 760
        image.height = 390

        worksheet.add_image(
            image,
            anchor,
        )


def write_scope_sheet(
    workbook: Workbook,
    used_sheet_names: set[str],
    sheet_name: str,
    scope_title: str,
    records: list[PairRecord],
    force_pairwise_tat: bool = False,
    include_same_day_columns: bool = False,
    include_different_day_time_columns: bool = False,
) -> str:
    actual_sheet_name = safe_sheet_name(
        sheet_name,
        used_sheet_names,
    )

    worksheet = workbook.create_sheet(
        actual_sheet_name
    )

    trip = trip_summary_frame(
        records,
        include_same_day_columns=(
            include_same_day_columns
        ),
        include_different_day_time_columns=(
            include_different_day_time_columns
        ),
    )

    time_frame = phase_time_frame(
        records
    )

    distance_frame = phase_distance_frame(
        records
    )

    current_row = 1

    current_row = write_dataframe_table(
        worksheet,
        current_row,
        trip,
        (
            f"Table 1: "
            f"{scope_title} - EV vs ICE Trip Summary"
        ),
        decimals=3,
    )

    current_row = write_dataframe_table(
        worksheet,
        current_row,
        time_frame,
        (
            f"Table 2: "
            f"{scope_title} - Phase Time Comparison (min)"
        ),
        decimals=3,
    )

    write_dataframe_table(
        worksheet,
        current_row,
        distance_frame,
        (
            f"Table 3: "
            f"{scope_title} - Phase Distance Comparison (km)"
        ),
        decimals=3,
    )

    worksheet.freeze_panes = "A4"

    set_reasonable_widths(
        worksheet,
        max(
            18,
            trip.shape[
                1
            ]
            if not trip.empty
            else 18,
        ),
    )

    charts = create_scope_charts(
        records,
        scope_title,
        force_pairwise_tat=force_pairwise_tat,
    )

    add_chart_images(
        worksheet,
        charts,
    )

    return actual_sheet_name


# ============================================================================
# ML TRAINING DATA
# ============================================================================

def phase_average_speed(
    phase_distance_km: float,
    phase_time_min: float,
) -> float:
    """
    Calculate phase average speed in km/h.

    Formula:
        speed = distance / time * 60

    If phase time is zero/non-positive, return 0 because there is no valid
    time interval over which to calculate an average speed.
    """
    distance = parse_num(
        phase_distance_km,
        default=0.0,
    )
    time_min = parse_num(
        phase_time_min,
        default=0.0,
    )

    if (
        not math.isfinite(distance)
        or not math.isfinite(time_min)
        or time_min <= 0
    ):
        return 0.0

    return float(
        distance
        / time_min
        * 60.0
    )


def build_ml_training_frame(
    records: list[PairRecord],
    required_pair_class: str,
) -> pd.DataFrame:
    """
    Build one wide ML-training table.

    One row = one EV-ICE pair.

    The table intentionally contains both EV and ICE values.  The later ML
    stage can then choose ICE columns as input features and EV columns as
    targets without having to reconstruct phase information from the
    Summary workbook.

    Only records whose Pair Class exactly matches required_pair_class are
    included.
    """
    required_pair_class = (
        clean(
            required_pair_class
        ).upper()
    )

    selected_records = [
        record
        for record in records
        if clean(
            record.pair_class
        ).upper()
        == required_pair_class
    ]

    selected_records.sort(
        key=lambda record: (
            record.pair_number,
            record.pair_id,
        )
    )

    rows = []

    for record in selected_records:
        row: dict[str, Any] = {
            # ---------------------------------------------------------------
            # Pair / route metadata
            # ---------------------------------------------------------------
            "Pair": record.pair_number,
            "Pair ID": record.pair_id,
            "Pair Class": record.pair_class,
            "Route Group": record.route_group,
            "Route Direction": canonical_intercity_direction(
                record.route_direction,
                record.route_group,
            ),
            "EV Date": record.ev_date,
            "ICE Date": record.ice_date,
            "EV Start Time": record.ev_start_time,
            "EV End Time": record.ev_end_time,
            "ICE Start Time": record.ice_start_time,
            "ICE End Time": record.ice_end_time,
            "Same Calendar Day": (
                "YES"
                if is_same_calendar_day(record)
                else "NO"
            ),
            "Start-Time Difference (min)": (
                start_time_difference_minutes(record)
            ),

            # ---------------------------------------------------------------
            # Overall EV parameters
            # ---------------------------------------------------------------
            "EV Total Distance (km)": record.ev_distance_km,
            "EV Total TAT (min)": record.ev_tat_min,
            "EV Average Moving Speed (km/h)": (
                record.ev_moving_speed_kmh
            ),

            # ---------------------------------------------------------------
            # Overall ICE parameters
            # ---------------------------------------------------------------
            "ICE Total Distance (km)": record.ice_distance_km,
            "ICE Total TAT (min)": record.ice_tat_min,
            "ICE Average Moving Speed (km/h)": (
                record.ice_moving_speed_kmh
            ),
        }

        # ---------------------------------------------------------------
        # Phase-level EV + ICE parameters
        #
        # For every phase:
        #   1. time
        #   2. distance
        #   3. average speed
        #
        # The phase average speed is derived consistently from the same
        # Stage-5 phase time and phase distance used elsewhere in this
        # Summary script.
        # ---------------------------------------------------------------
        for phase_name in PHASES:
            display_phase = (
                "INCHING"
                if phase_name == "INCH"
                else (
                    "CHARGING_REFUELING"
                    if phase_name
                    == "CHARGING / REFUELING"
                    else phase_name
                )
            )

            ev_time = record.phase_time_ev.get(
                phase_name,
                0.0,
            )
            ice_time = record.phase_time_ice.get(
                phase_name,
                0.0,
            )

            ev_distance = record.phase_distance_ev.get(
                phase_name,
                0.0,
            )
            ice_distance = record.phase_distance_ice.get(
                phase_name,
                0.0,
            )

            row[
                f"EV {display_phase} Time (min)"
            ] = ev_time

            row[
                f"ICE {display_phase} Time (min)"
            ] = ice_time

            row[
                f"EV {display_phase} Distance (km)"
            ] = ev_distance

            row[
                f"ICE {display_phase} Distance (km)"
            ] = ice_distance

            row[
                f"EV {display_phase} Average Speed (km/h)"
            ] = phase_average_speed(
                ev_distance,
                ev_time,
            )

            row[
                f"ICE {display_phase} Average Speed (km/h)"
            ] = phase_average_speed(
                ice_distance,
                ice_time,
            )

        rows.append(
            row
        )

    return pd.DataFrame(
        rows
    )


def write_training_workbook(
    frame: pd.DataFrame,
    output_path: Path,
    title: str,
) -> None:
    """
    Write one ML-training workbook.

    The workbook contains:
        01_Training_Data
        02_Feature_Groups

    01_Training_Data is the actual model-ready wide table.
    02_Feature_Groups documents the column groups so the later ML code can
    select ICE inputs and EV targets reliably.
    """
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    training_workbook = Workbook()

    data_sheet = training_workbook.active
    data_sheet.title = "01_Training_Data"

    # Reuse the Summary's existing table formatting.
    next_row = write_dataframe_table(
        data_sheet,
        1,
        frame,
        title,
        decimals=4,
    )

    data_sheet.freeze_panes = "A4"

    set_reasonable_widths(
        data_sheet,
        max(
            1,
            frame.shape[1],
        ),
    )

    # -----------------------------------------------------------------------
    # Feature-group documentation
    # -----------------------------------------------------------------------
    group_sheet = training_workbook.create_sheet(
        "02_Feature_Groups"
    )

    group_sheet["A1"] = (
        "ML Training Data Feature Groups"
    )

    group_sheet["A1"].font = Font(
        size=18,
        bold=True,
    )

    group_sheet["A3"] = "Feature Group"
    group_sheet["B3"] = "Meaning"
    group_sheet["C3"] = "Use in Later ML Stage"

    for cell in group_sheet[3]:
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_BLUE,
        )
        cell.font = Font(
            color=WHITE,
            bold=True,
        )
        cell.border = TABLE_BORDER
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    groups = [
        (
            "Pair Metadata",
            "Pair ID, route, direction, dates, times and pairing class",
            "Identification / optional categorical features",
        ),
        (
            "EV Overall",
            "EV total distance, total TAT and average moving speed",
            "Prediction targets / evaluation reference",
        ),
        (
            "ICE Overall",
            "ICE total distance, total TAT and average moving speed",
            "Primary candidate input features",
        ),
        (
            "EV Phase Time",
            "EV time spent in each phase",
            "Prediction targets",
        ),
        (
            "ICE Phase Time",
            "ICE time spent in each phase",
            "Primary candidate input features",
        ),
        (
            "EV Phase Distance",
            "EV distance covered in each phase",
            "Prediction targets",
        ),
        (
            "ICE Phase Distance",
            "ICE distance covered in each phase",
            "Primary candidate input features",
        ),
        (
            "EV Phase Average Speed",
            "EV average speed for each phase",
            "Prediction targets",
        ),
        (
            "ICE Phase Average Speed",
            "ICE average speed for each phase",
            "Primary candidate input features",
        ),
    ]

    for row_index, values in enumerate(
        groups,
        start=4,
    ):
        for column_index, value in enumerate(
            values,
            start=1,
        ):
            cell = group_sheet.cell(
                row=row_index,
                column=column_index,
                value=value,
            )

            cell.border = TABLE_BORDER

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor=(
                    LIGHT_BLUE
                    if row_index % 2 == 0
                    else ALT_BLUE
                ),
            )

            cell.alignment = Alignment(
                horizontal="left",
                vertical="top",
                wrap_text=True,
            )

    group_sheet.column_dimensions[
        "A"
    ].width = 28

    group_sheet.column_dimensions[
        "B"
    ].width = 60

    group_sheet.column_dimensions[
        "C"
    ].width = 42

    group_sheet["A15"] = "Total Rows"
    group_sheet["B15"] = len(frame)

    group_sheet["A16"] = "Total Features"
    group_sheet["B16"] = frame.shape[1]

    group_sheet["A18"] = (
        "Phase Average Speed Formula"
    )
    group_sheet["B18"] = (
        "Phase Distance (km) / Phase Time (min) × 60"
    )

    group_sheet["A19"] = (
        "Source"
    )
    group_sheet["B19"] = (
        "Stage-5 08_TAT_Phase_Comparison "
        "read by master_summary_intercity.py"
    )

    group_sheet["A20"] = (
        "Pair Selection"
    )
    group_sheet["B20"] = (
        "Only the pair class represented by this workbook is included."
    )

    # Bold metadata labels.
    for row_index in (
        15,
        16,
        18,
        19,
        20,
    ):
        group_sheet.cell(
            row=row_index,
            column=1,
        ).font = Font(
            bold=True,
        )

    # Make sure the workbook opens on the actual training table.
    training_workbook.active = 0

    try:
        training_workbook.save(
            output_path
        )
    except PermissionError:
        raise PermissionError(
            f"Cannot write {output_path}. "
            "Close the training workbook in Excel/WPS "
            "and run the Summary again."
        )


def create_training_workbooks(
    records: list[PairRecord],
) -> dict[str, int]:
    """
    Create independent STRICT and NONSTRICT training workbooks.

    The Summary workbook remains unchanged; these are additional outputs.
    """
    TRAINING_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    strict_frame = build_ml_training_frame(
        records,
        "STRICT",
    )

    nonstrict_frame = build_ml_training_frame(
        records,
        "NONSTRICT",
    )

    write_training_workbook(
        strict_frame,
        STRICT_TRAINING_WORKBOOK,
        "STRICT EV–ICE ML Training Data",
    )

    write_training_workbook(
        nonstrict_frame,
        NONSTRICT_TRAINING_WORKBOOK,
        "NONSTRICT EV–ICE ML Training Data",
    )

    print()
    print(
        "ML TRAINING DATA CREATED"
    )
    print(
        f"    STRICT rows: "
        f"{len(strict_frame)}"
    )
    print(
        f"    File: "
        f"{STRICT_TRAINING_WORKBOOK}"
    )
    print(
        f"    NONSTRICT rows: "
        f"{len(nonstrict_frame)}"
    )
    print(
        f"    File: "
        f"{NONSTRICT_TRAINING_WORKBOOK}"
    )
    print(
        f"    Features per row: "
        f"{strict_frame.shape[1] if not strict_frame.empty else nonstrict_frame.shape[1]}"
    )

    return {
        "STRICT": len(
            strict_frame
        ),
        "NONSTRICT": len(
            nonstrict_frame
        ),
    }


# ============================================================================
# INDEX / OVERVIEW SHEET
# ============================================================================

def write_index_sheet(
    workbook: Workbook,
    records: list[PairRecord],
    route_sheet_map: dict[str, str],
    direction_class_sheet_map: dict[tuple[str, str], str],
) -> None:
    worksheet = workbook.active
    worksheet.title = "00_Index"

    worksheet[
        "A1"
    ] = (
        "EV–ICE TAT Master Summary"
    )

    worksheet[
        "A1"
    ].font = Font(
        size=20,
        bold=True,
    )

    worksheet[
        "A3"
    ] = "Generated"

    worksheet[
        "B3"
    ] = datetime.now().strftime(
        "%d-%m-%Y %H:%M:%S"
    )

    worksheet[
        "A4"
    ] = "Total Pairs"

    worksheet[
        "B4"
    ] = len(
        records
    )

    intercity_count = sum(
        1
        for record in records
        if record.category == "INTERCITY"
    )

    worksheet[
        "A5"
    ] = "INTERCITY Pairs"

    worksheet[
        "B5"
    ] = intercity_count

    worksheet[
        "A6"
    ] = "Same-Day Pairs"

    worksheet[
        "B6"
    ] = sum(
        1
        for record in records
        if is_same_calendar_day(record)
    )

    worksheet[
        "A7"
    ] = "Different-Day / Similar-Time Pairs"

    worksheet[
        "B7"
    ] = sum(
        1
        for record in records
        if is_different_day_similar_time(record)
    )

    summary = route_index_frame(
        records
    )

    # Keep the original index-sheet layout unchanged.
    start_row = 10

    style_title(
        worksheet,
        start_row,
        "Route-wise Pair Index",
        5,
    )

    header_row = (
        start_row
        + 3
    )

    headers = [
        "Category",
        "Route",
        "Pair Count",
        "Route Sheet",
    ]

    for column, header in enumerate(
        headers,
        start=1,
    ):
        cell = worksheet.cell(
            row=header_row,
            column=column,
            value=header,
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_BLUE,
        )

        cell.font = Font(
            color=WHITE,
            bold=True,
        )

        cell.border = TABLE_BORDER

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for offset, row in enumerate(
        summary.itertuples(
            index=False,
        ),
        start=1,
    ):
        excel_row = (
            header_row
            + offset
        )

        values = [
            row.Category,
            row.Route,
            int(
                row.Pair_Count
            ),
            route_sheet_map.get(
                row.Route,
                "",
            ),
        ]

        for column, value in enumerate(
            values,
            start=1,
        ):
            cell = worksheet.cell(
                row=excel_row,
                column=column,
                value=value,
            )

            cell.border = TABLE_BORDER

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor=(
                    LIGHT_BLUE
                    if offset
                    % 2
                    else ALT_BLUE
                ),
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        sheet_name = route_sheet_map.get(
            row.Route
        )

        if sheet_name:
            cell = worksheet.cell(
                row=excel_row,
                column=4,
            )

            cell.hyperlink = (
                f"#'{sheet_name}'!A1"
            )

            cell.style = "Hyperlink"

    # ---------------------------------------------------------------
    # Route-direction + pair-class sheet index
    # ---------------------------------------------------------------
    direction_rows = []

    class_order = {
        "STRICT": 0,
        "NONSTRICT": 1,
    }

    for (
        direction,
        pair_class_value,
    ) in sorted(
        direction_class_sheet_map,
        key=lambda item: (
            item[0],
            class_order[item[1]],
        ),
    ):
        direction_records = [
            record
            for record in records
            if (
                record.category == "INTERCITY"
                and canonical_intercity_direction(
                    record.route_direction,
                    record.route_group,
                )
                == direction
                and clean(
                    record.pair_class
                ).upper()
                == pair_class_value
            )
        ]

        direction_rows.append(
            {
                "Direction": direction,
                "Pair Class": pair_class_value,
                "Pair Count": len(direction_records),
                "Sheet": direction_class_sheet_map[
                    (
                        direction,
                        pair_class_value,
                    )
                ],
            }
        )

    direction_start_row = (
        header_row
        + max(len(summary), 1)
        + 4
    )

    style_title(
        worksheet,
        direction_start_row,
        "Route-Direction + Pair-Class Sheet Index",
        5,
    )

    direction_header_row = (
        direction_start_row
        + 3
    )

    direction_headers = [
        "Direction",
        "Pair Class",
        "Pair Count",
        "Sheet",
    ]

    for column, header in enumerate(
        direction_headers,
        start=1,
    ):
        cell = worksheet.cell(
            row=direction_header_row,
            column=column,
            value=header,
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=HEADER_BLUE,
        )

        cell.font = Font(
            color=WHITE,
            bold=True,
        )

        cell.border = TABLE_BORDER

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for offset, row in enumerate(
        direction_rows,
        start=1,
    ):
        excel_row = (
            direction_header_row
            + offset
        )

        values = [
            row["Direction"],
            row["Pair Class"],
            row["Pair Count"],
            row["Sheet"],
        ]

        for column, value in enumerate(
            values,
            start=1,
        ):
            cell = worksheet.cell(
                row=excel_row,
                column=column,
                value=value,
            )

            cell.border = TABLE_BORDER

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor=(
                    LIGHT_BLUE
                    if offset % 2
                    else ALT_BLUE
                ),
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        cell = worksheet.cell(
            row=excel_row,
            column=4,
        )

        cell.hyperlink = (
            f"#'{row['Sheet']}'!A1"
        )

        cell.style = "Hyperlink"

    worksheet[
        "F3"
    ] = (
        "Workbook Structure"
    )

    worksheet[
        "F3"
    ].font = Font(
        bold=True,
        size=13,
    )

    structure = [
        (
            "01_All_Pairs",
            f"All {len(records)} INTERCITY pair-level tables + charts",
        ),
        (
            "02_INTERCITY",
            "All INTERCITY pair-level tables + charts",
        ),
        (
            "03_Same_Day_All",
            "All same-calendar-day INTERCITY pair tables + charts",
        ),
        (
            "04_Same_Day_INTERCITY",
            "Same-calendar-day INTERCITY pair tables + charts",
        ),
        (
            "05_Diff_Day_Time_All",
            "Different-day INTERCITY pairs within 2 hours",
        ),
        (
            "06_Diff_Day_Time_INTERCITY",
            "Different-day INTERCITY pairs within 2 hours",
        ),
        (
            "Route-group sheets",
            "One complete set of three tables + charts for every INTERCITY route group; both directions are retained together",
        ),
        (
            "Route-direction sheets",
            "One complete set of three tables + charts for every INTERCITY direction, such as Chennai → Pondicherry",
        ),
    ]

    for index, (
        name,
        description,
    ) in enumerate(
        structure,
        start=4,
    ):
        worksheet.cell(
            row=index,
            column=6,
            value=name,
        ).font = Font(
            bold=True,
        )

        worksheet.cell(
            row=index,
            column=7,
            value=description,
        )

    worksheet.column_dimensions[
        "A"
    ].width = 20

    worksheet.column_dimensions[
        "B"
    ].width = 34

    worksheet.column_dimensions[
        "C"
    ].width = 14

    worksheet.column_dimensions[
        "D"
    ].width = 26

    worksheet.column_dimensions[
        "F"
    ].width = 22

    worksheet.column_dimensions[
        "G"
    ].width = 50


# ============================================================================
# MAIN
# ============================================================================

def main() -> None:
    print(
        "="
        * 88
    )

    print(
        "MASTER SUMMARY: "
        "ALL PAIRS + INTERCITY + SAME-DAY + "
        "DIFFERENT-DAY/SIMILAR-TIME + ROUTE-WISE"
    )

    print(
        "="
        * 88
    )

    if not PAIRED_ROOT.exists():
        raise FileNotFoundError(
            f"Paired results folder not found: "
            f"{PAIRED_ROOT}"
        )

    OUTPUT_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    if CHART_ROOT.exists():
        shutil.rmtree(
            CHART_ROOT
        )

    CHART_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    pairs = discover_pairs()

    print(
        f"Pair workbooks found: "
        f"{len(pairs)}"
    )

    # Safety check: every discovered workbook must come from this
    # intercity project's Results\Paired folder. This prevents stale
    # workbooks from an older project path from entering the summary.
    invalid_sources = [
        path
        for _, _, path in pairs
        if PAIRED_ROOT not in path.parents
    ]
    if invalid_sources:
        raise RuntimeError(
            "Summary discovered pair workbooks outside the current "
            f"INTERCITY Paired folder: {invalid_sources}"
        )

    records: list[
        PairRecord
    ] = []

    errors = []

    for number, (
        pair_id,
        metadata,
        workbook_path,
    ) in enumerate(
        pairs,
        start=1,
    ):
        print(
            f"[{number}/{len(pairs)}] "
            f"{pair_id} | "
            f"{workbook_path.name}"
        )

        try:
            record = build_pair_record(
                pair_id,
                metadata,
                workbook_path,
            )

            records.append(
                record
            )

            print(
                "    COMPLETED | "
                f"{record.category} | "
                f"{record.route_group}"
            )

        except Exception as error:
            errors.append(
                {
                    "Pair ID": pair_id,
                    "Workbook": str(
                        workbook_path
                    ),
                    "Error": str(
                        error
                    ),
                }
            )

            print(
                f"    ERROR | "
                f"{error}"
            )

    if not records:
        raise RuntimeError(
            "No pair workbook could be summarized."
        )

    records.sort(
        key=lambda record: (
            record.pair_number,
            record.pair_id,
        )
    )

    # -----------------------------------------------------------------------
    # Additional ML-ready datasets
    # -----------------------------------------------------------------------
    # These are generated from the same PairRecord objects used by the
    # Summary workbook, so the training data and Summary remain consistent.
    training_counts = create_training_workbooks(
        records
    )

    print()

    print(
        f"Successfully extracted: "
        f"{len(records)}"
    )

    print(
        f"Errors: "
        f"{len(errors)}"
    )

    workbook = Workbook()

    used_sheet_names = {
        "00_Index"
    }

    # ---------------------------------------------------------------
    # All pairs
    # ---------------------------------------------------------------
    write_scope_sheet(
        workbook,
        used_sheet_names,
        "01_All_Pairs",
        "ALL PAIRS",
        records,
    )

    # ---------------------------------------------------------------
    # INTERCITY-only scope
    # ---------------------------------------------------------------
    # Every record reaching this stage belongs to the intercity pipeline.
    # Keep the explicit INTERCITY sheet because it is part of the Summary
    # specification, even though it contains the same records as All Pairs.
    intercity_records = [
        record
        for record in records
        if record.category == "INTERCITY"
    ]

    write_scope_sheet(
        workbook,
        used_sheet_names,
        "02_INTERCITY",
        "INTERCITY",
        intercity_records,
    )

    # ---------------------------------------------------------------
    # Same-calendar-day INTERCITY pairs
    # ---------------------------------------------------------------
    same_day_records = [
        record
        for record in intercity_records
        if is_same_calendar_day(record)
    ]

    write_scope_sheet(
        workbook,
        used_sheet_names,
        "03_Same_Day_All",
        "SAME-DAY EV-ICE PAIRS",
        same_day_records,
        force_pairwise_tat=True,
        include_same_day_columns=True,
    )

    # Kept as an explicit INTERCITY sheet for a stable, self-documenting
    # workbook structure. No CITY version is generated.
    write_scope_sheet(
        workbook,
        used_sheet_names,
        "04_Same_Day_INTERCITY",
        "SAME-DAY INTERCITY",
        same_day_records,
        force_pairwise_tat=True,
        include_same_day_columns=True,
    )

    # ---------------------------------------------------------------
    # Different-calendar-day pairs with start times within two hours
    # ---------------------------------------------------------------
    different_day_time_records = [
        record
        for record in intercity_records
        if is_different_day_similar_time(record)
    ]

    write_scope_sheet(
        workbook,
        used_sheet_names,
        "05_Diff_Day_Time_All",
        "DIFFERENT-DAY EV-ICE PAIRS WITHIN 2 HOURS",
        different_day_time_records,
        force_pairwise_tat=True,
        include_different_day_time_columns=True,
    )

    write_scope_sheet(
        workbook,
        used_sheet_names,
        "06_Diff_Day_Time_INTERCITY",
        "DIFFERENT-DAY INTERCITY WITHIN 2 HOURS",
        different_day_time_records,
        force_pairwise_tat=True,
        include_different_day_time_columns=True,
    )

    # ---------------------------------------------------------------
    # Route-wise grouped sheets
    # ---------------------------------------------------------------
    # These sheets preserve the existing behaviour: both directions of a
    # route group remain together (e.g. Chennai ↔ Pondicherry).
    routes = sorted(
        {
            record.route_group
            for record in records
        }
    )

    route_sheet_map = {}

    for route_number, route in enumerate(
        routes,
        start=1,
    ):
        route_records = [
            record
            for record in records
            if record.route_group == route
        ]

        base_sheet_name = (
            f"R{route_number:02d}_"
            + re.sub(
                r"[^A-Za-z0-9]+",
                "_",
                route,
            ).strip("_")
        )

        actual_name = write_scope_sheet(
            workbook,
            used_sheet_names,
            base_sheet_name,
            f"ROUTE GROUP: {route}",
            route_records,
        )

        route_sheet_map[route] = actual_name

    # ---------------------------------------------------------------
    # Route-direction + pair-class sheets
    # ---------------------------------------------------------------
    # Each actual travel direction is split into STRICT and NONSTRICT.
    #
    # Example:
    #   D01_Chennai_to_Pondicherry_STRICT
    #   D02_Chennai_to_Pondicherry_NONSTRICT
    #   D03_Pondicherry_to_Chennai_STRICT
    #   D04_Pondicherry_to_Chennai_NONSTRICT
    #
    # A sheet contains ONLY:
    #   1. the selected travel direction, and
    #   2. the selected pair class.
    #
    # Therefore a STRICT Chennai -> Pondicherry pair can never appear in
    # the NONSTRICT sheet or in the opposite-direction sheet.
    direction_class_groups: dict[
        tuple[str, str],
        list[PairRecord],
    ] = {}

    for record in records:
        if record.category != "INTERCITY":
            continue

        pair_class_value = clean(
            record.pair_class
        ).upper()

        if pair_class_value not in {
            "STRICT",
            "NONSTRICT",
        }:
            print(
                f"    WARNING | {record.pair_id}: "
                f"pair class {pair_class_value!r} is not STRICT/NONSTRICT; "
                "excluded from direction/class sheets."
            )
            continue

        direction = canonical_intercity_direction(
            record.route_direction,
            record.route_group,
        )

        direction_class_groups.setdefault(
            (
                direction,
                pair_class_value,
            ),
            [],
        ).append(record)

    direction_class_sheet_map: dict[
        tuple[str, str],
        str,
    ] = {}

    # Sort first by direction and then STRICT before NONSTRICT so the
    # workbook order is stable and easy to read.
    class_order = {
        "STRICT": 0,
        "NONSTRICT": 1,
    }

    sorted_direction_classes = sorted(
        direction_class_groups,
        key=lambda item: (
            item[0],
            class_order[item[1]],
        ),
    )

    for direction_number, (
        direction,
        pair_class_value,
    ) in enumerate(
        sorted_direction_classes,
        start=1,
    ):
        direction_records = sorted(
            direction_class_groups[
                (
                    direction,
                    pair_class_value,
                )
            ],
            key=lambda record: (
                record.pair_number,
                record.pair_id,
            ),
        )

        base_sheet_name = (
            f"D{direction_number:02d}_"
            f"{direction_slug(direction)}_"
            f"{pair_class_value}"
        )

        actual_name = write_scope_sheet(
            workbook,
            used_sheet_names,
            base_sheet_name,
            (
                f"ROUTE DIRECTION: {direction} | "
                f"{pair_class_value}"
            ),
            direction_records,
            force_pairwise_tat=True,
        )

        direction_class_sheet_map[
            (
                direction,
                pair_class_value,
            )
        ] = actual_name

    print(
        f"Route groups created: {len(routes)}"
    )

    print(
        "Route-direction / pair-class sheets created: "
        f"{len(direction_class_sheet_map)}"
    )

    # ---------------------------------------------------------------
    # Error log
    # ---------------------------------------------------------------
    if errors:
        error_sheet = workbook.create_sheet(
            "99_Errors"
        )

        error_frame = pd.DataFrame(
            errors
        )

        write_dataframe_table(
            error_sheet,
            1,
            error_frame,
            "Pair Workbooks Not Included",
            decimals=3,
        )

        set_reasonable_widths(
            error_sheet,
            error_frame.shape[
                1
            ],
        )

    # ---------------------------------------------------------------
    # Index last so hyperlinks know route sheet names.
    # ---------------------------------------------------------------
    write_index_sheet(
        workbook,
        records,
        route_sheet_map,
        direction_class_sheet_map,
    )

    # Put index first.
    index_sheet = workbook[
        "00_Index"
    ]

    workbook._sheets.remove(
        index_sheet
    )

    workbook._sheets.insert(
        0,
        index_sheet
    )

    # ---------------------------------------------------------------
    # Enforce final workbook sheet order.
    # ---------------------------------------------------------------
    # openpyxl creates sheets in call order, but route sheets and optional
    # error sheets are made explicit here so the final workbook is stable.
    fixed_order = [
        "00_Index",
        "01_All_Pairs",
        "02_INTERCITY",
        "03_Same_Day_All",
        "04_Same_Day_INTERCITY",
        "05_Diff_Day_Time_All",
        "06_Diff_Day_Time_INTERCITY",
    ]

    ordered_names = [
        name
        for name in fixed_order
        if name in workbook.sheetnames
    ]

    direction_names = sorted(
        [
            name
            for name in workbook.sheetnames
            if name.startswith("D")
        ]
    )

    ordered_names.extend(direction_names)

    route_names = sorted(
        [
            name
            for name in workbook.sheetnames
            if name.startswith("R")
        ]
    )

    ordered_names.extend(route_names)

    if "99_Errors" in workbook.sheetnames:
        ordered_names.append("99_Errors")

    workbook._sheets = [
        workbook[name]
        for name in ordered_names
        if name in workbook.sheetnames
    ]

    # ---------------------------------------------------------------
    # Save
    # ---------------------------------------------------------------
    try:
        workbook.save(
            OUTPUT_WORKBOOK
        )

    except PermissionError:
        raise PermissionError(
            f"Cannot write {OUTPUT_WORKBOOK}. "
            "Close the master workbook in Excel/WPS "
            "and run again."
        )

    print()

    print(
        "="
        * 88
    )

    print(
        "MASTER SUMMARY COMPLETED"
    )

    print(
        "="
        * 88
    )

    print(
        f"Workbook: "
        f"{OUTPUT_WORKBOOK}"
    )

    print(
        f"Charts: "
        f"{CHART_ROOT}"
    )

    print(
        f"Pairs included: "
        f"{len(records)}"
    )

    print(
        f"INTERCITY pairs: "
        f"{len(intercity_records)}"
    )

    print(
        f"Same-day INTERCITY pairs: "
        f"{len(same_day_records)}"
    )

    print(
        f"Different-day INTERCITY pairs within 2 hours: "
        f"{len(different_day_time_records)}"
    )

    print(
        f"Route groups: "
        f"{len(routes)}"
    )

    print(
        f"Route-direction / pair-class sheets: "
        f"{len(direction_class_sheet_map)}"
    )

    print(
        f"Errors: "
        f"{len(errors)}"
    )

    print(
        f"STRICT training rows: "
        f"{training_counts['STRICT']}"
    )

    print(
        f"NONSTRICT training rows: "
        f"{training_counts['NONSTRICT']}"
    )

    print(
        "="
        * 88
    )


if __name__ == "__main__":
    main()