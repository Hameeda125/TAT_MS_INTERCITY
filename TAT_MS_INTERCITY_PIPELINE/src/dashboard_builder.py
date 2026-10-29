"""
Standalone INTERCITY TAT dashboard builder.

Requires:
    Windows
    Microsoft Excel 2016 or later
    python -m pip install pywin32

Run:
    python src/dashboard_builder.py

Reads:
    Existing TAT_All_Data_Summary.xlsx -> 01_All_Pairs

Creates/rebuilds:
    Dashboard
    Trip_Input
    Model
    Calc (hidden)

Features:
    17 native Excel charts
    Live Corridor / Direction / Month / Counting basis selectors
    Editable model assumptions and calibration
    Existing dashboard manual inputs retained when keys match
    Backup before replacing the summary workbook

No reference/template workbook is required.
Existing summary sheets and Python pipeline scripts are retained.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import math
import os
import re
import shutil
import sys
import tempfile
import traceback

from datetime import date, datetime, time, timedelta
from pathlib import Path


# =====================================================================
# CONFIGURATION
# =====================================================================

PROJECT_ROOT = Path(
    r"D:\Hameeda_Projects\TAT_MS_INTERCITY"
    r"\TAT_MS_INTERCITY_PIPELINE"
)

SUMMARY_PATH = (
    PROJECT_ROOT.parent
    / "Data"
    / "Results"
    / "Master_Summary"
    / "TAT_All_Data_Summary.xlsx"
)

DASHBOARD_SHEETS = ("Dashboard", "Trip_Input", "Model", "Calc")
EXPECTED_CHARTS = 17
EXCEL_EPOCH = datetime(1899, 12, 30)

PHASES = [
    "CHARGING / REFUELING",
    "STOP",
    "IDLE",
    "INCHING",
    "ACC",
    "DEC",
    "CRUISE",
]

PHASE_LABELS = [
    "Charging / refuelling",
    "Stop",
    "Idle",
    "Inching",
    "Acceleration",
    "Deceleration",
    "Cruise",
]

# Approximate endpoint elevations from the reviewed reference.
# Add new routes here with appropriate endpoint elevations.
#
# Corridor, direction, origin, destination,
# origin elevation (m), destination elevation (m).
DEFAULT_ROUTES = [
    (
        "Chennai ↔ Pondicherry",
        "Chennai → Pondicherry",
        "Chennai", "Pondicherry", 10, 5,
    ),
    (
        "Chennai ↔ Pondicherry",
        "Pondicherry → Chennai",
        "Pondicherry", "Chennai", 5, 10,
    ),
    (
        "Maduravoyal ↔ Silk Board",
        "Maduravoyal → Silk Board",
        "Maduravoyal", "Silk Board", 20, 905,
    ),
    (
        "Maduravoyal ↔ Silk Board",
        "Silk Board → Maduravoyal",
        "Silk Board", "Maduravoyal", 905, 20,
    ),
    (
        "Maduravoyal ↔ Madiwala",
        "Maduravoyal → Madiwala",
        "Maduravoyal", "Madiwala", 20, 900,
    ),
    (
        "Maduravoyal ↔ Madiwala",
        "Madiwala → Maduravoyal",
        "Madiwala", "Maduravoyal", 900, 20,
    ),
]

# Rows 6:35 in Model.
# Parameter, unit, EV value/formula, Diesel value/formula.
PARAMETERS = [
    ("Kerb mass", "kg", 13780, 11780),
    ("Default passengers", "persons", 40, "=D7"),
    ("Passenger mass including luggage", "kg/person", 68, "=D8"),
    ("Default operating mass", "kg", "=D6+D7*D8", "=E6+D7*D8"),
    ("Rolling resistance coefficient", "-", 0.007, 0.007),
    ("Drag coefficient", "-", 0.65, 0.65),
    ("Frontal area", "m²", 8.2, 8.2),
    ("Air density", "kg/m³", 1.18, "=D13"),
    ("Gravitational acceleration", "m/s²", 9.81, "=D14"),
    (
        "Aerodynamic constant", "N/(m/s)²",
        "=0.5*D13*D11*D12", "=0.5*E13*E11*E12",
    ),
    ("Default mean acceleration", "m/s²", 0.25, 0.25),
    ("Motor/inverter or engine efficiency", "-", 0.9, 0.35),
    ("Driveline efficiency", "-", 0.95, 0.92),
    ("Battery discharge efficiency", "-", 0.97, None),
    ("Combined powertrain efficiency", "-", "=D17*D18*D19", "=E17*E18"),
    ("Regenerative recovery fraction", "-", 0.6, 0),
    ("Descent share offsetting road load", "-", 0.5, "=D22"),
    ("Base auxiliary power", "kW", 3, 3),
    ("HVAC load above setpoint", "kW/°C", 0.85, "=D24"),
    ("HVAC setpoint", "°C", 24, "=D25"),
    ("Default ambient temperature", "°C", 30, "=D26"),
    (
        "Auxiliary power at default temperature", "kW",
        "=D23+D24*MAX(0,D26-D25)",
        "=E23+E24*MAX(0,E26-E25)",
    ),
    ("Diesel idle fuel rate", "L/h", None, 2.5),
    ("Diesel lower heating value", "MJ/L", None, 35.9),
    ("Diesel lower heating value", "kWh/L", None, "=E29/3.6"),
    ("Charger efficiency", "-", 0.92, None),
    ("Electricity tariff", "INR/kWh", 9.5, None),
    ("Diesel price", "INR/L", None, 92),
    ("Grid emission factor", "kg/kWh", 0.71, None),
    ("Diesel emission factor", "kg/L", None, 2.68),
]

MODEL_NAMES = {
    "Kerb_EV": "D6",
    "Kerb_D": "E6",
    "Pax_Default": "D7",
    "Pax_Mass": "D8",
    "Crr_EV": "D10",
    "Crr_D": "E10",
    "G_acc": "D14",
    "Kaero_EV": "D15",
    "Kaero_D": "E15",
    "Acc_EV": "D16",
    "Acc_D": "E16",
    "EtaEng_D": "E17",
    "EtaPT_EV": "D20",
    "EtaPT_D": "E20",
    "Regen_EV": "D21",
    "Descent_Share": "D22",
    "AuxBase_EV": "D23",
    "AuxBase_D": "E23",
    "HVAC_k": "D24",
    "HVAC_Tset": "D25",
    "Temp_Default": "D26",
    "IdleFuel_D": "E28",
    "LHV_kWhL": "E30",
    "EtaCharger": "D31",
    "Tariff_EV": "D32",
    "Price_Diesel": "E33",
    "EF_Grid": "D34",
    "EF_Diesel": "E35",
    "Cal_Apply": "D38",
    "CF_EV": "D42",
    "CF_D": "E42",
}

OTHER_NAMES = {
    "Sel_Corridor", "Sel_Direction", "Sel_Month",
    "Sel_Basis", "Sel_Metric", "Sel_DirEff",
    "L_Corr", "L_Dir", "L_Month", "L_Basis",
    "L_Metric", "L_Bus", "L_RouteCorr", "L_RouteDir",
}

OWN_NAMES = set(MODEL_NAMES) | OTHER_NAMES

NAME_PATTERN = re.compile(
    r"\b("
    + "|".join(sorted(OWN_NAMES, key=len, reverse=True))
    + r")\b"
)

HEADERS = (
    "Trip ID|Pair ID|Date|Bus Type|Corridor|Direction|Start Time|"
    "t Charging (min)|t Stop (min)|t Idle (min)|t Inching (min)|"
    "t Acc (min)|t Dec (min)|t Cruise (min)|"
    "d Charging (km)|d Stop (km)|d Idle (km)|d Inching (km)|"
    "d Acc (km)|d Dec (km)|d Cruise (km)|"
    "Passengers|Ambient temp (C)|Accel measured (m/s2)|"
    "Measured EV energy (kWh)|Measured diesel (L)|"
    "Data check|Month|Unique trip|Flag route|Flag period|Flag all|"
    "Net elevation (m)|Mass (kg)|Accel used (m/s2)|Aux power (kW)|"
    "Distance (km)|Trip time incl charging (h)|Operating time (h)|"
    "Stationary time (h)|Moving time (h)|Moving speed (km/h)|"
    "v Inching (km/h)|v Acc (km/h)|v Dec (km/h)|v Cruise (km/h)|"
    "Road-load traction (kWh)|Road-load Dec (kWh)|"
    "Inertial energy (kWh)|Grade energy (kWh)|Road-load net (kWh)|"
    "Gradient wheel (kWh)|Braking energy (kWh)|Wheel energy (kWh)|"
    "Propulsion energy (kWh)|Auxiliary energy (kWh)|"
    "Idle fuel energy (kWh)|Regen recovered (kWh)|"
    "Modelled energy (kWh)|Measured energy (kWh eq)|Calibration factor|"
    "Energy consumed (kWh)|Energy consumed (MJ)|Diesel consumed (L)|"
    "EV grid energy (kWh)|SEC (kWh/km)|Efficiency (km/kWh)|"
    "Fuel economy (km/L)|Powertrain efficiency (%)|Energy cost (INR)|"
    "CO2 (kg)|Cost per km (INR/km)|CO2 per km (kg/km)|"
    "Model error vs measured|Chart SEC EV|Chart SEC Diesel"
).split("|")


# =====================================================================
# GENERAL HELPERS
# =====================================================================

def require(condition, message):
    if not condition:
        raise ValueError(message)


def text(value):
    return "" if value is None else str(value).strip()


def rgb(r, g, b):
    return r + 256 * g + 65536 * b


NAVY = rgb(25, 43, 65)
EV_COLOR = rgb(22, 119, 132)
DIESEL_COLOR = rgb(230, 126, 34)
GREEN = rgb(31, 143, 88)
WHITE = rgb(255, 255, 255)
BACKGROUND = rgb(242, 245, 249)
INPUT_BLUE = rgb(226, 237, 249)
GREY = rgb(100, 112, 128)

PALETTE = [
    rgb(108, 92, 231),
    rgb(150, 160, 175),
    rgb(195, 202, 212),
    rgb(73, 164, 195),
    rgb(230, 126, 34),
    rgb(66, 165, 113),
    rgb(25, 80, 120),
]


def formula(expression):
    return NAME_PATTERN.sub(
        lambda match: "TAT_" + match.group(0),
        expression,
    )


def put(sheet, address, value):
    try:
        cell = sheet.Range(address)
        if isinstance(value, str) and value.startswith("="):
            cell.Formula = formula(value)
        else:
            cell.Value2 = value
        return cell
    except Exception as exc:
        raise RuntimeError(
            f"Cannot write {sheet.Name}!{address}\n"
            f"Value/formula: {value!r}"
        ) from exc


def sheet_names(workbook):
    return [
        workbook.Worksheets.Item(index).Name
        for index in range(1, workbook.Worksheets.Count + 1)
    ]


def define(workbook, name, reference):
    workbook.Names.Add(
        Name="TAT_" + name,
        RefersTo=formula(reference),
    )


def file_hash(path):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def number(value, label, nonnegative=True):
    try:
        result = float(value)
    except (TypeError, ValueError):
        raise ValueError(
            f"{label}: missing or nonnumeric value {value!r}"
        )

    require(math.isfinite(result), f"{label}: nonfinite number")
    require(
        not nonnegative or result >= 0,
        f"{label}: negative value",
    )
    return result


def date_serial(value):
    if isinstance(value, (int, float)):
        require(
            math.isfinite(value) and value > 0,
            f"Invalid date: {value}",
        )
        return int(value)

    if isinstance(value, datetime):
        parsed = value.replace(tzinfo=None)
    elif isinstance(value, date):
        parsed = datetime.combine(value, time())
    else:
        parsed = None
        for fmt in (
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%Y-%m-%d %H:%M:%S",
        ):
            try:
                parsed = datetime.strptime(text(value), fmt)
                break
            except ValueError:
                pass

        require(
            parsed is not None,
            f"Unrecognised date: {value!r}",
        )

    return (parsed - EXCEL_EPOCH).days


def time_serial(value):
    if isinstance(value, (int, float)):
        require(math.isfinite(value), f"Invalid time: {value}")
        return float(value) % 1

    if isinstance(value, datetime):
        parsed = value.time()
    elif isinstance(value, time):
        parsed = value
    else:
        parsed = None
        for fmt in ("%H:%M:%S", "%H:%M:%S.%f", "%H:%M"):
            try:
                parsed = datetime.strptime(text(value), fmt).time()
                break
            except ValueError:
                pass

        require(
            parsed is not None,
            f"Unrecognised time: {value!r}",
        )

    return (
        parsed.hour * 3600
        + parsed.minute * 60
        + parsed.second
        + parsed.microsecond / 1_000_000
    ) / 86400


def normalise_direction(value):
    result = text(value).casefold()

    for old, new in (
        ("madhuravoyal", "maduravoyal"),
        ("maduravayol", "maduravoyal"),
        ("maduravoyil", "maduravoyal"),
        ("silkboard", "silk board"),
        ("puducherry", "pondicherry"),
        ("pondy", "pondicherry"),
        ("pondi", "pondicherry"),
    ):
        result = re.sub(
            r"\b" + re.escape(old) + r"\b",
            new,
            result,
        )

    result = re.sub(r"\bto\b|→|↔|->|–|—", " ", result)
    return " ".join(
        re.sub(r"[^a-z0-9]+", " ", result).split()
    )


def trip_key(pair, vehicle, trip_date, start):
    return (
        text(pair),
        text(vehicle),
        date_serial(trip_date),
        round(time_serial(start) * 86400, 3),
    )


def merge_text(
    sheet,
    address,
    value,
    size=11,
    fill=None,
    color=NAVY,
    bold=False,
):
    area = sheet.Range(address)
    area.Merge()
    cell = area.Cells(1, 1)

    if isinstance(value, str) and value.startswith("="):
        cell.Formula = formula(value)
    else:
        cell.Value2 = value

    area.Font.Name = "Segoe UI"
    area.Font.Size = size
    area.Font.Color = color
    area.Font.Bold = bold
    area.VerticalAlignment = -4108
    area.WrapText = True

    if fill is not None:
        area.Interior.Color = fill

    return area


def header(sheet, address):
    area = sheet.Range(address)
    area.Interior.Color = NAVY
    area.Font.Color = WHITE
    area.Font.Bold = True
    area.WrapText = True


def validation_list(sheet, address, source):
    """
    All dropdowns use ranges or names.
    No International() calls or literal list separators are needed.
    """
    require(
        str(source).startswith("="),
        f"Dropdown source must be a range/name: {source!r}",
    )

    try:
        rule = sheet.Range(address).Validation
        rule.Delete()

        # Type, AlertStyle, Operator, Formula1.
        rule.Add(3, 1, 1, formula(source))

        rule.IgnoreBlank = True
        rule.InCellDropdown = True
        rule.ShowError = True
        rule.ErrorTitle = "Select a listed value"
        rule.ErrorMessage = "Use the dropdown list."

    except Exception as exc:
        raise RuntimeError(
            f"Dropdown failed at {sheet.Name}!{address}\n"
            f"Source: {source}"
        ) from exc


def validation_number(
    sheet,
    address,
    minimum,
    maximum,
    whole=False,
):
    try:
        rule = sheet.Range(address).Validation
        rule.Delete()
        rule.Add(
            1 if whole else 2,
            1,
            1,
            str(minimum),
            str(maximum),
        )
        rule.IgnoreBlank = True
        rule.ShowError = True
        rule.ErrorTitle = "Invalid number"
        rule.ErrorMessage = (
            f"Enter a value from {minimum} to {maximum}."
        )
    except Exception as exc:
        raise RuntimeError(
            f"Numeric validation failed at "
            f"{sheet.Name}!{address}"
        ) from exc


def validation_status_format(sheet, last_row):
    area = sheet.Range(f"AA5:AA{last_row}")
    area.FormatConditions.Delete()

    try:
        # Explicit Operator argument avoids the missing-parameter COM error.
        rule = area.FormatConditions.Add(
            2,
            1,
            '=$AA5<>"OK"',
        )
        rule.Interior.Color = rgb(255, 220, 220)
        rule.Font.Color = rgb(170, 30, 30)
    except Exception as exc:
        raise RuntimeError(
            "Cannot create the Trip_Input status "
            "conditional-formatting rule."
        ) from exc


# =====================================================================
# READ SUMMARY AND EXISTING MANUAL INPUTS
# =====================================================================

def read_table(sheet, table_number):
    last_row = (
        sheet.UsedRange.Row
        + sheet.UsedRange.Rows.Count
        - 1
    )

    first_column = sheet.Range(f"A1:A{last_row}").Value2

    titles = [
        row_number
        for row_number, row in enumerate(first_column, 1)
        if text(row[0]).startswith(f"Table {table_number}:")
    ]

    require(
        len(titles) == 1,
        f"Cannot identify Table {table_number}.",
    )

    header_row = titles[0] + 3

    raw_headers = sheet.Range(
        sheet.Cells(header_row, 1),
        sheet.Cells(header_row, 100),
    ).Value2[0]

    columns = []
    for value in raw_headers:
        if not text(value):
            break
        columns.append(text(value))

    require(
        "Pair ID" in columns,
        f"Table {table_number}: missing Pair ID.",
    )
    require(
        last_row > header_row,
        f"Table {table_number}: no data.",
    )

    pair_column = columns.index("Pair ID")

    values = sheet.Range(
        sheet.Cells(header_row + 1, 1),
        sheet.Cells(last_row, len(columns)),
    ).Value2

    result = []
    for row in values:
        if not text(row[pair_column]):
            break
        result.append(dict(zip(columns, row)))

    require(result, f"Table {table_number}: no records.")
    return result


def snapshot_inputs(workbook):
    settings = {}
    optional = {}
    routes = list(DEFAULT_ROUTES)

    present = set(sheet_names(workbook)).intersection(
        DASHBOARD_SHEETS
    )

    require(
        not present or present == set(DASHBOARD_SHEETS),
        "Only some dashboard sheets exist. "
        "Restore a complete workbook before rebuilding.",
    )

    if not present:
        return settings, optional, routes

    model = workbook.Worksheets("Model")

    for row, parameter in enumerate(PARAMETERS, 6):
        for column, default in (
            ("D", parameter[2]),
            ("E", parameter[3]),
        ):
            if isinstance(default, (int, float)):
                value = model.Range(f"{column}{row}").Value2
                require(
                    isinstance(value, (int, float)),
                    f"Model!{column}{row} must be numeric.",
                )
                settings[f"{column}{row}"] = value

    calibration = text(model.Range("D38").Value2)
    require(
        calibration in ("Yes", "No"),
        "Model!D38 must be Yes or No.",
    )
    settings["D38"] = calibration

    table = model.ListObjects("Routes")
    if table.DataBodyRange is not None:
        routes = [
            tuple(row[:6])
            for row in table.DataBodyRange.Value2
        ]

    table = workbook.Worksheets("Trip_Input").ListObjects("Trips")

    if table.DataBodyRange is not None:
        for row in table.DataBodyRange.Value2:
            values = tuple(row[21:26])

            if all(value is None for value in values):
                continue

            key = trip_key(row[1], row[3], row[2], row[6])

            require(
                key not in optional or optional[key] == values,
                f"Conflicting optional inputs for {key}.",
            )
            optional[key] = values

    return settings, optional, routes


def extract_records(workbook, routes, optional):
    route_lookup = {}
    require(routes, "The route master is empty.")

    for route in routes:
        require(len(route) == 6, f"Invalid route: {route}")
        require(
            all(text(value) for value in route[:4]),
            "Blank route labels.",
        )

        corridor, direction, _, _, start_h, end_h = route

        number(start_h, f"{direction}: origin elevation", False)
        number(end_h, f"{direction}: destination elevation", False)

        key = normalise_direction(direction)

        require(
            key not in route_lookup,
            f"Duplicate direction: {direction}",
        )
        route_lookup[key] = (text(corridor), text(direction))

    source = workbook.Worksheets("01_All_Pairs")
    summary = read_table(source, 1)

    def phase_index(rows, label):
        result = {}

        for row in rows:
            key = (
                text(row["Pair ID"]),
                text(row["Vehicle"]).upper(),
            )
            require(
                key not in result,
                f"{label}: duplicate {key}",
            )
            result[key] = [
                number(
                    row.get(phase),
                    f"{label} {key}: {phase}",
                )
                for phase in PHASES
            ]

        return result

    times = phase_index(
        read_table(source, 2),
        "Phase times",
    )
    distances = phase_index(
        read_table(source, 3),
        "Phase distances",
    )

    pair_ids = [text(row["Pair ID"]) for row in summary]

    require(
        len(pair_ids) == len(set(pair_ids)),
        "Duplicate summary Pair ID.",
    )

    expected = {
        (pair_id, vehicle)
        for pair_id in pair_ids
        for vehicle in ("EV", "ICE")
    }

    require(
        set(times) == expected and set(distances) == expected,
        "Pair/vehicle keys do not match across the three tables.",
    )

    records = []
    retained_optional = set()

    for pair in summary:
        pair_id = text(pair["Pair ID"])
        direction_key = normalise_direction(pair["Trips"])

        require(
            direction_key in route_lookup,
            f"Unknown direction {pair['Trips']!r}. "
            "Add its endpoint elevations to DEFAULT_ROUTES "
            "or the existing Model/Routes table.",
        )

        corridor, direction = route_lookup[direction_key]

        for source_type, vehicle in (
            ("EV", "EV"),
            ("ICE", "Diesel"),
        ):
            key = (pair_id, source_type)
            phase_times = times[key]
            phase_distances = distances[key]

            total_time = number(
                pair[f"{source_type} TAT (min)"],
                f"{key}: TAT",
            )
            total_distance = number(
                pair[f"{source_type} Distance (km)"],
                f"{key}: distance",
            )

            require(
                abs(sum(phase_times) - total_time) <= 0.01,
                f"{key}: phase times do not reconcile.",
            )
            require(
                abs(sum(phase_distances) - total_distance) <= 0.01,
                f"{key}: phase distances do not reconcile.",
            )
            require(
                total_distance > 0,
                f"{key}: no positive distance.",
            )

            trip_date = date_serial(pair[f"{source_type} Date"])
            start = time_serial(pair[f"{source_type} Start Time"])

            optional_key = trip_key(
                pair_id, vehicle, trip_date, start
            )
            extra = optional.get(optional_key, (None,) * 5)

            for index, value in enumerate(extra):
                if value is not None:
                    number(
                        value,
                        f"{key}: optional input",
                        nonnegative=index != 1,
                    )

            if optional_key in optional:
                retained_optional.add(optional_key)

            records.append([
                f"{pair_id}_{source_type}",
                pair_id,
                trip_date,
                vehicle,
                corridor,
                direction,
                start,
                *phase_times,
                *phase_distances,
                *extra,
            ])

    require(records, "No trip records found.")
    require(
        len(records) <= 1_048_572,
        "Excel row limit exceeded.",
    )

    if len(retained_optional) != len(optional):
        print(
            "NOTE: unmatched old optional inputs "
            "remain available in the backup."
        )

    return records


# =====================================================================
# CREATE SHEETS AND MODEL
# =====================================================================

def recreate_sheets(workbook):
    dashboard_reference = re.compile(
        r"(?:^|[!' ])(?:Dashboard|Trip_Input|Model|Calc)'?!",
        re.IGNORECASE,
    )

    for index in range(workbook.Names.Count, 0, -1):
        item = workbook.Names.Item(index)
        name = text(item.Name)
        reference = text(item.RefersTo)

        if (
            name.split("!")[-1].startswith("TAT_")
            or dashboard_reference.search(reference)
            or "Trips[" in reference
            or "Routes[" in reference
        ):
            item.Delete()

    for name in DASHBOARD_SHEETS:
        if name in sheet_names(workbook):
            workbook.Worksheets(name).Delete()

    for index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets.Item(index)

        for table_index in range(1, sheet.ListObjects.Count + 1):
            table_name = sheet.ListObjects.Item(table_index).Name

            require(
                table_name not in ("Trips", "Routes"),
                f"Table {table_name} already exists on {sheet.Name}.",
            )

    created = {}

    for name in DASHBOARD_SHEETS:
        sheet = workbook.Worksheets.Add(
            After=workbook.Worksheets.Item(
                workbook.Worksheets.Count
            )
        )
        sheet.Name = name
        sheet.Cells.Font.Name = "Segoe UI"
        sheet.Cells.Font.Size = 10
        created[name] = sheet

    return created


def build_trip_table(sheet, records):
    require(
        len(HEADERS) == 76,
        "Incorrect Trips column definition.",
    )
    last_row = 4 + len(records)

    merge_text(
        sheet, "A1:U2", "Trip inputs and energy calculations",
        21, NAVY, WHITE, True,
    )
    merge_text(
        sheet, "A3:Z3",
        "A:U come from 01_All_Pairs. V:Z are optional inputs. "
        "AA:BX are formulas. ICE is displayed as Diesel.",
        10,
    )

    sheet.Range("A4:BX4").Value2 = (tuple(HEADERS),)
    sheet.Range(f"A5:Z{last_row}").Value2 = tuple(
        tuple(row) for row in records
    )

    table = sheet.ListObjects.Add(
        SourceType=1,
        Source=sheet.Range(f"A4:BX{last_row}"),
        XlListObjectHasHeaders=1,
    )
    table.Name = "Trips"
    table.TableStyle = "TableStyleMedium2"

    sheet.Columns("A:B").ColumnWidth = 20
    sheet.Columns("C:D").ColumnWidth = 14
    sheet.Columns("E:F").ColumnWidth = 32
    sheet.Columns("G:BX").ColumnWidth = 16

    sheet.Range("A4:BX4").RowHeight = 60
    sheet.Range("A4:BX4").WrapText = True
    sheet.Range(f"H5:BX{last_row}").NumberFormat = "0.000"
    sheet.Range(f"C5:C{last_row}").NumberFormat = "dd-mmm-yyyy"
    sheet.Range(f"G5:G{last_row}").NumberFormat = "hh:mm:ss"
    sheet.Range(f"AB5:AB{last_row}").NumberFormat = "mmm yyyy"
    sheet.Range(f"V5:Z{last_row}").Interior.Color = INPUT_BLUE

    return last_row


def build_model(workbook, sheet, settings, routes):
    sheet.Columns("A").ColumnWidth = 45
    sheet.Columns("B").ColumnWidth = 3
    sheet.Columns("C").ColumnWidth = 18
    sheet.Columns("D:E").ColumnWidth = 20
    sheet.Columns("F:G").ColumnWidth = 24

    merge_text(
        sheet, "A1:G2", "Energy model and route assumptions",
        21, NAVY, WHITE, True,
    )
    merge_text(
        sheet, "A3:G3",
        "Blue cells are editable. Defaults reproduce "
        "the reference's passenger/coach assumptions.",
        10,
    )

    for address, value in (
        ("A5", "Parameter"),
        ("C5", "Unit"),
        ("D5", "EV"),
        ("E5", "Diesel"),
    ):
        put(sheet, address, value)

    header(sheet, "A5:G5")

    for row, (label, unit, ev, diesel) in enumerate(PARAMETERS, 6):
        put(sheet, f"A{row}", label)
        put(sheet, f"C{row}", unit)
        sheet.Rows(row).RowHeight = 25

        for column, default in (("D", ev), ("E", diesel)):
            address = f"{column}{row}"
            put(sheet, address, settings.get(address, default))

            if isinstance(default, (int, float)):
                sheet.Range(address).Interior.Color = INPUT_BLUE

                minimum = -50 if row in (25, 26) else 0
                maximum = (
                    1 if row in (17, 18, 19, 21, 22, 31)
                    else 1_000_000
                )

                if row in (6, 13, 14, 16, 17, 18, 19, 29, 31):
                    minimum = 0.000001

                validation_number(
                    sheet, address, minimum, maximum
                )

    sheet.Range("D6:E35").NumberFormat = "0.000"
    sheet.Range("D17:E22").NumberFormat = "0.0%"
    sheet.Range("D31").NumberFormat = "0.0%"

    put(sheet, "A38", "Apply global vehicle-type calibration?")
    put(sheet, "D38", settings.get("D38", "Yes"))
    sheet.Range("D38").Interior.Color = INPUT_BLUE

    put(sheet, "H1", "Yes")
    put(sheet, "H2", "No")
    sheet.Columns("H").Hidden = True
    validation_list(sheet, "D38", "=$H$1:$H$2")

    for row, label in {
        39: "Trips with positive measured consumption",
        40: "Measured energy sum (kWh equivalent)",
        41: "Modelled energy for the same records",
        42: "Calibration factor",
    }.items():
        put(sheet, f"A{row}", label)

    for column, vehicle in (("D", "EV"), ("E", "Diesel")):
        criteria = (
            f'Trips[Bus Type],"{vehicle}",'
            'Trips[Measured energy (kWh eq)],">0"'
        )

        put(sheet, f"{column}39", f"=COUNTIFS({criteria})")
        put(
            sheet, f"{column}40",
            f"=SUMIFS(Trips[Measured energy (kWh eq)],{criteria})",
        )
        put(
            sheet, f"{column}41",
            f"=SUMIFS(Trips[Modelled energy (kWh)],{criteria})",
        )
        put(
            sheet, f"{column}42",
            f"=IF({column}40>0,{column}40/{column}41,1)",
        )

    notes = [
        "Methodology",
        "Distances are kilometres; phase times are minutes.",
        "Operating time excludes charging/refuelling. Moving phases "
        "are Inching, Acceleration, Deceleration and Cruise.",
        "Mass = kerb mass + passengers × passenger mass.",
        "Auxiliary power = base power + HVAC slope × "
        "MAX(0, ambient-setpoint).",
        "Road load uses rolling resistance and phase-average aerodynamic drag.",
        "Inertial energy = mass × acceleration × ACC distance / 3600.",
        "Grade energy = mass × gravity × endpoint elevation "
        "difference / 3600000.",
        "Downhill energy is divided between road-load offset and braking.",
        "EV auxiliaries use operating hours. Diesel auxiliaries use "
        "moving hours / engine efficiency. Stop + Idle uses idle fuel.",
        "EV energy = propulsion + auxiliaries - regeneration. Diesel "
        "energy = propulsion + auxiliaries + idle fuel energy.",
        "Measured EV kWh must be battery-side. Convert grid/charger "
        "measurements to the same boundary before entering them.",
        "Calibration is measured-energy sum / modelled-energy sum "
        "globally by vehicle type. Repeated instances are included.",
        "Fleet SEC, cost/km and CO2/km use ratios of sums.",
        "Unique trips retains the first Date + Vehicle + Start Time "
        "+ Direction occurrence.",
        "Monthly trend ignores Month. Corridor league ignores "
        "Corridor and Direction. Latest records ignores all selectors.",
        "Sensitivity uses the reference's acceleration approximation, "
        "not a full recalculation of braking thresholds.",
        "Endpoint elevations are approximate. Intermediate climbs "
        "and descents are not represented.",
        "Energy components may not reconcile for extreme downhill "
        "cases because total wheel energy is clipped at zero.",
        "Phase contribution = selected phase minutes / selected total TAT.",
        "Pair TAT differences require both selected vehicle records. "
        "Missing records are not treated as zero.",
        "Rerun after summary updates to refresh selector lists and chart ranges.",
    ]

    for row, note in enumerate(notes, 46):
        merge_text(sheet, f"A{row}:G{row}", note, 10)
        sheet.Rows(row).RowHeight = 30

    header(sheet, "A46:G46")

    sheet.Range("A80:G80").Value2 = ((
        "Corridor", "Direction", "Origin", "Destination",
        "Origin elevation (m)", "Destination elevation (m)",
        "Net elevation (m)",
    ),)

    last_route = 80 + len(routes)
    sheet.Range(f"A81:F{last_route}").Value2 = tuple(
        tuple(row) for row in routes
    )

    table = sheet.ListObjects.Add(
        SourceType=1,
        Source=sheet.Range(f"A80:G{last_route}"),
        XlListObjectHasHeaders=1,
    )
    table.Name = "Routes"
    table.TableStyle = "TableStyleMedium2"

    put(sheet, "G81", "=F81-E81")
    sheet.Range(f"G81:G{last_route}").FillDown()
    sheet.Range(f"E81:F{last_route}").Interior.Color = INPUT_BLUE
    sheet.Range(f"A80:G{last_route}").WrapText = True
    sheet.Range(f"A80:G{last_route}").RowHeight = 34

    for name, address in MODEL_NAMES.items():
        define(
            workbook,
            name,
            f"=Model!${address[0]}${address[1:]}",
        )


# =====================================================================
# SELECTORS AND TRIP FORMULAS
# =====================================================================

def build_lists(workbook, sheet, records, routes):
    corridors = sorted({row[4] for row in records})

    months = sorted({
        (
            EXCEL_EPOCH + timedelta(days=int(row[2]))
        ).replace(day=1)
        for row in records
    })

    route_corridors = sorted({text(row[0]) for row in routes})

    for address, value in (
        ("A1", "Corridors"),
        ("B1", "Dependent directions"),
        ("C1", "Month labels"),
        ("D1", "Month dates"),
        ("E1", "Basis"),
        ("F1", "Trend metric"),
        ("H1", "Vehicles"),
        ("I1", "Route corridors"),
    ):
        put(sheet, address, value)

    for column in ("A", "B", "C"):
        put(sheet, column + "2", "All")

    for row, corridor in enumerate(corridors, 3):
        put(sheet, f"A{row}", corridor)

    for row, month in enumerate(months, 3):
        put(sheet, f"D{row}", (month - EXCEL_EPOCH).days)
        put(sheet, f"C{row}", f'=TEXT(D{row},"mmm yyyy")')

    put(sheet, "E2", "Unique trips")
    put(sheet, "E3", "All pair instances")

    for row, metric in enumerate([
        "SEC (kWh/km)",
        "Efficiency (km/kWh)",
        "Energy cost (INR/km)",
        "CO2 (kg/km)",
    ], 2):
        put(sheet, f"F{row}", metric)

    put(sheet, "H2", "EV")
    put(sheet, "H3", "Diesel")

    for row, corridor in enumerate(route_corridors, 2):
        put(sheet, f"I{row}", corridor)

    definitions = {
        "L_Corr": f"=Calc!$A$2:$A${len(corridors)+2}",
        "L_Dir": (
            "=OFFSET(Calc!$B$2,0,0,1+"
            f'COUNTIF(Calc!$B$3:$B${len(routes)+2},"?*"),1)'
        ),
        "L_Month": f"=Calc!$C$2:$C${len(months)+2}",
        "L_Basis": "=Calc!$E$2:$E$3",
        "L_Metric": "=Calc!$F$2:$F$5",
        "L_Bus": "=Calc!$H$2:$H$3",
        "L_RouteCorr": (
            f"=Calc!$I$2:$I${len(route_corridors)+1}"
        ),
        "L_RouteDir": (
            f"=Model!$B$81:$B${80+len(routes)}"
        ),
        "Sel_Corridor": "=Dashboard!$B$8",
        "Sel_Direction": "=Dashboard!$F$8",
        "Sel_Month": "=Dashboard!$J$8",
        "Sel_Basis": "=Dashboard!$M$8",
        "Sel_Metric": "=Dashboard!$Q$8",
        "Sel_DirEff": "=Calc!$L$2",
    }

    for name, reference in definitions.items():
        define(workbook, name, reference)

    for row in range(3, len(routes) + 3):
        put(
            sheet,
            f"B{row}",
            '=IFERROR(INDEX(Routes[Direction],AGGREGATE(15,6,'
            '(ROW(Routes[Direction])-ROW(INDEX(Routes[Direction],1))+1)/'
            '(((Routes[Corridor]=Sel_Corridor)+(Sel_Corridor="All"))>0),'
            f'ROWS($B$3:B{row}))),"")',
        )

    put(
        sheet,
        "L2",
        '=IF(OR(Sel_Direction="All",Sel_Direction=""),"All",'
        'IF(Sel_Corridor="All",Sel_Direction,'
        'IF(COUNTIFS(Routes[Corridor],Sel_Corridor,'
        'Routes[Direction],Sel_Direction)=0,"All",Sel_Direction)))',
    )

    put(
        sheet,
        "L3",
        '=IFERROR(MATCH(Sel_Metric,$F$2:$F$5,0),1)',
    )

    return corridors, months


def write_trip_formulas(sheet, last_row):
    expressions = {
        "AA": (
            '=IF(AND(D5<>"EV",D5<>"Diesel"),"Invalid vehicle",'
            'IF(COUNTIFS(Routes[Corridor],E5,Routes[Direction],F5)=0,'
            '"Unknown route",IF(C5="","Missing date",'
            'IF(SUM(O5:U5)<=0,"No distance","OK"))))'
        ),
        "AB": '=IF(C5="","",DATE(YEAR(C5),MONTH(C5),1))',
        "AC": (
            '=--(COUNTIFS($C$5:C5,C5,$D$5:D5,D5,'
            '$G$5:G5,G5,$F$5:F5,F5)=1)'
        ),
        "AD": (
            '=--AND(AA5="OK",OR(Sel_Basis="All pair instances",AC5=1),'
            'OR(Sel_Corridor="All",E5=Sel_Corridor),'
            'OR(Sel_DirEff="All",F5=Sel_DirEff))'
        ),
        "AE": (
            '=--AND(AA5="OK",OR(Sel_Basis="All pair instances",AC5=1),'
            'OR(Sel_Month="All",TEXT(AB5,"mmm yyyy")=Sel_Month))'
        ),
        "AF": "=AD5*AE5",
        "AG": (
            "=SUMIFS(Routes[Net elevation (m)],"
            "Routes[Corridor],E5,Routes[Direction],F5)"
        ),
        "AH": (
            '=IF(D5="EV",Kerb_EV,Kerb_D)'
            '+IF(V5="",Pax_Default,V5)*Pax_Mass'
        ),
        "AI": '=IF(X5="",IF(D5="EV",Acc_EV,Acc_D),X5)',
        "AJ": (
            '=IF(D5="EV",AuxBase_EV,AuxBase_D)'
            '+HVAC_k*MAX(0,IF(W5="",Temp_Default,W5)-HVAC_Tset)'
        ),
        "AK": "=SUM(O5:U5)",
        "AL": "=SUM(H5:N5)/60",
        "AM": "=SUM(I5:N5)/60",
        "AN": "=SUM(I5:J5)/60",
        "AO": "=SUM(K5:N5)/60",
        "AP": "=IFERROR(SUM(R5:U5)/AO5,0)",
        "AQ": "=IFERROR(R5/(K5/60),0)",
        "AR": "=IFERROR(S5/(L5/60),0)",
        "AS": "=IFERROR(T5/(M5/60),0)",
        "AT": "=IFERROR(U5/(N5/60),0)",
        "AU": (
            '=(IF(D5="EV",Crr_EV,Crr_D)*AH5*G_acc*(AK5-T5)'
            '+IF(D5="EV",Kaero_EV,Kaero_D)*'
            '((AQ5/3.6)^2*R5+(AR5/3.6)^2*S5+(AT5/3.6)^2*U5))/3600'
        ),
        "AV": (
            '=(IF(D5="EV",Crr_EV,Crr_D)*AH5*G_acc'
            '+IF(D5="EV",Kaero_EV,Kaero_D)*(AS5/3.6)^2)*T5/3600'
        ),
        "AW": "=AH5*AI5*S5/3600",
        "AX": "=AH5*G_acc*AG5/3600000",
        "AY": "=AU5+MAX(0,AV5-AW5)",
        "AZ": "=MAX(0,AX5)-Descent_Share*MAX(0,-AX5)",
        "BA": "=MAX(0,AW5-AV5)+(1-Descent_Share)*MAX(0,-AX5)",
        "BB": "=MAX(0,AY5+AW5+AZ5)",
        "BC": '=BB5/IF(D5="EV",EtaPT_EV,EtaPT_D)',
        "BD": '=IF(D5="EV",AJ5*AM5,AJ5*AO5/EtaEng_D)',
        "BE": '=IF(D5="EV",0,IdleFuel_D*AN5*LHV_kWhL)',
        "BF": '=IF(D5="EV",Regen_EV*EtaPT_EV*BA5,0)',
        "BG": "=BC5+BD5+BE5-BF5",
        "BH": (
            '=IF(D5="EV",IF(Y5="","",Y5),'
            'IF(Z5="","",Z5*LHV_kWhL))'
        ),
        "BI": '=IF(Cal_Apply="Yes",IF(D5="EV",CF_EV,CF_D),1)',
        "BJ": "=BG5*BI5",
        "BK": "=BJ5*3.6",
        "BL": '=IF(D5="EV","",BJ5/LHV_kWhL)',
        "BM": '=IF(D5="EV",BJ5/EtaCharger,"")',
        "BN": '=IFERROR(BJ5/AK5,"")',
        "BO": '=IFERROR(AK5/BJ5,"")',
        "BP": '=IF(D5="EV","",IFERROR(AK5/BL5,""))',
        "BQ": '=IFERROR(BB5/BJ5,"")',
        "BR": '=IF(D5="EV",BM5*Tariff_EV,BL5*Price_Diesel)',
        "BS": '=IF(D5="EV",BM5*EF_Grid,BL5*EF_Diesel)',
        "BT": '=IFERROR(BR5/AK5,"")',
        "BU": '=IFERROR(BS5/AK5,"")',
        "BV": '=IF(AND(ISNUMBER(BH5),BH5>0),BG5/BH5-1,"")',
        "BW": '=IF(AND(AF5=1,D5="EV"),BN5,NA())',
        "BX": '=IF(AND(AF5=1,D5="Diesel"),BN5,NA())',
    }

    for column, expression in expressions.items():
        put(sheet, f"{column}5", expression)

    sheet.Range(f"AA5:BX{last_row}").FillDown()
    sheet.Range(f"BQ5:BQ{last_row}").NumberFormat = "0.0%"
    sheet.Range(f"BV5:BV{last_row}").NumberFormat = "0.0%"

    for address, source in (
        (f"D5:D{last_row}", "=L_Bus"),
        (f"E5:E{last_row}", "=L_RouteCorr"),
        (f"F5:F{last_row}", "=L_RouteDir"),
    ):
        validation_list(sheet, address, source)

    for address, minimum, maximum, whole in (
        (f"C5:C{last_row}", 1, 2958465, True),
        (f"G5:G{last_row}", 0, 0.999999999, False),
        (f"H5:U{last_row}", 0, 1_000_000, False),
        (f"V5:V{last_row}", 0, 200, True),
        (f"W5:W{last_row}", -50, 70, False),
        (f"X5:X{last_row}", 0, 10, False),
        (f"Y5:Z{last_row}", 0, 1_000_000, False),
    ):
        validation_number(
            sheet, address, minimum, maximum, whole
        )

    validation_status_format(sheet, last_row)


# =====================================================================
# FLEET AGGREGATIONS
# =====================================================================

def sum_trips(metric, vehicle, flag="Flag all", extra=""):
    return (
        f'SUMIFS(Trips[{metric}],Trips[Bus Type],"{vehicle}",'
        f'Trips[{flag}],1{extra})'
    )


def build_calculations(sheet, corridors, months):
    for address, value in (
        ("K20", "Selected fleet"),
        ("L20", "EV"),
        ("M20", "Diesel"),
        ("K21", "Selected trips"),
    ):
        put(sheet, address, value)

    totals = {
        22: "Distance (km)",
        23: "Energy consumed (kWh)",
        24: "Energy consumed (MJ)",
        25: "Diesel consumed (L)",
        26: "EV grid energy (kWh)",
        27: "Wheel energy (kWh)",
        28: "Energy cost (INR)",
        29: "CO2 (kg)",
        30: "Operating time (h)",
        31: "t Charging (min)",
        32: "Trip time incl charging (h)",
    }

    ratios = {
        33: (23, 22),
        34: (22, 23),
        35: (22, 25),
        36: (27, 23),
        37: (28, 22),
        38: (29, 22),
        39: (31, 21),
        40: (32, 21),
        41: (23, 21),
        42: (22, 30),
    }

    for column, vehicle in (("L", "EV"), ("M", "Diesel")):
        put(
            sheet, f"{column}21",
            f'=COUNTIFS(Trips[Bus Type],"{vehicle}",'
            'Trips[Flag all],1)',
        )

        for row, metric in totals.items():
            put(sheet, f"K{row}", metric)
            put(
                sheet, f"{column}{row}",
                "=" + sum_trips(metric, vehicle),
            )

        for row, (numerator, denominator) in ratios.items():
            put(
                sheet, f"{column}{row}",
                f'=IFERROR({column}{numerator}/'
                f'{column}{denominator},"")',
            )

        put(
            sheet, f"{column}48",
            f'=COUNTIFS(Trips[Bus Type],"{vehicle}",'
            'Trips[Measured energy (kWh eq)],">0")',
        )

    labels = {
        33: "SEC kWh/km",
        34: "Efficiency km/kWh",
        35: "Economy km/L",
        36: "Wheel / consumed energy",
        37: "Cost/km",
        38: "CO2/km",
        39: "Average charging/refuel minutes",
        40: "Average trip hours",
        41: "Average trip energy",
        42: "Operating speed",
        43: "Energy saving",
        44: "Cost saving",
        45: "CO2 saving",
        46: "Efficiency multiplier",
        47: "Full TAT difference, minutes",
        48: "Global measured trips",
    }

    for row, label in labels.items():
        put(sheet, f"K{row}", label)

    put(sheet, "L35", '=""')

    for row, metric in ((43, 33), (44, 37), (45, 38)):
        put(
            sheet, f"L{row}",
            f'=IFERROR(1-L{metric}/M{metric},"")',
        )

    put(sheet, "L46", '=IFERROR(L34/M34,"")')
    put(
        sheet, "L47",
        '=IF(AND(L21>0,M21>0),(L40-M40)*60,"")',
    )

    components = [
        ("Road load", "Road-load net (kWh)", True, 1),
        ("Acceleration", "Inertial energy (kWh)", True, 1),
        ("Gradient", "Gradient wheel (kWh)", True, 1),
        ("Auxiliaries / HVAC", "Auxiliary energy (kWh)", False, 1),
        ("Engine idling", "Idle fuel energy (kWh)", False, 1),
        ("Regeneration", "Regen recovered (kWh)", False, -1),
    ]

    for row, (label, metric, wheel_side, sign) in enumerate(
        components, 51
    ):
        put(sheet, f"K{row}", label)

        for column, vehicle, suffix in (
            ("L", "EV", "EV"),
            ("M", "Diesel", "D"),
        ):
            divisor = f"/EtaPT_{suffix}" if wheel_side else ""

            put(
                sheet, f"{column}{row}",
                f'=IFERROR({sign}*{sum_trips(metric, vehicle)}'
                f'*IF(Cal_Apply="Yes",CF_{suffix},1)'
                f'{divisor}/{column}$22,0)',
            )

    for row, (label, metric) in enumerate(
        zip(PHASE_LABELS, HEADERS[7:14]), 60
    ):
        put(sheet, f"K{row}", label)

        for column, vehicle in (("L", "EV"), ("M", "Diesel")):
            put(
                sheet, f"{column}{row}",
                f"=IFERROR({sum_trips(metric, vehicle)}/"
                f"{column}$21,0)",
            )

    for row, acceleration in enumerate(
        (0.15, 0.20, 0.25, 0.30, 0.35, 0.40), 71
    ):
        put(sheet, f"K{row}", acceleration)

        for column, vehicle, suffix in (
            ("L", "EV", "EV"),
            ("M", "Diesel", "D"),
        ):
            adjustment = (
                "(1/EtaPT_EV-Regen_EV*EtaPT_EV)"
                if vehicle == "EV"
                else "(1/EtaPT_D)"
            )

            put(
                sheet, f"{column}{row}",
                f'=IFERROR(({column}$23+($K{row}/Acc_{suffix}-1)*'
                f'{sum_trips("Inertial energy (kWh)", vehicle)}'
                f'*IF(Cal_Apply="Yes",CF_{suffix},1)*'
                f'{adjustment})/{column}$22,NA())',
            )

        put(
            sheet, f"N{row}",
            f"=IFERROR(1-L{row}/M{row},NA())",
        )
        put(
            sheet, f"O{row}",
            f'=TEXT(K{row},"0.00")&'
            f'IF(ABS(K{row}-Acc_EV)<0.001," *","")',
        )

    for address, value in (
        ("X50", "Vehicle"),
        ("X51", "EV"),
        ("X52", "Diesel"),
        ("X60", "Vehicle"),
        ("X61", "EV"),
        ("X62", "Diesel"),
        ("X78", "EV trip"),
        ("Y78", "Charging"),
        ("Z78", "Rest of trip"),
        ("X79", "EV"),
        ("Y79", "=L60"),
        ("Z79", "=MAX(0,IFERROR(L40*60,0)-L60)"),
    ):
        put(sheet, address, value)

    for column, row, component in zip(
        ("Y", "Z", "AA", "AB", "AC", "AD"),
        range(51, 57),
        components,
    ):
        put(sheet, f"{column}50", component[0])
        put(sheet, f"{column}51", f"=L{row}")
        put(sheet, f"{column}52", f"=M{row}")

    for column, row, label in zip(
        ("Y", "Z", "AA", "AB", "AC", "AD", "AE"),
        range(60, 67),
        PHASE_LABELS,
    ):
        put(sheet, f"{column}60", label)
        put(sheet, f"{column}61", f"=L{row}")
        put(sheet, f"{column}62", f"=M{row}")

    # Corridor league: Month and Basis only.
    for row, corridor in enumerate(corridors, 83):
        put(sheet, f"K{row}", corridor)
        extra = f",Trips[Corridor],K{row}"

        for column, vehicle, metric in (
            ("L", "EV", "Distance (km)"),
            ("M", "EV", "Energy consumed (kWh)"),
            ("N", "Diesel", "Distance (km)"),
            ("O", "Diesel", "Energy consumed (kWh)"),
        ):
            put(
                sheet, f"{column}{row}",
                "=" + sum_trips(
                    metric, vehicle, "Flag period", extra
                ),
            )

        put(sheet, f"P{row}", f"=IFERROR(M{row}/L{row},NA())")
        put(sheet, f"Q{row}", f"=IFERROR(O{row}/N{row},NA())")
        put(sheet, f"R{row}", f"=IFERROR(1-P{row}/Q{row},NA())")

    # Monthly trend ignores the selected Month.
    month_start = max(102, 86 + len(corridors))

    for row, month in enumerate(months, month_start):
        put(sheet, f"K{row}", (month - EXCEL_EPOCH).days)
        sheet.Range(f"K{row}").NumberFormat = "mmm yyyy"

        extra = f",Trips[Month],K{row}"

        metrics = [
            "Distance (km)",
            "Energy consumed (kWh)",
            "Energy cost (INR)",
            "CO2 (kg)",
        ]

        for columns, vehicle in (
            (("L", "M", "N", "O"), "EV"),
            (("P", "Q", "R", "S"), "Diesel"),
        ):
            for column, metric in zip(columns, metrics):
                put(
                    sheet, f"{column}{row}",
                    "=" + sum_trips(
                        metric, vehicle, "Flag route", extra
                    ),
                )

        put(
            sheet, f"T{row}",
            f"=IFERROR(CHOOSE($L$3,M{row}/L{row},L{row}/M{row},"
            f"N{row}/L{row},O{row}/L{row}),NA())",
        )
        put(
            sheet, f"U{row}",
            f"=IFERROR(CHOOSE($L$3,Q{row}/P{row},P{row}/Q{row},"
            f"R{row}/P{row},S{row}/P{row}),NA())",
        )

    # Direct fleet comparison charts.
    sheet.Range("BT1:BV1").Value2 = ((
        "Metric", "EV", "Diesel",
    ),)

    for row, label, source_row in (
        (2, "Charging / refuelling min", 39),
        (3, "Energy efficiency km/kWh", 34),
        (4, "Energy cost INR/km", 37),
        (5, "CO2 intensity kg/km", 38),
    ):
        put(sheet, f"BT{row}", label)

        for target, source in (("BU", "L"), ("BV", "M")):
            put(
                sheet, f"{target}{row}",
                f'=IF({source}21>0,{source}{source_row},NA())',
            )

    return month_start, month_start + len(months) - 1


# =====================================================================
# PAIR-WISE TAT AND PHASE CHART DATA
# =====================================================================

def build_pair_calculations(sheet, records):
    pairs = list(dict.fromkeys(row[1] for row in records))
    require(pairs, "No pair IDs found.")
    last_pair = len(pairs) + 1

    sheet.Range("AK1:AQ1").Value2 = ((
        "Candidate pair",
        "Selected rows",
        "Selected index",
        "Selected pair",
        "EV TAT min",
        "Diesel TAT min",
        "EV minus Diesel min",
    ),)

    sheet.Range(f"AK2:AK{last_pair}").Value2 = tuple(
        (pair,) for pair in pairs
    )

    for row in range(2, last_pair + 1):
        put(
            sheet, f"AL{row}",
            f'=COUNTIFS(Trips[Pair ID],AK{row},'
            'Trips[Flag all],1)',
        )
        put(
            sheet, f"AM{row}",
            '=IFERROR(AGGREGATE(15,6,'
            f'(ROW($AK$2:$AK${last_pair})-ROW($AK$2)+1)/'
            f'($AL$2:$AL${last_pair}>0),'
            f'ROWS($AM$2:AM{row})),0)',
        )
        put(
            sheet, f"AN{row}",
            f'=IF(AM{row}=0,"",'
            f'INDEX($AK$2:$AK${last_pair},AM{row}))',
        )

        for column, vehicle in (("AO", "EV"), ("AP", "Diesel")):
            criteria = (
                f'Trips[Pair ID],AN{row},'
                f'Trips[Bus Type],"{vehicle}",'
                'Trips[Flag all],1'
            )

            put(
                sheet, f"{column}{row}",
                f'=IF(AND(AN{row}<>"",COUNTIFS({criteria})>0),'
                f'SUMIFS(Trips[Trip time incl charging (h)],'
                f'{criteria})*60,NA())',
            )

        put(
            sheet, f"AQ{row}",
            f"=IFERROR(AO{row}-AP{row},NA())",
        )

    for label_column, phase_columns, vehicle in (
        (
            "BA",
            ("BB", "BC", "BD", "BE", "BF", "BG", "BH"),
            "EV",
        ),
        (
            "BK",
            ("BL", "BM", "BN", "BO", "BP", "BQ", "BR"),
            "Diesel",
        ),
    ):
        put(sheet, f"{label_column}1", f"{vehicle} pair")

        for column, label in zip(phase_columns, PHASE_LABELS):
            put(sheet, f"{column}1", label)

        for row in range(2, last_pair + 1):
            put(sheet, f"{label_column}{row}", f"=AN{row}")

            criteria = (
                f'Trips[Pair ID],AN{row},'
                f'Trips[Bus Type],"{vehicle}",'
                'Trips[Flag all],1'
            )

            for column, metric in zip(
                phase_columns, HEADERS[7:14]
            ):
                put(
                    sheet, f"{column}{row}",
                    f'=IF(AND(AN{row}<>"",'
                    f'COUNTIFS({criteria})>0),'
                    f'SUMIFS(Trips[{metric}],{criteria}),NA())',
                )

    # Sheet-local dynamic names keep filtered chart ranges compact.
    dynamic_columns = [
        "AN", "AO", "AP", "AQ",
        "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH",
        "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR",
    ]

    for column in dynamic_columns:
        sheet.Names.Add(
            Name="TAT_Pair_" + column,
            RefersTo=(
                f"=OFFSET(Calc!${column}$2,0,0,"
                f'MAX(1,COUNTIF(Calc!$AL$2:$AL${last_pair},">0")),1)'
            ),
        )

    sheet.Columns("AK:BV").ColumnWidth = 18
    return len(pairs)


# =====================================================================
# DASHBOARD LAYOUT
# =====================================================================

def build_dashboard(workbook, sheet):
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("B:V").ColumnWidth = 10
    sheet.Range("A1:W111").Interior.Color = BACKGROUND
    sheet.Range("A1:W111").RowHeight = 20

    merge_text(
        sheet, "B2:V3",
        "EV and Diesel TAT and Energy Dashboard",
        24, NAVY, WHITE, True,
    )
    merge_text(
        sheet, "B4:V4",
        "Live comparison of trip TAT, phase contributions, "
        "energy, cost and CO2.",
        11,
    )

    for address, label, target in (
        ("B5", "Trip inputs", "'Trip_Input'!A1"),
        ("F5", "Model assumptions", "'Model'!A1"),
        ("K5", "Route elevations", "'Model'!A80"),
        ("Q5", "Source summary", "'01_All_Pairs'!A1"),
    ):
        # Anchor, Address, SubAddress, ScreenTip, TextToDisplay.
        sheet.Hyperlinks.Add(
            sheet.Range(address), "", target, label, label
        )

    selectors = [
        ("B7:E7", "B8:E9", "Corridor", "All", "L_Corr"),
        ("F7:I7", "F8:I9", "Direction", "All", "L_Dir"),
        ("J7:L7", "J8:L9", "Month", "All", "L_Month"),
        (
            "M7:P7", "M8:P9", "Counting basis",
            "All pair instances", "L_Basis",
        ),
        (
            "Q7:V7", "Q8:V9", "Trend metric",
            "SEC (kWh/km)", "L_Metric",
        ),
    ]

    for label_area, input_area, label, default, _ in selectors:
        merge_text(sheet, label_area, label, 10, bold=True)
        merge_text(
            sheet, input_area, default,
            11, INPUT_BLUE, NAVY, True,
        )

    sheet.Application.CalculateFullRebuild()

    for _, input_area, _, _, source in selectors:
        validation_list(sheet, input_area, "=" + source)

    merge_text(
        sheet, "B10:V10",
        '=IF(AND(Sel_Direction<>"All",'
        'Sel_Direction<>Sel_DirEff),'
        '"Direction is not on this corridor: showing All directions.",'
        'IF(SUM(Calc!L21:M21)=0,'
        '"No trips match the current selection.",'
        '"Results use the selected trip population."))',
        10,
    )

    def card(label_area, value_area, label, expression, color=NAVY):
        merge_text(
            sheet, label_area, label,
            10, WHITE, GREY, True,
        )
        merge_text(
            sheet, value_area, expression,
            17, WHITE, color, True,
        )

    def comparison(row, fmt):
        return (
            '=IF(AND(Calc!L21>0,Calc!M21>0),'
            f'"EV "&TEXT(Calc!L{row},"{fmt}")&'
            f'" | Diesel "&TEXT(Calc!M{row},"{fmt}"),'
            '"Comparison unavailable")'
        )

    card(
        "B12:F12", "B13:F15",
        "Energy intensity · kWh/km",
        comparison(33, "0.000"),
        EV_COLOR,
    )
    card(
        "G12:K12", "G13:K15",
        "Energy cost · INR/km",
        comparison(37, "0.00"),
    )
    card(
        "L12:P12", "L13:P15",
        "CO2 intensity · kg/km",
        comparison(38, "0.000"),
    )

    for label_area, value_area, label, calc_row in (
        ("Q12:V12", "Q13:V15", "Fleet energy saving per km", 43),
        ("B17:F17", "B18:F20", "Energy cost saving per km", 44),
        ("G17:K17", "G18:K20", "CO2 saving per km", 45),
    ):
        card(
            label_area,
            value_area,
            label,
            '=IF(AND(Calc!L21>0,Calc!M21>0),'
            f'TEXT(Calc!L{calc_row},"0.0%"),'
            '"Comparison unavailable")',
            GREEN if calc_row != 45 else NAVY,
        )

    card(
        "L17:P17", "L18:P20",
        "Full TAT difference · EV minus Diesel",
        '=IF(AND(Calc!L21>0,Calc!M21>0),'
        'TEXT(Calc!L47,"+0.00;-0.00;0.00")&" min",'
        '"Comparison unavailable")',
    )
    card(
        "Q17:V17", "Q18:V20",
        "Charging/refuelling difference",
        '=IF(AND(Calc!L21>0,Calc!M21>0),'
        'TEXT(Calc!L39-Calc!M39,"+0.00;-0.00;0.00")&'
        '" min","Comparison unavailable")',
    )

    merge_text(
        sheet, "B22:J22", "Selected fleet profile",
        13, NAVY, WHITE, True,
    )
    merge_text(
        sheet, "B23:D23", "Metric",
        10, WHITE, NAVY, True,
    )
    merge_text(
        sheet, "E23:G23", "EV",
        10, WHITE, EV_COLOR, True,
    )
    merge_text(
        sheet, "H23:J23", "Diesel",
        10, WHITE, DIESEL_COLOR, True,
    )

    fleet_rows = [
        ("Trips", 21, "0"),
        ("Distance · km", 22, "0.00"),
        ("Energy · kWh equivalent", 23, "0.00"),
        ("Efficiency · km/kWh", 34, "0.000"),
        ("Diesel economy · km/L", 35, "0.000"),
        ("Average trip · min", 40, "0.00"),
        ("Average charging/refuel · min", 39, "0.00"),
        ("Operating speed · km/h", 42, "0.00"),
        ("Wheel / consumed energy", 36, "0.0%"),
    ]

    for row, (label, source_row, fmt) in enumerate(fleet_rows, 24):
        merge_text(
            sheet, f"B{row}:D{row}", label, 10, WHITE
        )

        for area, column in (
            (f"E{row}:G{row}", "L"),
            (f"H{row}:J{row}", "M"),
        ):
            expression = f"=Calc!{column}{source_row}"

            if source_row == 40:
                expression = (
                    f'=IF(Calc!{column}21>0,'
                    f'Calc!{column}40*60,"")'
                )
            elif source_row != 21:
                expression = (
                    f'=IF(Calc!{column}21>0,'
                    f'Calc!{column}{source_row},"")'
                )

            merge_text(
                sheet, area, expression, 10, WHITE
            )
            sheet.Range(area).NumberFormat = fmt

    merge_text(
        sheet, "L22:V22",
        "Latest table records · independent of filters",
        12, NAVY, WHITE, True,
    )

    latest_columns = [
        ("L", "M", "Pair", "Pair ID"),
        ("N", "N", "Type", "Bus Type"),
        ("O", "P", "Date", "Date"),
        ("Q", "S", "Direction", "Direction"),
        ("T", "T", "SEC", "SEC (kWh/km)"),
        ("U", "V", "Check", "Data check"),
    ]

    for left, right, label, _ in latest_columns:
        merge_text(
            sheet, f"{left}23:{right}23",
            label, 9, WHITE, NAVY, True,
        )

    for offset in range(10):
        row = 24 + offset
        sheet.Rows(row).RowHeight = 29

        for left, right, _, column in latest_columns:
            expression = (
                f'=IF(ROWS(Trips[Trip ID])>{offset},'
                f'INDEX(Trips[{column}],'
                f'ROWS(Trips[Trip ID])-{offset}),"")'
            )

            area = f"{left}{row}:{right}{row}"
            merge_text(
                sheet, area, expression, 9, WHITE
            )

            if column == "Date":
                sheet.Range(area).NumberFormat = "dd-mmm-yy"
            elif column == "SEC (kWh/km)":
                sheet.Range(area).NumberFormat = "0.000"

    merge_text(
        sheet, "B35:V35",
        '="Source rows: "&ROWS(Trips[Trip ID])&'
        '" | Latest date: "&TEXT(MAX(Trips[Date]),"dd-mmm-yyyy")&'
        '" | Failed checks: "&COUNTIF(Trips[Data check],"<>OK")&'
        '" | Measured trips EV/Diesel: "&Calc!L48&"/"&Calc!M48&'
        '" | Calibration factors: "&TEXT(CF_EV,"0.000")&'
        '"/"&TEXT(CF_D,"0.000")',
        9,
    )

    merge_text(
        sheet, "B102:V103",
        "Negative saving means EV is higher. Full TAT includes "
        "all phases. Energy uses editable assumptions. "
        "Sensitivity is approximate. * marks EV default acceleration.",
        10,
    )

    card(
        "B105:F105", "B106:F108",
        "Energy efficiency · km/kWh",
        comparison(34, "0.000"),
        EV_COLOR,
    )
    card(
        "G105:K105", "G106:K108",
        "EV energy efficiency multiplier",
        '=IF(AND(Calc!L21>0,Calc!M21>0),'
        'TEXT(Calc!L46,"0.00")&" x",'
        '"Comparison unavailable")',
        GREEN,
    )
    card(
        "L105:P105", "L106:P108",
        "Energy cost saved · INR/km",
        '=IF(AND(Calc!L21>0,Calc!M21>0),'
        'TEXT(Calc!M37-Calc!L37,"0.00"),'
        '"Comparison unavailable")',
        GREEN,
    )
    card(
        "Q105:V105", "Q106:V108",
        "Average total TAT · minutes",
        '=IF(AND(Calc!L21>0,Calc!M21>0),'
        '"EV "&TEXT(Calc!L40*60,"0.00")&'
        '" | Diesel "&TEXT(Calc!M40*60,"0.00"),'
        '"Comparison unavailable")',
    )

    merge_text(
        sheet, "B110:V111",
        "Trip charts follow Corridor, Direction, Month and Counting basis. "
        "Pair differences require both vehicles in the selection. "
        "Positive TAT difference means EV takes longer.",
        10,
    )

    sheet.Tab.Color = NAVY


# =====================================================================
# NATIVE EXCEL CHARTS
# =====================================================================

def add_chart(
    dashboard,
    anchor,
    title,
    chart_type,
    definitions,
    y_title="",
    x_title="",
):
    area = dashboard.Range(anchor)

    chart_object = dashboard.ChartObjects().Add(
        area.Left,
        area.Top,
        area.Width,
        area.Height,
    )

    chart = chart_object.Chart
    chart.ChartType = chart_type

    while chart.SeriesCollection().Count:
        chart.SeriesCollection(1).Delete()

    for index, definition in enumerate(definitions, 1):
        label, xref, yref, color, secondary = definition
        series = chart.SeriesCollection().NewSeries()
        safe_label = label.replace('"', '""')

        try:
            series.Formula = (
                f'=SERIES("{safe_label}",'
                f'{xref},{yref},{index})'
            )

            if secondary:
                series.ChartType = 4
                series.AxisGroup = 2

        except Exception as exc:
            raise RuntimeError(
                f"Cannot create {label!r} series "
                f"in chart {title!r}."
            ) from exc

        try:
            series.Format.Fill.ForeColor.RGB = color
            series.Format.Line.ForeColor.RGB = color

            if chart_type == -4169:
                series.MarkerStyle = 8
                series.MarkerSize = 6
                series.MarkerForegroundColor = color
                series.MarkerBackgroundColor = color
                series.Format.Line.Visible = 0

        except Exception:
            pass

    chart.HasTitle = True
    chart.ChartTitle.Text = title
    chart.HasLegend = len(definitions) > 1

    if chart.HasLegend:
        chart.Legend.Position = -4107
        chart.Legend.Font.Size = 9

    chart.PlotVisibleOnly = False
    chart.DisplayBlanksAs = 1
    chart.ChartArea.Font.Name = "Segoe UI"
    chart.ChartArea.Font.Size = 10

    try:
        chart.ChartArea.Format.Fill.ForeColor.RGB = WHITE
        chart.ChartArea.Format.Line.Visible = 0
        chart.ChartTitle.Font.Size = 13
        chart.ChartTitle.Font.Bold = True

        if y_title:
            chart.Axes(2, 1).HasTitle = True
            chart.Axes(2, 1).AxisTitle.Text = y_title

        if x_title:
            chart.Axes(1, 1).HasTitle = True
            chart.Axes(1, 1).AxisTitle.Text = x_title

        if any(item[4] for item in definitions):
            chart.Axes(2, 2).TickLabels.NumberFormat = "0%"
            chart.Axes(2, 2).HasTitle = True
            chart.Axes(2, 2).AxisTitle.Text = "EV saving"

    except Exception:
        pass

    return chart


def build_charts(
    dashboard,
    corridors,
    month_bounds,
    last_trip,
    pair_count,
):
    first_month, last_month = month_bounds

    def cr(address):
        return f"'Calc'!{address}"

    def tr(address):
        return f"'Trip_Input'!{address}"

    def pr(column):
        return f"'Calc'!TAT_Pair_{column}"

    # 1. Monthly trend.
    categories = cr(
        f"$K${first_month}:$K${last_month}"
    )

    add_chart(
        dashboard, "B37:K51",
        "Monthly trend · selected metric", 51,
        [
            (
                "EV", categories,
                cr(f"$T${first_month}:$T${last_month}"),
                EV_COLOR, False,
            ),
            (
                "Diesel", categories,
                cr(f"$U${first_month}:$U${last_month}"),
                DIESEL_COLOR, False,
            ),
        ],
    )

    # 2. EV charging share.
    add_chart(
        dashboard, "M37:V51",
        "EV charging versus rest of trip", 59,
        [
            (
                "Charging",
                cr("$X$79"), cr("$Y$79"),
                PALETTE[0], False,
            ),
            (
                "Rest of trip",
                cr("$X$79"), cr("$Z$79"),
                EV_COLOR, False,
            ),
        ],
        "Share of total trip time",
    )

    # 3. Corridor league.
    last_corridor = 82 + len(corridors)
    categories = cr(f"$K$83:$K${last_corridor}")

    add_chart(
        dashboard, "B54:K68",
        "Corridor league · month and counting basis", 51,
        [
            (
                "EV", categories,
                cr(f"$P$83:$P${last_corridor}"),
                EV_COLOR, False,
            ),
            (
                "Diesel", categories,
                cr(f"$Q$83:$Q${last_corridor}"),
                DIESEL_COLOR, False,
            ),
            (
                "Saving", categories,
                cr(f"$R$83:$R${last_corridor}"),
                GREEN, True,
            ),
        ],
        "kWh/km",
    )

    # 4. Energy components.
    component_labels = [
        "Road load",
        "Acceleration",
        "Gradient",
        "Auxiliaries/HVAC",
        "Engine idling",
        "Regeneration",
    ]

    add_chart(
        dashboard, "M54:V68",
        "Where the energy goes", 52,
        [
            (
                label,
                cr("$X$51:$X$52"),
                cr(f"${column}$51:${column}$52"),
                PALETTE[index],
                False,
            )
            for index, (label, column) in enumerate(zip(
                component_labels,
                ("Y", "Z", "AA", "AB", "AC", "AD"),
            ))
        ],
        "kWh/km",
    )

    # 5. Average phase minutes.
    add_chart(
        dashboard, "B71:K85",
        "Turnaround by phase", 58,
        [
            (
                label,
                cr("$X$61:$X$62"),
                cr(f"${column}$61:${column}$62"),
                PALETTE[index],
                False,
            )
            for index, (label, column) in enumerate(zip(
                PHASE_LABELS,
                ("Y", "Z", "AA", "AB", "AC", "AD", "AE"),
            ))
        ],
        "Average minutes per trip",
    )

    # 6. Acceleration sensitivity.
    add_chart(
        dashboard, "M71:V85",
        "Acceleration sensitivity · approximation", 51,
        [
            (
                "EV",
                cr("$O$71:$O$76"),
                cr("$L$71:$L$76"),
                EV_COLOR, False,
            ),
            (
                "Diesel",
                cr("$O$71:$O$76"),
                cr("$M$71:$M$76"),
                DIESEL_COLOR, False,
            ),
            (
                "Saving",
                cr("$O$71:$O$76"),
                cr("$N$71:$N$76"),
                GREEN, True,
            ),
        ],
        "kWh/km",
        "Acceleration · m/s²",
    )

    # 7. Moving speed versus SEC.
    add_chart(
        dashboard, "B88:V100",
        "Trip scatter · moving speed and energy intensity",
        -4169,
        [
            (
                "EV",
                tr(f"$AP$5:$AP${last_trip}"),
                tr(f"$BW$5:$BW${last_trip}"),
                EV_COLOR, False,
            ),
            (
                "Diesel",
                tr(f"$AP$5:$AP${last_trip}"),
                tr(f"$BX$5:$BX${last_trip}"),
                DIESEL_COLOR, False,
            ),
        ],
        "SEC · kWh/km",
        "Average moving speed · km/h",
    )

    # Expand trip chart height to make all pair labels readable.
    pair_height = max(18, math.ceil(pair_count * 1.25))
    start = 114

    def chart_area(left, right, row, height):
        end = row + height - 1
        dashboard.Range(
            f"{left}{row}:{right}{end}"
        ).RowHeight = 20
        return f"{left}{row}:{right}{end}"

    # 8. Pair-wise total TAT.
    add_chart(
        dashboard,
        chart_area("B", "K", start, pair_height),
        "Trip-by-trip total TAT", 57,
        [
            ("EV", pr("AN"), pr("AO"), EV_COLOR, False),
            ("Diesel", pr("AN"), pr("AP"), DIESEL_COLOR, False),
        ],
        "Total TAT · minutes",
    )

    # 9. Pair-wise TAT difference.
    add_chart(
        dashboard,
        chart_area("M", "V", start, pair_height),
        "TAT difference · EV minus Diesel", 57,
        [
            (
                "EV minus Diesel",
                pr("AN"), pr("AQ"),
                DIESEL_COLOR, False,
            ),
        ],
        "Difference · minutes",
    )

    start += pair_height + 3

    # 10–11. Individual-trip phase breakdown.
    for left, right, vehicle, label_column, phase_columns in (
        (
            "B", "K", "EV", "BA",
            ("BB", "BC", "BD", "BE", "BF", "BG", "BH"),
        ),
        (
            "M", "V", "Diesel", "BK",
            ("BL", "BM", "BN", "BO", "BP", "BQ", "BR"),
        ),
    ):
        add_chart(
            dashboard,
            chart_area(left, right, start, pair_height),
            f"{vehicle} trip TAT by phase",
            58,
            [
                (
                    label,
                    pr(label_column),
                    pr(column),
                    PALETTE[index],
                    False,
                )
                for index, (label, column) in enumerate(
                    zip(PHASE_LABELS, phase_columns)
                )
            ],
            "Phase time · minutes",
        )

    start += pair_height + 3

    # 12. Direct phase comparison.
    add_chart(
        dashboard,
        chart_area("B", "K", start, 16),
        "Phase-by-phase time comparison",
        57,
        [
            (
                "EV",
                cr("$K$60:$K$66"),
                cr("$L$60:$L$66"),
                EV_COLOR, False,
            ),
            (
                "Diesel",
                cr("$K$60:$K$66"),
                cr("$M$60:$M$66"),
                DIESEL_COLOR, False,
            ),
        ],
        "Average minutes per selected trip",
    )

    # 13. Percentage contribution to TAT.
    add_chart(
        dashboard,
        chart_area("M", "V", start, 16),
        "Phase contribution to total TAT · percent",
        59,
        [
            (
                label,
                cr("$X$61:$X$62"),
                cr(f"${column}$61:${column}$62"),
                PALETTE[index],
                False,
            )
            for index, (label, column) in enumerate(zip(
                PHASE_LABELS,
                ("Y", "Z", "AA", "AB", "AC", "AD", "AE"),
            ))
        ],
        "Share of total TAT",
    )

    start += 19

    # 14–17. Fleet comparisons.
    for left, right, offset, title, source_row, unit in (
        (
            "B", "K", 0,
            "Charging versus refuelling", 2,
            "Minutes per trip",
        ),
        (
            "M", "V", 0,
            "Energy efficiency", 3,
            "km/kWh",
        ),
        (
            "B", "K", 19,
            "Energy cost comparison", 4,
            "INR/km",
        ),
        (
            "M", "V", 19,
            "CO2 intensity comparison", 5,
            "kg/km",
        ),
    ):
        add_chart(
            dashboard,
            chart_area(left, right, start + offset, 16),
            title,
            51,
            [
                (
                    "EV",
                    cr(f"$BT${source_row}"),
                    cr(f"$BU${source_row}"),
                    EV_COLOR, False,
                ),
                (
                    "Diesel",
                    cr(f"$BT${source_row}"),
                    cr(f"$BV${source_row}"),
                    DIESEL_COLOR, False,
                ),
            ],
            unit,
        )

    end_row = start + 37

    merge_text(
        dashboard,
        f"B{end_row}:V{end_row+1}",
        "Phase percentages use selected phase minutes divided by "
        "selected total TAT. Energy and cost savings are compared per km. "
        "Missing sides of a pair are not treated as zero.",
        10,
    )

    dashboard.PageSetup.PrintArea = (
        f"$B$2:$V${end_row+1}"
    )


# =====================================================================
# VALIDATION
# =====================================================================

def validate_parameters(model):
    for row, parameter in enumerate(PARAMETERS, 6):
        for column, default in (
            ("D", parameter[2]),
            ("E", parameter[3]),
        ):
            if isinstance(default, (int, float)):
                number(
                    model.Range(f"{column}{row}").Value2,
                    f"Model!{column}{row}",
                    nonnegative=row not in (25, 26),
                )

    for address in (
        "D6", "E6", "D13", "D14", "D16", "E16",
        "D17", "E17", "D18", "E18", "D19", "E29", "D31",
    ):
        require(
            float(model.Range(address).Value2) > 0,
            f"Model!{address} must be positive.",
        )

    for address in (
        "D17", "E17", "D18", "E18",
        "D19", "D21", "D22", "D31",
    ):
        require(
            0 <= float(model.Range(address).Value2) <= 1,
            f"Model!{address} must be between 0 and 1.",
        )


def validate_workbook(workbook, records):
    trips = workbook.Worksheets("Trip_Input")
    calc = workbook.Worksheets("Calc")
    last_row = 4 + len(records)

    require(
        trips.ListObjects("Trips").ListRows.Count == len(records),
        "Trips table size mismatch.",
    )

    failures = [
        row_number
        for row_number, row in enumerate(
            trips.Range(f"AA5:AA{last_row}").Value2,
            5,
        )
        if row[0] != "OK"
    ]

    require(
        not failures,
        f"Trip checks failed at rows: {failures[:20]}",
    )

    require(
        workbook.Worksheets("Dashboard").ChartObjects().Count
        == EXPECTED_CHARTS,
        f"Expected {EXPECTED_CHARTS} native charts.",
    )

    ev_count = sum(row[3] == "EV" for row in records)

    require(
        calc.Range("L21").Value2 == ev_count
        and calc.Range("M21").Value2 == len(records) - ev_count,
        "Default selector counts do not match source records.",
    )

    errors = []

    for name in DASHBOARD_SHEETS:
        sheet = workbook.Worksheets(name)

        try:
            cells = sheet.UsedRange.SpecialCells(-4123, 16)
        except Exception:
            continue

        for cell in cells.Cells:
            displayed = text(cell.Text)

            intentional_na = (
                displayed == "#N/A"
                and (
                    name == "Calc"
                    or (
                        name == "Trip_Input"
                        and cell.Column in (75, 76)
                    )
                )
            )

            if not intentional_na:
                errors.append(
                    f"{name}!{cell.Address(False, False)}: "
                    f"{displayed}"
                )

            if len(errors) >= 20:
                break

    require(
        not errors,
        "Formula errors:\n" + "\n".join(errors),
    )

    require(
        not workbook.LinkSources(1),
        "External Excel workbook links found.",
    )

    # Check pair TAT values for the default All-pair-instances selection.
    pair_totals = {}
    for row in records:
        pair_totals.setdefault(row[1], {})[row[3]] = sum(row[7:14])

    last_pair = len(pair_totals) + 1
    actual_rows = calc.Range(
        f"AN2:AQ{last_pair}"
    ).Value2

    for pair_id, ev_tat, diesel_tat, difference in actual_rows:
        require(pair_id in pair_totals, "Pair chart ID mismatch.")
        expected = pair_totals[pair_id]

        require(
            abs(float(ev_tat) - expected["EV"]) <= 0.01,
            f"{pair_id}: EV chart TAT mismatch.",
        )
        require(
            abs(float(diesel_tat) - expected["Diesel"]) <= 0.01,
            f"{pair_id}: Diesel chart TAT mismatch.",
        )
        require(
            abs(
                float(difference)
                - (expected["EV"] - expected["Diesel"])
            ) <= 0.01,
            f"{pair_id}: chart TAT difference mismatch.",
        )


def check_live_filters(excel, workbook, records):
    dashboard = workbook.Worksheets("Dashboard")
    calc = workbook.Worksheets("Calc")

    corridor, direction = records[0][4], records[0][5]

    put(dashboard, "B8", corridor)
    put(dashboard, "F8", direction)
    excel.CalculateFullRebuild()

    for column, vehicle in (("L", "EV"), ("M", "Diesel")):
        expected = sum(
            row[3] == vehicle
            and row[4] == corridor
            and row[5] == direction
            for row in records
        )

        require(
            calc.Range(f"{column}21").Value2 == expected,
            f"Corridor/direction filter failed for {vehicle}.",
        )

    put(dashboard, "B8", "All")
    put(dashboard, "F8", "All")
    put(dashboard, "M8", "Unique trips")
    excel.CalculateFullRebuild()

    seen = set()
    counts = {"EV": 0, "Diesel": 0}

    for row in records:
        key = (row[2], row[3], row[6], row[5])

        if key not in seen:
            counts[row[3]] += 1
            seen.add(key)

    require(
        calc.Range("L21").Value2 == counts["EV"]
        and calc.Range("M21").Value2 == counts["Diesel"],
        "Unique-trip filter check failed.",
    )

    put(dashboard, "M8", "All pair instances")
    excel.CalculateFullRebuild()


# =====================================================================
# SAFE BUILD TRANSACTION
# =====================================================================

def build(summary_path):
    summary_path = summary_path.resolve()

    require(
        os.name == "nt",
        "Windows and desktop Excel are required.",
    )
    require(
        summary_path.is_file(),
        f"Summary not found:\n{summary_path}",
    )

    import pythoncom
    import win32com.client

    lock_file = summary_path.with_name(
        "~$" + summary_path.name
    )

    require(
        not lock_file.exists(),
        "Close the summary workbook in Excel before running.",
    )

    original_hash = file_hash(summary_path)

    descriptor, temporary_name = tempfile.mkstemp(
        prefix="tat_dashboard_working_",
        suffix=".xlsx",
        dir=summary_path.parent,
    )
    os.close(descriptor)

    temporary_path = Path(temporary_name)
    shutil.copy2(summary_path, temporary_path)

    excel = None
    workbook = None
    sheets = None
    window = None
    completed = False

    pythoncom.CoInitialize()

    try:
        excel = win32com.client.DispatchEx(
            "Excel.Application"
        )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AutomationSecurity = 3

        require(
            float(excel.Version) >= 16,
            "Excel 2016 or later is required.",
        )

        workbook = excel.Workbooks.Open(
            str(temporary_path),
            UpdateLinks=0,
            ReadOnly=False,
        )

        require(
            not workbook.ReadOnly,
            "Working workbook is read-only.",
        )
        require(
            not workbook.Date1904,
            "1904-date-system workbooks are unsupported.",
        )
        require(
            "01_All_Pairs" in sheet_names(workbook),
            "Missing 01_All_Pairs.",
        )

        excel.Calculation = -4135

        print(
            "[1/8] Reading source tables and existing manual inputs..."
        )
        settings, optional, routes = snapshot_inputs(workbook)
        records = extract_records(
            workbook, routes, optional
        )

        print(
            f"Validated {len(records)//2} pairs "
            f"and {len(records)} trips."
        )

        print("[2/8] Creating the four dashboard sheets...")
        sheets = recreate_sheets(workbook)

        last_trip = build_trip_table(
            sheets["Trip_Input"],
            records,
        )

        print(
            "[3/8] Building model assumptions and route master..."
        )
        build_model(
            workbook,
            sheets["Model"],
            settings,
            routes,
        )
        validate_parameters(sheets["Model"])

        print(
            "[4/8] Building dropdown lists and trip formulas..."
        )
        corridors, months = build_lists(
            workbook,
            sheets["Calc"],
            records,
            routes,
        )
        write_trip_formulas(
            sheets["Trip_Input"],
            last_trip,
        )

        print(
            "[5/8] Building fleet, pair and phase comparisons..."
        )
        month_bounds = build_calculations(
            sheets["Calc"],
            corridors,
            months,
        )
        pair_count = build_pair_calculations(
            sheets["Calc"],
            records,
        )
        build_dashboard(
            workbook,
            sheets["Dashboard"],
        )

        print(
            f"[6/8] Creating {EXPECTED_CHARTS} native Excel charts..."
        )
        excel.CalculateFullRebuild()

        build_charts(
            sheets["Dashboard"],
            corridors,
            month_bounds,
            last_trip,
            pair_count,
        )

        sheets["Trip_Input"].Tab.Color = EV_COLOR
        sheets["Model"].Tab.Color = DIESEL_COLOR

        for name in ("Trip_Input", "Model"):
            sheets[name].Activate()

            window = workbook.Windows.Item(1)
            window.DisplayGridlines = False
            window.FreezePanes = False

            sheets[name].Range(
                "H5" if name == "Trip_Input" else "D6"
            ).Select()

            window.FreezePanes = True

        sheets["Dashboard"].Activate()
        sheets["Dashboard"].Range("B8").Select()

        window = workbook.Windows.Item(1)
        window.DisplayGridlines = False
        window.FreezePanes = False
        window.Zoom = 80

        sheets["Calc"].Visible = 0

        excel.Calculation = -4105
        excel.CalculateFullRebuild()

        print(
            "[7/8] Checking formulas, TAT values and live filters..."
        )
        validate_workbook(workbook, records)
        check_live_filters(excel, workbook, records)
        validate_workbook(workbook, records)

        print("[8/8] Saving and creating the backup...")
        workbook.Save()
        workbook.Close(SaveChanges=False)

        window = None
        sheets = None
        workbook = None

        require(
            file_hash(summary_path) == original_hash,
            "The source changed during generation. "
            "It has not been replaced.",
        )
        require(
            not lock_file.exists(),
            "The source workbook was opened during generation. "
            "Close it and rerun.",
        )

        backup_folder = summary_path.parent / "Backups"
        backup_folder.mkdir(
            parents=True,
            exist_ok=True,
        )

        stamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S_%f"
        )
        backup_path = backup_folder / (
            f"{summary_path.stem}_before_dashboard_{stamp}.xlsx"
        )

        shutil.copy2(summary_path, backup_path)
        os.replace(temporary_path, summary_path)
        completed = True

        print("\n" + "=" * 78)
        print("DASHBOARD COMPLETED")
        print(f"Workbook: {summary_path}")
        print(f"Backup:   {backup_path}")
        print(
            "Sheets: Dashboard, Trip_Input, Model, Calc (hidden)"
        )
        print(f"Native charts: {EXPECTED_CHARTS}")
        print("No external template workbook is required.")
        print("=" * 78)

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

        window = None
        sheets = None
        workbook = None

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        excel = None
        gc.collect()
        pythoncom.CoUninitialize()

        if not completed and temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                print(
                    "Temporary file could not be removed:\n"
                    f"{temporary_path}",
                    file=sys.stderr,
                )


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Build four live TAT dashboard sheets "
            "with seventeen charts and no template."
        )
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=SUMMARY_PATH,
    )
    arguments = parser.parse_args()
    build(arguments.summary)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(
            "\nDashboard generation cancelled.",
            file=sys.stderr,
        )
        raise SystemExit(130)
    except Exception:
        print(
            "\nDASHBOARD BUILD FAILED\n",
            file=sys.stderr,
        )
        traceback.print_exc()
        raise SystemExit(1)