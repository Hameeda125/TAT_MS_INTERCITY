"""Tests for Stage 4 phase segmentation using Python's built-in unittest.

Run from the project root with:

    python tests/test_phase_segmentation.py

or:

    python -m unittest tests.test_phase_segmentation -v

The tests use synthetic data only. They do not modify real pair workbooks,
pair indexes, trip logs, or the Results folder.
"""

from __future__ import annotations

import math
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from phase_segmentation import (  # noqa: E402
    ACCELERATION_COLUMN,
    HAVERSINE_DISTANCE_COLUMN,
    LATITUDE_COLUMN,
    LONGITUDE_COLUMN,
    SPEED_KMH_COLUMN,
    SPEED_MPS_COLUMN,
    SPEED_RANGE_COLUMN,
    TIMESTAMP_COLUMN,
    TIME_DIFFERENCE_COLUMN,
    REVERSE_GEOCODED_ADDRESS_COLUMN,
    StationEvidence,
    StationResolver,
    _classify_moving,
    _parse_datetime_value,
    _phase_order,
    _phase_rules,
    segment_trip,
)


def build_rules() -> dict[str, Any]:
    config = {
        "project": {
            "results_folder": "Results",
        },
        "phase_segmentation": {
            "idle_gap_min_seconds": 30.0,
            "stop_gap_seconds": 180.0,
            "stationary_displacement_m": 50.0,
            "station_radius_m": 100.0,
            "inching_speed_upper_kmh": 5.0,
            "cruise_acceleration_abs_mps2": 0.14,
            "manual_event_tolerance_seconds": 60.0,
            "zero_speed_tolerance_kmh": 0.01,
            "online_station_lookup": False,
            "strict_station_lookup": False,
            "time_tolerance_seconds": 0.001,
            "distance_tolerance_m": 0.001,
        },
    }
    return _phase_rules(config)


class FakeStationResolver:
    def resolve(
        self,
        latitude: float,
        longitude: float,
        vehicle_type: str,
    ) -> StationEvidence:
        near_station = (
            abs(latitude - 13.0005) < 0.001
            and abs(longitude - 80.0005) < 0.001
        )

        if near_station:
            return StationEvidence(
                address="Synthetic Test Station",
                station_matched=True,
                station_type="CHARGING" if vehicle_type == "EV" else "REFUELING",
                station_name="Synthetic Station",
                station_latitude=13.0005,
                station_longitude=80.0005,
                station_distance_m=5.0,
                evidence_source="VERIFIED_STATION_REGISTRY",
                lookup_status="MATCHED",
            )

        return StationEvidence(
            address="Synthetic Road",
            station_matched=False,
            evidence_source="TEST_RESOLVER",
            lookup_status="NO_COMPATIBLE_STATION_WITHIN_RADIUS",
        )


def make_feature_frame(
    timestamps: list[str],
    speeds_kmh: list[float],
    distances_m: list[float],
    accelerations_mps2: list[float],
    latitudes: list[float] | None = None,
    longitudes: list[float] | None = None,
) -> pd.DataFrame:
    parsed_timestamps = pd.to_datetime(timestamps)
    row_count = len(parsed_timestamps)

    if len(speeds_kmh) != row_count:
        raise ValueError("speeds_kmh length must equal timestamps length")
    if len(distances_m) != row_count:
        raise ValueError("distances_m length must equal timestamps length")
    if len(accelerations_mps2) != row_count:
        raise ValueError("accelerations_mps2 length must equal timestamps length")

    if latitudes is None:
        latitudes = [13.0 + index * 0.001 for index in range(row_count)]
    if longitudes is None:
        longitudes = [80.0 + index * 0.001 for index in range(row_count)]

    time_differences = (
        pd.Series(parsed_timestamps).diff().dt.total_seconds().fillna(0.0)
    )

    return pd.DataFrame(
        {
            TIMESTAMP_COLUMN: parsed_timestamps,
            LATITUDE_COLUMN: latitudes,
            LONGITUDE_COLUMN: longitudes,
            SPEED_MPS_COLUMN: np.asarray(speeds_kmh, dtype=float) / 3.6,
            SPEED_KMH_COLUMN: speeds_kmh,
            TIME_DIFFERENCE_COLUMN: time_differences,
            HAVERSINE_DISTANCE_COLUMN: distances_m,
            SPEED_RANGE_COLUMN: ["TEST"] * row_count,
            ACCELERATION_COLUMN: accelerations_mps2,
        }
    )


class TestPhaseSegmentation(unittest.TestCase):
    def setUp(self) -> None:
        self.rules = build_rules()

    def test_parse_datetime_value_returns_timestamp_or_none(self) -> None:
        parsed = _parse_datetime_value("06-08-2026 09:15:00")

        self.assertIsInstance(parsed, pd.Timestamp)
        self.assertEqual(parsed, pd.Timestamp("2026-08-06 09:15:00"))
        self.assertIsNone(_parse_datetime_value(""))
        self.assertIsNone(_parse_datetime_value("not a date"))

    def test_reverse_geocoded_address_column_name(self) -> None:
        self.assertEqual(
            REVERSE_GEOCODED_ADDRESS_COLUMN,
            "Reverse Geocoded Address",
        )

    def test_reverse_geocode_cache_key_is_vehicle_independent(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            resolver = StationResolver(
                self.rules,
                Path(temporary_directory),
            )
            try:
                first_key = resolver._reverse_cache_key(13.0827001, 80.2707001)
                second_key = resolver._reverse_cache_key(13.0827002, 80.2707002)

                # With the configured coordinate rounding, effectively identical
                # places reuse one reverse-geocoded address cache record.
                self.assertEqual(first_key, second_key)
                self.assertTrue(
                    str(resolver.reverse_cache_path).endswith(
                        "reverse_geocode_cache.csv"
                    )
                )
            finally:
                resolver.close()

    def test_phase_order_contains_exactly_seven_vehicle_phases(self) -> None:
        self.assertEqual(
            _phase_order("EV"),
            ["CHARGING", "STOP", "IDLE", "INCH", "ACC", "DEC", "CRUISE"],
        )
        self.assertEqual(
            _phase_order("ICE"),
            ["REFUELING", "STOP", "IDLE", "INCH", "ACC", "DEC", "CRUISE"],
        )

    def test_classify_moving(self) -> None:
        cases = [
            (2.0, 0.50, "INCH"),
            (20.0, 0.20, "ACC"),
            (20.0, -0.20, "DEC"),
            (20.0, 0.14, "CRUISE"),
            (20.0, -0.14, "CRUISE"),
            (20.0, 0.00, "CRUISE"),
            (20.0, math.nan, "CRUISE"),
        ]

        for speed_kmh, acceleration_mps2, expected_phase in cases:
            with self.subTest(
                speed_kmh=speed_kmh,
                acceleration_mps2=acceleration_mps2,
            ):
                actual = _classify_moving(
                    speed_kmh,
                    acceleration_mps2,
                    self.rules,
                )
                self.assertEqual(actual, expected_phase)

    def test_segment_trip_classifies_idle_stop_and_motion(self) -> None:
        frame = make_feature_frame(
            timestamps=[
                "2026-08-06 08:00:00",
                "2026-08-06 08:00:10",
                "2026-08-06 08:00:20",
                "2026-08-06 08:01:20",
                "2026-08-06 08:05:20",
                "2026-08-06 08:05:30",
                "2026-08-06 08:05:40",
            ],
            speeds_kmh=[0.0, 2.0, 20.0, 0.0, 0.0, 20.0, 20.0],
            distances_m=[0.0, 5.0, 50.0, 5.0, 5.0, 50.0, 50.0],
            accelerations_mps2=[np.nan, 0.0, 0.20, 0.0, 0.0, -0.20, 0.0],
        )

        intervals, summary, event_audit, metrics = segment_trip(
            frame,
            "EV",
            self.rules,
            trip_log_record={},
            station_resolver=None,
        )

        self.assertEqual(intervals.loc[0, "Phase"], "")
        self.assertEqual(intervals.loc[1, "Phase"], "INCH")
        self.assertEqual(intervals.loc[2, "Phase"], "ACC")
        self.assertEqual(intervals.loc[3, "Phase"], "IDLE")
        self.assertEqual(intervals.loc[4, "Phase"], "STOP")
        self.assertEqual(intervals.loc[5, "Phase"], "DEC")
        self.assertEqual(intervals.loc[6, "Phase"], "CRUISE")

        positive = intervals[TIME_DIFFERENCE_COLUMN] > 0
        self.assertFalse(intervals.loc[positive, "Phase"].eq("").any())

        elapsed = (
            pd.Timestamp("2026-08-06 08:05:40")
            - pd.Timestamp("2026-08-06 08:00:00")
        ).total_seconds()

        self.assertAlmostEqual(metrics["tat_seconds"], elapsed)
        total_time = summary.loc[
            summary["Phase"] == "TOTAL",
            "Time (s)",
        ].iloc[0]
        self.assertAlmostEqual(total_time, elapsed)
        self.assertTrue(event_audit.empty)

    def test_ev_stationary_station_interval_becomes_charging(self) -> None:
        frame = make_feature_frame(
            timestamps=[
                "2026-08-06 09:00:00",
                "2026-08-06 09:04:00",
            ],
            speeds_kmh=[0.0, 0.0],
            distances_m=[0.0, 5.0],
            accelerations_mps2=[np.nan, 0.0],
            latitudes=[13.0005, 13.0005],
            longitudes=[80.0005, 80.0005],
        )

        intervals, summary, _, metrics = segment_trip(
            frame,
            "EV",
            self.rules,
            trip_log_record={},
            station_resolver=FakeStationResolver(),
        )

        self.assertTrue(bool(intervals.loc[1, "Stationary Candidate"]))
        self.assertTrue(bool(intervals.loc[1, "Station Matched"]))
        self.assertEqual(intervals.loc[1, "Phase"], "CHARGING")
        self.assertEqual(
            intervals.loc[1, REVERSE_GEOCODED_ADDRESS_COLUMN],
            "Synthetic Test Station",
        )
        self.assertEqual(metrics["replenishment_intervals"], 1)

        charging_time = summary.loc[
            summary["Phase"] == "CHARGING",
            "Time (s)",
        ].iloc[0]
        self.assertAlmostEqual(charging_time, 240.0)

    def test_ice_stationary_station_interval_becomes_refueling(self) -> None:
        frame = make_feature_frame(
            timestamps=[
                "2026-08-06 10:00:00",
                "2026-08-06 10:04:00",
            ],
            speeds_kmh=[0.0, 0.0],
            distances_m=[0.0, 5.0],
            accelerations_mps2=[np.nan, 0.0],
            latitudes=[13.0005, 13.0005],
            longitudes=[80.0005, 80.0005],
        )

        intervals, summary, _, metrics = segment_trip(
            frame,
            "ICE",
            self.rules,
            trip_log_record={},
            station_resolver=FakeStationResolver(),
        )

        self.assertEqual(intervals.loc[1, "Phase"], "REFUELING")
        self.assertEqual(metrics["replenishment_intervals"], 1)

        refueling_time = summary.loc[
            summary["Phase"] == "REFUELING",
            "Time (s)",
        ].iloc[0]
        self.assertAlmostEqual(refueling_time, 240.0)

    def test_manual_stop_overrides_idle_when_timestamp_matches(self) -> None:
        frame = make_feature_frame(
            timestamps=[
                "2026-08-06 11:00:00",
                "2026-08-06 11:01:00",
            ],
            speeds_kmh=[0.0, 0.0],
            distances_m=[0.0, 5.0],
            accelerations_mps2=[np.nan, 0.0],
        )

        trip_log_record = {
            "log_date": "06-08-2026",
            "log_additional": "Stop at 11:00",
        }

        intervals, _, event_audit, metrics = segment_trip(
            frame,
            "EV",
            self.rules,
            trip_log_record=trip_log_record,
            station_resolver=None,
        )

        self.assertTrue(bool(intervals.loc[1, "Manual Stop Match"]))
        self.assertEqual(intervals.loc[1, "Phase"], "STOP")
        self.assertEqual(metrics["manual_events"], 1)
        self.assertEqual(metrics["manual_events_matched"], 1)
        self.assertTrue(
            str(event_audit.loc[0, "Match Status"]).startswith("MATCHED")
        )

    def test_missing_required_feature_column_raises_error(self) -> None:
        frame = make_feature_frame(
            timestamps=[
                "2026-08-06 12:00:00",
                "2026-08-06 12:00:10",
            ],
            speeds_kmh=[0.0, 10.0],
            distances_m=[0.0, 20.0],
            accelerations_mps2=[np.nan, 0.2],
        ).drop(columns=[ACCELERATION_COLUMN])

        with self.assertRaisesRegex(
            ValueError,
            "Missing Stage 3 columns",
        ):
            segment_trip(frame, "EV", self.rules)

    def test_time_difference_must_match_consecutive_timestamps(self) -> None:
        frame = make_feature_frame(
            timestamps=[
                "2026-08-06 13:00:00",
                "2026-08-06 13:00:10",
            ],
            speeds_kmh=[0.0, 10.0],
            distances_m=[0.0, 20.0],
            accelerations_mps2=[np.nan, 0.2],
        )

        frame.loc[1, TIME_DIFFERENCE_COLUMN] = 9.0

        with self.assertRaisesRegex(
            ValueError,
            r"Time Difference \(s\) does not match consecutive timestamps",
        ):
            segment_trip(frame, "EV", self.rules)


if __name__ == "__main__":
    unittest.main(verbosity=2)