"""Tests for heats_perfect.py — Perfect-N heat generation."""

import sys

import pandas as pd
import pytest

from heats_perfect import generate_heats, get_cli_args, validate_heats


# ---------------------------------------------------------------------------
# validate_heats
# ---------------------------------------------------------------------------
class TestValidateHeats:
    def _make_heats(self, rows):
        return pd.DataFrame(rows, columns=["Heat", "Car", "Lane"])

    def test_valid_heats(self):
        df = self._make_heats(
            [
                (1, "A", "L1"),
                (1, "B", "L2"),
                (2, "A", "L2"),
                (2, "B", "L1"),
            ]
        )
        assert validate_heats(df, expected_races_per_car=2, all_cars={"A", "B"}) is True

    def test_wrong_race_count(self):
        df = self._make_heats(
            [
                (1, "A", "L1"),
                (1, "B", "L2"),
                (2, "A", "L2"),
            ]
        )
        assert validate_heats(df, expected_races_per_car=2, all_cars={"A", "B"}) is False

    def test_duplicate_car_lane(self):
        df = self._make_heats(
            [
                (1, "A", "L1"),
                (1, "B", "L2"),
                (2, "A", "L1"),  # A in L1 again
                (2, "B", "L2"),
            ]
        )
        assert validate_heats(df, expected_races_per_car=2) is False

    def test_missing_car(self):
        df = self._make_heats(
            [
                (1, "A", "L1"),
                (1, "B", "L2"),
            ]
        )
        assert validate_heats(df, expected_races_per_car=1, all_cars={"A", "B", "C"}) is False

    def test_duplicate_car_in_heat(self):
        df = self._make_heats(
            [
                (1, "A", "L1"),
                (1, "A", "L2"),
            ]
        )
        assert validate_heats(df, expected_races_per_car=2) is False

    def test_heat_too_small(self):
        df = self._make_heats([(1, "A", "L1")])
        assert validate_heats(df, expected_races_per_car=1, min_cars_per_heat=2) is False


# ---------------------------------------------------------------------------
# generate_heats
# ---------------------------------------------------------------------------
class TestGenerateHeats:
    def test_basic_generation(self):
        cars = [f"Car{i}" for i in range(6)]
        df = generate_heats(cars, num_lanes=4, runs_per_car=4)
        assert set(df.columns) >= {"Heat", "Car", "Lane"}
        # All cars should appear in the result
        assert set(df["Car"].unique()) == set(cars)
        # Each car races at least 2 times (rebalancing may reduce some)
        for car in cars:
            assert len(df[df["Car"] == car]) >= 2

    def test_no_duplicate_car_lane(self):
        cars = [f"C{i}" for i in range(8)]
        df = generate_heats(cars, num_lanes=4, runs_per_car=4)
        dupes = df.groupby(["Car", "Lane"]).size()
        assert (dupes > 1).sum() == 0

    def test_no_duplicate_car_in_heat(self):
        cars = [f"C{i}" for i in range(8)]
        df = generate_heats(cars, num_lanes=4, runs_per_car=4)
        assert not df.groupby(["Heat", "Car"]).size().gt(1).any()

    def test_min_two_cars_per_heat(self):
        cars = [f"C{i}" for i in range(6)]
        df = generate_heats(cars, num_lanes=3, runs_per_car=3)
        heat_sizes = df.groupby("Heat").size()
        assert (heat_sizes >= 2).all()

    def test_runs_per_car_clamped(self):
        """runs_per_car=1 should be clamped to 2 internally."""
        cars = ["A", "B", "C", "D", "E"]
        df = generate_heats(cars, num_lanes=4, runs_per_car=1)
        # All cars should appear; rebalancing may shift counts
        assert set(df["Car"].unique()) == set(cars)
        # No car should exceed 2 races (the clamped value)
        for car in cars:
            assert len(df[df["Car"] == car]) <= 2


# ---------------------------------------------------------------------------
# get_cli_args
# ---------------------------------------------------------------------------
class TestGetCliArgs:
    def test_minimal_args(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["heats_perfect.py", "race.xlsx"])
        filename, lanes, runs = get_cli_args()
        assert filename == "race.xlsx"
        assert lanes == 4
        assert runs == 4

    def test_custom_lanes_and_runs(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["heats_perfect.py", "race.xlsx", "6", "3"])
        filename, lanes, runs = get_cli_args()
        assert lanes == 6
        assert runs == 3

    def test_no_args_exits(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["heats_perfect.py"])
        with pytest.raises(SystemExit):
            get_cli_args()

    def test_invalid_lanes_exits(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["heats_perfect.py", "race.xlsx", "10"])
        with pytest.raises(SystemExit):
            get_cli_args()
