"""Tests for heats_runoff.py — Round-robin runoff heat generation."""

import sys
from itertools import combinations

import pandas as pd
import pytest

from heats_runoff import (
    _generate_small_group_heats,
    generate_round_robin_heats,
    get_cli_args,
    validate_runoff_heats,
)


# ---------------------------------------------------------------------------
# _generate_small_group_heats
# ---------------------------------------------------------------------------
class TestGenerateSmallGroupHeats:
    def test_each_car_in_each_lane(self):
        cars = ["A", "B", "C"]
        lanes = ["L1", "L2", "L3", "L4"]
        df = _generate_small_group_heats(cars, lanes)

        for car in cars:
            car_rows = df[df["Car"] == car]
            assert car_rows["Lane"].nunique() == len(cars)
            assert len(car_rows) == len(cars)

    def test_correct_heat_count(self):
        cars = ["A", "B"]
        lanes = ["L1", "L2"]
        df = _generate_small_group_heats(cars, lanes)
        assert df["Heat"].nunique() == len(cars)

    def test_no_duplicate_car_in_heat(self):
        cars = ["X", "Y", "Z"]
        df = _generate_small_group_heats(cars, ["A", "B", "C"])
        for _, group in df.groupby("Heat"):
            assert not group["Car"].duplicated().any()

    def test_no_duplicate_lane_in_heat(self):
        cars = ["X", "Y", "Z"]
        df = _generate_small_group_heats(cars, ["A", "B", "C"])
        for _, group in df.groupby("Heat"):
            assert not group["Lane"].duplicated().any()


# ---------------------------------------------------------------------------
# generate_round_robin_heats
# ---------------------------------------------------------------------------
class TestGenerateRoundRobinHeats:
    def test_small_group_delegates(self):
        """When cars <= lanes, should use small-group logic."""
        cars = ["A", "B", "C"]
        df = generate_round_robin_heats(cars, num_lanes=4)
        # Small group: each car races in each lane
        for car in cars:
            assert df[df["Car"] == car]["Lane"].nunique() == len(cars)

    def test_all_matchups_covered(self):
        cars = ["A", "B", "C", "D", "E"]
        df = generate_round_robin_heats(cars, num_lanes=4)
        expected = {frozenset(pair) for pair in combinations(cars, 2)}
        actual = set()
        for _, group in df.groupby("Heat"):
            actual.update(frozenset(p) for p in combinations(group["Car"].tolist(), 2))
        assert expected.issubset(actual)

    def test_no_duplicate_car_in_heat(self):
        cars = ["A", "B", "C", "D", "E"]
        df = generate_round_robin_heats(cars, num_lanes=3)
        for _, group in df.groupby("Heat"):
            assert not group["Car"].duplicated().any()

    def test_lane_limit_respected(self):
        cars = ["A", "B", "C", "D", "E", "F"]
        df = generate_round_robin_heats(cars, num_lanes=4)
        for _, group in df.groupby("Heat"):
            assert len(group) <= 4


# ---------------------------------------------------------------------------
# validate_runoff_heats
# ---------------------------------------------------------------------------
class TestValidateRunoffHeats:
    def test_valid_small_group(self):
        cars = ["A", "B", "C"]
        df = _generate_small_group_heats(cars, ["L1", "L2", "L3"])
        matchups = {frozenset(p) for p in combinations(cars, 2)}
        assert validate_runoff_heats(df, matchups, cars=cars, num_lanes=3) is True

    def test_missing_matchup(self):
        df = pd.DataFrame(
            {
                "Heat": [1, 1],
                "Car": ["A", "B"],
                "Lane": ["L1", "L2"],
            }
        )
        matchups = {frozenset(("A", "B")), frozenset(("A", "C"))}
        assert validate_runoff_heats(df, matchups) is False


# ---------------------------------------------------------------------------
# get_cli_args
# ---------------------------------------------------------------------------
class TestRunoffCliArgs:
    def test_minimal(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["heats_runoff.py", "file.xlsx"])
        filename, lanes = get_cli_args()
        assert filename == "file.xlsx"
        assert lanes == 4

    def test_no_args_exits(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["heats_runoff.py"])
        with pytest.raises(SystemExit):
            get_cli_args()
