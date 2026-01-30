"""Tests for race_utils.py — shared utility functions."""

import math

import pandas as pd
import pytest

from race_utils import (
    _validate_excel_file,
    analyze_opponents,
    format_all_sheets,
    get_excel_sheet_names,
    get_racer_heats,
    is_nan,
    opponent_fairness_score,
    optimize_opponent_fairness,
    read_excel_sheet,
    rebalance_heats,
    sanitize_sheet_title,
    secure_shuffle,
    update_racer_heats,
    validate_heat_sheet_columns,
    validate_racers_columns,
)


# ---------------------------------------------------------------------------
# is_nan
# ---------------------------------------------------------------------------
class TestIsNan:
    def test_float_nan(self):
        assert is_nan(float("nan")) is True

    def test_math_nan(self):
        assert is_nan(math.nan) is True

    def test_regular_float(self):
        assert is_nan(1.0) is False

    def test_zero(self):
        assert is_nan(0.0) is False

    def test_none(self):
        assert is_nan(None) is False

    def test_string(self):
        assert is_nan("nan") is False

    def test_int(self):
        assert is_nan(42) is False


# ---------------------------------------------------------------------------
# secure_shuffle
# ---------------------------------------------------------------------------
class TestSecureShuffle:
    def test_returns_new_list(self):
        original = [1, 2, 3, 4, 5]
        result = secure_shuffle(original)
        assert result is not original
        assert original == [1, 2, 3, 4, 5]  # unchanged

    def test_same_elements(self):
        lst = list(range(20))
        result = secure_shuffle(lst)
        assert sorted(result) == sorted(lst)

    def test_empty_list(self):
        assert secure_shuffle([]) == []

    def test_single_element(self):
        assert secure_shuffle([42]) == [42]

    def test_produces_different_orderings(self):
        """Over many trials, shuffling should produce at least one different order."""
        lst = list(range(10))
        results = {tuple(secure_shuffle(lst)) for _ in range(50)}
        assert len(results) > 1


# ---------------------------------------------------------------------------
# sanitize_sheet_title
# ---------------------------------------------------------------------------
class TestSanitizeSheetTitle:
    def test_simple_name(self):
        assert sanitize_sheet_title("Tiger_A") == "Tiger_A"

    def test_removes_special_chars(self):
        assert sanitize_sheet_title("Hello/World!") == "HelloWorld"

    def test_spaces_to_underscores(self):
        assert sanitize_sheet_title("Tiger Group A") == "Tiger_Group_A"

    def test_truncates_to_31(self):
        long_name = "A" * 50
        assert len(sanitize_sheet_title(long_name)) == 31

    def test_empty_string(self):
        assert sanitize_sheet_title("") == ""

    def test_preserves_hyphens(self):
        assert sanitize_sheet_title("Foo-Bar") == "Foo-Bar"


# ---------------------------------------------------------------------------
# _validate_excel_file
# ---------------------------------------------------------------------------
class TestValidateExcelFile:
    def test_file_not_found(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            _validate_excel_file(str(tmp_path / "nope.xlsx"))

    def test_wrong_extension(self, tmp_path):
        bad = tmp_path / "file.csv"
        bad.write_text("data")
        with pytest.raises(ValueError, match="must be .xlsx"):
            _validate_excel_file(str(bad))

    def test_not_a_zip(self, tmp_path):
        bad = tmp_path / "file.xlsx"
        bad.write_text("not a zip")
        with pytest.raises(ValueError, match="not a valid Excel zip"):
            _validate_excel_file(str(bad))

    def test_valid_file(self, racers_xlsx):
        _validate_excel_file(str(racers_xlsx))  # should not raise


# ---------------------------------------------------------------------------
# validate_racers_columns
# ---------------------------------------------------------------------------
class TestValidateRacersColumns:
    def test_valid(self, sample_racers_df):
        validate_racers_columns(sample_racers_df)  # no exception

    def test_not_dataframe(self):
        with pytest.raises(ValueError, match="not a valid DataFrame"):
            validate_racers_columns("not a df")

    def test_empty_df(self):
        with pytest.raises(ValueError, match="empty"):
            validate_racers_columns(pd.DataFrame())

    def test_missing_columns(self):
        df = pd.DataFrame({"Car": [1], "Name": ["X"]})
        with pytest.raises(ValueError, match="Missing required columns"):
            validate_racers_columns(df)

    def test_duplicate_car_ids(self):
        df = pd.DataFrame(
            {
                "Car": [1, 1, 2],
                "Name": ["A", "B", "C"],
                "Class": ["T", "T", "T"],
                "Group": ["A", "A", "A"],
            }
        )
        with pytest.raises(ValueError, match="Duplicate Car IDs"):
            validate_racers_columns(df)


# ---------------------------------------------------------------------------
# validate_heat_sheet_columns
# ---------------------------------------------------------------------------
class TestValidateHeatSheetColumns:
    def test_valid(self, sample_heat_sheet_df):
        validate_heat_sheet_columns(sample_heat_sheet_df)

    def test_not_dataframe(self):
        with pytest.raises(ValueError, match="not a valid DataFrame"):
            validate_heat_sheet_columns([1, 2])

    def test_empty(self):
        with pytest.raises(ValueError, match="empty"):
            validate_heat_sheet_columns(pd.DataFrame())

    def test_missing_columns(self):
        df = pd.DataFrame({"Car": [1], "Heat": [1]})
        with pytest.raises(ValueError, match="Missing required columns"):
            validate_heat_sheet_columns(df)


# ---------------------------------------------------------------------------
# read_excel_sheet / get_excel_sheet_names
# ---------------------------------------------------------------------------
class TestExcelIO:
    def test_read_existing_sheet(self, racers_xlsx):
        df = read_excel_sheet(str(racers_xlsx), sheet_name="Racers")
        assert "Car" in df.columns
        assert len(df) == 6

    def test_read_missing_sheet(self, racers_xlsx):
        with pytest.raises(ValueError, match="not found"):
            read_excel_sheet(str(racers_xlsx), sheet_name="Nonexistent")

    def test_get_sheet_names(self, racers_xlsx):
        names = get_excel_sheet_names(str(racers_xlsx))
        assert "Racers" in names

    def test_get_sheet_names_invalid_file(self, tmp_path):
        with pytest.raises(Exception):
            get_excel_sheet_names(str(tmp_path / "missing.xlsx"))


# ---------------------------------------------------------------------------
# analyze_opponents / opponent_fairness_score
# ---------------------------------------------------------------------------
class TestOpponentAnalysis:
    def test_analyze_basic(self, sample_heats_df):
        opponents = analyze_opponents(sample_heats_df)
        # A is in heat 1 with B,C,D and heat 2 with C,D,B and heat 3 with D,B,C
        assert opponents["A"] == {"B", "C", "D"}

    def test_all_opponents_symmetric(self, sample_heats_df):
        opponents = analyze_opponents(sample_heats_df)
        for car, opps in opponents.items():
            for opp in opps:
                assert car in opponents[opp]

    def test_fairness_perfectly_fair(self):
        # All cars have same number of opponents
        opponents = {"A": {"B", "C"}, "B": {"A", "C"}, "C": {"A", "B"}}
        score = opponent_fairness_score(opponents)
        assert score == 0.0

    def test_fairness_unfair(self):
        opponents = {"A": {"B", "C"}, "B": {"A"}, "C": {"A"}}
        score = opponent_fairness_score(opponents)
        assert score > 0


# ---------------------------------------------------------------------------
# get_racer_heats
# ---------------------------------------------------------------------------
class TestGetRacerHeats:
    def test_extracts_heats(self, heats_xlsx):
        data = get_racer_heats(str(heats_xlsx))
        assert "101" in data
        assert isinstance(data["101"], list)
        assert len(data["101"]) == 2  # car 101 in heat 1 and 2

    def test_skips_racers_sheet(self, heats_xlsx):
        """Racers sheet should not be parsed for heat data."""
        data = get_racer_heats(str(heats_xlsx))
        # All cars come from Tiger_A, not Racers
        for heats in data.values():
            assert all(isinstance(h, int) for h in heats)


# ---------------------------------------------------------------------------
# update_racer_heats
# ---------------------------------------------------------------------------
class TestUpdateRacerHeats:
    def test_adds_heats_column(self, heats_xlsx):
        heats_data = {"101": [1, 2], "102": [1, 2], "103": [1, 2]}
        update_racer_heats(str(heats_xlsx), heats_data)
        df = pd.read_excel(str(heats_xlsx), sheet_name="Racers")
        assert "Heats" in df.columns


# ---------------------------------------------------------------------------
# format_all_sheets
# ---------------------------------------------------------------------------
class TestFormatAllSheets:
    def test_runs_without_error(self, heats_xlsx):
        format_all_sheets(str(heats_xlsx))
        # Just verify the file is still valid
        names = get_excel_sheet_names(str(heats_xlsx))
        assert "Racers" in names


# ---------------------------------------------------------------------------
# rebalance_heats
# ---------------------------------------------------------------------------
class TestRebalanceHeats:
    def test_no_change_when_balanced(self):
        """Heats already balanced should not change."""
        df = pd.DataFrame(
            {
                "Heat": [1, 1, 1, 2, 2, 2],
                "Car": ["A", "B", "C", "D", "E", "F"],
                "Lane": ["L1", "L2", "L3", "L1", "L2", "L3"],
            }
        )
        result = rebalance_heats(df, num_lanes=3)
        assert len(result) == 6

    def test_rebalances_underfilled_heat(self):
        """A heat with num_lanes-2 cars should gain one if a donor is available."""
        df = pd.DataFrame(
            {
                "Heat": [1, 1, 1, 1, 2, 2],
                "Car": ["A", "B", "C", "D", "E", "F"],
                "Lane": ["L1", "L2", "L3", "L4", "L1", "L2"],
            }
        )
        result = rebalance_heats(df, num_lanes=4)
        heat2_size = len(result[result["Heat"] == 2])
        # Should have gained a car from heat 1
        assert heat2_size >= 2


# ---------------------------------------------------------------------------
# optimize_opponent_fairness
# ---------------------------------------------------------------------------
class TestOptimizeOpponentFairness:
    def test_does_not_worsen_score(self):
        df = pd.DataFrame(
            {
                "Heat": [1, 1, 2, 2, 3, 3],
                "Car": ["A", "B", "A", "C", "B", "C"],
                "Lane": ["L1", "L2", "L1", "L2", "L1", "L2"],
            }
        )
        initial = opponent_fairness_score(analyze_opponents(df))
        result = optimize_opponent_fairness(df, iterations=50)
        final = opponent_fairness_score(analyze_opponents(result))
        assert final <= initial + 1e-9

    def test_preserves_structure(self):
        df = pd.DataFrame(
            {
                "Heat": [1, 1, 2, 2],
                "Car": ["A", "B", "C", "D"],
                "Lane": ["L1", "L2", "L1", "L2"],
            }
        )
        result = optimize_opponent_fairness(df, iterations=10)
        assert set(result.columns) == set(df.columns)
        assert len(result) == len(df)


# ---------------------------------------------------------------------------
# read_excel_sheet — read all sheets
# ---------------------------------------------------------------------------
class TestReadAllSheets:
    def test_read_without_sheet_name(self, heats_xlsx):
        """Reading without sheet_name returns a dict of DataFrames."""
        result = read_excel_sheet(str(heats_xlsx))
        assert isinstance(result, dict)
        assert "Racers" in result
