"""Tests for results.py — results processing and ranking."""

import sys

import pandas as pd
import pytest

from results import (
    add_runoff_tab,
    calculate_opponent_uniqueness,
    get_cli_args,
    process_results,
)


# ---------------------------------------------------------------------------
# calculate_opponent_uniqueness
# ---------------------------------------------------------------------------
class TestCalculateOpponentUniqueness:
    def test_basic(self):
        df = pd.DataFrame(
            {
                "Heat": [1, 1, 1, 2, 2],
                "Car": ["A", "B", "C", "A", "B"],
            }
        )
        result = calculate_opponent_uniqueness(df)
        assert set(result.columns) == {"Car", "Opponent_Uniqueness_Pct"}
        # A has opponents {B, C} out of {B, C} = 100%
        a_pct = result[result["Car"] == "A"]["Opponent_Uniqueness_Pct"].iloc[0]
        assert a_pct == 100.0

    def test_partial_uniqueness(self):
        df = pd.DataFrame(
            {
                "Heat": [1, 1, 2, 2],
                "Car": ["A", "B", "A", "C"],
            }
        )
        result = calculate_opponent_uniqueness(df)
        # A sees B (heat 1) and C (heat 2) → 100% of 2 possible
        a_pct = result[result["Car"] == "A"]["Opponent_Uniqueness_Pct"].iloc[0]
        assert a_pct == 100.0
        # B sees only A → 50% of {A, C}
        b_pct = result[result["Car"] == "B"]["Opponent_Uniqueness_Pct"].iloc[0]
        assert b_pct == 50.0

    def test_single_car(self):
        df = pd.DataFrame({"Heat": [1], "Car": ["A"]})
        result = calculate_opponent_uniqueness(df)
        assert result[result["Car"] == "A"]["Opponent_Uniqueness_Pct"].iloc[0] == 0.0


# ---------------------------------------------------------------------------
# process_results (integration-ish, uses tmp Excel)
# ---------------------------------------------------------------------------
class TestProcessResults:
    def test_creates_rankings_sheet(self, heats_xlsx):
        process_results(str(heats_xlsx))
        sheets = pd.ExcelFile(str(heats_xlsx)).sheet_names
        assert "Tiger_A_Rankings" in sheets

    def test_rankings_have_rank_column(self, heats_xlsx):
        process_results(str(heats_xlsx))
        df = pd.read_excel(str(heats_xlsx), sheet_name="Tiger_A_Rankings")
        assert "Rank" in df.columns
        assert "Total_Points" in df.columns


# ---------------------------------------------------------------------------
# add_runoff_tab
# ---------------------------------------------------------------------------
class TestAddRunoffTab:
    def test_no_ties_no_runoff(self, heats_xlsx):
        process_results(str(heats_xlsx))
        add_runoff_tab(str(heats_xlsx))
        sheets = pd.ExcelFile(str(heats_xlsx)).sheet_names
        # With 3 cars and no ties, no runoff tab
        assert "Runoff" not in sheets

    def test_ties_create_runoff(self, tmp_path):
        """Create a scenario with tied rankings, verify Runoff tab is created."""
        path = tmp_path / "tied.xlsx"
        racers = pd.DataFrame(
            {
                "Car": [1, 2, 3],
                "Name": ["A", "B", "C"],
                "Class": ["T", "T", "T"],
                "Group": ["G", "G", "G"],
                "Description": ["d1", "d2", "d3"],
            }
        )
        # Create rankings where cars 1 and 2 tie at rank 1
        rankings = pd.DataFrame(
            {
                "Car": [1, 2, 3],
                "Name": ["A", "B", "C"],
                "Total_Points": [3, 3, 6],
                "First_Place_Count": [1, 1, 0],
                "Heat_Count": [3, 3, 3],
                "Avg_Heat_Size": [3.0, 3.0, 3.0],
                "Heat_Size_Pct": [100.0, 100.0, 100.0],
                "Opponent_Uniqueness_Pct": [100.0, 100.0, 100.0],
                "Rank": [1, 1, 3],
            }
        )
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            racers.to_excel(writer, sheet_name="Racers", index=False)
            rankings.to_excel(writer, sheet_name="T_G_Rankings", index=False)

        add_runoff_tab(str(path))
        sheets = pd.ExcelFile(str(path)).sheet_names
        assert "Runoff" in sheets


# ---------------------------------------------------------------------------
# get_cli_args
# ---------------------------------------------------------------------------
class TestResultsCliArgs:
    def test_returns_filename(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["results.py", "data.xlsx"])
        assert get_cli_args() == "data.xlsx"

    def test_no_args_exits(self, monkeypatch):
        monkeypatch.setattr(sys, "argv", ["results.py"])
        with pytest.raises(SystemExit):
            get_cli_args()
