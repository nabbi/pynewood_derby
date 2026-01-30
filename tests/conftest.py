"""Shared fixtures for pynewood_derby tests."""

import pandas as pd
import pytest


@pytest.fixture
def sample_racers_df():
    """A minimal valid Racers DataFrame."""
    return pd.DataFrame(
        {
            "Car": [101, 102, 103, 104, 105, 106],
            "Name": ["Alice", "Bob", "Charlie", "Diana", "Eve", "Frank"],
            "Class": ["Tiger", "Tiger", "Tiger", "Bear", "Bear", "Bear"],
            "Group": ["A", "A", "A", "A", "A", "A"],
            "Description": ["Red", "Blue", "Green", "Yellow", "Black", "White"],
        }
    )


@pytest.fixture
def sample_heats_df():
    """A simple heats DataFrame with 3 heats, 4 lanes."""
    return pd.DataFrame(
        {
            "Heat": [1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3],
            "Car": ["A", "B", "C", "D", "A", "C", "D", "B", "A", "D", "B", "C"],
            "Lane": ["L1", "L2", "L3", "L4", "L2", "L1", "L3", "L4", "L3", "L1", "L2", "L4"],
        }
    )


@pytest.fixture
def sample_heat_sheet_df():
    """A heat sheet DataFrame with Place column (for results processing)."""
    return pd.DataFrame(
        {
            "Heat": [1, 1, 1, 2, 2, 2],
            "Car": ["101", "102", "103", "101", "103", "102"],
            "Name": ["Alice", "Bob", "Charlie", "Alice", "Charlie", "Bob"],
            "Lane": ["A", "B", "C", "B", "A", "C"],
            "Place": [1, 2, 3, 2, 1, 3],
        }
    )


@pytest.fixture
def racers_xlsx(tmp_path, sample_racers_df):
    """Create a temporary Excel workbook with a Racers sheet."""
    path = tmp_path / "racers.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sample_racers_df.to_excel(writer, sheet_name="Racers", index=False)
    return path


@pytest.fixture
def heats_xlsx(tmp_path, sample_racers_df, sample_heat_sheet_df):
    """Create a temporary Excel workbook with Racers and a heat sheet."""
    path = tmp_path / "heats.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sample_racers_df.to_excel(writer, sheet_name="Racers", index=False)
        sample_heat_sheet_df.to_excel(writer, sheet_name="Tiger_A", index=False)
    return path
