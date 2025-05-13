#!/usr/bin/env python3
"""
heats_runoff.py — Pinewood Derby Runoff Heat Generator (Round-Robin)

Generates head-to-head runoff heats for tied cars using full round-robin logic.
Reads from a manually created workbook containing only a 'Racers' tab.

Usage:
    python heats_runoff.py <racers_file.xlsx> [num_lanes]

Output:
    Updates the input file with:
    - 'Racers' tab (reused from original)
    - One heat tab per Class/Group (e.g. "Tiger_A")
"""

import sys
from itertools import combinations
import pandas as pd

from race_utils import (
    read_excel_sheet,
    sanitize_sheet_title,
    secure_shuffle,
    format_all_sheets,
    get_racer_heats,
    update_racer_heats,
)


def _generate_small_group_heats(cars, lane_labels):
    """
    Generate a set of race heats for a small group of cars, where the number
    of cars is less than or equal to the number of available lanes.

    Each car will run exactly once in each lane, with their positions rotating
    between heats to ensure that every car gets an opportunity to race in every lane.

    Parameters:
    ----------
    cars : list
        A list of car identifiers (e.g., numbers or names).

    lane_labels : list
        A list of lane labels (e.g., ["Lane 1", "Lane 2", ...]).
        The function will only use as many lanes as there are cars.

    Returns:
    -------
    pandas.DataFrame
        A DataFrame representing the scheduled heats. Each row contains:
            - "Heat": the heat number (starting from 1),
            - "Car": the car identifier,
            - "Lane": the lane the car is assigned to in that heat.

    Notes:
    ------
    - If the number of lanes exceeds the number of cars, only the first `len(cars)` lanes are used.
    - Each car rotates through each lane exactly once.

    Example:
    -------
    >>> _generate_small_group_heats(["CarA", "CarB", "CarC"], ["L1", "L2", "L3", "L4"])
       Heat   Car Lane
    0     1  CarA   L1
    1     1  CarB   L2
    2     1  CarC   L3
    3     2  CarB   L1
    4     2  CarC   L2
    5     2  CarA   L3
    6     3  CarC   L1
    7     3  CarA   L2
    8     3  CarB   L3
    """
    used_lanes = lane_labels[: len(cars)]
    car_rotations = [cars[i:] + cars[:i] for i in range(len(cars))]
    lane_zips = zip(*car_rotations)
    heats = []
    for heat_num, lane_rotation in enumerate(lane_zips, start=1):
        for car, lane in zip(lane_rotation, used_lanes):
            heats.append({"Heat": heat_num, "Car": car, "Lane": lane})
    return pd.DataFrame(heats)


def generate_round_robin_heats(cars, num_lanes):
    """
    Generate a round-robin heat schedule for a group of cars, assigning them to
    available lanes across multiple heats.

    If the number of cars is less than or equal to the number of lanes, each car
    runs once in each lane (handled by a helper function). Otherwise, cars are
    paired off in a round-robin style so that each pair races exactly once.

    Parameters:
    ----------
    cars : list
        A list of car identifiers (e.g., strings or integers).

    num_lanes : int
        Total number of available lanes for each heat.

    Returns:
    -------
    pandas.DataFrame
        A DataFrame where each row represents a car's participation in a heat.
        Columns include:
            - "Heat": the heat number,
            - "Car": the car identifier,
            - "Lane": the lane assigned for that heat.

    Logic Overview:
    ---------------
    - If the number of cars is small enough, use a fixed rotation where each car
      runs in each lane exactly once.
    - For larger groups, generate all unique car pairings.
    - In each heat:
        - Try to fit as many non-overlapping car pairs as possible, respecting
          the lane limit.
        - Assign each car in that heat to a shuffled lane (for fairness).
    - Continue until all unique pairings have raced at least once.

    Example:
    --------
    >>> generate_round_robin_heats(["C1", "C2", "C3", "C4"], 3)
       Heat Car Lane
    0     1  C1    A
    1     1  C2    C
    2     1  C3    B
    3     2  C1    B
    4     2  C4    A
    5     2  C2    C
    ...

    Notes:
    ------
    - Uses a helper `_generate_small_group_heats` for small groups.
    - Relies on a secure shuffle (assumed from `secure_shuffle()`) for fair and
      randomized lane assignments.
    - If no full pairing can be made in a heat (due to lane limits vs remaining cars),
      an arbitrary leftover pair is forcibly added to ensure all matchups are completed.
    """
    lane_labels = [chr(ord("A") + i) for i in range(num_lanes)]

    if len(cars) <= num_lanes:
        return _generate_small_group_heats(cars, lane_labels)

    matchups = {frozenset(pair) for pair in combinations(cars, 2)}
    unmatched = matchups.copy()
    heats = []
    heat_num = 1

    while unmatched:
        heat_cars = set()
        used_pairs = set()

        for pair in sorted(unmatched):
            car1, car2 = tuple(pair)
            if car1 in heat_cars or car2 in heat_cars:
                continue
            if len(heat_cars) + 2 > num_lanes:
                continue
            heat_cars.update([car1, car2])
            used_pairs.add(pair)

        if len(heat_cars) < 2:
            leftover = list(unmatched)[0]
            heat_cars.update(leftover)
            used_pairs.add(leftover)

        unmatched -= used_pairs
        assigned_lanes = secure_shuffle(lane_labels)[: len(heat_cars)]
        for car, lane in zip(sorted(heat_cars), assigned_lanes):
            heats.append({"Heat": heat_num, "Car": car, "Lane": lane})
        heat_num += 1

    return pd.DataFrame(heats)


def validate_runoff_heats(heats_df, expected_matchups, cars=None, num_lanes=None):
    """
    Validates a set of generated runoff heats against a set of race rules and expectations.

    This function performs multiple checks to ensure that:
    - Each heat has valid structure (no duplicate cars/lanes, enough participants).
    - For small groups (cars <= lanes), each car runs once in each lane.
    - All expected car pair matchups occur exactly once.

    Parameters:
    ----------
    heats_df : pd.DataFrame
        DataFrame containing heat results with columns: "Heat", "Car", "Lane".

    expected_matchups : set of frozenset
        Set of all expected unique car pairings (e.g., {frozenset(["CarA", "CarB"])}).

    cars : list, optional
        Full list of car identifiers involved in the heats. Required for small group validation.

    num_lanes : int, optional
        Number of lanes used in the track. Required for small group validation.

    Returns:
    -------
    bool
        True if all validations pass; False if any rule is violated.

    Notes:
    ------
    - Prints detailed error messages to help identify validation failures.
    - Only used when number of cars is less than or equal to the number of lanes.
    """

    def _validate_heat(group, heat_num):
        """
        Validates a single heat group by checking:
        - There are at least 2 cars in the heat.
        - No duplicate cars appear.
        - No lane is assigned to more than one car.

        Parameters:
        ----------
        group : pd.DataFrame
            A DataFrame slice corresponding to a single heat (grouped by "Heat").

        heat_num : int
            The number/index of the current heat (for logging purposes).

        Returns:
        -------
        bool
            True if this heat passes all validation checks; False otherwise.
        """
        ok = True
        if len(group) < 2:
            print(f"[ERROR] Heat {heat_num} has fewer than 2 cars")
            ok = False
        if group["Car"].duplicated().any():
            print(f"[ERROR] Duplicate cars in Heat {heat_num}")
            ok = False
        if group["Lane"].duplicated().any():
            print(f"[ERROR] Duplicate lanes in Heat {heat_num}")
            ok = False
        return ok

    def _validate_small_group(heats_df, cars):
        """
        Validates that for small groups (cars <= lanes), each car:
        - Runs once in every lane.
        - Appears in the expected number of heats.

        Parameters:
        ----------
        heats_df : pd.DataFrame
            The full heats DataFrame.

        cars : list
            List of all car identifiers.

        Returns:
        -------
        bool
            True if all small-group rules are satisfied; False otherwise.

        Notes:
        ------
        This is only used when the number of cars is less than or equal to the number of lanes.
        """
        ok = True
        unique_heats = heats_df["Heat"].nunique()
        if unique_heats != len(cars):
            print(f"[ERROR] Expected {len(cars)} heats, but got {unique_heats}")
            ok = False
        for car in cars:
            car_rows = heats_df[heats_df["Car"] == car]
            lanes_run = car_rows["Lane"].nunique()
            total_runs = car_rows.shape[0]
            if lanes_run != len(cars):
                print(
                    f"[ERROR] Car {car} ran in {lanes_run} unique lanes, expected {len(cars)}"
                )
                ok = False
            if total_runs != len(cars):
                print(f"[ERROR] Car {car} has {total_runs} heats, expected {len(cars)}")
                ok = False
        return ok

    def _check_matchups(heats_df, expected_matchups):
        """
        Checks that all expected car matchups (pairs) occurred across the heats.

        Parameters:
        ----------
        heats_df : pd.DataFrame
            The full heats DataFrame.

        expected_matchups : set of frozenset
            Set of all expected unique car pairings.

        Returns:
        -------
        bool
            True if all expected matchups occurred at least once; False otherwise.

        Notes:
        ------
        Prints a list of any missing matchups.
        """
        actual_matchups = set()
        for _, group in heats_df.groupby("Heat"):
            car_list = group["Car"].tolist()
            actual_matchups.update(
                frozenset(pair) for pair in combinations(car_list, 2)
            )
        missing = expected_matchups - actual_matchups
        if missing:
            print("[ERROR] Missing matchups:")
            for pair in missing:
                print(f"  {sorted(list(pair))}")
            return False
        return True

    valid = all(
        _validate_heat(group, heat_num) for heat_num, group in heats_df.groupby("Heat")
    )

    if cars and num_lanes and len(cars) <= num_lanes:
        valid = _validate_small_group(heats_df, cars) and valid

    if not _check_matchups(heats_df, expected_matchups):
        valid = False

    if valid:
        print("[OK] All runoff heat validations passed.")
    return valid


def process_class_group(writer, cls, grp, group_df, num_lanes):
    """
    Generate, validate, and export round-robin race heats for a specific class/group of cars.

    This function handles the full pipeline for a race group:
    - Randomly shuffles the list of participating cars.
    - Attempts to generate valid heats using round-robin logic (with a retry loop).
    - Validates the generated heats to ensure all rules are satisfied.
    - Maps car names and prepares final race sheet formatting.
    - Writes the result to a sheet in the provided Excel writer.

    Parameters:
    ----------
    writer : pandas.ExcelWriter
        An Excel writer instance used to output the resulting heats to a spreadsheet.

    cls : str
        The name of the class/category this group belongs to (used in sheet naming and logs).

    grp : str
        The name/identifier of the group within the class (used in sheet naming and logs).

    group_df : pandas.DataFrame
        DataFrame containing car data for the class/group. Expected columns:
        - "Car" (required): Unique car identifiers.
        - "Name" (optional): Optional human-readable names for the cars.

    num_lanes : int
        The number of lanes available on the track. This affects heat generation logic.

    Behavior:
    --------
    - If fewer than 2 cars are present, the group is skipped with a log message.
    - For valid groups, it tries up to 200 times to generate heats that pass validation.
    - The heat schedule includes Heat, Car, Name, Lane, and Place columns.
    - Small groups (cars <= lanes) use `_generate_small_group_heats`, larger groups use pairing logic.
    - Final output is sorted by heat and lane and written to an Excel sheet with a sanitized name.

    Sheet Naming:
    -------------
    The output sheet is named based on the class and group, e.g., "Stock_A" or "Open_General".

    Example:
    --------
    >>> process_class_group(writer, "Stock", "A", group_df, 4)
    [INFO] Generating round-robin heats for: Class 'Stock' / Group 'A'
    [INFO] Using standard heat logic
    [OK] Heats written to sheet: Stock_A

    Notes:
    ------
    - Depends on helper functions: `secure_shuffle`, `generate_round_robin_heats`,
      `validate_runoff_heats`, and `sanitize_sheet_title`.
    - Validation ensures no duplicate lanes or cars, complete matchups, and correct heat counts.
    """
    cars = secure_shuffle(group_df["Car"].dropna().tolist())
    name_map = dict(zip(group_df["Car"], group_df.get("Name", [""] * len(group_df))))

    if len(cars) < 2:
        print(f"[SKIP] Not enough cars for {cls} / {grp}")
        return

    print(
        f"[INFO] Generating round-robin heats for: Class '{cls}' / Group '{grp}'"
    )  # noqa: E501
    expected_matchups = set(frozenset(pair) for pair in combinations(cars, 2))

    for _ in range(200):
        heats_df = generate_round_robin_heats(cars, num_lanes)
        mode = "small" if len(cars) <= num_lanes else "standard"
        print(f"[INFO] Using {mode} heat logic")
        if validate_runoff_heats(
            heats_df, expected_matchups, cars=cars, num_lanes=num_lanes
        ):
            break
    else:
        print(f"[FAIL] Validation failed for {cls} / {grp} after 200 attempts")
        return

    heats_df["Name"] = heats_df["Car"].map(name_map)
    heats_df["Place"] = ""
    heats_df = heats_df[["Heat", "Car", "Name", "Lane", "Place"]]

    sheet_name = sanitize_sheet_title(f"{cls}_{'General' if pd.isna(grp) else grp}")
    heats_df.sort_values(by=["Heat", "Lane"]).to_excel(
        writer, sheet_name=sheet_name, index=False
    )
    print(f"[OK] Heats written to sheet: {sheet_name}")


def main():
    """
    Main entry point for the runoff heat generator script.
    Reads the Racers tab from an Excel file, generates round-robin heats
    per class/group, validates them, and appends them to the file.
    """
    if len(sys.argv) < 2:
        print("Usage: python heats_runoff.py <racers_file.xlsx> [num_lanes]")
        sys.exit(1)

    filename = sys.argv[1]
    num_lanes = int(sys.argv[2]) if len(sys.argv) > 2 else 4

    try:
        df = read_excel_sheet(filename, sheet_name="Racers")
    except ValueError as err:
        print(f"[ERROR] {err}")
        sys.exit(1)

    df["Car"] = df["Car"].astype(str)
    grouped = df.groupby(["Class", "Group"], dropna=False)

    with pd.ExcelWriter(
        filename, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name="Racers", index=False)
        for (cls, grp), group_df in grouped:
            process_class_group(writer, cls, grp, group_df, num_lanes)

    heats_data = get_racer_heats(filename)
    update_racer_heats(filename, heats_data)

    format_all_sheets(filename)
    print(f"[DONE] Runoff heats updated in: {filename}")


if __name__ == "__main__":
    main()
