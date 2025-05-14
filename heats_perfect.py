#!/usr/bin/env python3
"""
heats_perfect.py — Pinewood Derby Runoff Heat Generator (Perfect-N)

This script generates fair and valid heats for a Pinewood Derby race. It ensures:
- Each car races a set number of times (default: once per lane)
- No car repeats a lane
- No car or lane repeats within a heat
- Each heat has a minimum number of cars
- Matchups between cars are fairly distributed

Usage:
    python heats_perfect.py <racers_file.xlsx> [num_lanes] [runs_per_car]
"""

import sys
import pandas as pd
from race_utils import (
    read_excel_sheet,
    is_nan,
    secure_shuffle,
    rebalance_heats,
    optimize_opponent_fairness,
    sanitize_sheet_title,
    get_racer_heats,
    update_racer_heats,
    format_all_sheets,
    validate_unique_car_ids,
)


def validate_heats(
    heat_df: pd.DataFrame,
    expected_races_per_car: int,
    min_cars_per_heat: int = 2,
    all_cars: set | list | None = None,
) -> bool:
    """
    Validates a set of heats for fairness, consistency, and completeness
    based on several configurable criteria.

    This function checks the following:
    - Each car appears in the correct number of heats (expected_races_per_car).
    - No car is assigned to the same lane more than once.
    - No car appears more than once in a single heat.
    - No lane is assigned to more than one car in a single heat.
    - Each heat has at least the minimum required number of cars.
    - All expected cars (if provided) appear in the heat data.

    Parameters:
    ----------
    heat_df : pd.DataFrame
        A DataFrame containing race heat assignments.
        Expected columns:
        - "Heat": Identifier for each heat.
        - "Car": Identifier for each car.
        - "Lane": The lane each car is assigned to in that heat.

    expected_races_per_car : int
        The expected number of heats (races) that each car should participate in.

    min_cars_per_heat : int, optional (default=2)
        The minimum number of cars required in each heat.
        Any heat with fewer cars is considered invalid.

    all_cars : set | list | None, optional
        The complete set or list of car IDs that should appear in the heats.
        If provided, the function checks for any missing cars.

    Returns:
    -------
    bool
        True if all validations pass; False otherwise.

    Notes:
    ------
    - Prints detailed error messages for each type of validation failure.
    - Returns False as soon as any validation rule fails.
    - This function is suitable for enforcing fairness in race scheduling, such as
      pinewood derby heats, robotics trials, or any other round-robin competition.

    Example:
    --------
    >>> validate_heats(df, expected_races_per_car=3, all_cars={"A", "B", "C"})
    [OK] All heat validations passed.
    """
    valid = True

    car_lane_counts = heat_df.groupby(["Car", "Lane"]).size()
    duplicates = car_lane_counts[car_lane_counts > 1]
    if not duplicates.empty:
        print("[ERROR] Duplicate (Car, Lane) combinations:")
        for (car, lane), count in duplicates.items():
            print(f"  Car {car}, Lane {lane} → {count} times")
        valid = False

    total_races = heat_df.groupby("Car").size()
    invalid_race_counts = total_races[total_races != expected_races_per_car]
    if not invalid_race_counts.empty:
        print("[ERROR] Incorrect race counts:")
        for car, count in invalid_race_counts.items():
            print(f"  Car {car} raced {count} times")
        valid = False

    if all_cars is not None:
        missing_cars = set(all_cars) - set(heat_df["Car"].unique())
        if missing_cars:
            print("[ERROR] Missing cars from heats:")
            for car in sorted(missing_cars):
                print(f"  Car {car}")
            valid = False

    if heat_df.groupby(["Heat", "Car"]).size().gt(1).any():
        print("[ERROR] Duplicate cars found within heats")
        valid = False
    if heat_df.groupby(["Heat", "Lane"]).size().gt(1).any():
        print("[ERROR] Duplicate lanes found within heats")
        valid = False

    if heat_df.groupby("Heat").size().lt(min_cars_per_heat).any():
        print(f"[ERROR] Some heats have fewer than {min_cars_per_heat} cars")
        valid = False

    if valid:
        print("[OK] All heat validations passed.")
    return valid


def generate_heats(entry_list, num_lanes=4, runs_per_car=None) -> pd.DataFrame:
    """
    Generates race heats using a Perfect-N style lane assignment strategy with
    additional fairness optimizations.

    This function ensures that:
    - Each car runs in multiple lanes (based on `runs_per_car` or default `num_lanes`).
    - No car or lane is repeated in a single heat.
    - Each heat contains up to `num_lanes` cars,
      and only if at least 2 cars are assignable does the heat qualify as valid.
    - Heats are further optimized to improve fairness of opponent distribution.

    Parameters:
    ----------
    entry_list : list
        A list of unique car identifiers (strings, ints, etc.).

    num_lanes : int, optional (default=4)
        The number of lanes available in each heat (max 8, due to lane labeling).

    runs_per_car : int | None, optional
        The number of times each car should race.
        If None, defaults to `num_lanes`.
        Clamped between 2 and `num_lanes` for practical lane diversity.

    Returns:
    -------
    pd.DataFrame
        A flattened DataFrame representing the heat schedule with columns:
        - "Heat": Heat number (starting from 1)
        - "Car": Car identifier
        - "Lane": Lane label (e.g., "A", "B", ...)

    Processing Steps:
    -----------------
    1. Randomly assign each car to a shuffled subset of lanes.
    2. Build a pool of car-lane assignments (run pool).
    3. Iteratively form heats by selecting non-conflicting car-lane pairs.
       - No car or lane repeats within the same heat.
       - Skipped runs are retried in later heats.
    4. Only heats with 2 or more cars are considered valid.
    5. Resulting heat list is rebalanced and optimized:
       - `rebalance_heats`: Improves distribution of car appearances.
       - `optimize_opponent_fairness`: Attempts to balance car matchups across heats.

    Notes:
    ------
    - Depends on helper functions `secure_shuffle`, `rebalance_heats`,
      and `optimize_opponent_fairness`.
    - Lane labels are drawn from ["A" to "H"], clamped by `num_lanes`.
    - Fairness is approximate; due to randomness and constraints, perfect balance is not guaranteed.

    Example:
    --------
    >>> generate_heats(["Car1", "Car2", "Car3", "Car4", "Car5"], num_lanes=4)
       Heat   Car Lane
    0     1  Car1    A
    1     1  Car2    B
    2     1  Car3    D
    3     2  Car4    B
    4     2  Car5    C
    ...
    """
    lane_labels = ["A", "B", "C", "D", "E", "F", "G", "H"][:num_lanes]

    if runs_per_car is None:
        runs_per_car = num_lanes

    runs_per_car = max(2, min(runs_per_car, num_lanes))

    run_pool = []
    for car in entry_list:
        chosen_lanes = secure_shuffle(lane_labels)[:runs_per_car]
        for lane in chosen_lanes:
            run_pool.append({"Car": car, "Lane": lane})

    run_pool = secure_shuffle(run_pool)

    heats = []
    heat_num = 1
    while run_pool:
        heat = []
        used_cars, used_lanes = set(), set()

        i = 0
        skipped_runs = []
        while i < len(run_pool) and len(heat) < num_lanes:
            run = run_pool[i]
            if run["Car"] not in used_cars and run["Lane"] not in used_lanes:
                heat.append({"Heat": heat_num, "Car": run["Car"], "Lane": run["Lane"]})
                used_cars.add(run["Car"])
                used_lanes.add(run["Lane"])
                run_pool.pop(i)
            else:
                skipped_runs.append(run_pool.pop(i))

        run_pool.extend(skipped_runs)

        if len(heat) >= 2:
            heats.append(heat)
            heat_num += 1
        else:
            run_pool = secure_shuffle(run_pool)

    heats_raw = pd.DataFrame([r for heat in heats for r in heat])
    balanced_df = rebalance_heats(heats_raw, num_lanes)
    optimized_df = optimize_opponent_fairness(balanced_df)

    return optimized_df


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(
            "Usage: python heats_perfect.py <racers_file.xlsx> [num_lanes] [runs_per_car]"
        )
        sys.exit(1)

    filename = sys.argv[1]
    lanes = int(sys.argv[2]) if len(sys.argv) > 2 else 4
    runs = int(sys.argv[3]) if len(sys.argv) > 3 else lanes

    if not 2 <= lanes <= 8:
        print("[ERROR] Number of lanes must be between 2 and 8.")
        sys.exit(1)

    try:
        df = read_excel_sheet(filename, sheet_name="Racers")
    except ValueError as err:
        print(f"[ERROR] {err}")
        sys.exit(1)

    validate_unique_car_ids(df)

    classes = df["Class"].dropna().unique()
    for c_class in classes:
        dfclass = df[df["Class"] == c_class]
        groups = dfclass["Group"].unique()

        for group in groups:
            dfgroup = (
                dfclass[dfclass["Group"].isna()]
                if is_nan(group)
                else dfclass[dfclass["Group"] == group]
            ).copy()
            group_name = "General" if is_nan(group) else group

            print(f"\nGenerating heats for: Class '{c_class}' / Group '{group_name}'")
            car_entries = secure_shuffle(dfgroup["Car"].dropna().astype(str).tolist())

            if len(car_entries) <= 1:
                print(f"[Tropy] Winner by default: {car_entries[0]}")
                continue

            # With how we are randomlly stuff cars into heats,
            # we might have instances where validations fail
            # Loop until we have success
            for attempt in range(200):
                gen_heats_df = generate_heats(
                    car_entries, num_lanes=lanes, runs_per_car=runs
                )

                if validate_heats(
                    gen_heats_df, expected_races_per_car=runs, all_cars=car_entries
                ):
                    gen_heats_df["Car"] = gen_heats_df["Car"].astype(str)
                    dfgroup["Car"] = dfgroup["Car"].astype(str)

                    gen_heats_df = gen_heats_df.merge(
                        dfgroup[["Car", "Name"]],
                        left_on="Car",
                        right_on="Car",
                        how="left",
                    )
                    gen_heats_df = gen_heats_df[["Heat", "Car", "Name", "Lane"]]
                    gen_heats_df["Place"] = ""

                    try:
                        sheet_title = sanitize_sheet_title(f"{c_class}_{group_name}")
                        with pd.ExcelWriter(
                            filename,
                            engine="openpyxl",
                            mode="a",
                            if_sheet_exists="replace",
                        ) as writer:
                            gen_heats_df.sort_values(by=["Heat", "Lane"]).to_excel(
                                writer, sheet_name=sheet_title, index=False
                            )
                        print(f"[OK] Heats written to sheet: {sheet_title}")
                        break
                    except OSError as exc:
                        print(f"[ERROR] Could not write heats to Excel: {exc}")
            else:
                print(
                    f"[FAIL] Could not generate valid heats for {c_class} / {group_name} "
                    "after 200 attempts."
                )

    heats_data = get_racer_heats(filename)
    update_racer_heats(filename, heats_data)
    format_all_sheets(filename)
