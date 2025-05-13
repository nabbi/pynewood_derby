# 🌟 Pynewood Derby Race Management

![Pynewood Derby](pynewood.png)

A Python-based toolchain for generating, simulating, and processing Pinewood Derby heats and results using Excel workbooks.  
Designed for fairness, traceability, and ease of use in large-scale events.

- ✅ Perfect-N & Partial-Perfect-N Scheduling Format  
- 🤝 Head-to-Head / Round-Robin Scheduling Format

---

## 📁 Project Files

| File               | Purpose                                                                 |
|--------------------|-------------------------------------------------------------------------|
| `raceday.xlsx`     | Main workbook containing racer data, heats, results, and rankings       |
| `heats_perfect.py` | Generates Perfect-N and Partial-Perfect-N heats                         |
| `heats_runoff.py`  | Generates head-to-head runoff heats (used for tiebreakers)              |
| `sim_results.py`   | Simulates randomized results for testing and demo purposes              |
| `results.py`       | Processes heat results and generates rankings and summaries             |
| `race_utils.py`    | Shared utilities for heat validation, opponent fairness optimization, secure shuffling, file handling, and formatting     |

---

## 🔄 Workflow

1. **Generate Heats**  
   Create balanced heats by class/group:
   ```bash
   python heats_perfect.py raceday.xlsx [num_lanes] [runs_per_car]
   ```

2. **Simulate Results (Optional)**  
   Fill in randomized placements for testing/demo:
   ```bash
   python sim_results.py raceday.xlsx
   ```

3. **Process Results**  
   Tally placements, compute rankings, and identify ties:
   ```bash
   python results.py raceday.xlsx
   ```

4. **Generate Runoff Heats**  
   Build fair runoff heats from tied top-3 racers (typically run after results.py into a new workbook):
   ```bash
   python heats_runoff.py raceday_runoff.xlsx [num_lanes]
   ```

---

## 🛠️ Features

- ✅ Heat validation (no duplicate cars/lane conflicts)
- 🎲 Secure shuffling for unbiased heat generation
- 📊 Competition-style ranking (1, 2, 2, 4…)
- 🔄 Auto-formatted Excel output with rankings & runoff tracking
- 🧐 Supports classes and sub-groups for flexible race organization
- 🧮 Optimized heat balancing and opponent fairness scoring
- 📅 Runoff generator based on unresolved top-3 ties

---

## 📂 Excel Workbook Structure

- **Racers** – Primary list of participants with class/group info
- **[Class_Group]** – Heat assignments per racing group
- **[Class_Group]_Rankings** – Auto-generated rankings from heat results
- **Runoff** – Auto-created if ties are detected in top-3 standings

### Racers Tab

The `Racers` sheet should include:

| Column      | Description                                              |
|-------------|----------------------------------------------------------|
| `Car`       | Unique identifier for each racer (e.g., 101, 202)        |
| `Name`      | Racer's display name                                     |
| `Class`     | Main category or division (e.g., Stock, Open)            |
| `Group`     | Sub-group within the class (e.g., Den 3, Tigers)         |
| `Description` | Optional notes (e.g., theme, color, builder notes)    |

This tab is used throughout the toolchain to generate, simulate, and evaluate results.

---

## 📌 Requirements

- Python 3.9+
- Libraries: `pandas`, `openpyxl`, `numpy`, `scipy`

### Install via virtual environment:
```bash
python3 -m venv venv-pynewood_derby
source venv-pynewood_derby/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 📣 Notes

- All ranking sheets are auto-named using the format: `[Group]_Rankings`
- Runoff logic only triggers if there are ties among ranks 1–3
- Can be used for real events or simulated/testing environments

---
