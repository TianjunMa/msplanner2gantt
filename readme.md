# MS Planner to Gantt Chart Generator

An automated Python tool that transforms standard Microsoft Planner Excel exports into a clean, presentation-ready visual Gantt chart with automated milestone lines, checklist tracking, and schedule shading.

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![pandas](https://img.shields.io/badge/pandas-1.5%2B-150458?style=flat&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.0%2B-217346?style=flat&logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)

---

## Visual Color Scheme & Status Guide

The generated workbook applies a distinct, clean palette across stages, bars, and headers:

| Indicator / Element | Color Code | Preview / Hex | Description |
| :--- | :---: | :---: | :--- |
| **Completed Tasks** | Green | `🟩 #28A745` | Tasks marked as completed (positioned at the top) |
| **In Progress (Current / Past)** | Blue | `🟦 #3B82F6` | Active workstreams in elapsed or current calendar months |
| **Future Planned Work** | Slate Gray | `⬜ #94A3B8` | Active workstreams scheduled after the current month |
| **Actual Project Start Marker** | Solid Red | `🟥 #FF0000` | Vertical guide line denoting project start milestone |
| **Table Headers** | Navy Blue | `🟦 #1B365D` | Primary metadata header band |
| **Timeline Year / Month Headers** | Slate Dark | `⬛ #34495E` | Chronological timeline header band |
| **Completed Checklist Step** | Dark Green | `[✓] #155724` | Checklist milestones fulfilled |
| **Pending Checklist Step** | Muted Gray | `[ ] #7F8C8D` | Checklist milestones pending completion |

---

## Key Features

* **Interactive File Selection:** Opens your native system file explorer dialog (`tkinter`) to select any Planner `.xlsx` export without typing file paths.
* **Automatic Export Management:** Writes the output workbook directly into the same directory as the source file, appending an automated `_YYYY-MM-DD` date suffix.
* **Completed Tasks on Top:** Finished workstreams automatically sort to the top of the schedule, followed by active and planned tasks sorted chronologically.
* **Customizable Project Start Marker:** Prompts for the project start date (e.g., `2026-07`) and renders a solid red line (`#FF0000`) down that column. If left blank, it defaults to the earliest task date.
* **Time-Aware Bar Colors:** Active work differentiates past/current months (`#3B82F6`) from future planned periods (`#94A3B8`).
* **Checklist Progress Parsing:** Parses semicolon-delimited Planner checklists into multiline checkboxes (`[✓]` Done vs. `[ ]` Pending) mapped to completion ratios.
* **Expense Column:** Injects a dedicated, accounting-formatted (`$#,##0.00`) blank **Expense** column between *Due Date* and *Days* ready for manual cost entry.
* **Frozen Panes:** Pins metadata columns (Task ID, Name, Bucket, Status, Priority, Dates, Expense, Days, Checklist) on the left (`M6`) while timeline months scroll horizontally.

---

## Directory Structure

```text
ms-planner-gantt/
├── .gitignore
├── LICENSE
├── README.md
├── requirements.txt
└── msplanner2gantt.py
```

---

## Installation

### 1. Clone the repository
```bash
git clone https://github.com/TianjunMa/ms-planner-gantt.git
cd ms-planner-gantt
```

### 2. Set up a virtual environment (Recommended)
```bash
# macOS / Linux
python3 -m venv venv
source venv/bin/activate

# Windows (Command Prompt / PowerShell)
python -m venv venv
venv\Scripts\activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

> **Note for Linux users:** If running on a headless or bare Linux environment, ensure `python3-tk` is installed for the GUI file dialog:
> ```bash
> sudo apt-get update && sudo apt-get install -y python3-tk
> ```

---

## Usage

### Interactive Mode (Default)
Run the script directly from your terminal:
```bash
python msplanner2gantt.py
```
1. A file browser dialog will appear. Select your exported Microsoft Planner `.xlsx` file.
2. A prompt asks for the **Actual Project Start Date** (`YYYY-MM` or `YYYY-MM-DD`). Leave it blank to default to the earliest start date found in the file.
3. The script processes the schedule and writes `<Original_Filename>_Gantt_Chart_YYYY-MM-DD.xlsx` into the same folder.

### Programmatic / Scripting Mode
Import and execute the function in custom automated workflows:
```python
from msplanner2gantt import generate_gantt_from_planner

# Direct call with explicit paths and start date
generate_gantt_from_planner(
    planner_excel_path="Project Management.xlsx",
    project_start_date="2026-07"
)
```

---

## Expected Planner Export Format

The script processes standard exports generated from Microsoft Planner (*Export plan to Excel*), specifically targeting the **`Consolidated Data`** worksheet (with automatic fallback to sheet 1).

Supported headers:
* `Task Name`
* `Bucket`
* `Status` (`Completed`, `In progress`, `Not started`)
* `Priority` (`Urgent`, `Important`, `Medium`, `Low`)
* `Assigned To`
* `Start date`
* `Due date`
* `Checklist Items`
* `Completed Checklist Items` (e.g., `4/6` or integer)

---

## Configuration & Dependencies

Minimum tested versions (`requirements.txt`):
```text
pandas>=1.5.0
openpyxl>=3.0.0
```

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.