# Axis MF Monthly Portfolio Automation

## Data Model and Assumptions
- Each monthly portfolio Excel file contains multiple scheme-level sheets.
- The index sheet is ignored during processing and a csv file is considered.
- One consolidated CSV is generated per month.
- Each CSV row represents a single instrument holding.
- The output schema includes:
  - AMC name
  - Scheme name
  - Instrument name (in what the fund is invested in)
  - ISIN
  - Market value
  - Percentage of portfolio
  - Reporting date
  - quantity (no of instruments)
  - rating
- Rows without ISIN are filtered.
- Additional fields (quantity, industry/rating) are retained for analytical completeness.

---

## Automation Approach
- The script automatically downloads monthly portfolio Excel files from the Axis Mutual Fund statutory disclosures website.
- Downloaded Excel files are converted into standardized CSV files.
- State tracking is maintained using JSON files to avoid duplicate processing.
- If a file is unreleased or has a naming mismatch, the month is marked as pending.
- The solution supports:
  - Interactive mode with manual fallback
  - Scheduled mode (non-interactive) for automation

---

# Steps to Run the Code

## 1. Prerequisites
- Python 3.9 or above installed
- Internet connection
- Windows OS (for scheduled execution)

---

## 2. Set Up the Environment

### Create and activate virtual environment
```bash```
*python -m venv venv*
*venv\Scripts\activate*

---

## 3. Install Dependencies
*pip install -r requirements.txt*

## 4. INTERACTIVE MODE
This mode is used for:
    - Historical backfill
    - Manual fallback when files are missing (DUE TO NAMING CONVENTION)
*python auto_monthly_download.py*
Behavior:
    - Scans all months from the earliest available date
    - Downloads available Excel files
    - Converts them into CSV
    - Prompts user only if a file is missing or naming does not match

## 5. SCHEDULED MODE (AUTOMATION)
This mode is used for unattended execution (Task Scheduler).
## Scheduled Execution (PowerShell Based)

## Overview
The script supports non-interactive execution using an environment variable (`RUN_MODE=scheduled`).  
This makes it safe to run via Windows Task Scheduler or automated scripts.

---

## 1. Open PowerShell
- Press **Win + R**
- Type `powershell`
- Press **Enter**

---

## 2. Navigate to Project Directory
```powershell``
cd "C:\Path\To\Project"

## 3. Set Scheduled Mode
*$env:RUN_MODE="scheduled"*

## 4. Execute the Sctipt
*python auto_monthly_download.py*

# NOTE
- Scheduled mode checks only recent months
- No browser or user prompt is triggered
- Missing files are logged in pending_months.json
- The script exits with status code 0 for successful runs

# IF WANT TO ADD A SCHEDULE IN TASK SCHEDULER PLEASE FOLLOW THE FOLLOWING
# Steps to Schedule the Script in Windows Task Scheduler

This document explains how to schedule the execution of the
`auto_monthly_download.py` script using **Windows Task Scheduler**.

The script is designed to run in **non-interactive (scheduled) mode**
using an environment variable.

---

## 1. Open Task Scheduler
1. Press **Win + R**
2. Type `taskschd.msc`
3. Press **Enter**

---

## 2. Create a New Task
1. Click **Create Task** (do not use “Create Basic Task”)

---

## 3. General Tab
- **Name**: Axis MF Monthly Portfolio Automation
- **Description**: Automatically downloads and converts Axis MF monthly portfolio files
- **Configure for**: Windows 10
- Check:
  - ☑ Run whether user is logged on or not
  - ☑ Run with highest privileges

---

## 4. Triggers Tab
1. Click **New**
2. Choose when the task should run (Daily / Weekly / One Time)
3. Set date and time
4. Click **OK**

---

## 5. Actions Tab
1. Click **New**
2. **Action**: Start a program

### Program/script
- give the pathway of the .bat file
### Start in
- fill with the project file path


(Example: the folder where `auto_monthly_download.py` is located)

---

## 6. Conditions Tab
- Uncheck:
  - ❌ Start the task only if the computer is on AC power
- Leave other options as default

---

## 7. Settings Tab
- Check:
  - ☑ Allow task to be run on demand
  - ☑ If the task fails, restart every 5 minutes
- Uncheck:
  - ❌ Stop the task if it runs longer than

---

## 8. Save the Task
- Click **OK**
- Enter Windows credentials if prompted

---

## 9. Verify Execution
1. Right-click the task
2. Click **Run**
3. Confirm:
   - **Last Run Result** = `0x0` (Success)

---

(Example: the folder where `auto_monthly_download.py` is located)

---

## 6. Conditions Tab
- Uncheck:
  - ❌ Start the task only if the computer is on AC power
- Leave other options as default

---

## 7. Settings Tab
- Check:
  - ☑ Allow task to be run on demand
  - ☑ If the task fails, restart every 5 minutes
- Uncheck:
  - ❌ Stop the task if it runs longer than

---

## 8. Save the Task
- Click **OK**
- Enter Windows credentials if prompted

---

## 9. Verify Execution
1. Right-click the task
2. Click **Run**
3. Confirm:
   - **Last Run Result** = `0x0` (Success)

---

## 6. Conditions Tab
- Uncheck:
  - ❌ Start the task only if the computer is on AC power
- Leave other options as default

---

## 7. Settings Tab
- Check:
  - ☑ Allow task to be run on demand
  - ☑ If the task fails, restart every 5 minutes
- Uncheck:
  - ❌ Stop the task if it runs longer than

---

## 8. Save the Task
- Click **OK**
- Enter Windows credentials if prompted

---

## 9. Verify Execution
1. Right-click the task
2. Click **Run**
3. Confirm:
   - **Last Run Result** = `0x0` (Success)

---

## Notes
- The task runs in **scheduled mode** without user input.
- Missing or unreleased files are recorded in `pending_months.json`.
- The script exits safely even when files are unavailable.

# Output
- Downloaded Excel files are stored in axis_monthly_portfolios/
- Generated CSV files are stored in axis_monthly_csv/
- State tracking is maintained using JSON files:
  - downloaded_files.json
  - generated_csv_files.json
  - pending_months.json
