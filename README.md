# NSQIP Risk Calculator Automation

Automates data entry into the [ACS NSQIP Surgical Risk Calculator](https://riskcalculator.facs.org/RiskCalculator/PatientInfo.jsp). Reads patient data from an Excel workbook, fills the calculator form in a real Chrome browser, scrapes the predicted risk values, and saves results to a JSON file.

## Requirements

- **Google Chrome** (installed normally on your computer)
- **Python 3.9 or later** -- download from [python.org](https://www.python.org/downloads/)
  - **Windows users:** check "Add Python to PATH" during installation

Everything else is installed automatically on first run.

## How to Run

### macOS / Linux

Open Terminal, navigate to this folder, and run:

```
./run.sh
```

Or double-click `run.sh` in Finder (you may need to right-click > Open the first time).

### Windows

Double-click `run.bat`, or open Command Prompt in this folder and run:

```
run.bat
```

## What the Program Asks

When you launch the program, it will prompt you for:

1. **Excel file path** -- drag and drop the file into the terminal, or type the path
2. **Sheet name** -- the program lists all sheets in the workbook; type the name or number
3. **Start row** -- the first data row to process (default: 2, since row 1 is headers)
4. **End row** -- the last data row to process (default: all rows)

After confirming, the program opens Chrome and begins filling the calculator automatically.

## Excel Format

The workbook must have this column layout (row 1 = headers, data starts at row 2):

| Column | Field |
|--------|-------|
| B | CASEID (used as the key in JSON output) |
| C | CPT code |
| D | Age |
| E | Sex |
| F | Functional Status |
| G | Emergency Case |
| H | ASA Class |
| I | Steroid Use |
| J | Ascites |
| K | Systemic Sepsis |
| L | Ventilator Dependent |
| M | Disseminated Cancer |
| N | Diabetes |
| O | Hypertension |
| P | CHF |
| Q | Oxygen Support |
| R | Current Smoker |
| S | History of COPD |
| T | Dialysis |
| U | Acute Renal Failure |
| V | Height (inches) |
| W | Weight (pounds) |

Special values:
- Height/Weight of `-99` is treated as blank (unknown)
- "Unknown" functional status is mapped to "Independent"
- CPT `20969` is automatically replaced with `20962`

## Output

Results are saved as a JSON file next to your Excel workbook, named:

```
<excel_filename>_<sheetname>_risks.json
```

For example, if your file is `data.xlsx` and the sheet is `Anahita`, the output is `data_anahita_risks.json`.

Each entry contains 15 risk values (5 outcomes x 3 surgeon adjustment levels):

- Serious Complication
- Any Complication
- Pneumonia
- Surgical Site Infection
- Return to OR

Each with: no adjustment, somewhat higher, significantly higher.

The JSON is saved after every row, so progress is never lost if the program is interrupted.

## CAPTCHA Handling

The NSQIP website sometimes shows a CAPTCHA challenge. When this happens:

1. The script pauses and displays a message in the terminal
2. Complete the CAPTCHA in the Chrome window that opened
3. The script automatically resumes once you pass it

The program uses a persistent browser profile (`browser_profile/` folder) that remembers your session between runs, so you typically only need to solve the CAPTCHA once.

## Advanced Usage

You can also run the batch script directly with command-line arguments:

```bash
source venv/bin/activate

python nsqip_batch.py --excel data.xlsx --sheet Anahita --start-row 2 --end-row 50

python nsqip_batch.py --excel data.xlsx --sheet Charbel --end-row 9999 --dry-run
```

Run `python nsqip_batch.py --help` for all available options.

## File Overview

| File | Purpose |
|------|---------|
| `run.sh` | macOS/Linux launcher (setup + run) |
| `run.bat` | Windows launcher (setup + run) |
| `launcher.py` | Interactive prompts for Excel path, sheet, row range |
| `nsqip_batch.py` | Core automation script |
| `fill_row.py` | Fill a single row for manual testing |
| `json_to_excel.py` | Write JSON results back into the Excel sheet |
| `requirements.txt` | Python dependencies |
