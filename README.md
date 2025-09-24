# Category Match Report

Generate an Excel report comparing Spanish `CATEGORY` vs English `NEW_RESULT['category']` across multiple CSV files.

- Summary tab: one row per input CSV with totals and match rate.
- One tab per CSV: all mismatched rows with the following columns only:
  - `ADMINISTRATIVE_CODE`
  - `ID`
  - `CREATED_AT`
  - `END_AT`
  - `JSON`
  - `CATEGORY_ES`
  - `CATEGORY_EN`

Input CSVs can be placed either in the project root or under `reports/`. The script automatically picks up both locations and ignores the generated `category_match_report.csv` if present.

## Requirements

- Python 3.9+ (tested with Python 3.13)
- Dependencies in `requirements.txt` (currently: `openpyxl`)

## Setup a virtual environment

### Windows (PowerShell)

```powershell
# From the project root (c:\Users\jmonterrosol\Downloads\new-reports)
py -m venv .venv
.\.venv\Scripts\Activate.ps1

# Upgrade pip (optional)
python -m pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt
```

If PowerShell blocks script execution, you may need to run PowerShell as Administrator and allow running local scripts temporarily:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### macOS (or Linux)

```bash
# From the project root
python3 -m venv .venv
source .venv/bin/activate

# Upgrade pip (optional)
python -m pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt
```

## Prepare your input data

- Put your CSV files in one of these locations:
  - Project root: `./*.csv`
  - Subfolder: `./reports/*.csv`
- CSV delimiter must be `;` (as in your samples).
- Required columns: `ADMINISTRATIVE_CODE`, `CREATED_AT`, `END_AT`, `ID`, `CATEGORY`, `JSON`, `NEW_RESULT`.
  - `NEW_RESULT` must be a Python-dict-like string containing `'category': '<english_slug>'`.

## Run the report

### Windows

```powershell
py report_compare_categories.py
```

### macOS/Linux

```bash
python3 report_compare_categories.py
```

Output:
- Excel: `category_match_report.xlsx` in the project root.
  - If the file is open/locked, the script will write a fallback like `category_match_report_YYYYMMDD_HHMMSS.xlsx` and print the path.
- Console summary of totals and match rates per file.

## Category mapping

`CATEGORY` (Spanish) is mapped to `NEW_RESULT['category']` (English slug) using the dictionary in `report_compare_categories.py` (`DEFAULT_MAPPING_RAW`). If you see unexpected mismatches for a valid category, update or add the mapping there.

## Troubleshooting

- "Permission denied" when saving Excel: close the file in Excel and rerun. The script already tries a fallback filename if the file is locked.
- Empty mismatch tabs: that file had 100% matches.
- Wrong delimiter: ensure your CSVs use `;` and are UTF-8 encoded (with or without BOM).
- Virtual environment not activating on Windows: see the PowerShell execution policy note above.
