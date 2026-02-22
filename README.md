# QuickBooks Check Vendor Updater (Windows Desktop App)

A simple Windows desktop app built with **Python + PySide6 + pandas**.

It updates vendor/payee names in a QuickBooks upload CSV by matching check numbers from a second reference CSV.

## Features

- Upload two CSV files:
  - **QuickBooks Upload CSV** (target file to update)
  - **Check Reference CSV** (lookup file with check number + vendor name)
- Flexible column mapping via dropdowns
- Automatic best-effort detection of common column names
- Matching by check number with normalization options:
  - trim spaces
  - convert to string
  - remove trailing `.0`
  - optionally extract check number from text like `Check 101` / `Cheque #101`
- Replaces vendor/payee only for matched check numbers
- Leaves unmatched rows unchanged
- Duplicate handling in reference CSV:
  - warns user
  - uses first match by default
- Preview first 100 rows before saving
- Save updated CSV as a new file (default `QuickBooks_Upload_Updated.csv`)
- Optional unmatched report export (`*_Unmatched.csv`)
- Summary metrics:
  - total rows
  - matched rows
  - unmatched rows
  - vendor names replaced

## Setup

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
python app.py
```

## Typical Workflow

1. Select QuickBooks Upload CSV.
2. Select Check Reference CSV.
3. Confirm/adjust column mappings.
4. Keep **Extract check number from text** enabled if your bank exports values like `Check 101`.
5. Click **Update Vendor Names**.
6. Review summary and preview.
7. Click **Save Updated CSV**.

## Notes

- Input files are not overwritten unless you explicitly save to the same path.
- This tool expects valid CSV files. Friendly errors are shown for missing files, bad CSVs, or missing mappings.
