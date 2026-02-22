# QuickBooks Check Vendor Updater (Windows Desktop App)

A simple Windows desktop app built with **Python + PySide6 + pandas** with a polished startup splash and clean step-by-step interface.

It updates vendor/payee names in a QuickBooks upload file by matching check numbers from a second reference file (CSV or Excel .xlsx).

## Features

- Upload two files (CSV or Excel `.xlsx`):
  - **QuickBooks Upload file** (target file to update)
  - **Check Reference file** (lookup file with check number + vendor name)
- Friendly step-by-step GUI with status label, reset button, and clear action flow
- Startup splash screen (2 seconds) with branded title: **Jose's CSV Check Merger**
- Login screen appears after splash (User: `Kiri`, Password: `Jcr16331878`)
- Modernized visual styling with rounded panels, improved spacing, and polished controls
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
- Save updated file as CSV or Excel (default `QuickBooks_Upload_Updated.csv`)
- Optional unmatched report export (`*_Unmatched.csv` / `*_Unmatched.xlsx`)
- Optional backup of the original QuickBooks source file on save (`*_Backup.csv` or `*_Backup.xlsx`)
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

1. Open app and sign in (User: `Kiri`, Password: `Jcr16331878`).
2. Select QuickBooks Upload file (CSV/Excel).
3. Select Check Reference file (CSV/Excel).
4. Confirm/adjust column mappings.
5. Keep **Extract check number from text** enabled if your bank exports values like `Check 101`.
6. Click **Update Vendor Names**.
7. Review summary and preview.
8. (Optional) Keep **Create backup of original QuickBooks file** enabled.
9. Click **Save Updated CSV** (or choose Excel `.xlsx`).

## Notes

- Input files are not overwritten unless you explicitly save to the same path.
- This tool expects valid CSV files. Friendly errors are shown for missing files, bad CSVs, or missing mappings.


## Build a Windows Executable (.exe)

You can package the app as a Windows executable using **PyInstaller**.

### Option A (recommended): batch script

```bat
build_windows_exe.bat
```

### Option B: manual command

```bat
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --clean --noconfirm quickbooks_merger.spec
```

After build, use:

- `dist\QuickBooksCheckVendorUpdater\QuickBooksCheckVendorUpdater.exe`

> Note: if you add `assets/splash_bg.jpg`, it will be bundled automatically.

