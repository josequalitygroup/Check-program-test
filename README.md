# Check Extractor Desktop App (Windows MVP)

A Windows desktop app (Tkinter + Python) to:
- Upload scanned check files (`.pdf`, `.jpg`, `.jpeg`, `.png`)
- Extract **Check Number** and **Vendor/Payee Name** using OpenAI vision
- Capture **Confidence** and **Notes** when extraction is uncertain
- Review/edit extracted rows in a table
- Filter by **Month** + **Year**
- Export filtered rows to **Excel (.xlsx)**

## 1) Prerequisites
- Windows 10/11
- Python 3.10+
- OpenAI API key in environment variable

## 2) Install
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 3) Set OpenAI API key (required)
PowerShell:
```powershell
setx OPENAI_API_KEY "your_api_key_here"
```
Reopen terminal so the variable is loaded.

## 4) Run
```bash
python app.py
```

## 5) Usage
1. Choose **Month** and **Year**.
2. Click **Upload Files** and select PDF/image files.
3. Wait for processing progress to finish.
4. Double-click table cells to edit values manually.
5. Click **Export Filtered to Excel**.

## 6) Data extraction behavior
- Multi-page PDFs are fully processed.
- A page can return zero, one, or multiple checks.
- If extraction fails for a file/page, the app still creates a row with an error note.

## 7) Export format
Exported filename pattern:
- `Checks_YYYY_MM.xlsx`

Exported columns:
- Check Number
- Vendor Name
- Month
- Year
- Source File
- Confidence
- Notes

## 8) Build Windows executable (optional)
```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name CheckExtractor app.py
```
Output:
- `dist\CheckExtractor.exe`

## Notes
- Never hardcode API keys in source code.
- OCR quality depends on image/PDF scan quality.
