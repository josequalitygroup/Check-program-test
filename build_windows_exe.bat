@echo off
setlocal

REM Build Windows executable for QuickBooks Check Vendor Updater
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

pyinstaller --clean --noconfirm quickbooks_merger.spec

echo.
echo Build finished.
echo Executable: dist\QuickBooksCheckVendorUpdater\QuickBooksCheckVendorUpdater.exe

endlocal
