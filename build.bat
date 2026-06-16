@echo off
setlocal
cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 exit /b %errorlevel%

python -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name InventoryJsonViewer ^
  --icon icon.ico ^
  --add-data "icon.ico;." ^
  --collect-all openpyxl ^
  inventory_json_viewer.py
if errorlevel 1 exit /b %errorlevel%

endlocal
