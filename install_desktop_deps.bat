@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and run this again.
  if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
  exit /b 1
)

echo Installing desktop UI dependency: PySide6
python -m pip install PySide6
if errorlevel 1 (
  echo PySide6 installation failed.
  if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
  exit /b 1
)

echo Done. You can now run run_app.bat.
if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
exit /b 0
