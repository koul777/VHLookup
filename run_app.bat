@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and run this again.
  if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
  exit /b 1
)

set PYTHONPATH=%CD%\src
python -m vhlookup_app.tk_main
if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
exit /b 0
