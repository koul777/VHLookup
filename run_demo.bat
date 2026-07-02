@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and run this again.
  if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
  exit /b 1
)

echo Generating public-admin sample Excel reports...
python scripts\run_public_admin_demo.py
if errorlevel 1 (
  echo Demo generation failed.
  if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
  exit /b 1
)

echo.
echo Done. Reports are in:
echo %CD%\demo_output

if not "%VHLOOKUP_NO_OPEN%"=="1" start "" "%CD%\demo_output"
if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
exit /b 0
