@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and run this again.
  if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
  exit /b 1
)

set OUT=%CD%\cli_output
if not exist "%OUT%" mkdir "%OUT%"
set PYTHONPATH=%CD%\src

echo Running CLI examples with public-admin sample data...

python -m vhlookup_cli.main templates --out "%OUT%\00_templates_cli.xlsx"
if errorlevel 1 goto fail

python -m vhlookup_cli.main inspect --path samples\public_admin\03_merge_files\row_merge_school_submissions --template school_submission_consolidation --out "%OUT%\00_inspect_cli.xlsx"
if errorlevel 1 goto fail

python -m vhlookup_cli.main consolidate --folder samples\public_admin\03_merge_files\row_merge_school_submissions --template school_submission_consolidation --out "%OUT%\01_consolidation_cli.xlsx"
if errorlevel 1 goto fail

python -m vhlookup_cli.main lookup --reference samples\public_admin\03_merge_files\column_merge_hr_training\hr_employee_master.csv --target samples\public_admin\03_merge_files\column_merge_hr_training\hr_training_completion.csv --out "%OUT%\02_lookup_cli.xlsx"
if errorlevel 1 goto fail

python -m vhlookup_cli.main reconcile --reference samples\public_admin\90_extra_cli_samples\submission_reconciliation\expected_submitters.csv --target samples\public_admin\90_extra_cli_samples\submission_reconciliation\received_submitters.csv --reference-label expected --target-label received --out "%OUT%\03_reconcile_cli.xlsx"
if errorlevel 1 goto fail

python -m vhlookup_cli.main horizontal --file samples\public_admin\90_extra_cli_samples\horizontal_table\monthly_budget_wide.csv --out "%OUT%\04_horizontal_cli.xlsx"
if errorlevel 1 goto fail

echo.
echo Done. CLI reports are in:
echo %OUT%

if not "%VHLOOKUP_NO_OPEN%"=="1" start "" "%OUT%"
if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
exit /b 0

:fail
echo CLI example failed.
if not "%VHLOOKUP_NO_PAUSE%"=="1" pause
exit /b 1
