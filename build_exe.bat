@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python 3.10 or later and run this again.
  pause
  exit /b 1
)

set BUILD_PY=%CD%\.build_venv\Scripts\python.exe
if not exist "%BUILD_PY%" (
  echo Creating clean build environment...
  python -m venv "%CD%\.build_venv"
  if errorlevel 1 (
    echo Failed to create build environment.
    pause
    exit /b 1
  )
)

echo Installing build dependencies...
"%BUILD_PY%" -m pip install --disable-pip-version-check --upgrade pip >nul
"%BUILD_PY%" -m pip install --disable-pip-version-check pyinstaller pandas openpyxl >nul
if errorlevel 1 (
  echo Failed to install build dependencies.
  pause
  exit /b 1
)

echo Building VHLookupLocal_pivot.exe...
"%BUILD_PY%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  --name VHLookupLocal_pivot ^
  --paths "%CD%\src" ^
  --add-data "%CD%\samples;samples" ^
  --exclude-module PySide6 ^
  --exclude-module PyQt5 ^
  --exclude-module PyQt6 ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  --exclude-module pyarrow ^
  --exclude-module IPython ^
  --exclude-module pytest ^
  --exclude-module sphinx ^
  --exclude-module numba ^
  --exclude-module llvmlite ^
  --exclude-module zmq ^
  --exclude-module pygame ^
  "%CD%\src\vhlookup_app\tk_main.py"

if errorlevel 1 (
  echo EXE build failed.
  pause
  exit /b 1
)

echo.
echo Done:
echo %CD%\dist\VHLookupLocal_pivot.exe
pause
exit /b 0
