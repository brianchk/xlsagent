@echo off
REM Excel Analyzer - Windows Setup Script
REM Checks dependencies and sets up the environment

echo ============================================================
echo Excel Analyzer - Windows Setup
echo ============================================================
echo.

REM Check Python version
echo Checking Python...
python --version 2>NUL
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.11 or later.
    echo Download from: https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2 delims= " %%i in ('python --version') do set PYVER=%%i
echo Found Python %PYVER%

REM Check if Python version is 3.11+
python -c "import sys; exit(0 if sys.version_info >= (3, 11) else 1)" 2>NUL
if errorlevel 1 (
    echo ERROR: Python 3.11 or later required. Found %PYVER%
    pause
    exit /b 1
)
echo Python version OK

REM Check for Excel
echo.
echo Checking for Microsoft Excel...
reg query "HKEY_CLASSES_ROOT\Excel.Application" >NUL 2>&1
if errorlevel 1 (
    echo WARNING: Microsoft Excel not detected in registry.
    echo Screenshots require Excel to be installed.
    echo Analysis will still work without Excel.
) else (
    echo Excel found
)

REM Check for git (optional)
echo.
echo Checking for Git...
git --version 2>NUL
if errorlevel 1 (
    echo WARNING: Git not found. You can still run the analyzer.
) else (
    echo Git found
)

REM Create virtual environment
echo.
echo Setting up virtual environment...
cd /d "%~dp0skills\excel-analyzer"

if exist .venv (
    echo Virtual environment already exists
) else (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
)

REM Activate and install dependencies
echo.
echo Installing dependencies...
call .venv\Scripts\activate.bat

pip install --upgrade pip
pip install -e ".[dev]"
if errorlevel 1 (
    echo ERROR: Failed to install base dependencies
    pause
    exit /b 1
)

REM Install Windows-specific dependencies for screenshots
echo.
echo Installing Windows screenshot dependencies...
pip install pywin32 pillow
if errorlevel 1 (
    echo WARNING: Failed to install screenshot dependencies
    echo Screenshots may not work, but analysis will still function.
)

REM Verify installation
echo.
echo ============================================================
echo Verifying installation...
echo ============================================================
python -c "import openpyxl; print('openpyxl:', openpyxl.__version__)"
python -c "import xlwings; print('xlwings:', xlwings.__version__)"
python -c "import win32gui; print('pywin32: OK')" 2>NUL || echo pywin32: NOT INSTALLED
python -c "from PIL import Image; print('pillow: OK')"

echo.
echo ============================================================
echo Setup complete!
echo ============================================================
echo.
echo To run the analyzer:
echo   1. Open a command prompt
echo   2. cd %~dp0skills\excel-analyzer
echo   3. .venv\Scripts\activate
echo   4. python -m src.main "path\to\your\file.xlsx" -o output_folder
echo.
echo Example with sample file:
echo   python -m src.main "..\..\files\shop-sales\shop-sales.xlsm" -o output
echo.
pause
