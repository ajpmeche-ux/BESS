@echo off
REM Build script for Windows
REM Run this on a Windows machine to create the Windows executable

echo ============================================================
echo BESS Analyzer - Windows Build
echo ============================================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

REM Check if we're in the right directory
if not exist "main.py" (
    echo ERROR: main.py not found. Please run this script from the bess_analyzer directory.
    pause
    exit /b 1
)

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

REM Build the executable
echo.
echo Building executable...
python build_exe.py --clean

echo.
echo ============================================================
if exist "dist\BESS_Analyzer\BESS_Analyzer.exe" (
    echo Build successful!
    echo.
    echo Executable: dist\BESS_Analyzer\BESS_Analyzer.exe
    echo.
    echo To create a single-file .exe, run:
    echo     python build_exe.py --onefile --clean
) else (
    echo Build may have failed. Check the output above for errors.
)
echo ============================================================

pause
