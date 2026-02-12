@echo off
REM ============================================================================
REM BESS Analyzer - Windows Standalone Build Script
REM ============================================================================
REM
REM Prerequisites:
REM   1. Install Python 3.10+ from https://www.python.org/downloads/
REM      (Check "Add Python to PATH" during install)
REM   2. Open Command Prompt or PowerShell in this directory
REM   3. Run this script: build_windows.bat
REM
REM Output:
REM   dist\BESS_Analyzer\BESS_Analyzer.exe  (standalone application)
REM
REM ============================================================================

echo.
echo ============================================
echo   BESS Analyzer - Windows Build
echo ============================================
echo.

REM Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.10+ and add to PATH.
    echo Download from: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Creating virtual environment...
if not exist "venv" (
    python -m venv venv
)
call venv\Scripts\activate.bat

echo [2/4] Installing dependencies...
pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt >nul 2>&1
pip install pyinstaller >nul 2>&1

echo [3/4] Building standalone application...
pyinstaller --noconfirm bess_analyzer.spec

echo [4/4] Copying data files...
if not exist "dist\BESS_Analyzer\resources\libraries" (
    mkdir "dist\BESS_Analyzer\resources\libraries"
)
copy /Y resources\libraries\*.json "dist\BESS_Analyzer\resources\libraries\" >nul 2>&1

echo.
echo ============================================
echo   Build Complete!
echo ============================================
echo.
echo   Executable: dist\BESS_Analyzer\BESS_Analyzer.exe
echo.
echo   To distribute:
echo     - Zip the entire dist\BESS_Analyzer folder
echo     - Send the zip to users
echo     - No Python installation needed to run
echo.
pause
