@echo off
REM ============================================================
REM EML to PST Converter - Build Script
REM This script builds the executable using the current Python
REM ============================================================

echo.
echo ============================================================
echo   EML to PST Converter - Build Script
echo ============================================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python and try again.
    pause
    exit /b 1
)

REM Install requirements
echo Installing requirements...
pip install pyinstaller pywin32 >nul 2>&1

REM Run the build script
python build_exe.py

echo.
echo Build complete! Check the 'dist' folder for the executable.
echo.
pause
