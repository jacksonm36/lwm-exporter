@echo off
REM ============================================================
REM EML to PST Converter - Windows 7 Compatible Build Script
REM ============================================================
REM
REM This script builds Windows 7 compatible executables.
REM Requires Python 3.8.x (the last version supporting Windows 7)
REM
REM Download Python 3.8.10 from:
REM   https://www.python.org/downloads/release/python-3810/
REM
REM ============================================================

echo.
echo ============================================================
echo   EML to PST Converter - Windows 7 Build Script
echo ============================================================
echo.

set PYTHON38_32=C:\Python38-32\python.exe
set PYTHON38_64=C:\Python38\python.exe

REM Check for 32-bit Python 3.8
if exist "%PYTHON38_32%" (
    echo Found 32-bit Python 3.8 at %PYTHON38_32%
    echo Building 32-bit Windows 7 executable...
    "%PYTHON38_32%" -m pip install pyinstaller pywin32 --quiet
    "%PYTHON38_32%" build_exe.py --win7
) else (
    echo 32-bit Python 3.8 not found at %PYTHON38_32%
    echo To build 32-bit version, install Python 3.8.x 32-bit
)

echo.

REM Check for 64-bit Python 3.8
if exist "%PYTHON38_64%" (
    echo Found 64-bit Python 3.8 at %PYTHON38_64%
    echo Building 64-bit Windows 7 executable...
    "%PYTHON38_64%" -m pip install pyinstaller pywin32 --quiet
    "%PYTHON38_64%" build_exe.py --win7
) else (
    echo 64-bit Python 3.8 not found at %PYTHON38_64%
    echo To build 64-bit version, install Python 3.8.x 64-bit
)

echo.
echo ============================================================
echo   Build complete! Check the 'dist' folder.
echo ============================================================
echo.
echo Download Python 3.8.10 for Windows 7 compatibility:
echo   https://www.python.org/downloads/release/python-3810/
echo.
pause
