@echo off
title Company Website Scraper Pro - Starting...

REM Change to the script directory
cd /d "%~dp0"

echo ===============================================
echo   🚀 Company Website Scraper Pro
echo ===============================================
echo.
echo Starting GUI application...
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Python is not installed or not in PATH
    echo Please install Python and try again.
    echo.
    pause
    exit /b 1
)

REM Check if the GUI script exists
if not exist "scraper_gui.py" (
    echo ❌ Error: scraper_gui.py not found
    echo Please ensure you're running this from the correct directory.
    echo.
    pause
    exit /b 1
)

REM Start the GUI application
echo ✅ Starting the GUI...
python scraper_gui.py

REM If the script exits with an error, show error message
if errorlevel 1 (
    echo.
    echo ❌ The application encountered an error.
    echo Check the debug.log file for more details.
    echo.
)

REM Keep window open to see any messages
echo.
echo Press any key to close this window...
pause >nul
