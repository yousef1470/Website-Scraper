@echo off
title Company Website Scraper Pro - Setup Installation
color 0A

echo ===============================================
echo   🚀 Company Website Scraper Pro - Setup
echo ===============================================
echo.
echo This setup will install all required dependencies
echo to run the Company Website Scraper on this device.
echo.
echo Dependencies to be installed:
echo   - Python packages (selenium, openpyxl)
echo   - Chrome WebDriver (if needed)
echo.
echo ===============================================
echo.

REM Change to the script directory
cd /d "%~dp0"

REM Check if Python is installed
echo 🔍 Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: Python is not installed or not in PATH
    echo.
    echo 📥 Please install Python first:
    echo    1. Go to https://python.org/downloads/
    echo    2. Download Python 3.8 or newer
    echo    3. Install with "Add Python to PATH" checked
    echo    4. Restart this setup after installation
    echo.
    pause
    exit /b 1
)

python --version
echo ✅ Python is installed!
echo.

REM Check if pip is available
echo 🔍 Checking pip (Python package manager)...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Error: pip is not available
    echo Installing pip...
    curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
    python get-pip.py
    del get-pip.py
)
echo ✅ pip is available!
echo.

REM Upgrade pip to latest version
echo 📦 Upgrading pip to latest version...
python -m pip install --upgrade pip
echo.

REM Install required Python packages
echo 📦 Installing required Python packages...
echo.

echo Installing selenium (web automation)...
python -m pip install selenium
if errorlevel 1 (
    echo ❌ Failed to install selenium
    goto :error
)
echo ✅ selenium installed successfully!
echo.

echo Installing openpyxl (Excel file handling)...
python -m pip install openpyxl
if errorlevel 1 (
    echo ❌ Failed to install openpyxl
    goto :error
)
echo ✅ openpyxl installed successfully!
echo.

echo Installing webdriver-manager (automatic ChromeDriver management)...
python -m pip install webdriver-manager
if errorlevel 1 (
    echo ❌ Failed to install webdriver-manager
    goto :error
)
echo ✅ webdriver-manager installed successfully!
echo.

REM ChromeDriver will be automatically managed
echo 🔍 ChromeDriver Management...
echo ✅ ChromeDriver will be automatically downloaded and managed!
echo 💡 No manual ChromeDriver setup required - webdriver-manager handles this automatically.
echo.

REM Create requirements.txt for future reference
echo 📝 Creating requirements.txt file...
echo selenium>requirements.txt
echo openpyxl>>requirements.txt
echo webdriver-manager>>requirements.txt
echo ✅ requirements.txt created!
echo.

REM Verify installation by testing imports
echo 🧪 Testing installation...
python -c "import selenium; print('✅ selenium import successful')" 2>nul
if errorlevel 1 (
    echo ❌ selenium import failed
    goto :error
)

python -c "import openpyxl; print('✅ openpyxl import successful')" 2>nul
if errorlevel 1 (
    echo ❌ openpyxl import failed
    goto :error
)

python -c "import tkinter; print('✅ tkinter (GUI) available')" 2>nul
if errorlevel 1 (
    echo ❌ tkinter not available (needed for GUI)
    echo 💡 tkinter should come with Python. Try reinstalling Python.
    goto :error
)

echo.
echo ===============================================
echo   🎉 SETUP COMPLETED SUCCESSFULLY!
echo ===============================================
echo.
echo ✅ All Python packages installed successfully
echo ✅ Dependencies verified
echo.
echo 📋 What's installed:
echo    - selenium (web automation)
echo    - openpyxl (Excel file handling)
echo    - webdriver-manager (automatic ChromeDriver management)
echo    - tkinter (GUI - built-in with Python)
echo.
echo 🚀 You can now run the scraper using:
echo    - Double-click: start_GUI.bat (for GUI)
echo    - Command: python scraper_gui.py
echo    - Command: python scraper_fixed.py
echo.
echo 📁 Make sure you have:
echo    - Your Excel file ready for processing
echo    - Chrome browser installed (ChromeDriver will be auto-managed)
echo.
echo Setup completed successfully!
echo.
pause
exit /b 0

:error
echo.
echo ===============================================
echo   ❌ SETUP FAILED
echo ===============================================
echo.
echo Something went wrong during installation.
echo Please check the error messages above.
echo.
echo 💡 Common solutions:
echo    - Ensure you have internet connection
echo    - Run as Administrator if needed
echo    - Check firewall/antivirus settings
echo    - Try installing packages manually:
echo      python -m pip install selenium openpyxl
echo.
pause
exit /b 1
