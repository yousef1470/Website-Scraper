@echo off
title Quick Install - WebDriver Manager
color 0A

echo ===============================================
echo   📦 Installing WebDriver Manager
echo ===============================================
echo.
echo Installing webdriver-manager for automatic ChromeDriver management...
echo.

python -m pip install webdriver-manager

if errorlevel 1 (
    echo ❌ Failed to install webdriver-manager
    echo.
    echo 💡 Try running as Administrator or check your internet connection
    pause
    exit /b 1
) else (
    echo ✅ webdriver-manager installed successfully!
    echo.
    echo 🎉 Your scraper can now automatically manage ChromeDriver!
    echo 💡 No manual ChromeDriver setup needed anymore.
    echo.
)

pause
