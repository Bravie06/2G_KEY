@echo off
title 2G KPI Reporter Generator

echo Checking for Python...
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not added to PATH.
    echo Please install Python and try again. Make sure to check "Add Python to PATH" during installation.
    pause
    goto :eof
)

echo.
echo Setting up the virtual environment...
IF NOT EXIST "venv" (
    echo Creating virtual environment with --system-site-packages...
    python -m venv --system-site-packages venv
)

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Installing dependencies...
python -m pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [WARNING] Dependency installation failed!
    echo This is likely due to a corporate firewall blocking pip ^(Error 10013^).
    echo The application may not start if the required modules ^(pandas, openpyxl, customtkinter^) are not installed globally or offline.
    echo Check the "Troubleshooting / Offline Installation" section in the README.md for solutions.
    echo.
)

echo.
echo Starting the application...
python app.py
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application crashed. This is usually because a required module is missing.
    echo Please ensure you are connected to the internet, not blocked by a firewall, or manually install the packages.
)

echo.
pause
