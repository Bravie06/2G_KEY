@echo off
title 2G KPI Reporter Generator

echo Checking for Python...
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed or not added to PATH. Please install Python and try again.
    pause
    goto :eof
)

echo.
echo Setting up the virtual environment...
IF NOT EXIST "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Installing dependencies...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Starting the application...
python app.py

echo.
pause
