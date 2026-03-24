@echo off
echo ==============================================
echo 2G KPI Report Generator Setup and Runner
echo ==============================================

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in the system PATH.
    echo Please install Python and try again.
    pause
    exit /b
)

:: Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

:: Activate the virtual environment
call venv\Scripts\activate

:: Install required packages
echo Installing requirements...
pip install -r requirements.txt

:: Run the application
echo Starting Application...
python main.py

pause
