# 2G KPI Report Generator

A desktop application designed to automate the generation of 2G KPI Excel reports from raw data files. The application generates a professional and structured Excel sheet replicating a specific template's format and applying color-coded rules based on KPI thresholds.

## Features
- Complete data extraction and transformation of the last 10 hours for 2G KPIs.
- Export to Excel using an exact template layout.
- Application of color code standards.
- User-friendly GUI interface.
- Complete data privacy: No APIs, everything runs locally.

## Requirements
- Python 3.8 or above installed on your Windows machine.
- Optional: VS Code.

## Setup & Running the Application

1. **Extract or Clone** this repository to your local machine.
2. **Double-click** the `run.bat` file.
   - The script will automatically create a virtual environment (`venv`).
   - It will install all the necessary dependencies listed in `requirements.txt` (`pandas`, `openpyxl`, `customtkinter`).
   - It will then open the graphical user interface.

## How to use

1. Open the application (via `run.bat`).
2. Click **Select Raw Data** and choose your raw Excel data file (e.g. `Performance Management...`).
3. Click **Select Output Folder** to choose where the generated report should be saved.
4. Input the desired name for your output file.
5. Click **Generate Excel Report**.
6. A success message will appear indicating your output file has been created in the destination folder.

## Color Codes Logic
- **Availability**: >= 98.5 (Green), < 98.5 (Faded Red)
- **CSSR CS**: >= 98.5 (Green), < 98.5 (Faded Red)
- **Call Drop CS**: <= 0.7 (Green), > 0.7 (Red)
- **Traffic CS, SDCCH Blocking, TCH Blocking**: White (No color)

---
*Confidentiality guaranteed. All processes are completed offline on your local machine.*