# 2G KPI Automation Tool

This tool automates the process of generating 2G KPI reports based on a specific template. It processes raw Excel exports (by site or by city), filters the data for the last 10 hours, applies color-coded conditional formatting, corrects percentage representation, and outputs a formatted report to the user's `Downloads` directory.

Everything runs locally on your machine with absolutely no data sent to external APIs or cloud services, guaranteeing full confidentiality.

## Prerequisites

- **Python**: Make sure Python is installed on your machine. You can download it from [python.org](https://www.python.org/downloads/). During installation, be sure to check the box that says **"Add Python to PATH"**.

## Installation & Usage (Windows)

1. Extract the project files to a folder on your computer.
2. Double-click the **`run.bat`** file.
3. The script will automatically:
   - Create a local virtual environment (`venv`).
   - Install the required dependencies (pandas, openpyxl, customtkinter) locally.
   - Launch the application's graphical user interface.

## Using the Application

1. **Raw Data File:** Click **Browse** to select the raw data Excel file (.xlsx) exported from the system (either by site or by city).
2. **Event Name:** Enter the event name in the provided text field. This will be used to name the output file (e.g., `KEA_2G_YourEventName_...xlsx`).
3. Click the **Generate Report** button.
4. The application will process the data and display a message indicating the exact path where the report was saved in your `Downloads` directory.

## Report Rules and Logic

- The tool filters the provided raw data to only include the last 10 hours of available data.
- **Percentages:** Percentage data is automatically scaled appropriately (multiplied by 100) prior to being displayed.
- **Color Logic:**
  - `Average of ORA_2G_TCH Availability(%)`: `≥ 98.5` is Green, `< 98.5` is Rose (Pink)
  - `Average of ORA_2G_CSSR_CS_New(%)`: `≥ 98.5` is Green, `< 98.5` is Rose (Pink)
  - `Average of ORA_2G_Call_Drop_CS_New(%)`: `≤ 0.7` is Green, `> 0.7` is Rose (Pink)
  - All other metrics are given a standard white background.
- Values are uniformly truncated to two decimal places.

## Troubleshooting

- **Python not recognized:** If running `run.bat` results in "Python is not installed or not added to PATH", ensure you have installed Python and ticked the box to add it to your PATH during installation.
- **Missing modules:** Ensure you use the `run.bat` script rather than clicking `app.py` directly, as the bat file manages dependency installation in the virtual environment.
