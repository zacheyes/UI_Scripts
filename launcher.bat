@echo off
setlocal

:: --- Configuration ---
set PYTHON_SCRIPT=GUI.py

title RF Renamer Tool Launcher

:: --- 1. Check for Python installation ---
echo Checking for Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not found in your system's PATH.
    echo Please install Python 3 from python.org and ensure it is added to your PATH.
    pause
    exit /b
)

:: --- 2. Install required packages ---
echo Installing required packages...
pip install pandas openpyxl requests Pillow --quiet --disable-pip-version-check
if %errorlevel% neq 0 (
    echo ERROR: Failed to install required packages. Please check your internet connection.
    pause
    exit /b
)

:: --- 3. Launch the UI and Exit ---
echo.
echo Launching the RF Renamer Tool...

:: Use "pythonw.exe" to run the script without a console window.
:: The "start" command ensures this batch script doesn't wait for the UI to close.
:: The empty "" is a required placeholder for the start command's window title.
start "" pythonw "%PYTHON_SCRIPT%"

endlocal
exit /b