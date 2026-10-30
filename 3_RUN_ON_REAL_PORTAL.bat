@echo off
cd /d "%~dp0"
title 3. Run Automation On Swanidhi Portal
echo ======================================================================
echo PM SVANIDHI - VENDOR SCHEME TRACKER AUTOMATION
echo ======================================================================
echo.
echo Check before running:
echo 1. You already started Chrome using 1_START_CHROME.bat.
echo 2. You logged into the portal and opened "Vendor Scheme Tracker".
echo 3. The Excel file is CLOSED in Microsoft Excel (no lock file).
echo.

set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist %PYTHON_EXE% (
    if exist "C:\Users\hp\AppData\Local\Programs\Python\Python311\python.exe" (
        set PYTHON_EXE="C:\Users\hp\AppData\Local\Programs\Python\Python311\python.exe"
    ) else (
        set PYTHON_EXE=python
    )
)

set /p SHEET="Enter Sheet name (press Enter for DFS, or type Labour, Health, Food, WCD): "
if "%SHEET%"=="" set SHEET=DFS

set /p LIMIT="Enter number of applications to process (press Enter for 2, or type all): "
if "%LIMIT%"=="" set LIMIT=2

set /p WORKERS="Enter number of parallel tabs on this PC (press Enter for 3, or type 1): "
if "%WORKERS%"=="" set WORKERS=3

echo.
if /i "%LIMIT%"=="all" (
    echo [RUNNING FULL BATCH ON SHEET: %SHEET% WITH %WORKERS% PARALLEL TABS]
    %PYTHON_EXE% svanidhi_automation.py --excel data.xlsx.xlsx --sheet %SHEET% --workers %WORKERS% --allow-extra-portal
) else (
    echo [RUNNING WITH LIMIT = %LIMIT% ON SHEET: %SHEET% WITH %WORKERS% PARALLEL TABS]
    %PYTHON_EXE% svanidhi_automation.py --excel data.xlsx.xlsx --sheet %SHEET% --limit %LIMIT% --workers %WORKERS% --allow-extra-portal
)

echo.
pause
