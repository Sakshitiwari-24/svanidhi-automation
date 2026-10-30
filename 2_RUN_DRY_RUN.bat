@echo off
cd /d "%~dp0"
title 2. Dry Run - Vendor Scheme Tracker
echo ======================================================================
echo PM SVANIDHI VENDOR SCHEME TRACKER - DRY RUN (READ ONLY)
echo ======================================================================
echo.
echo This searches and compares the portal with your Excel data.
echo NO changes are made to the portal or Excel in Dry Run mode.
echo.
echo Requirements:
echo 1. You started Chrome using 1_START_CHROME.bat and logged in.
echo 2. You are on the "Vendor Scheme Tracker" page.
echo 3. The Excel file is closed in Microsoft Excel.
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

set /p LIMIT="Enter number of applications to test (press Enter for 5, or type all): "
if "%LIMIT%"=="" set LIMIT=5

echo.
if /i "%LIMIT%"=="all" (
    echo [RUNNING DRY RUN - ENTIRE SHEET %SHEET%]
    %PYTHON_EXE% svanidhi_automation.py --excel data.xlsx.xlsx --sheet %SHEET% --dry-run --allow-extra-portal
) else (
    echo [RUNNING DRY RUN - %LIMIT% APPLICATIONS ON SHEET %SHEET%]
    %PYTHON_EXE% svanidhi_automation.py --excel data.xlsx.xlsx --sheet %SHEET% --dry-run --limit %LIMIT% --allow-extra-portal
)

echo.
pause
