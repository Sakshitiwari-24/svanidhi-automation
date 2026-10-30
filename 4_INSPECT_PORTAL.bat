@echo off
cd /d "%~dp0"
title 4. Inspect Portal Selectors
echo ======================================================================
echo PM SVANIDHI - INSPECT PORTAL PAGE STRUCTURE
echo ======================================================================
echo.
echo This command searches for an application on the Vendor Scheme Tracker
echo and prints out the HTML table structure, column headers, and input fields.
echo.

set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist %PYTHON_EXE% (
    if exist "C:\Users\hp\AppData\Local\Programs\Python\Python311\python.exe" (
        set PYTHON_EXE="C:\Users\hp\AppData\Local\Programs\Python\Python311\python.exe"
    ) else (
        set PYTHON_EXE=python
    )
)

set /p APP_NO="Enter Application No to inspect (e.g. PMS072615318872): "
if "%APP_NO%"=="" set APP_NO=PMS072615318872

echo.
echo Inspecting %APP_NO% on sheet DFS...
%PYTHON_EXE% svanidhi_automation.py --excel data.xlsx.xlsx --sheet DFS --inspect %APP_NO%
echo.
pause
