@echo off
title 2. Test On Mock Portal
echo ======================================================================
echo RUNNING AUTOMATION TEST ON LOCAL MOCK FORM
echo ======================================================================
echo.
echo This runs the script in DRY RUN mode against mock_portal.html.
echo It fills the fields, checks the values, and lets you inspect the screen.
echo.
set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist %PYTHON_EXE% (
    if exist "C:\Users\hp\AppData\Local\Programs\Python\Python311\python.exe" (
        set PYTHON_EXE="C:\Users\hp\AppData\Local\Programs\Python\Python311\python.exe"
    ) else (
        set PYTHON_EXE=python
    )
)

%PYTHON_EXE% fill_form.py --mock
echo.
pause
