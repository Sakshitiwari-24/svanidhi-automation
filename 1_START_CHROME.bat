@echo off
cd /d "%~dp0"
title 1. Start Chrome for Swanidhi Portal
echo ======================================================================
echo STARTING CHROME IN DEBUG MODE (PORT 9222)
echo ======================================================================
echo.
echo IMPORTANT INSTRUCTIONS:
echo 1. Close any existing open Chrome windows first.
echo 2. Chrome will launch automatically in debug mode.
echo 3. In the Chrome window that opens:
echo    - Log in to your PM SVANidhi portal (solve CAPTCHA / OTP).
echo    - Open the "Vendor Scheme Tracker" page.
echo    - KEEP THAT TAB OPEN!
echo 4. Leave this command prompt window open in the background.
echo.
echo ======================================================================

set CHROME_EXE="C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist %CHROME_EXE% (
    if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
        set CHROME_EXE="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ) else if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
        set CHROME_EXE="%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
    ) else (
        set CHROME_EXE=chrome.exe
    )
)

%CHROME_EXE% --remote-debugging-port=9222 --user-data-dir="C:\chrome_debug_profile"
pause
