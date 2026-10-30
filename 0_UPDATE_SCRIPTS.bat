@echo off
cd /d "%~dp0"
title 0. Update Scripts from GitHub
echo ======================================================================
echo PM SVANIDHI - ONE-CLICK SCRIPT UPDATER
echo ======================================================================
echo.
echo Fetching latest svanidhi_automation.py and batch scripts from GitHub...
echo Repository: https://github.com/Sakshitiwari-24/svanidhi-automation
echo.

where git >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    if exist ".git" (
        git pull origin main
    ) else (
        powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Sakshitiwari-24/svanidhi-automation/main/svanidhi_automation.py' -OutFile 'svanidhi_automation.py'; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Sakshitiwari-24/svanidhi-automation/main/2_RUN_DRY_RUN.bat' -OutFile '2_RUN_DRY_RUN.bat'; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Sakshitiwari-24/svanidhi-automation/main/3_RUN_ON_REAL_PORTAL.bat' -OutFile '3_RUN_ON_REAL_PORTAL.bat'"
    )
) else (
    powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Sakshitiwari-24/svanidhi-automation/main/svanidhi_automation.py' -OutFile 'svanidhi_automation.py'; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Sakshitiwari-24/svanidhi-automation/main/2_RUN_DRY_RUN.bat' -OutFile '2_RUN_DRY_RUN.bat'; Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Sakshitiwari-24/svanidhi-automation/main/3_RUN_ON_REAL_PORTAL.bat' -OutFile '3_RUN_ON_REAL_PORTAL.bat'"
)

echo.
echo [SUCCESS] Scripts updated to the latest version!
echo.
pause
