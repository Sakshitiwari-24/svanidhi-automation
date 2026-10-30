================================================================================
PM SVANIDHI - VENDOR SCHEME TRACKER AUTOMATION GUIDE
================================================================================

Folder: c:\Users\hp\Desktop\SwanidhinData

FILES OVERVIEW:
--------------------------------------------------------------------------------
1_START_CHROME.bat        -> Launches Chrome in automation port 9222.
2_RUN_DRY_RUN.bat         -> Safe test (Read-only). Checks Excel & portal.
3_RUN_ON_REAL_PORTAL.bat  -> Live automation! Applies, sanctions, and updates Excel.
4_INSPECT_PORTAL.bat      -> Debug helper to inspect any application's table structure.
svanidhi_automation.py    -> Main automation script.
data.xlsx.xlsx            -> Your real data workbook (DFS, Labour, Health, etc.).

--------------------------------------------------------------------------------
HOW TO USE (STEP-BY-STEP):
--------------------------------------------------------------------------------

STEP 0: CLOSE EXCEL
   Make sure "data.xlsx" or "data.xlsx.xlsx" is NOT open in Excel.
   If Excel has it open, Windows locks the file and prevents saving.

STEP 1: LAUNCH CHROME
   - Double-click "1_START_CHROME.bat".
   - In the Chrome window that opens:
     1. Log in to PM SVANidhi with your login/password/OTP/CAPTCHA.
     2. Navigate to "Vendor Scheme Tracker" page.
     3. LEAVE THAT TAB OPEN. Keep the black command window open in background.

STEP 2: RUN DRY RUN (RECOMMENDED)
   - Double-click "2_RUN_DRY_RUN.bat".
   - Press Enter for Sheet (default: DFS).
   - Press Enter for limit (default: 5).
   - It will search the applications, check schemes, and show what it would do.
   - NO changes will be made to portal or Excel in Dry Run mode.

STEP 3: TEST ON 2 APPLICATIONS (LIVE)
   - Double-click "3_RUN_ON_REAL_PORTAL.bat".
   - Press Enter for Sheet (default: DFS).
   - Press Enter for Limit (default: 2).
   - The script will:
     * Create an automatic backup of your Excel file.
     * Apply & Sanction the schemes on the portal.
     * Update the Status column in Excel to "Sanction" and save.
   - Verify the 2 applications in your Excel file and on the portal.

STEP 4: FULL BATCH RUN
   - Double-click "3_RUN_ON_REAL_PORTAL.bat".
   - Enter Sheet (e.g., DFS, Labour, Health, Food, or WCD).
   - Type "all" and press Enter.
   - You can stop or restart anytime; it safely skips already sanctioned rows.
================================================================================
