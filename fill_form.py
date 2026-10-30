"""
================================================================================
GOVERNMENT PORTAL (SVANIDHI) DATA ENTRY AUTOMATION SCRIPT
================================================================================
CRITICAL SAFETY NOTICE:
- Accuracy is paramount. The portal has NO edit/undo capability.
- Once submitted, records are final.
- This script implements pre-validation, read-back verification before clicking
  Submit, atomic row-by-row Excel updates, and immediate halt on any anomaly.
================================================================================
"""

import sys
import os
import re
import time
import datetime
from pathlib import Path
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ==============================================================================
# 1. SAFETY SWITCHES (Read carefully before changing)
# ==============================================================================
# DRY_RUN: If True, the script fills and verifies every form field on screen
# but NEVER clicks the Submit button. Keep True until you have verified the mock
# and are doing initial checks on the real portal.
DRY_RUN = True

# ASK_BEFORE_SUBMIT: If True, the script pauses before clicking Submit, shows
# the Name & Reference on screen, and waits for you to press Enter (or 'q' to quit).
ASK_BEFORE_SUBMIT = True

# MAX_ROWS_THIS_RUN: Safety throttle. Limits how many records are processed
# in a single execution. Start with 1-2 rows during rollout, then increase.
MAX_ROWS_THIS_RUN = 2

# ==============================================================================
# 2. RUNTIME MODE & BROWSER CONNECTION
# ==============================================================================
# Set USE_MOCK_PORTAL = True to test locally against mock_portal.html.
# Set USE_MOCK_PORTAL = False to connect to your real Chrome window with the portal.
USE_MOCK_PORTAL = False
if "--mock" in sys.argv:
    USE_MOCK_PORTAL = True

# Chrome DevTools Protocol URL (when attaching to your manually opened Chrome)
CDP_URL = "http://localhost:9222"

# Keyword/substring to identify the correct portal tab in Chrome (case-insensitive).
# If blank, the script uses the active or first open tab.
PORTAL_TAB_HINT = "swanidhi"

# ==============================================================================
# 3. PORTAL SELECTORS & CONFIGURATION
# ==============================================================================
# Change these when switching from the mock page to the real Swanidhi portal.
# For the real portal: Right click the input -> Inspect -> copy CSS selector / ID.

if USE_MOCK_PORTAL:
    FORM_URL = Path(__file__).resolve().parent.joinpath("mock_portal.html").as_uri()
    NAME_SELECTOR = "#applicantName"
    REF_SELECTOR = "#refNumber"
    SUBMIT_BUTTON_SELECTOR = "#btnSubmit"
    SUCCESS_SELECTOR = "#successMessage"
    ERROR_SELECTOR = "#errorMessage"
else:
    # --------------------------------------------------------------------------
    # REAL PORTAL CONFIGURATION (Update with your inspected portal elements)
    # --------------------------------------------------------------------------
    FORM_URL = ""  # Optional: full URL of the entry page if auto-navigation is desired
    NAME_SELECTOR = "#applicantName"          # <-- Replace with real portal Name field selector
    REF_SELECTOR = "#refNumber"              # <-- Replace with real portal Reference field selector
    SUBMIT_BUTTON_SELECTOR = "#btnSubmit"    # <-- Replace with real portal Submit button selector
    SUCCESS_SELECTOR = ".alert-success"      # <-- Replace with real portal Success banner/message selector
    ERROR_SELECTOR = ".alert-danger"         # <-- Replace with real portal Error banner selector (or None)

# How many milliseconds to wait for the success indicator after clicking Submit
# Government portals can be slow; 15000 ms = 15 seconds.
SUBMIT_TIMEOUT_MS = 15000

# Action after successful submission:
# "reload"   -> reloads current page (page.reload()) to clear form
# "navigate" -> re-navigates to FORM_URL
# "none"     -> portal automatically clears form or stays on page
AFTER_SUBMIT_ACTION = "reload"

# ==============================================================================
# 4. FILE PATHS & VALIDATION RULES
# ==============================================================================
EXCEL_FILE = "data.xlsx"
LOG_FILE = "log.txt"

# Expected format for Reference numbers.
# Example: r'^\d{12}$' enforces exactly 12 numeric digits.
# Set to None or "" if reference numbers have no fixed length/format.
EXPECTED_REF_REGEX = r'^\d+$'  # Enforces numeric digits (any length) with leading zeros preserved


# ==============================================================================
# LOGGING HELPER
# ==============================================================================
def write_log(message: str):
    """Appends a timestamped log entry to LOG_FILE and prints to console."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {message}"
    print(log_line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_line + "\n")
    except Exception as e:
        print(f"WARNING: Could not write to log file: {e}")


# ==============================================================================
# DATA CLEANING & UPFRONT VALIDATION
# ==============================================================================
def clean_spaces(text: str) -> str:
    """Trims leading/trailing spaces and collapses double/multiple internal spaces to one."""
    if not text:
        return ""
    # Strip leading/trailing whitespaces, then replace multiple spaces with single space
    return re.sub(r'\s+', ' ', str(text).strip())


def validate_and_load_data(file_path: str):
    """
    Validates ALL rows in the Excel file before touching the portal.
    Checks:
      1. Required columns exist ('Name', 'Reference', 'Status')
      2. No empty Name or Reference for pending rows
      3. No duplicate Reference numbers across the file
      4. Detects and normalizes extra spaces, reporting changes
      5. Matches EXPECTED_REF_REGEX (if set)
    Returns:
      (rows_to_process, all_data_stats, changes_made)
      Raises SystemExit if any critical validation error is found.
    """
    if not os.path.exists(file_path):
        write_log(f"FATAL ERROR: Excel file '{file_path}' not found in current folder.")
        sys.exit(1)

    print("=" * 70)
    print("STEP 1: UPFRONT DATA VALIDATION")
    print(f"Reading '{file_path}' (all values read as exact text)...")
    print("=" * 70)

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
    except PermissionError:
        write_log(f"FATAL ERROR: Cannot open '{file_path}'. It appears to be open in Microsoft Excel.")
        write_log("Please close Excel and run the script again.")
        sys.exit(1)
    except Exception as e:
        write_log(f"FATAL ERROR: Failed to read '{file_path}': {e}")
        sys.exit(1)

    # Read header row
    headers = [cell.value for cell in ws[1]]
    header_map = {}
    for idx, h in enumerate(headers, start=1):
        if h:
            header_map[str(h).strip()] = idx

    required_cols = ["Name", "Reference", "Status"]
    missing_cols = [col for col in required_cols if col not in header_map]
    if missing_cols:
        write_log(f"FATAL VALIDATION ERROR: Missing required column(s) in row 1: {missing_cols}")
        write_log(f"Found columns: {list(header_map.keys())}")
        sys.exit(1)

    name_col = header_map["Name"]
    ref_col = header_map["Reference"]
    status_col = header_map["Status"]

    validation_errors = []
    space_cleaning_notes = []
    seen_references = {}  # ref -> list of row numbers
    records = []
    total_rows = 0
    done_count = 0

    # Scan every data row (starting from row 2)
    max_row = ws.max_row
    excel_needs_update = False

    for r_idx in range(2, max_row + 1):
        name_cell = ws.cell(row=r_idx, column=name_col)
        ref_cell = ws.cell(row=r_idx, column=ref_col)
        status_cell = ws.cell(row=r_idx, column=status_col)

        raw_name = str(name_cell.value) if name_cell.value is not None else ""
        raw_ref = str(ref_cell.value) if ref_cell.value is not None else ""
        raw_status = str(status_cell.value) if status_cell.value is not None else ""

        # Skip completely blank lines at the bottom of sheet
        if raw_name.strip() == "" and raw_ref.strip() == "" and raw_status.strip() == "":
            continue

        total_rows += 1
        cleaned_status = raw_status.strip()
        is_done = cleaned_status.lower() == "done"

        if is_done:
            done_count += 1

        # Check spaces for Name
        clean_n = clean_spaces(raw_name)
        if raw_name != clean_n:
            space_cleaning_notes.append(
                f"Row {r_idx} [Name]: space normalized from '{raw_name}' -> '{clean_n}'"
            )
            # Update cell in memory
            name_cell.value = clean_n
            excel_needs_update = True

        # Check spaces for Reference
        clean_r = clean_spaces(raw_ref)
        if raw_ref != clean_r:
            space_cleaning_notes.append(
                f"Row {r_idx} [Reference]: space normalized from '{raw_ref}' -> '{clean_r}'"
            )
            ref_cell.value = clean_r
            ref_cell.number_format = '@'
            excel_needs_update = True

        # Validation Rule A: Empty values
        if not clean_n:
            validation_errors.append(f"Row {r_idx}: 'Name' is empty.")
        if not clean_r:
            validation_errors.append(f"Row {r_idx}: 'Reference' is empty.")

        # Validation Rule B: Reference format regex check
        if clean_r and EXPECTED_REF_REGEX:
            if not re.match(EXPECTED_REF_REGEX, clean_r):
                validation_errors.append(
                    f"Row {r_idx}: Reference '{clean_r}' does not match expected pattern ({EXPECTED_REF_REGEX})."
                )

        # Validation Rule C: Duplicate reference check
        if clean_r:
            if clean_r in seen_references:
                seen_references[clean_r].append(r_idx)
            else:
                seen_references[clean_r] = [r_idx]

        records.append({
            "row_idx": r_idx,
            "name": clean_n,
            "ref": clean_r,
            "status": cleaned_status,
            "is_done": is_done
        })

    # Check for duplicates
    for ref_val, row_nums in seen_references.items():
        if len(row_nums) > 1:
            validation_errors.append(
                f"Duplicate Reference '{ref_val}' found across rows: {row_nums}. Portal allows no duplicates!"
            )

    # Print Validation Report
    print(f"\n--- VALIDATION SUMMARY REPORT ---")
    print(f"Total rows inspected: {total_rows}")
    print(f"Already completed (Status = Done): {done_count}")
    print(f"Pending to process: {total_rows - done_count}")

    if space_cleaning_notes:
        print(f"\n[INFO] {len(space_cleaning_notes)} whitespace adjustments made:")
        for note in space_cleaning_notes:
            print(f"  • {note}")

    if validation_errors:
        print("\n" + "!" * 70)
        print("FATAL: VALIDATION FAILED! The following issues must be fixed in Excel:")
        print("!" * 70)
        for err in validation_errors:
            print(f"  [ERROR] {err}")
        print("!" * 70)
        print("Automation STOPPED immediately. No portal actions were taken.")
        wb.close()
        sys.exit(1)

    print("\n[SUCCESS] Upfront data validation passed with 0 errors!")

    # If spaces were cleaned, save the cleaned strings back to Excel
    if excel_needs_update:
        try:
            wb.save(file_path)
            print(f"[INFO] Saved cleaned space adjustments back to '{file_path}'.")
        except PermissionError:
            write_log(f"FATAL ERROR: Could not save cleaned spaces to '{file_path}'. File is open in Excel!")
            sys.exit(1)

    wb.close()
    return records, status_col


# ==============================================================================
# PRE-FLIGHT WRITE PERMISSION CHECK
# ==============================================================================
def check_excel_writable(file_path: str):
    """Verifies that data.xlsx is NOT locked by Excel before any form is filled."""
    try:
        # Attempt to open in append mode to test OS lock
        with open(file_path, "a+b"):
            pass
    except PermissionError:
        print("\n" + "!" * 70)
        print("CRITICAL PRE-FLIGHT FAILURE: EXCEL FILE IS LOCKED!")
        print(f"'{file_path}' is currently open in Microsoft Excel or another program.")
        print("Please CLOSE the Excel file immediately and re-run this script.")
        print("The script will NOT proceed because it cannot write 'Done' to a locked file.")
        print("!" * 70 + "\n")
        sys.exit(1)


# ==============================================================================
# ATOMIC EXCEL UPDATE
# ==============================================================================
def mark_row_done(file_path: str, row_idx: int, status_col: int, name: str, ref: str):
    """
    Atomically writes 'Done' into the Status cell of data.xlsx and saves immediately.
    Guarantees no work is lost and duplicate submissions cannot happen.
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        status_cell = ws.cell(row=row_idx, column=status_col)
        status_cell.value = "Done"
        wb.save(file_path)
        wb.close()
    except PermissionError:
        print("\n" + "!" * 70)
        print(f"CRITICAL PERMISSION ERROR WHILE SAVING ROW {row_idx}!")
        print(f"Record was SUBMITTED to portal successfully:")
        print(f"  Name:      {name}")
        print(f"  Reference: {ref}")
        print(f"HOWEVER, '{file_path}' could NOT be updated because Excel is open!")
        print("ACTION REQUIRED:")
        print("1. Close Microsoft Excel.")
        print(f"2. Manually enter 'Done' in the Status column for Row {row_idx}.")
        print("3. Do NOT re-run until Row {row_idx} has Status='Done'!")
        print("!" * 70 + "\n")
        write_log(f"[FILE_LOCK_ERROR] Row {row_idx} submitted but Excel save failed due to file lock.")
        sys.exit(1)
    except Exception as e:
        write_log(f"FATAL ERROR: Failed saving Excel row {row_idx}: {e}")
        sys.exit(1)


# ==============================================================================
# BROWSER ATTACHMENT / PAGE SELECTION
# ==============================================================================
def get_browser_and_page(playwright_instance):
    """
    Connects to the user's existing Chrome window via CDP (port 9222)
    or launches a test browser if USE_MOCK_PORTAL is configured for standalone testing.
    """
    if USE_MOCK_PORTAL and not os.environ.get("USE_CDP_FOR_MOCK"):
        is_headless = os.environ.get("HEADLESS", "0") == "1"
        browser = playwright_instance.chromium.launch(headless=is_headless)
        context = browser.new_context()
        page = context.new_page()
        page.goto(FORM_URL)
        return browser, page

    print(f"[INFO] Connecting to Chrome via CDP on {CDP_URL}...")
    try:
        browser = playwright_instance.chromium.connect_over_cdp(CDP_URL)
    except Exception as e:
        print("\n" + "!" * 70)
        print(f"FATAL: COULD NOT CONNECT TO CHROME ON {CDP_URL}!")
        print(f"Details: {e}")
        print("\nPlease ensure you started Chrome using this exact Windows command:")
        print(r'  & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome_debug_profile"')
        print("!" * 70 + "\n")
        sys.exit(1)

    contexts = browser.contexts
    if not contexts:
        write_log("FATAL: No browser contexts found in connected Chrome window.")
        sys.exit(1)

    pages = contexts[0].pages
    if not pages:
        write_log("FATAL: No open tabs found in the connected Chrome window.")
        write_log("Please open a tab, log in to the portal, and navigate to the entry page.")
        sys.exit(1)

    target_page = None
    if PORTAL_TAB_HINT:
        for pg in pages:
            url_match = PORTAL_TAB_HINT.lower() in pg.url.lower()
            title_match = PORTAL_TAB_HINT.lower() in pg.title().lower()
            if url_match or title_match:
                target_page = pg
                break

    if target_page is None:
        # If hint not matched or not set, use the active/first page
        target_page = pages[0]
        print(f"[INFO] Using Chrome tab: '{target_page.title()}' ({target_page.url})")
    else:
        print(f"[INFO] Found matching portal tab: '{target_page.title()}' ({target_page.url})")

    return browser, target_page


# ==============================================================================
# MAIN FORM SUBMISSION WORKFLOW
# ==============================================================================
def run_automation():
    print("\n" + "=" * 70)
    print("PM SVANIDHI DATA ENTRY AUTOMATION")
    print("=" * 70)
    print(f"Configured Safety Switches:")
    print(f"  • DRY_RUN:           {DRY_RUN} (Submit button will {'NEVER' if DRY_RUN else 'BE'} clicked)")
    print(f"  • ASK_BEFORE_SUBMIT: {ASK_BEFORE_SUBMIT}")
    print(f"  • MAX_ROWS_THIS_RUN: {MAX_ROWS_THIS_RUN}")
    print(f"  • Target Mode:       {'MOCK PORTAL' if USE_MOCK_PORTAL else 'REAL SVANIDHI PORTAL'}")
    print("=" * 70 + "\n")

    # Step 1: Pre-flight file lock check
    check_excel_writable(EXCEL_FILE)

    # Step 2: Validate all rows up front
    records, status_col = validate_and_load_data(EXCEL_FILE)

    # Filter out records already marked Done
    pending_records = [r for r in records if not r["is_done"]]

    if not pending_records:
        write_log("All records in data.xlsx are already marked 'Done'! Nothing to do.")
        return

    # Apply MAX_ROWS_THIS_RUN throttle
    batch = pending_records[:MAX_ROWS_THIS_RUN]
    print(f"\n[BATCH PLAN] Will process {len(batch)} row(s) in this run (out of {len(pending_records)} pending).")
    for item in batch:
        print(f"  -> Row {item['row_idx']}: Name='{item['name']}', Reference='{item['ref']}'")

    # Connect to browser
    with sync_playwright() as p:
        browser, page = get_browser_and_page(p)

        processed_in_this_run = 0

        for record in batch:
            row_idx = record["row_idx"]
            name = record["name"]
            ref = record["ref"]

            print("\n" + "-" * 70)
            print(f"PROCESSING ROW {row_idx}: Name = '{name}' | Reference = '{ref}'")
            print("-" * 70)

            # Double-check file writability before filling each row
            check_excel_writable(EXCEL_FILE)

            # ------------------------------------------------------------------
            # Form Field Population
            # ------------------------------------------------------------------
            try:
                name_field = page.locator(NAME_SELECTOR)
                ref_field = page.locator(REF_SELECTOR)

                name_field.wait_for(state="visible", timeout=10000)
                ref_field.wait_for(state="visible", timeout=10000)

                # Clear and fill Name
                name_field.fill("")
                name_field.fill(name)

                # Clear and fill Reference
                ref_field.fill("")
                ref_field.fill(ref)

            except Exception as e:
                write_log(f"[LOCATOR_ERROR] Row {row_idx}: Could not find or fill form fields: {e}")
                print("\nFATAL: Form fields were not accessible. Stopping script to prevent errors.")
                sys.exit(1)

            # ------------------------------------------------------------------
            # MANDATORY READ-BACK VERIFICATION
            # ------------------------------------------------------------------
            page_name_val = name_field.input_value()
            page_ref_val = ref_field.input_value()

            print(f"[READ-BACK CHECK]")
            print(f"  Excel Name:      '{name}'  | Page Field Name:      '{page_name_val}'")
            print(f"  Excel Reference: '{ref}'  | Page Field Reference: '{page_ref_val}'")

            if page_name_val != name:
                msg = (f"READ-BACK MISMATCH for Row {row_idx} [Name]!\n"
                       f"Excel has: '{name}', but page contains: '{page_name_val}'")
                write_log(f"[READ_BACK_MISMATCH] Row {row_idx} | Name: {name} | Ref: {ref} | {msg}")
                print("\n" + "!" * 70)
                print(msg)
                print("ABORTING SUBMISSION IMMEDIATELY!")
                print("!" * 70 + "\n")
                sys.exit(1)

            if page_ref_val != ref:
                msg = (f"READ-BACK MISMATCH for Row {row_idx} [Reference]!\n"
                       f"Excel has: '{ref}', but page contains: '{page_ref_val}'")
                write_log(f"[READ_BACK_MISMATCH] Row {row_idx} | Name: {name} | Ref: {ref} | {msg}")
                print("\n" + "!" * 70)
                print(msg)
                print("ABORTING SUBMISSION IMMEDIATELY!")
                print("!" * 70 + "\n")
                sys.exit(1)

            print("[SUCCESS] Read-back verification passed! Values on page match Excel character-for-character.")

            # ------------------------------------------------------------------
            # DRY RUN CHECK
            # ------------------------------------------------------------------
            if DRY_RUN:
                print(f"[DRY_RUN] Form filled and verified successfully for Row {row_idx}.")
                print(f"[DRY_RUN] Submit button was NOT clicked. Status was NOT modified in Excel.")
                write_log(f"[DRY_RUN] Row {row_idx} | Name: {name} | Ref: {ref} | Result: Verified on page, Submit skipped.")
                processed_in_this_run += 1
                if ASK_BEFORE_SUBMIT and sys.stdin.isatty():
                    user_input = input("--> [DRY RUN] Inspect the browser window now. Press ENTER to continue (or 'q' to quit): ").strip()
                    if user_input.lower() == 'q':
                        print("[INFO] Quitting.")
                        break
                else:
                    time.sleep(2)
                continue

            # ------------------------------------------------------------------
            # USER CONFIRMATION PROMPT (if ASK_BEFORE_SUBMIT is True)
            # ------------------------------------------------------------------
            if ASK_BEFORE_SUBMIT:
                print("\n" + "=" * 60)
                print(f"READY TO SUBMIT ROW {row_idx}:")
                print(f"  Applicant Name:   {name}")
                print(f"  Reference Number: {ref}")
                print("=" * 60)
                user_choice = input("Press ENTER to submit this record, or type 'q' and ENTER to quit: ").strip()
                if user_choice.lower() == 'q':
                    write_log(f"[USER_ABORT] User cancelled submission at Row {row_idx}. Exiting safely.")
                    sys.exit(0)

            # ------------------------------------------------------------------
            # SUBMIT RECORD & AWAIT CONFIRMATION
            # ------------------------------------------------------------------
            try:
                submit_btn = page.locator(SUBMIT_BUTTON_SELECTOR)
                submit_btn.wait_for(state="visible", timeout=5000)
                print("[INFO] Clicking Submit button...")
                submit_btn.click()

                # Await success banner/indicator
                print(f"[INFO] Awaiting confirmation indicator '{SUCCESS_SELECTOR}' (timeout {SUBMIT_TIMEOUT_MS/1000}s)...")
                page.wait_for_selector(SUCCESS_SELECTOR, state="visible", timeout=SUBMIT_TIMEOUT_MS)

            except PlaywrightTimeoutError:
                err_msg = (
                    f"TIMEOUT waiting for submission confirmation on Row {row_idx}!\n"
                    f"Name: {name} | Reference: {ref}\n"
                    f"The portal did NOT display the success indicator '{SUCCESS_SELECTOR}' within {SUBMIT_TIMEOUT_MS/1000}s."
                )
                print("\n" + "!" * 70)
                print("CRITICAL SUBMISSION TIMEOUT / UNEXPECTED STATE")
                print(err_msg)
                print("\nDO NOT RE-RUN THE SCRIPT YET!")
                print("Check the government portal manually RIGHT NOW to see if this record was saved.")
                print(f"Row {row_idx} was NOT marked 'Done' in Excel.")
                print("!" * 70 + "\n")
                write_log(f"[SUBMISSION_TIMEOUT] Row {row_idx} | Name: {name} | Ref: {ref} | Result: Timeout waiting for success")
                sys.exit(1)
            except Exception as e:
                err_msg = f"Unexpected error during submission of Row {row_idx}: {e}"
                print("\n" + "!" * 70)
                print("CRITICAL SUBMISSION ERROR")
                print(err_msg)
                print("Check the government portal manually before proceeding.")
                print("!" * 70 + "\n")
                write_log(f"[SUBMISSION_ERROR] Row {row_idx} | Name: {name} | Ref: {ref} | Result: {e}")
                sys.exit(1)

            # ------------------------------------------------------------------
            # SUCCESS CONFIRMED: ATOMIC UPDATE & LOGGING
            # ------------------------------------------------------------------
            print(f"[CONFIRMED] Success indicator confirmed on page for Row {row_idx}!")
            mark_row_done(EXCEL_FILE, row_idx, status_col, name, ref)
            write_log(f"[SUCCESS] Row {row_idx} | Name: {name} | Ref: {ref} | Result: Submitted successfully and marked Done")
            processed_in_this_run += 1

            # ------------------------------------------------------------------
            # READY FORM FOR NEXT ENTRY
            # ------------------------------------------------------------------
            if processed_in_this_run < len(batch):
                print(f"[INFO] Preparing form for next row (Action: {AFTER_SUBMIT_ACTION})...")
                if AFTER_SUBMIT_ACTION == "reload":
                    page.reload()
                    time.sleep(1)
                elif AFTER_SUBMIT_ACTION == "navigate" and FORM_URL:
                    page.goto(FORM_URL)
                    time.sleep(1)

        print("\n" + "=" * 70)
        print(f"RUN FINISHED: Processed {processed_in_this_run} record(s) in this session.")
        print("=" * 70 + "\n")


if __name__ == "__main__":
    exit_code = 0
    try:
        run_automation()
    except SystemExit as e:
        exit_code = e.code if isinstance(e.code, int) else (1 if e.code else 0)
    except Exception as e:
        print(f"\n[UNEXPECTED ERROR] {e}")
        exit_code = 1
    finally:
        # Prevent window from disappearing if opened by double-click in an interactive terminal
        if sys.stdin.isatty():
            try:
                input("\n[Script finished] Press ENTER to close this window...")
            except (EOFError, KeyboardInterrupt):
                pass
    if exit_code != 0:
        sys.exit(exit_code)
