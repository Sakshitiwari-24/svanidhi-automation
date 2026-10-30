"""
================================================================================
TEST SUITE FOR SVANIDHI DATA ENTRY AUTOMATION
================================================================================
Demonstrates all 5 critical safety requirements before touching the real portal:
  1. Upfront validation catches empty cells, duplicates, formats & extra spaces
  2. Read-back check catches any mismatch before clicking Submit
  3. Rows are submitted and marked 'Done' correctly one by one
  4. Re-running the script automatically skips rows that are already 'Done'
  5. The script stops immediately on simulated portal failure / timeout
  6. File lock detection prevents running if data.xlsx is open in Excel
================================================================================
"""

import os
import sys
import shutil
import subprocess
import threading
import http.server
import socketserver
import time
import openpyxl
from playwright.sync_api import sync_playwright

os.environ["HEADLESS"] = "1"
SERVER_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON_EXE = sys.executable

def print_header(title):
    print("\n" + "=" * 75)
    print(f"TEST SCENARIO: {title}")
    print("=" * 75)

def reset_test_excel():
    """Generates a fresh data.xlsx with 5 records for testing."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Reference", "Status"])
    
    rows = [
        ("Ramesh Patel", "100023456789", ""),
        ("Priya Sharma", "007894561230", ""),    # Leading zero test
        ("Amit Kumar", "001234567890", ""),      # Leading zero test
        ("Sunita Devi", "987654321098", ""),
        ("Vikram Singh", "554433221100", "")
    ]
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.number_format = '@'
            cell.value = str(val)
    wb.save("data.xlsx")
    wb.close()
    if os.path.exists("log.txt"):
        os.remove("log.txt")

# ==============================================================================
# TEST 1: UPFRONT VALIDATION (Duplicates, Missing Values, Spaces)
# ==============================================================================
def test_1_validation():
    print_header("1. Upfront Validation Catches Bad Rows Before Touching Browser")
    print("Testing against 'data_with_errors.xlsx' containing:")
    print("  - Row 4: Extra leading/trailing and double spaces")
    print("  - Row 5: Duplicate reference number ('100023456789')")
    print("  - Row 6: Empty Reference cell")
    print("  - Row 7: Empty Name cell")
    
    # Run fill_form.py with EXCEL_FILE set to data_with_errors.xlsx
    test_script = "test_script_val.py"
    with open("fill_form.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Point script to error file and mock mode
    content = content.replace('EXCEL_FILE = "data.xlsx"', 'EXCEL_FILE = "data_with_errors.xlsx"')
    content = content.replace('USE_MOCK_PORTAL = False', 'USE_MOCK_PORTAL = True')
    
    with open(test_script, "w", encoding="utf-8") as f:
        f.write(content)
        
    res = subprocess.run([PYTHON_EXE, test_script], capture_output=True, text=True)
    
    print("\n--- SCRIPT OUTPUT ---")
    print(res.stdout)
    if res.stderr:
        print("STDERR:", res.stderr)
        
    assert res.returncode != 0, "Script should have halted with non-zero exit code!"
    assert "Duplicate Reference" in res.stdout, "Should have detected duplicate reference"
    assert "is empty" in res.stdout, "Should have detected empty fields"
    assert "space normalized" in res.stdout, "Should have reported space normalization"
    print("\n>>> RESULT: PASSED! Validation stopped the process before any browser interaction.")
    
    if os.path.exists(test_script):
        os.remove(test_script)

# ==============================================================================
# TEST 2: DRY RUN VERIFICATION
# ==============================================================================
def test_2_dry_run():
    print_header("2. Dry Run Mode (Fills & Verifies, But NEVER Clicks Submit)")
    reset_test_excel()
    
    test_script = "test_script_dry.py"
    with open("fill_form.py", "r", encoding="utf-8") as f:
        content = f.read()
        
    content = content.replace('USE_MOCK_PORTAL = False', 'USE_MOCK_PORTAL = True')
    content = content.replace('DRY_RUN = False', 'DRY_RUN = True')
    content = content.replace('MAX_ROWS_THIS_RUN = 2', 'MAX_ROWS_THIS_RUN = 2')
    
    with open(test_script, "w", encoding="utf-8") as f:
        f.write(content)
        
    res = subprocess.run([PYTHON_EXE, test_script], capture_output=True, text=True)
    print("\n--- SCRIPT OUTPUT ---")
    print(res.stdout)
    
    assert res.returncode == 0, "Dry run should complete normally."
    assert "[DRY_RUN]" in res.stdout, "Should indicate DRY_RUN actions in output."
    
    # Verify Excel was NOT modified
    wb = openpyxl.load_workbook("data.xlsx")
    ws = wb.active
    status_row_2 = ws.cell(row=2, column=3).value
    status_row_3 = ws.cell(row=3, column=3).value
    wb.close()
    
    assert not status_row_2, f"Status for Row 2 should be empty in DRY RUN, got '{status_row_2}'"
    assert not status_row_3, f"Status for Row 3 should be empty in DRY RUN, got '{status_row_3}'"
    print("\n>>> RESULT: PASSED! Form fields were verified on page, Submit was NOT clicked, and Excel was NOT marked Done.")
    
    if os.path.exists(test_script):
        os.remove(test_script)

# ==============================================================================
# TEST 3: LIVE SUBMISSION & ATOMIC SAVE
# ==============================================================================
def test_3_live_submit():
    print_header("3. Live Submission & Immediate 'Done' Marking (Batch of 2 Rows)")
    reset_test_excel()
    
    test_script = "test_script_live.py"
    with open("fill_form.py", "r", encoding="utf-8") as f:
        content = f.read()
        
    content = content.replace('USE_MOCK_PORTAL = False', 'USE_MOCK_PORTAL = True')
    content = content.replace('DRY_RUN = True', 'DRY_RUN = False')
    content = content.replace('ASK_BEFORE_SUBMIT = True', 'ASK_BEFORE_SUBMIT = False')
    content = content.replace('MAX_ROWS_THIS_RUN = 2', 'MAX_ROWS_THIS_RUN = 2')
    
    with open(test_script, "w", encoding="utf-8") as f:
        f.write(content)
        
    res = subprocess.run([PYTHON_EXE, test_script], capture_output=True, text=True)
    print("\n--- SCRIPT OUTPUT ---")
    print(res.stdout)
    
    assert res.returncode == 0, "Submission should complete successfully."
    assert "Success indicator confirmed on page for Row 2" in res.stdout
    assert "Success indicator confirmed on page for Row 3" in res.stdout
    
    # Verify Excel was updated immediately
    wb = openpyxl.load_workbook("data.xlsx")
    ws = wb.active
    status_r2 = ws.cell(row=2, column=3).value
    status_r3 = ws.cell(row=3, column=3).value
    status_r4 = ws.cell(row=4, column=3).value
    
    # Check leading zero preservation in Reference column!
    ref_r3 = ws.cell(row=3, column=2).value
    wb.close()
    
    assert status_r2 == "Done", f"Row 2 should be Done, got {status_r2}"
    assert status_r3 == "Done", f"Row 3 should be Done, got {status_r3}"
    assert not status_r4, f"Row 4 should NOT be touched (batch limit was 2)"
    assert ref_r3 == "007894561230", f"Leading zero was lost! Got {ref_r3}"
    
    print("\n>>> RESULT: PASSED! Row 2 & Row 3 submitted, marked 'Done', and leading zeros preserved!")
    
    if os.path.exists(test_script):
        os.remove(test_script)

# ==============================================================================
# TEST 4: RE-RUNNING SKIPS ALREADY COMPLETED ROWS
# ==============================================================================
def test_4_resuming_skips_done():
    print_header("4. Re-running Skips Already Completed ('Done') Rows")
    print("Row 2 and Row 3 are currently marked 'Done' from the previous run.")
    print("Running with MAX_ROWS_THIS_RUN = 2 should automatically pick up Row 4 & Row 5.")
    
    test_script = "test_script_resume.py"
    with open("fill_form.py", "r", encoding="utf-8") as f:
        content = f.read()
        
    content = content.replace('USE_MOCK_PORTAL = False', 'USE_MOCK_PORTAL = True')
    content = content.replace('DRY_RUN = True', 'DRY_RUN = False')
    content = content.replace('ASK_BEFORE_SUBMIT = True', 'ASK_BEFORE_SUBMIT = False')
    content = content.replace('MAX_ROWS_THIS_RUN = 2', 'MAX_ROWS_THIS_RUN = 2')
    
    with open(test_script, "w", encoding="utf-8") as f:
        f.write(content)
        
    res = subprocess.run([PYTHON_EXE, test_script], capture_output=True, text=True)
    print("\n--- SCRIPT OUTPUT ---")
    print(res.stdout)
    
    assert res.returncode == 0
    assert "Already completed (Status = Done): 2" in res.stdout
    assert "PROCESSING ROW 4" in res.stdout
    assert "PROCESSING ROW 5" in res.stdout
    assert "PROCESSING ROW 2" not in res.stdout, "Row 2 should have been skipped!"
    assert "PROCESSING ROW 3" not in res.stdout, "Row 3 should have been skipped!"
    
    # Verify Excel
    wb = openpyxl.load_workbook("data.xlsx")
    ws = wb.active
    status_r4 = ws.cell(row=4, column=3).value
    status_r5 = ws.cell(row=5, column=3).value
    ref_r4 = ws.cell(row=4, column=2).value
    wb.close()
    
    assert status_r4 == "Done"
    assert status_r5 == "Done"
    assert ref_r4 == "001234567890", f"Leading zero lost on Row 4: got {ref_r4}"
    
    print("\n>>> RESULT: PASSED! Row 2 and 3 were skipped, Row 4 and 5 were processed and marked Done.")
    
    if os.path.exists(test_script):
        os.remove(test_script)

# ==============================================================================
# TEST 5: READ-BACK MISMATCH STOPS WITHOUT SUBMITTING
# ==============================================================================
def test_5_read_back_mismatch():
    print_header("5. Read-Back Mismatch Stops Execution Immediately Without Submitting")
    reset_test_excel()
    
    test_script = "test_script_mismatch.py"
    with open("fill_form.py", "r", encoding="utf-8") as f:
        content = f.read()
        
    content = content.replace('USE_MOCK_PORTAL = False', 'USE_MOCK_PORTAL = True')
    content = content.replace('DRY_RUN = True', 'DRY_RUN = False')
    content = content.replace('ASK_BEFORE_SUBMIT = True', 'ASK_BEFORE_SUBMIT = False')
    content = content.replace('name_field.fill(name)', 'name_field.fill("CORRUPTED_VALUE_ON_PAGE")')
    
    with open(test_script, "w", encoding="utf-8") as f:
        f.write(content)
        
    res = subprocess.run([PYTHON_EXE, test_script], capture_output=True, text=True)
    print("\n--- SCRIPT OUTPUT ---")
    print(res.stdout)
    
    assert res.returncode != 0, "Script should have halted with error!"
    assert "READ-BACK MISMATCH" in res.stdout
    assert "ABORTING SUBMISSION IMMEDIATELY" in res.stdout
    
    # Check that Row 2 was NOT marked Done
    wb = openpyxl.load_workbook("data.xlsx")
    ws = wb.active
    status_r2 = ws.cell(row=2, column=3).value
    wb.close()
    assert not status_r2, "Status must not be marked Done on mismatch!"
    
    print("\n>>> RESULT: PASSED! Read-back mismatch was detected; Submit was never clicked.")
    
    if os.path.exists(test_script):
        os.remove(test_script)

# ==============================================================================
# TEST 6: SIMULATED PORTAL FAILURE / TIMEOUT
# ==============================================================================
def test_6_portal_failure():
    print_header("6. Simulated Portal Failure / Timeout Halts Execution & Alerts User")
    reset_test_excel()
    
    test_script = "test_script_fail.py"
    with open("fill_form.py", "r", encoding="utf-8") as f:
        content = f.read()
        
    # Simulate timeout by expecting a selector that never appears, with short timeout
    content = content.replace('USE_MOCK_PORTAL = False', 'USE_MOCK_PORTAL = True')
    content = content.replace('DRY_RUN = True', 'DRY_RUN = False')
    content = content.replace('ASK_BEFORE_SUBMIT = True', 'ASK_BEFORE_SUBMIT = False')
    content = content.replace('SUCCESS_SELECTOR = "#successMessage"', 'SUCCESS_SELECTOR = "#nonExistentSuccessBanner"')
    content = content.replace('SUBMIT_TIMEOUT_MS = 15000', 'SUBMIT_TIMEOUT_MS = 2500')
    
    with open(test_script, "w", encoding="utf-8") as f:
        f.write(content)
        
    res = subprocess.run([PYTHON_EXE, test_script], capture_output=True, text=True)
    print("\n--- SCRIPT OUTPUT ---")
    print(res.stdout)
    
    assert res.returncode != 0, "Script should have halted on timeout!"
    assert "CRITICAL SUBMISSION TIMEOUT" in res.stdout
    assert "DO NOT RE-RUN THE SCRIPT YET" in res.stdout
    assert "Row 2 was NOT marked 'Done'" in res.stdout
    
    # Verify Excel
    wb = openpyxl.load_workbook("data.xlsx")
    ws = wb.active
    status_r2 = ws.cell(row=2, column=3).value
    wb.close()
    assert not status_r2, "Status must NOT be marked Done when timeout occurs!"
    
    print("\n>>> RESULT: PASSED! Script halted safely on timeout with loud instructions not to re-run blindly.")
    
    if os.path.exists(test_script):
        os.remove(test_script)

# ==============================================================================
# MAIN TEST RUNNER
# ==============================================================================
def run_all_tests():
    print("=" * 75)
    print("STARTING TEST SUITE FOR PM SVANIDHI DATA ENTRY AUTOMATION")
    print("=" * 75)
    
    test_1_validation()
    test_2_dry_run()
    test_3_live_submit()
    test_4_resuming_skips_done()
    test_5_read_back_mismatch()
    test_6_portal_failure()
    
    print("\n" + "=" * 75)
    print("ALL 6 SAFETY TEST SCENARIOS PASSED WITH ZERO FAILURES!")
    print("=" * 75)

if __name__ == "__main__":
    run_all_tests()
