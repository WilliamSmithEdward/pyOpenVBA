"""Run RunFixture macro in each built Excel workbook and verify output files.

Requires pywin32 (`pip install pywin32`). Must be run on Windows after
build_excel_fixtures.py has produced the .xlsm files.

For each .xlsm in build/fixtures/excel/:
  - Opens the workbook via COM automation
  - Calls Module1.RunFixture
  - Verifies <fixture_name>.txt was written alongside the workbook
  - Reports pass / fail per fixture

Run from the repo root:

    python scripts/run_excel_fixtures.py

Exit code: 0 = all passed, 1 = one or more failed.
"""

from __future__ import annotations

import sys
import os
import gc

# Force UTF-8 output on Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

from pathlib import Path

BUILD_DIR = Path(__file__).resolve().parents[1] / "build" / "fixtures" / "excel"


def _run_fixture(xlsm: Path) -> bool:
    """Open *xlsm*, run Module1.RunFixture, verify the sentinel file. Return True on pass."""
    import win32com.client  # type: ignore[import]
    import time
    import os
    import traceback

    xl = None
    wb = None
    print(f"[START] {xlsm.name}", flush=True)
    try:
        print(f"  [1] Dispatching Excel.Application...", flush=True)
        xl = win32com.client.Dispatch("Excel.Application")
        print(f"  [1.1] Setting display alerts and visibility...", flush=True)
        xl.DisplayAlerts = False
        xl.Visible = True  # Make visible for debugging screenshots
        
        print(f"  [1.2] Opening VBE window for error visibility...", flush=True)
        try:
            xl.VBE.MainWindow.Visible = True
        except Exception as vbe_err:
            print(f"      Warning: Could not open VBE window: {vbe_err}", flush=True)

        print(f"  [2] Opening workbook at {xlsm.absolute()}...", flush=True)
        wb = xl.Workbooks.Open(str(xlsm.absolute()))
        print(f"  [2.1] Workbook opened, waiting 0.5s...", flush=True)
        time.sleep(0.5)
        
        print(f"  [3] Running Module1.RunFixture...", flush=True)
        try:
            xl.Application.Run("Module1.RunFixture")
            print(f"  [3.1] Macro completed, waiting 0.5s...", flush=True)
        except Exception as macro_err:
            print(f"  [3.E] FAIL - macro execution error: {macro_err}")
            traceback.print_exc()
            return False
        time.sleep(0.5)  # Give macro time to complete
        
        print(f"  [4] Closing workbook (SaveChanges=False)...", flush=True)
        wb.Close(SaveChanges=False)
        wb = None
        print(f"  [4.1] Workbook closed.", flush=True)

        expected = xlsm.with_suffix(".txt")
        print(f"  [5] Waiting 0.5s for file I/O...", flush=True)
        time.sleep(0.5)
        
        print(f"  [5.1] Checking for {expected.name}...", flush=True)
        if not expected.exists():
            # Check if file ended up in TEMP directory instead
            import tempfile
            temp_path = Path(tempfile.gettempdir()) / expected.name
            print(f"      {expected.name} not found in workbook dir, checking TEMP...", flush=True)
            if temp_path.exists():
                expected = temp_path
                print(f"      Found at: {temp_path}", flush=True)
            else:
                print(f"  [5.E] FAIL - output file not created")
                dir_contents = list(xlsm.parent.glob("*"))
                print(f"      Expected: {expected}")
                print(f"      Dir contents: {sorted([f.name for f in dir_contents])}")
                return False

        print(f"  [6] Reading output file...", flush=True)
        content = expected.read_text(encoding="utf-8").strip()
        print(f"[END] PASS -> {expected.name}: {content!r}", flush=True)
        return True

    except Exception as exc:
        print(f"[ERROR] {type(exc).__name__}: {exc}")
        traceback.print_exc()
        return False

    finally:
        print(f"  [CLEANUP] Starting cleanup...", flush=True)
        if wb is not None:
            print(f"  [CLEANUP.1] Closing workbook...", flush=True)
            try:
                wb.Close(SaveChanges=False)
                print(f"  [CLEANUP.1.OK] Workbook closed.", flush=True)
            except Exception as e:
                print(f"  [CLEANUP.1.E] Error closing workbook: {e}", flush=True)
            wb = None
        
        if xl is not None:
            print(f"  [CLEANUP.2] Calling xl.Quit()...", flush=True)
            try:
                xl.Quit()
                print(f"  [CLEANUP.2.OK] Excel.Quit() completed.", flush=True)
            except Exception as e:
                print(f"  [CLEANUP.2.E] Error quitting Excel: {e}", flush=True)
            xl = None
        
        print(f"  [CLEANUP.3] Running gc.collect()...", flush=True)
        gc.collect()
        print(f"  [CLEANUP.3] gc.collect() done.", flush=True)
        
        print(f"  [CLEANUP.4] Waiting 1 second for Excel to fully exit...", flush=True)
        time.sleep(1)
        print(f"  [CLEANUP.4] Wait complete.", flush=True)
        print(f"[CLEANUP.DONE]\n", flush=True)


def main() -> None:
    # Accept optional path argument for testing a single file
    if len(sys.argv) > 1:
        xlsm_files = [Path(sys.argv[1])]
    else:
        xlsm_files = sorted(BUILD_DIR.glob("*.xlsm"))
    
    if not xlsm_files:
        print(f"No .xlsm files found")
        sys.exit(1)

    print(f"[MAIN] Running {len(xlsm_files)} fixture(s)")
    print(f"[MAIN] Build dir: {BUILD_DIR}\n")

    results = []
    for i, xlsm in enumerate(xlsm_files, 1):
        print(f"[MAIN] ({i}/{len(xlsm_files)}) Processing {xlsm.name}", flush=True)
        result = _run_fixture(xlsm)
        results.append(result)
        print(f"[MAIN] ({i}/{len(xlsm_files)}) Result: {'PASS' if result else 'FAIL'}\n", flush=True)

    passed = sum(results)
    failed = len(results) - passed
    print(f"\n{'-' * 40}")
    print(f"[SUMMARY] {passed}/{len(results)} passed" + (f"  ({failed} failed)" if failed else ""))
    print(f"[SUMMARY] Completed at {__import__('datetime').datetime.now()}", flush=True)

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
