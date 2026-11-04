"""What a macro's save keeps of a file's calculation settings, when the file is opened with a workbook already open: the
application's settings do not change, so does the save write those or the file's own? Rule A: the application's. Rule
B: the file's own. Each file is tests/fixtures/silent_loss/calc_*.xlsx, opened, edited in a cell, saved and read as XML;
and the same for a file whose settings a macro changes while it is open, to see which wins.

    python scripts/measure_calc_settings_kept.py

writes tests/fixtures/silent_loss/calc_kept.json, which tests/test_excel_silent_loss.py replays.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / "tests/fixtures/silent_loss"
FILES = ["calc_manual", "calc_iteration", "calc_not_before_save", "calc_semi", "tab_color"]
#: What is done with the file open; every row's Application settings are put back.
ACTIONS = {
    "kept": "",
    "app_manual": "Application.Calculation = xlCalculationManual",
    "app_automatic": "Application.Calculation = xlCalculationAutomatic",
    "app_iteration": "Application.Iteration = True\nApplication.MaxIterations = 7",
    "app_no_iteration": "Application.Iteration = False",
    "app_before_save": "Application.CalculateBeforeSave = False",
}
PROBE = '''Function Go(ByVal file As String) As String
    Dim wb As Workbook, s As String
    Dim calc As Long, iter As Boolean, maxI As Long, maxC As Double, before As Boolean
    Application.DisplayAlerts = False
    calc = Application.Calculation: iter = Application.Iteration: maxI = Application.MaxIterations: maxC = Application.MaxChange
    before = Application.CalculateBeforeSave
    Set wb = Workbooks.Open(file)
    s = Application.Calculation & "|" & Application.Iteration & "|" & Application.MaxIterations
    {action}
    wb.Worksheets(1).Range("H20").Value = 7
    wb.Save
    wb.Close False
    Application.Calculation = calc: Application.Iteration = iter: Application.MaxIterations = maxI: Application.MaxChange = maxC
    Application.CalculateBeforeSave = before
    Go = s
End Function
'''


def calc(file: Path) -> str:
    with zipfile.ZipFile(file) as package:
        found = re.search(r"<calcPr\b[^>]*/>", package.read("xl/workbook.xml").decode("utf-8"))
    return found.group(0) if found else ""


def main() -> None:
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_kept_"))
    rows: list[dict[str, str]] = []
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            for name in FILES:
                for label, action in ACTIONS.items():
                    file = folder / f"{name}_{label}.xlsx"
                    shutil.copy(FIXTURES / f"{name}.xlsx", file)
                    original = calc(file)
                    result = excel.run_vba(PROBE.format(action=action), "Go", (str(file),), timeout=60.0)
                    row = {"file": name, "action": label, "original": original,
                           "reads": str(result.value) if result.ok else f"failed {result.outcome}", "saved": calc(file)}
                    rows.append(row)
                    print(name, label, row["reads"], row["saved"], flush=True)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    (FIXTURES / "calc_kept.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE,
                                                         "actions": ACTIONS, "rows": rows}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
