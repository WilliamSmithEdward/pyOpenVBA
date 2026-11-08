"""When Excel raises Application.AfterCalculate, against SheetChange and SheetCalculate, in live Excel.

An application's handlers log SheetChange, SheetCalculate and AfterCalculate while a macro calculates, edits cells and writes
formulas on a new workbook. The log is read twice: at the end of the macro, and by a second macro run after it, in case Excel
raises AfterCalculate only once the macro is done; the second puts calculation back to automatic only after it has read the log.

Predictions, written before the run: AfterCalculate follows every calculation that works anything out, once, after the last
SheetCalculate, inside the macro: Application.Calculate, CalculateFull and CalculateFullRebuild with or without formulas, an edit
that a formula reads or with a volatile formula on any sheet, and a formula written. An edit with no formula anywhere, and any edit
in manual calculation, raise none. Range.Calculate and Worksheet.Calculate raise it as Application.Calculate does.

It holds live.lock while it runs, since it adds a class module to the harness's project, and touches no member that could send
mail or acts on the machine.

    python scripts/measure_after_calculate.py

writes tests/fixtures/after_calculate.json, which tests/test_excel_after_calculate.py replays.
"""

from __future__ import annotations

import json
import os
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "after_calculate.json"

LOG = "Public Log As String\nPublic Sink As Object\n"
#: An application's handlers, which log the events each case raises.
SINK = '''Public WithEvents App As Application
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "change:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
Log = Log & "calculate:" & Sh.Name & "|"
End Sub
Private Sub App_AfterCalculate()
Log = Log & "after|"
End Sub
'''

WATCH = '''Public Function Watch(how As Long) As String
    Dim wb As Object, s1 As Object, s2 As Object, made As New AfterSink
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set s1 = wb.Worksheets(1)
    s1.Name = "s1"
    Set s2 = wb.Worksheets.Add(After:=s1)
    s2.Name = "s2"
    s1.Range("A1").Value = 1
    Select Case how
        Case 2, 5, 18, 22
            s1.Range("C1").Formula = "=RAND()"
        Case 4, 6, 10, 11, 13, 14, 16, 17, 19, 23
            s1.Range("B1").Formula = "=A1*2"
        Case 20
            s2.Range("C1").Formula = "=RAND()"
            s1.Range("C1").Formula = "=RAND()"
        Case 21
            s2.Range("B1").Formula = "=s1!A1*2"
    End Select
    If how = 10 Or how = 11 Or how = 14 Then Application.Calculation = xlCalculationManual
    Set made.App = Application
    Set Sink = made
    Application.EnableEvents = True
    Log = ""
    Select Case how
        Case 1, 2, 20
            Application.Calculate
        Case 3, 4, 5, 10, 21
            s1.Range("A1").Value = 5
        Case 6
            s1.Range("B1").Calculate
        Case 7, 22
            s1.Calculate
        Case 8
            Application.CalculateFull
        Case 9
            Application.CalculateFullRebuild
        Case 11
            s1.Range("A1").Value = 5
            Log = Log & "then|"
            Application.Calculate
        Case 12
            s1.Range("B1").Formula = "=1+1"
        Case 13
            s1.Range("A1").ClearContents
        Case 14
            s1.Range("A1").Value = 5
            Log = Log & "then|"
            Application.Calculation = xlCalculationAutomatic
        Case 15
            s1.Range("A2").Value = 5
            s1.Range("A3").Value = 6
        Case 16
            s1.Range("A1").Value = 5
            s1.Range("A1").Value = 6
        Case 17
            s1.Range("A1:A3").Value = Array(7, 8, 9)
        Case 18
            s1.Range("A1").Calculate
        Case 19
            s1.Range("A1").Value = 1
        Case 23
            DoEvents
            s1.Range("A1").Value = 5
            DoEvents
    End Select
    Watch = Log & "end|"
End Function

Public Function Report() As String
    Report = Log
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Set Sink = Nothing
    Do While Workbooks.Count > 1
        Workbooks(Workbooks.Count).Close False
    Loop
    Application.EnableEvents = True
    Log = ""
End Function
'''

#: name -> (the case's number in Watch, what it does).
CASES: dict[str, tuple[int, str]] = {
    "calculate_constants": (1, "Application.Calculate with no formula"),
    "calculate_volatile": (2, "Application.Calculate with =RAND() on s1"),
    "edit_constants": (3, "an edit with no formula"),
    "edit_precedent": (4, "an edit to A1, which B1 =A1*2 reads"),
    "edit_with_volatile": (5, "an edit to A1 with =RAND() in C1"),
    "range_calculate": (6, "B1.Calculate on =A1*2"),
    "sheet_calculate_constants": (7, "s1.Calculate with no formula"),
    "calculate_full": (8, "Application.CalculateFull with no formula"),
    "calculate_full_rebuild": (9, "Application.CalculateFullRebuild with no formula"),
    "manual_edit": (10, "manual calculation; an edit to A1, which B1 reads"),
    "manual_edit_calculate": (11, "manual calculation; an edit to A1, which B1 reads, then Application.Calculate"),
    "formula_written": (12, "B1.Formula = \"=1+1\""),
    "clear_precedent": (13, "A1.ClearContents, which B1 reads"),
    "manual_then_automatic": (14, "manual; an edit to A1, which B1 reads, then Calculation = xlCalculationAutomatic"),
    "two_edits_constants": (15, "two edits with no formula"),
    "two_edits_precedent": (16, "two edits to A1, which B1 reads"),
    "block_edit_precedent": (17, "A1:A3 set at once, B1 reading A1"),
    "range_calculate_constant": (18, "A1.Calculate on a constant, =RAND() in C1"),
    "same_value_precedent": (19, "A1 set to the 1 it holds, B1 reading A1"),
    "calculate_two_sheets": (20, "Application.Calculate with =RAND() on s1 and s2"),
    "edit_precedent_other_sheet": (21, "an edit to s1!A1, which s2!B1 reads"),
    "sheet_calculate_volatile": (22, "s1.Calculate with =RAND() in C1"),
    "edit_precedent_do_events": (23, "DoEvents, an edit to A1 which B1 reads, DoEvents"),
}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    common = Path(subprocess.run(["git", "rev-parse", "--git-common-dir"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip())
    lock = (common if common.is_absolute() else ROOT / common) / "excel-coordination" / "live.lock"
    for _ in range(900):
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(1)
    else:
        raise TimeoutError("live.lock stayed held")
    rows: list[dict[str, object]] = []
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            excel.add_module("EventLog", LOG)
            excel.add_module("AfterSink", SINK, kind="class")
            excel.add_module("Watcher", WATCH)
            # The first macro run in a new project loses its module's state before the next: a run to warm it.
            assert excel.run_macro("Watch", 1, timeout=120.0).ok and excel.run_macro("Report", timeout=120.0).ok
            for name, (how, what) in CASES.items():
                inside = excel.run_macro("Watch", how, timeout=120.0)
                assert inside.ok, f"{name}: {inside.outcome}: {inside.message} {inside.error}"
                after = excel.run_macro("Report", timeout=120.0)
                assert after.ok, f"{name}: {after.outcome}: {after.message} {after.error}"
                rows.append({"name": name, "how": how, "what": what, "inside": str(inside.value), "after": str(after.value)})
                print("==", name, rows[-1]["inside"], "||", rows[-1]["after"], flush=True)
    finally:
        os.rmdir(lock)
    record = {"log": LOG, "sink": SINK, "watch": WATCH, "cases": rows, "measured_at": datetime.now(timezone.utc).isoformat()}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
