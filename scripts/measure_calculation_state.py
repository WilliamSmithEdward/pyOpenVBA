"""Calculation state and controls in live Excel: CalculationState, CalculationVersion, CalculationInterruptKey,
CalculateUntilAsyncQueriesDone, Workbook.ForceFullCalculation, Range.Dirty and Range.CalculateRowMajorOrder, and
which files Excel works out again as it opens them.

Each session runs in an Excel of its own, since a workbook's ForceFullCalculation set True holds every workbook to
full calculation until Excel quits. A function Counted, in the probe's module, counts how often formulas calling
it are worked out. The opens session makes a workbook, writes copies of it with another calcPr and cached values
the formulas would not give (VARIANTS), and opens each in manual calculation and then automatic, saving a copy.

    python scripts/measure_calculation_state.py

writes tests/fixtures/calculation_state.json, which tests/test_excel_calculation_state.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "calculation_state.json"

#: Each copy of the made workbook the opens session writes, by name: what replaces its calcPr ("" for none, None to
#: keep it).
VARIANTS: dict[str, str | None] = {
    "same": None, "older": '<calcPr calcId="152511"/>', "just_below": '<calcPr calcId="191027"/>',
    "lowest_kept": '<calcPr calcId="191028"/>', "newer": '<calcPr calcId="999999"/>', "no_calc_id": "<calcPr/>",
    "no_calc_pr": "", "full_on_load": '<calcPr calcId="191029" fullCalcOnLoad="1"/>',
    "tool_written": '<calcPr calcId="124519" fullCalcOnLoad="1"/>', "forced": '<calcPr calcId="191029" forceFullCalc="1"/>',
}

COMMON = r'''Public calls As Long

Public Function Counted(x As Variant) As Variant
    calls = calls + 1
    Counted = x
End Function

Public Function StateNow(x As Variant) As Variant
    StateNow = Application.CalculationState
End Function

Private Function Shown(value As Variant) As String
    If IsError(value) Then
        Shown = "Error:" & CLng(value)
    ElseIf IsNull(value) Then
        Shown = "Null"
    ElseIf IsEmpty(value) Then
        Shown = "Empty"
    Else
        Shown = TypeName(value) & ":" & CStr(value)
    End If
End Function
'''

RULES = COMMON + r'''
Private Function After(wb As Workbook, what As String) As String
    Dim ws As Worksheet, before As Long
    Set ws = wb.Worksheets(1)
    Application.Calculate
    Select Case what
        Case "isolated_value": ws.Range("Z50").Value = 5
        Case "same_value": ws.Range("A1").Value = ws.Range("A1").Value
        Case "format": ws.Range("A1").Font.Bold = True
        Case "clear_empty": ws.Range("Z60").ClearContents
        Case "clear_value": ws.Range("Z50").ClearContents
        Case "other_sheet_value": wb.Worksheets(2).Range("A1").Value = 9
        Case "formula_unread": ws.Range("Z70").Formula = "=1+1"
        Case "number_format": ws.Range("A1").NumberFormat = "0.00"
        Case "name_add": wb.Names.Add Name:="Zed", RefersTo:="=1"
        Case "row_height": ws.Rows(5).RowHeight = 30
        Case "insert_row": ws.Rows(40).Insert
        Case "sheet_add": wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
        Case "dirty_formula": ws.Range("B1").Dirty
        Case "dirty_read_value": ws.Range("A1").Dirty
        Case "dirty_isolated": ws.Range("Z99").Dirty
        Case "name_comment": wb.Names("Used").Comment = "noted"
        Case "name_hidden": wb.Names("Used").Visible = False
        Case "name_shown": wb.Names("Used").Visible = True
        Case "name_refers_unused": wb.Names("Zed").RefersTo = "=2"
        Case "name_refers_used": wb.Names("Used").RefersTo = "=5"
        Case "name_add_used": wb.Names.Add Name:="Later", RefersTo:="=3"
        Case "name_rename_unused": wb.Names("Zed").Name = "Zee"
        Case "name_delete_unused": wb.Names("Zee").Delete
        Case "name_delete_used": wb.Names("Later").Delete
    End Select
    After = what & "=" & Application.CalculationState
    before = calls
    Application.Calculate
    After = After & ":" & (calls - before)
End Function

Private Function Orders(ws As Worksheet, rowMajor As Boolean) As String
    ws.Range("A5:C9").ClearContents
    ws.Range("C5").Value = 1: ws.Range("B6").Formula = "=C5": ws.Range("A5").Formula = "=B6*10"
    ws.Range("B5").Value = 0: ws.Range("A6").Value = 0
    ws.Range("C8").Value = 1: ws.Range("B8").Formula = "=C8": ws.Range("A9").Formula = "=B8*10"
    ws.Range("A8").Value = 0: ws.Range("B9").Value = 0
    Application.Calculate
    ws.Range("C5").Value = 2
    ws.Range("C8").Value = 2
    If rowMajor Then
        ws.Range("A5:B6").CalculateRowMajorOrder
        ws.Range("A8:B9").CalculateRowMajorOrder
    Else
        ws.Range("A5:B6").Calculate
        ws.Range("A8:B9").Calculate
    End If
    Orders = "A5=" & ws.Range("A5").Value & ",B6=" & ws.Range("B6").Value & ",A9=" & ws.Range("A9").Value & _
        ",B8=" & ws.Range("B8").Value
End Function

Private Function TryKey(value As Variant) As String
    On Error Resume Next
    Application.CalculationInterruptKey = value
    If Err.Number <> 0 Then TryKey = "E" & Err.Number Else TryKey = CStr(Application.CalculationInterruptKey)
    Err.Clear
    Application.CalculationInterruptKey = xlAnyKey
End Function

Public Function Probe() As String
    Dim out As String, ws As Worksheet, wb As Workbook, before As Long, plain As Workbook, what As Variant
    Dim late As Object, app As Object, got As Variant
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)
    If wb.Worksheets.Count < 2 Then wb.Worksheets.Add After:=ws
    ws.Activate
    out = "state=" & Application.CalculationState
    out = out & "|application_version=" & Shown(Application.CalculationVersion)
    out = out & "|book_version=" & Shown(wb.CalculationVersion)
    Set plain = Workbooks.Add
    out = out & "|new_book_version=" & Shown(plain.CalculationVersion) & ",force=" & Shown(plain.ForceFullCalculation)
    plain.Worksheets(1).Range("A1").Value = 1
    out = out & "|no_formulas_value=" & Application.CalculationState
    plain.Close False
    wb.Activate
    out = out & "|interrupt=" & Shown(Application.CalculationInterruptKey)
    out = out & "|keys=" & TryKey(xlEscKey) & "," & TryKey(xlNoKey) & "," & TryKey(5) & "," & TryKey(-1) & "," & _
        TryKey("1") & "," & TryKey(True) & "," & TryKey(Null) & "," & TryKey(3) & "," & TryKey(2.6) & "," & _
        TryKey(1.4) & "," & TryKey(0.5) & "," & TryKey(1.5) & "," & TryKey("x") & "," & TryKey(Empty) & "," & _
        TryKey(False) & "," & TryKey("") & "," & TryKey(-0.4)
    ws.Range("A1").Value = 1
    ws.Range("B1").Formula = "=Counted(A1)"
    ws.Range("D1").Formula = "=Counted(1)"
    ws.Range("C1").Formula = "=StateNow(A1)"
    wb.Names.Add Name:="Used", RefersTo:="=1"
    ws.Range("F1").Formula = "=Counted(Used)"
    ws.Range("G1").Formula = "=Counted(Later)"
    out = out & "|while_calculating=" & ws.Range("C1").Value
    before = calls
    ws.Range("B1").Dirty
    out = out & "|automatic_dirty calls=" & (calls - before) & ",state=" & Application.CalculationState
    before = calls
    ws.Range("Z99").Dirty
    out = out & "|automatic_dirty_isolated calls=" & (calls - before)
    Set late = ws.Range("B1")
    got = late.Dirty
    out = out & "|dirty_answers=" & Shown(got)
    Set app = Application
    before = calls
    got = app.CalculateUntilAsyncQueriesDone
    out = out & "|automatic_until_async=" & Shown(got) & ",calls=" & (calls - before)
    Application.Calculation = xlCalculationManual
    out = out & "|manual=" & Application.CalculationState
    out = out & "|"
    For Each what In Array("isolated_value", "same_value", "format", "clear_empty", "clear_value", _
                           "other_sheet_value", "formula_unread", "number_format", "name_add", "row_height", _
                           "insert_row", "sheet_add", "dirty_formula", "dirty_read_value", "dirty_isolated", _
                           "name_comment", "name_hidden", "name_shown", "name_refers_unused", "name_refers_used", _
                           "name_add_used", "name_rename_unused", "name_delete_unused", "name_delete_used")
        out = out & After(wb, CStr(what)) & ","
    Next
    Application.Calculate
    before = calls
    ws.Range("E1").Formula = "=Counted(2)"
    out = out & "|entered calls=" & (calls - before) & ",state=" & Application.CalculationState
    Application.Calculate
    out = out & ",then=" & (calls - before) & ",state=" & Application.CalculationState
    ws.Range("A1").Value = 2
    out = out & "|edited=" & Application.CalculationState & ",B1=" & ws.Range("B1").Value
    before = calls
    got = app.CalculateUntilAsyncQueriesDone
    out = out & "|manual_until_async=" & Shown(got) & ",calls=" & (calls - before) & ",B1=" & _
        ws.Range("B1").Value & ",state=" & Application.CalculationState
    before = calls
    got = late.CalculateRowMajorOrder
    out = out & "|row_major_answers=" & Shown(got) & ",calls=" & (calls - before) & ",state=" & _
        Application.CalculationState
    before = calls
    got = late.Calculate
    out = out & "|calculate_answers=" & Shown(got) & ",calls=" & (calls - before)
    Application.Calculate
    out = out & "|column_order " & Orders(ws, False)
    out = out & "|row_order " & Orders(ws, True)
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic
    ws.Range("A1:Z99").Clear
    Probe = out
End Function
'''

OPENS = COMMON + r'''
Public Function MakeFile(path As String) As String
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Worksheets(1).Range("A1").Value = 2
    wb.Worksheets(1).Range("A2").Formula = "=A1*10"
    wb.Worksheets(1).Range("A3").Formula = "=A2+1"
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    wb.Close False
    MakeFile = "made"
End Function

Public Function Opened(path As String, copy As String) As String
    Dim wb As Workbook, out As String
    Application.Calculation = xlCalculationManual
    Set wb = Workbooks.Open(path)
    out = "manual version=" & wb.CalculationVersion & ",state=" & Application.CalculationState & _
        ",A2=" & wb.Worksheets(1).Range("A2").Value & ",A3=" & wb.Worksheets(1).Range("A3").Value & _
        ",saved=" & wb.Saved
    Application.Calculate
    out = out & " calculated version=" & wb.CalculationVersion & ",state=" & Application.CalculationState & _
        ",A2=" & wb.Worksheets(1).Range("A2").Value & ",saved=" & wb.Saved
    wb.Close False
    Application.Calculation = xlCalculationAutomatic
    Set wb = Workbooks.Open(path)
    out = out & " | automatic version=" & wb.CalculationVersion & ",state=" & Application.CalculationState & _
        ",A2=" & wb.Worksheets(1).Range("A2").Value & ",A3=" & wb.Worksheets(1).Range("A3").Value & _
        ",saved=" & wb.Saved & ",force=" & wb.ForceFullCalculation
    Application.DisplayAlerts = False
    wb.SaveCopyAs copy
    wb.Close False
    Opened = out
End Function
'''

FORCED = COMMON + r'''
Private Function TryForce(wb As Workbook, value As Variant) As String
    On Error Resume Next
    wb.ForceFullCalculation = False
    wb.ForceFullCalculation = value
    If Err.Number <> 0 Then TryForce = "E" & Err.Number Else TryForce = CStr(wb.ForceFullCalculation)
    Err.Clear
    wb.ForceFullCalculation = False
End Function

Public Function Probe(folder As String) As String
    Dim out As String, ws As Worksheet, wb As Workbook, before As Long
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = 1
    ws.Range("B1").Formula = "=Counted(A1)"
    ws.Range("D1").Formula = "=Counted(1)"
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    wb.Worksheets(1).Range("A1").Value = 2
    wb.Worksheets(1).Range("A2").Formula = "=A1*10"
    wb.SaveAs folder & "\plain.xlsx", 51
    wb.Close False
    Set wb = Workbooks.Add
    wb.Saved = True
    out = "saved_before=" & wb.Saved
    out = out & "|values=" & TryForce(wb, True) & "," & TryForce(wb, 1) & "," & TryForce(wb, "True") & "," & _
        TryForce(wb, Null) & "," & TryForce(wb, 0) & "," & TryForce(wb, "x") & "," & TryForce(wb, Empty) & "," & _
        TryForce(wb, 2)
    out = out & "|saved_after=" & wb.Saved & ",state=" & Application.CalculationState
    wb.Close False
    ThisWorkbook.Activate
    before = calls
    ws.Range("Z1").Value = 1
    out = out & "|automatic_unrelated_edit calls=" & (calls - before) & ",state=" & Application.CalculationState
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Saved = True
    before = calls
    Application.Calculate
    out = out & "|manual_calculate calls=" & (calls - before) & ",state=" & Application.CalculationState & _
        ",saved=" & ThisWorkbook.Saved
    before = calls
    ws.Calculate
    out = out & "|sheet_calculate calls=" & (calls - before)
    Application.CalculateFull
    out = out & "|after_full state=" & Application.CalculationState
    Application.CalculateFullRebuild
    out = out & "|after_rebuild state=" & Application.CalculationState
    Application.Calculation = xlCalculationAutomatic
    Set wb = Workbooks.Open(folder & "\plain.xlsx")
    out = out & "|plain_opened saved=" & wb.Saved & ",state=" & Application.CalculationState & ",version=" & _
        wb.CalculationVersion & ",A2=" & wb.Worksheets(1).Range("A2").Value
    wb.Close False
    Set wb = Workbooks.Add
    wb.ForceFullCalculation = True
    wb.SaveAs folder & "\forced.xlsx", 51
    wb.Close False
    out = out & "|closed state=" & Application.CalculationState
    before = calls
    ws.Range("Z2").Value = 1
    out = out & "|closed_unrelated_edit calls=" & (calls - before)
    Set wb = Workbooks.Open(folder & "\forced.xlsx")
    out = out & "|reopened force=" & wb.ForceFullCalculation & ",state=" & Application.CalculationState
    wb.ForceFullCalculation = False
    wb.SaveAs folder & "\unforced.xlsx", 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\unforced.xlsx")
    out = out & "|unforced force=" & wb.ForceFullCalculation
    wb.Close False
    ws.Range("A1:Z9").Clear
    Probe = out
End Function
'''

FIRST = COMMON + r'''
Public Function Probe(path As String) As String
    Dim wb As Workbook, out As String, ws As Worksheet, before As Long
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value = 1
    ws.Range("B1").Formula = "=Counted(A1)"
    ws.Range("D1").Formula = "=Counted(1)"
    out = "start=" & Application.CalculationState
    Set wb = Workbooks.Open(path)
    out = out & ",opened=" & Application.CalculationState & ",force=" & wb.ForceFullCalculation & ",saved=" & wb.Saved
    ThisWorkbook.Activate
    before = calls
    ws.Range("Z1").Value = 1
    out = out & ",unrelated_edit calls=" & (calls - before)
    wb.Close False
    out = out & ",closed=" & Application.CalculationState
    ws.Range("A1:Z9").Clear
    Probe = out
End Function
'''


def crafted(made: Path, target: Path, calc: str | None) -> None:
    """A copy of ``made`` with ``calc`` for its calcPr, and the cached values 99 and 77 its formulas would not give."""
    with zipfile.ZipFile(made) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as copy:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/workbook.xml" and calc is not None:
                data = re.sub(r"<calcPr\b[^>]*?(?:/>|>.*?</calcPr>)", calc, data.decode("utf-8"), flags=re.DOTALL)
                if calc and "<calcPr" not in data:
                    data = data.replace("</workbook>", calc + "</workbook>")
                data = data.encode("utf-8")
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                text = re.sub(r'(<c r="A2"[^>]*>\s*<f>A1\*10</f>\s*<v>)20(</v>)', r"\g<1>99\g<2>", text)
                text = re.sub(r'(<c r="A3"[^>]*>\s*<f>A2\+1</f>\s*<v>)21(</v>)', r"\g<1>77\g<2>", text)
                data = text.encode("utf-8")
            copy.writestr(item, data)


def saved_calc(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        found = re.search(r"<calcPr\b[^>]*>", package.read("xl/workbook.xml").decode("utf-8"))
    return found.group(0) if found else ""


def main() -> None:
    folder = Path(tempfile.mkdtemp(prefix="calculation_state_"))
    record: dict[str, object] = {"rules": RULES, "opens": OPENS, "forced": FORCED, "first": FIRST,
                                 "variants": VARIANTS}
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(RULES, "Probe", timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            record["rules_answer"] = str(result.value)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = folder / "made.xlsx"
            assert excel.run_vba(OPENS, "MakeFile", [str(made)], timeout=120.0).ok
            opens: dict[str, list[str]] = {}
            for name, calc in VARIANTS.items():
                target, copy = folder / f"{name}.xlsx", folder / f"{name}_copy.xlsx"
                crafted(made, target, calc)
                result = excel.run_vba(OPENS, "Opened", [str(target), str(copy)], timeout=120.0)
                assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
                opens[name] = [str(result.value), saved_calc(copy)]
            record["opens_answer"] = opens
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(FORCED, "Probe", [str(folder)], timeout=300.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            record["forced_answer"] = str(result.value)
            record["forced_file"] = saved_calc(folder / "forced.xlsx")
            record["unforced_file"] = saved_calc(folder / "unforced.xlsx")
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(FIRST, "Probe", [str(folder / "forced.xlsx")], timeout=120.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            record["first_answer"] = str(result.value)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for key in ("rules_answer", "forced_answer", "forced_file", "unforced_file", "first_answer"):
        print(key, str(record[key]).replace("|", "\n  "))
    for name, (answer, calc) in record["opens_answer"].items():  # type: ignore[union-attr]
        print(name, answer, calc)


if __name__ == "__main__":
    main()
