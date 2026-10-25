"""The order Excel raises Calculate over several sheets, the numbers it saves the sheets with, and the calculation
that ends a pivot refresh, in live Excel.

A macro makes five sheets one after another but places them out of order
(tabs s5 s3 s1 s2 s4), or moves or deletes one, and an application's
handlers log the Calculate events of Application.Calculate, an edit and
Worksheet.Calculate over volatile formulas and over formulas written while
events were off, which leave each sheet a notification held for the next
calculation; once with the workbook saved and opened again. Then a pivot
table on s2 is refreshed: with held notifications, with a volatile formula
or one reading the table written with events on, in manual calculation and
by RefreshAll over two caches. Last, workbooks whose sheets were added
before others, deleted or copied, some saved and opened again first, are
saved, and the name and sheetId of each sheet their workbook part lists
are kept.

    python scripts/measure_calculate_order.py

writes tests/fixtures/calculate_order.json, which
tests/test_excel_calculate_order.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "calculate_order.json"
FILES = ROOT / "tests" / "fixtures"

#: An application's handlers, which log the events each case raises.
SINK = '''Public WithEvents App As Application
Private Sub App_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
Log = Log & "update:" & Sh.Name & "/" & Target.Name & "|"
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "change:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
Log = Log & "calculate:" & Sh.Name & "|"
End Sub
'''

WATCH = '''Private Function Sheet(wb As Object, name As String, before As String, after As String) As Object
    Dim ws As Object
    If before <> "" Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(before))
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(after))
    End If
    ws.Name = name
    Set Sheet = ws
End Function

Private Function Book(kind As Long) As Object
    Dim wb As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "s1"
    Select Case kind
        Case 1
            Sheet wb, "s2", "", "s1"
            Sheet wb, "s3", "s1", ""
            Sheet wb, "s4", "", "s2"
            Sheet wb, "s5", "s3", ""
        Case 2, 3
            Sheet wb, "s2", "", "s1"
            Sheet wb, "s3", "", "s2"
            Sheet wb, "s4", "", "s3"
            If kind = 2 Then
                wb.Worksheets("s1").Move After:=wb.Worksheets("s4")
            Else
                wb.Worksheets("s3").Move Before:=wb.Worksheets("s1")
            End If
        Case 4
            Sheet wb, "s2", "", "s1"
            Sheet wb, "s3", "", "s2"
            wb.Worksheets("s1").Delete
            Sheet wb, "s4", "s2", ""
    End Select
    Set Book = wb
End Function

Private Sub Formulas(wb As Object, formula As String)
    Dim ws As Object
    For Each ws In wb.Worksheets
        ws.Range("Z1").Formula = formula
    Next
End Sub

Private Sub Lists(ws As Object)
    Dim r As Long
    ws.Range("A1:B1").Value = Array("Region", "Amount")
    For r = 0 To 7
        ws.Cells(r + 2, 1).Value = Array("East", "North", "South", "West")(r Mod 4)
        ws.Cells(r + 2, 2).Value = (r + 1) * 10
    Next
    ws.Range("H1:I1").Value = Array("Kind", "Count")
    For r = 0 To 3
        ws.Cells(r + 2, 8).Value = Array("New", "Used")(r Mod 2)
        ws.Cells(r + 2, 9).Value = (r + 1) * 3
    Next
End Sub

Private Sub Report(wb As Object, source As String, at As Object, name As String, rows As String, values As String)
    Dim pt As Object
    Set pt = wb.PivotCaches.Create(xlDatabase, source).CreatePivotTable(at, name)
    pt.PivotFields(rows).Orientation = xlRowField
    pt.AddDataField pt.PivotFields(values), "Sum of " & values, xlSum
End Sub

Public Function Watch(how As Long, folder As String) As String
    Dim wb As Object, s2 As Object, sink As New CalculateSink, out As String, name As String
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic
    Select Case how
        Case 6
            Set wb = Book(2)
        Case 7
            Set wb = Book(3)
        Case 8
            Set wb = Book(4)
        Case Else
            Set wb = Book(1)
    End Select
    Set s2 = wb.Worksheets("s2")
    If how >= 10 Then
        Lists s2
        Report wb, "s2!R1C1:R9C2", s2.Range("E3"), "T", "Region", "Amount"
        If how = 16 Then Report wb, "s2!R1C8:R5C9", wb.Worksheets("s4").Range("A3"), "U", "Kind", "Count"
        If how >= 18 Then
            With s2.PivotTables("T").PivotCache.CreatePivotTable(wb.Worksheets("s4").Range("A3"), "V")
                .PivotFields("Region").Orientation = xlRowField
                .AddDataField .PivotFields("Amount"), "Sum of Amount", xlSum
            End With
        End If
    End If
    Select Case how
        Case 3 To 8, 10, 16, 17, 18
            Formulas wb, "=1+1"
        Case 9
            Formulas wb, "=RAND()"
        Case 11, 12, 19
            s2.Range("Z1").Formula = "=2+2"
    End Select
    If how = 9 Or how = 17 Then
        name = "calculate_order.xlsx"
        wb.SaveAs folder & "\\" & name, 51
        wb.Close False
        Set wb = Workbooks.Open(folder & "\\" & name)
        Set s2 = wb.Worksheets("s2")
    End If
    If how >= 10 Then s2.Range("B3").Value = 1000
    If how = 16 Then s2.Range("I3").Value = 100
    If how = 11 Or how = 19 Then Application.Calculation = xlCalculationManual
    Set sink.App = Application
    Application.EnableEvents = True
    Select Case how
        Case 1, 2
            Formulas wb, "=RAND()"
        Case 13
            wb.Worksheets("s4").Range("Z1").Formula = "=RAND()"
        Case 14
            wb.Worksheets("s4").Range("Z1").Formula = "=SUM(s2!F4:F7)"
        Case 15
            wb.Worksheets("s4").Range("Z1").Formula = "=GETPIVOTDATA(""Sum of Amount"",s2!E3)"
    End Select
    Log = ""
    Select Case how
        Case 1, 3, 6, 7, 8, 9
            Application.Calculate
        Case 2, 4
            s2.Range("B1").Value = 7
        Case 5
            s2.Calculate
        Case 10 To 15, 17
            s2.PivotTables("T").PivotCache.Refresh
        Case 16
            wb.RefreshAll
        Case 18, 19
            s2.PivotTables("T").ClearTable
    End Select
    If how = 11 Or how = 12 Or how = 19 Then
        Log = Log & "then|"
        Application.Calculate
    End If
    out = Log
    Application.EnableEvents = False
    If how = 14 Or how = 15 Then out = out & "value=" & wb.Worksheets("s4").Range("Z1").Value
    Application.Calculation = xlCalculationAutomatic
    wb.Close False
    Watch = out
End Function
'''

#: name -> (the case's number in Watch, what it does).
CASES: dict[str, tuple[int, str]] = {
    "calculate_volatile": (1, "=RAND() on each sheet, written with events on; Application.Calculate"),
    "edit_volatile": (2, "=RAND() on each sheet, written with events on; an edit to s2!B1"),
    "calculate_held": (3, "=1+1 on each sheet, written with events off; Application.Calculate"),
    "edit_held": (4, "=1+1 on each sheet, written with events off; an edit to s2!B1"),
    "sheet_calculate_held": (5, "=1+1 on each sheet, written with events off; Worksheets(\"s2\").Calculate"),
    "moved_after": (6, "s1 to s4 one after another, s1 moved after s4; held =1+1; Application.Calculate"),
    "moved_before": (7, "s1 to s4 one after another, s3 moved before s1; held =1+1; Application.Calculate"),
    "deleted": (8, "s1 to s3 one after another, s1 deleted, s4 before s2; held =1+1; Application.Calculate"),
    "reopened": (9, "=RAND() on each sheet, the workbook saved and opened again; Application.Calculate"),
    "refresh_held": (10, "a table on s2 refreshed, =1+1 held on each sheet"),
    "refresh_held_manual": (11, "s2!Z1 =2+2 held; manual calculation; the table refreshed, then Application.Calculate"),
    "refresh_held_automatic": (12, "s2!Z1 =2+2 held; the table refreshed, then Application.Calculate"),
    "refresh_volatile": (13, "s4!Z1 =RAND() written with events on; the table refreshed"),
    "refresh_dependent": (14, "s4!Z1 =SUM(s2!F4:F7) over the table, written with events on; the table refreshed"),
    "refresh_getpivotdata": (15, "s4!Z1 =GETPIVOTDATA over the table, written with events on; the table refreshed"),
    "refresh_all_once": (16, "a second cache's table on s4; =1+1 held on each sheet; RefreshAll"),
    "refresh_reopened": (17, "=1+1 on each sheet, the workbook saved and opened again; the table refreshed"),
    "clear_table": (18, "a second table V of the cache on s4; =1+1 held on each sheet; T.ClearTable"),
    "clear_table_manual": (19, "a second table V of the cache on s4; s2!Z1 =2+2 held; manual calculation; "
                               "T.ClearTable, then Application.Calculate"),
}


MAKE = '''Public Function Make(how As Long, folder As String) As String
    Dim wb As Object, name As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Select Case how
        Case 1
            wb.Worksheets.Add
            wb.Worksheets.Add
            wb.Worksheets.Add
        Case 2
            wb.Worksheets(1).Name = "s1"
            wb.Worksheets.Add(After:=wb.Worksheets("s1")).Name = "s2"
            wb.Worksheets.Add(Before:=wb.Worksheets("s1")).Name = "s3"
            wb.Worksheets.Add(After:=wb.Worksheets("s2")).Name = "s4"
            wb.Worksheets.Add(Before:=wb.Worksheets("s3")).Name = "s5"
        Case 3
            wb.Worksheets(1).Name = "s1"
            wb.Worksheets.Add(After:=wb.Worksheets("s1")).Name = "s2"
            wb.Worksheets.Add(After:=wb.Worksheets("s2")).Name = "s3"
            wb.Worksheets("s3").Delete
            wb.Worksheets.Add(After:=wb.Worksheets("s2")).Name = "s4"
        Case 4
            wb.Worksheets(1).Name = "s1"
            wb.Worksheets.Add(After:=wb.Worksheets("s1")).Name = "s2"
            wb.Worksheets("s1").Delete
            wb.Worksheets.Add(After:=wb.Worksheets("s2")).Name = "s3"
        Case 5
            wb.Worksheets(1).Name = "s1"
            wb.Worksheets.Add(After:=wb.Worksheets("s1")).Name = "s2"
            wb.Worksheets("s1").Copy Before:=wb.Worksheets("s1")
            wb.Worksheets.Add(After:=wb.Worksheets("s2")).Name = "s3"
        Case 6, 7
            wb.Worksheets(1).Name = "s1"
            wb.Worksheets.Add(After:=wb.Worksheets("s1")).Name = "s2"
            wb.Worksheets.Add(Before:=wb.Worksheets("s1")).Name = "s3"
            name = "numbers_first.xlsx"
            wb.SaveAs folder & "\\" & name, 51
            wb.Close False
            Set wb = Workbooks.Open(folder & "\\" & name)
            If how = 7 Then wb.Worksheets("s3").Delete
            wb.Worksheets.Add(After:=wb.Worksheets("s2")).Name = "s4"
            wb.Worksheets.Add(Before:=wb.Worksheets("s1")).Name = "s5"
    End Select
    name = "numbers.xlsx"
    wb.SaveAs folder & "\\" & name, 51
    wb.Close False
    Make = "ok"
End Function
'''

#: name -> (the case's number in Make, what it does).
NUMBERS: dict[str, tuple[int, str]] = {
    "added_before": (1, "Worksheets.Add three times, each new sheet before the active one"),
    "added_out_of_order": (2, "s1, s2 after s1, s3 before s1, s4 after s2, s5 before s3"),
    "newest_deleted": (3, "s1, s2, s3 one after another, s3 deleted, s4 after s2"),
    "oldest_deleted": (4, "s1, s2 one after another, s1 deleted, s3 after s2"),
    "copied": (5, "s1, s2 after s1, s1 copied before s1, s3 after s2"),
    "reopened_added": (6, "s1, s2 after s1, s3 before s1, saved and opened again, s4 after s2, s5 before s1"),
    "reopened_highest_deleted": (7, "as reopened_added, s3 deleted after the file is opened again"),
}


def numbers(path: Path) -> list[list[object]]:
    """Each sheet's name and sheetId, in the order the workbook part lists them."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
    return [[name, int(number)] for name, number in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\bsheetId="(\d+)"',
                                                                 workbook)]


def install(code: str, name: str) -> str:
    """A macro adding a class module of ``code`` to the workbook the harness runs in."""
    lines = ["Function Install() As String", "Dim text As String, part As Object",
             "Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)", f'part.Name = "{name}"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
    return "\n".join([*lines, "part.CodeModule.AddFromString text", 'Install = "ok"', "End Function"])


def main() -> None:
    logs: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.add_module("EventLog", "Public Log As String\n")
        installed = excel.run_vba(install(SINK, "CalculateSink"), "Install", timeout=60.0)
        assert installed.ok, f"{installed.outcome}: {installed.message}"
        for name, (how, _what) in CASES.items():
            result = excel.run_vba(WATCH, "Watch", (how, str(FILES)), timeout=300.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            logs[name] = str(result.value)
            print("==", name, logs[name])
        saved: dict[str, list[list[object]]] = {}
        for name, (how, _what) in NUMBERS.items():
            result = excel.run_vba(MAKE, "Make", (how, str(FILES)), timeout=300.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            saved[name] = numbers(FILES / "numbers.xlsx")
            print("==", name, saved[name])
    for leftover in ("calculate_order.xlsx", "numbers.xlsx", "numbers_first.xlsx"):
        (FILES / leftover).unlink(missing_ok=True)
    record = {"sink": SINK, "watch": WATCH, "make": MAKE,
              "cases": [{"name": name, "how": how, "what": what, "log": logs[name]}
                        for name, (how, what) in CASES.items()],
              "numbers": [{"name": name, "how": how, "what": what, "sheets": saved[name]}
                          for name, (how, what) in NUMBERS.items()]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
