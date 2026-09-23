"""What a cell shift -- Range.Delete or Range.Insert with a Shift -- does to the references around it.

Every layout gets a new workbook: a grid of numbers in A1:E10 of a sheet
named Grid, formulas in column H reading single cells, blocks, whole
columns and rows of it, formulas in B11:C12 that move with the cells
under them, formulas on a second sheet, and two names; then one delete or
insert. The dump lists every formula in Grid!A1:J17 by address, Grid's
cells A1:E12 as they stand, the other sheet's formulas, the names, the
AutoFilter's range where there is one, and the error the edit raised.

    python scripts/measure_cell_shifts.py

writes tests/fixtures/cell_shifts.json, which tests/test_excel_cell_shifts.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "cell_shifts.json"

HELPER = '''
Private Function Dump(ws As Object) As String
    Dim out As String, c As Object, other As Object
    For Each c In ws.Range("A1:J17").Cells
        If c.HasFormula Then out = out & c.Address(False, False) & "=" & c.Formula & ";"
    Next
    out = out & " cells="
    For Each c In ws.Range("A1:E12").Cells
        out = out & c.Formula & ","
    Next
    Set other = ws.Parent.Worksheets("Other")
    out = out & " other=" & other.Range("A1").Formula & ";" & other.Range("A2").Formula
    out = out & " names=" & ws.Parent.Names("Tracked").RefersTo & ";" & ws.Parent.Names("Cell").RefersTo
    On Error Resume Next
    Dim f As String, db As String
    f = ""
    f = ws.AutoFilter.Range.Address(False, False)
    db = ""
    db = ws.Names("_FilterDatabase").RefersTo
    On Error GoTo 0
    If f <> "" Then out = out & " filter=" & f
    If db <> "" Then out = out & " db=" & db
    out = out & " fmt=" & ws.Range("B3").NumberFormat & "," & ws.Range("B4").NumberFormat & "," & _
        ws.Range("C3").NumberFormat & "," & ws.Range("B3").Font.Bold & "," & ws.Range("B5").NumberFormat
    Dump = out
End Function
'''

#: The grid: numbers in A1:E10, formulas reading it, formulas that move with it, another sheet, names.
GRID = '''For r = 1 To 10
    For c = 1 To 5
        ws.Cells(r, c).Value = r * 10 + c
    Next
Next
ws.Range("H1").Formula = "=B3"
ws.Range("H2").Formula = "=B5"
ws.Range("H3").Formula = "=$B$3"
ws.Range("H4").Formula = "=B$5"
ws.Range("H5").Formula = "=$B5"
ws.Range("H6").Formula = "=C5"
ws.Range("H7").Formula = "=SUM(B2:B6)"
ws.Range("H8").Formula = "=SUM(B4:B8)"
ws.Range("H9").Formula = "=SUM(A2:C6)"
ws.Range("H10").Formula = "=SUM(B2:C6)"
ws.Range("H11").Formula = "=SUM(B3:B4)"
ws.Range("H12").Formula = "=SUM(B:B)"
ws.Range("H13").Formula = "=SUM(3:3)"
ws.Range("H14").Formula = "=B10"
ws.Range("H15").Formula = "=SUM(A5:E5)"
ws.Range("H16").Formula = "=D3"
ws.Range("H17").Formula = "=SUM(B1:B10)"
ws.Range("B11").Formula = "=B2+1"
ws.Range("B12").Formula = "=B11*2"
ws.Range("C11").Formula = "=C5"
other.Range("A1").Formula = "=Grid!B5"
other.Range("A2").Formula = "=SUM(Grid!B2:B6)"
ws.Parent.Names.Add Name:="Tracked", RefersTo:="=Grid!$B$2:$B$6"
ws.Parent.Names.Add Name:="Cell", RefersTo:="=Grid!$B$5"
'''

#: name -> the edit, run on Grid (ws) after the grid.
LAYOUTS: dict[str, str] = {
    "del_up_b3b4": 'ws.Range("B3:B4").Delete Shift:=xlUp',
    "del_up_b3c4": 'ws.Range("B3:C4").Delete Shift:=xlUp',
    "del_up_b1b2": 'ws.Range("B1:B2").Delete Shift:=xlUp',
    "del_up_b2b6": 'ws.Range("B2:B6").Delete Shift:=xlUp',
    "del_up_b2b3": 'ws.Range("B2:B3").Delete Shift:=xlUp',
    "del_up_b6b7": 'ws.Range("B6:B7").Delete Shift:=xlUp',
    "del_up_b5": 'ws.Range("B5").Delete Shift:=xlUp',
    "del_up_a3e4": 'ws.Range("A3:E4").Delete Shift:=xlUp',
    "del_up_a3c4": 'ws.Range("A3:C4").Delete Shift:=xlUp',
    "del_up_b10": 'ws.Range("B10").Delete Shift:=xlUp',
    "ins_down_b3b4": 'ws.Range("B3:B4").Insert Shift:=xlDown',
    "ins_down_b3c4": 'ws.Range("B3:C4").Insert Shift:=xlDown',
    "ins_down_b2": 'ws.Range("B2").Insert Shift:=xlDown',
    "ins_down_b6": 'ws.Range("B6").Insert Shift:=xlDown',
    "ins_down_b7": 'ws.Range("B7").Insert Shift:=xlDown',
    "ins_down_a2c2": 'ws.Range("A2:C2").Insert Shift:=xlDown',
    "del_left_b3c3": 'ws.Range("B3:C3").Delete Shift:=xlToLeft',
    "del_left_b3c5": 'ws.Range("B3:C5").Delete Shift:=xlToLeft',
    "del_left_a5": 'ws.Range("A5").Delete Shift:=xlToLeft',
    "del_left_a5e5": 'ws.Range("A5:E5").Delete Shift:=xlToLeft',
    "ins_right_b3c3": 'ws.Range("B3:C3").Insert Shift:=xlToRight',
    "ins_right_a5": 'ws.Range("A5").Insert Shift:=xlToRight',
    "del_default_b3": 'ws.Range("B3").Delete',
    "del_default_b3b4": 'ws.Range("B3:B4").Delete',
    "del_default_b3c3": 'ws.Range("B3:C3").Delete',
    "del_default_b3d4": 'ws.Range("B3:D4").Delete',
    "del_default_b3c5": 'ws.Range("B3:C5").Delete',
    "ins_default_b3b4": 'ws.Range("B3:B4").Insert',
    "ins_default_b3c3": 'ws.Range("B3:C3").Insert',
    "ins_default_b3c5": 'ws.Range("B3:C5").Insert',
    "filter_del_b3": 'ws.Range("A1:C10").AutoFilter\nws.Range("B3").Delete Shift:=xlUp',
    "filter_del_a3c3": 'ws.Range("A1:C10").AutoFilter\nws.Range("A3:C3").Delete Shift:=xlUp',
    "filter_ins_a3c3": 'ws.Range("A1:C10").AutoFilter\nws.Range("A3:C3").Insert Shift:=xlDown',
    "filter_del_left_a3": 'ws.Range("A1:C10").AutoFilter\nws.Range("A3").Delete Shift:=xlToLeft',
    "filter_ins_right_a1": 'ws.Range("A1:C10").AutoFilter\nws.Range("A1:A10").Insert Shift:=xlToRight',
    "filter_del_below": 'ws.Range("A1:C10").AutoFilter\nws.Range("A11:C11").Delete Shift:=xlUp',
    "del_default_b3c4": 'ws.Range("B3:C4").Delete',
    "del_default_b3d5": 'ws.Range("B3:D5").Delete',
    "ins_default_b3c4": 'ws.Range("B3:C4").Insert',
    "filter_del_header": 'ws.Range("A1:C10").AutoFilter\nws.Range("A1:C1").Delete Shift:=xlUp',
    "filter_ins_header": 'ws.Range("A1:C10").AutoFilter\nws.Range("A1:C1").Insert Shift:=xlDown',
    "filter_del_whole": 'ws.Range("A1:C10").AutoFilter\nws.Range("A1:C10").Delete Shift:=xlUp',
    "filter_del_left_header": 'ws.Range("A1:C10").AutoFilter\nws.Range("A1").Delete Shift:=xlToLeft',
    "ins_down_formats": 'ws.Range("B2").NumberFormat = "0.00"\nws.Range("B2").Font.Bold = True\n'
                        'ws.Range("B5").NumberFormat = "0%"\nws.Range("B3:B4").Insert Shift:=xlDown',
    "ins_right_formats": 'ws.Range("A3").NumberFormat = "0.00"\nws.Range("A3").Font.Bold = True\n'
                         'ws.Range("B3:C3").Insert Shift:=xlToRight',
    "del_up_formats": 'ws.Range("B5").NumberFormat = "0%"\nws.Range("B3:B4").Delete Shift:=xlUp',
    "ins_down_formats_top": 'ws.Range("B1").NumberFormat = "0.00"\nws.Range("B1").Insert Shift:=xlDown',
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object, other As Object) As String",
             "Dim failed As String, r As Long, c As Long", *GRID.splitlines(), "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f'Case{index} = failed & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, other As Object",
             "Dim out As String", "Application.DisplayAlerts = False"]
    bodies: list[str] = []
    for index, setup in enumerate(LAYOUTS.values()):
        build += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Name = "Grid"',
                  "Set other = wb.Worksheets.Add(After:=ws)", 'other.Name = "Other"', "ws.Activate",
                  f'out = out & Case{index}(ws, other) & "|"', "wb.Close False"]
        bodies.append(case_code(index, setup))
    build += ["Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "grid": GRID,
              "layouts": [{"name": name, "setup": setup, "answers": answer}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        print(f"{name}: {answer}")


if __name__ == "__main__":
    main()
