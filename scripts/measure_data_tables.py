"""What-if data tables: Range.Table, the cells it makes, how they calculate and what a file keeps of them.

Each probe adds a sheet holding a small loan model, a rate in A1, years in
A2 and a principal in A3, its future value in A4 and interest in A5, and
makes data tables over it with Range.Table: rates down column C against
both formulas (C2:E6, column input A1), years along row 2 against the
future value (G2:K3, row input A2), and both at once (M2:P5). It records
what the table's cells hold and say, what an edit does to them, and what
Range.Table refuses; then deletes the sheet. A last part saves a workbook
holding the three tables, the same in semi-automatic calculation, and the
same again with row 1 deleted, taking two input cells with it, and keeps
the cells Excel wrote for each; then opens the same three as the model
saves them, and keeps what Excel finds in their cells, before and after
the model changes.

    python scripts/measure_data_tables.py

writes tests/fixtures/data_tables.json, which tests/test_excel_data_tables.py
replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/data_tables.json'

#: Builds the model and its three tables on ws.
BUILD = '''Sub Build(ws As Worksheet)
ws.Range("A1").Value = 0.05: ws.Range("A2").Value = 10: ws.Range("A3").Value = 1000
ws.Range("A4").Formula = "=A3*(1+A1)^A2": ws.Range("A5").Formula = "=A4-A3"
ws.Range("C3:C6").Value = Application.Transpose(Array(0.01, 0.02, 0.03, 0.04))
ws.Range("D2").Formula = "=A4": ws.Range("E2").Formula = "=A5"
ws.Range("C2:E6").Table ColumnInput:=ws.Range("A1")
ws.Range("H2:K2").Value = Array(1, 2, 3, 4)
ws.Range("G3").Formula = "=A4"
ws.Range("G2:K3").Table RowInput:=ws.Range("A2")
ws.Range("M2").Formula = "=A4"
ws.Range("N2:P2").Value = Array(5, 10, 15)
ws.Range("M3:M5").Value = Application.Transpose(Array(0.01, 0.02, 0.03))
ws.Range("M2:P5").Table RowInput:=ws.Range("A2"), ColumnInput:=ws.Range("A1")
End Sub

Function Cells2(ws As Worksheet, addresses As String) As String
Dim one As Variant, s As String
For Each one In Split(addresses, " ")
    s = s & "|" & CStr(ws.Range(one).Value)
Next
Cells2 = Mid$(s, 2)
End Function

Function Says(cell As Range) As String
Dim s As String
On Error Resume Next
s = cell.Formula & "~" & cell.FormulaR1C1 & "~" & cell.HasFormula & "~" & cell.HasArray
s = s & "~" & cell.FormulaArray & "~" & cell.CurrentArray.Address & "~" & cell.Formula2
Says = s
End Function

Function Tried(ws As Worksheet, target As String, value As Variant) As String
On Error Resume Next
Err.Clear
ws.Range(target).Value = value
Tried = Err.Number & ":" & CStr(ws.Range(target).Value)
End Function
'''
#: Each probe's body, ws being a sheet Build has filled and s what it returns.
BODIES = {
    'values': '''s = Cells2(ws, "D3 D4 D5 D6 E3 E4 E5 E6") & "#" & Cells2(ws, "H3 I3 J3 K3")
s = s & "#" & Cells2(ws, "N3 O3 P3 N4 O4 P4 N5 O5 P5") & "#" & Cells2(ws, "A1 A2 A4 C2 G2")''',
    'formulas': '''s = Says(ws.Range("D3")) & "#" & Says(ws.Range("E6")) & "#" & Says(ws.Range("I3")) & "#" & Says(ws.Range("O4"))
s = s & "#" & Says(ws.Range("D2")) & "#" & Says(ws.Range("C3")) & "#" & Says(ws.Range("C2")) & "#" & Says(ws.Range("M2"))''',
    'model_changes': '''ws.Range("A3").Value = 2000
s = Cells2(ws, "D3 E3 H3 N3 P5")''',
    'value_tried_changes': '''ws.Range("C3").Value = 0.1: ws.Range("H2").Value = 20: ws.Range("N2").Value = 1
s = Cells2(ws, "D3 E3 H3 N3 N4 N5")''',
    'formula_changes': '''ws.Range("D2").Formula = "=A4*2": ws.Range("M2").Formula = "=A5"
s = Cells2(ws, "D3 D6 E3 N3 P5")''',
    'blank_text_error_tried': '''ws.Range("C3").ClearContents: ws.Range("C4").Value = "x": ws.Range("C5").Formula = "=1/0"
ws.Range("C6").Formula = "=0.07"
s = Cells2(ws, "D3 D4 D5 D6")''',
    'write_into_table': '''Application.DisplayAlerts = False
s = Tried(ws, "D3", 5) & "#" & Tried(ws, "O4", 5)
On Error Resume Next
Err.Clear: ws.Range("D3:D4").ClearContents: s = s & "#" & Err.Number
Err.Clear: ws.Range("D3:E6").ClearContents: s = s & "#" & Err.Number & ":" & CStr(ws.Range("E6").Value) & ":" & ws.Range("D2").Formula
Err.Clear: ws.Range("N3:P5").ClearContents: s = s & "#" & Err.Number & ":" & CStr(ws.Range("O4").Value)
On Error GoTo 0
Application.DisplayAlerts = True''',
    'refusals': '''Dim other As Worksheet
On Error Resume Next
Set other = Worksheets.Add(After:=ws)
Err.Clear: ws.Range("R2:S5").Table: s = Err.Number
Err.Clear: ws.Range("R2:S5").Table ColumnInput:=other.Range("A1"): s = s & "|" & Err.Number
Err.Clear: ws.Range("R2:S5").Table ColumnInput:=ws.Range("A1:A2"): s = s & "|" & Err.Number
Err.Clear: ws.Range("R2").Table ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number
Err.Clear: ws.Range("R2:S5").Table ColumnInput:=ws.Range("S3"): s = s & "|" & Err.Number
ws.Range("R2:S5").Clear
Err.Clear: ws.Range("U2:V5").Table ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number & ":" & ws.Range("V3").Formula
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'semiautomatic': '''Application.Calculation = xlCalculationSemiautomatic
ws.Range("A3").Value = 2000
s = Cells2(ws, "A4 D3 H3 N3")
ws.Calculate: s = s & "#" & Cells2(ws, "D3 H3 N3")
Application.Calculate: s = s & "#" & Cells2(ws, "D3 H3 N3")
ws.Range("A3").Value = 3000
Application.Calculation = xlCalculationAutomatic: s = s & "#" & Cells2(ws, "D3 H3 N3")''',
    'semiautomatic_reads': '''Application.Calculation = xlCalculationSemiautomatic
s = CStr(Application.Calculation)
Application.Calculation = xlCalculationAutomatic''',
    'manual': '''Application.Calculation = xlCalculationManual
ws.Range("A3").Value = 2000
s = Cells2(ws, "A4 D3 H3")
ws.Calculate: s = s & "#" & Cells2(ws, "A4 D3 H3")
Application.Calculation = xlCalculationAutomatic''',
    'other_sheet_formula': '''Dim other As Worksheet
Set other = Worksheets.Add(After:=ws)
other.Range("A1").Formula = "='" & ws.Name & "'!A4/2"
ws.Range("R2").Formula = "='" & other.Name & "'!A1"
ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:R4").Table ColumnInput:=ws.Range("A1")
s = Cells2(ws, "R3 R4") & "#" & CStr(other.Range("A1").Value)
Application.DisplayAlerts = False: other.Delete: Application.DisplayAlerts = True''',
    'branch_on_input': '''ws.Range("R2").Formula = "=IF(A1>0.025,A4,A3)"
ws.Range("Q3:Q5").Value = Application.Transpose(Array(0.01, 0.03, 0.05))
ws.Range("Q2:R5").Table ColumnInput:=ws.Range("A1")
s = Cells2(ws, "R3 R4 R5")''',
    'input_is_formula': '''ws.Range("A6").Formula = "=0.05"
ws.Range("A7").Formula = "=A3*A6"
ws.Range("R2").Formula = "=A7"
ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:R4").Table ColumnInput:=ws.Range("A6")
s = Cells2(ws, "R3 R4 A6 A7") & "#" & ws.Range("A6").Formula''',
    'table_reads_table': '''ws.Range("R2").Formula = "=D3+1"
ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:R4").Table ColumnInput:=ws.Range("A1")
s = Cells2(ws, "R3 R4 D3")''',
    'sum_of_table': '''ws.Range("D8").Formula = "=SUM(D3:D6)"
s = CStr(ws.Range("D8").Value)
ws.Range("A3").Value = 2000: s = s & "|" & CStr(ws.Range("D8").Value)''',
    'table_returns': '''Dim back As Variant
ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("R2").Formula = "=A4"
back = ws.Range("Q2:R4").Table(ColumnInput:=ws.Range("A1"))
s = TypeName(back) & ":" & CStr(back)''',
    'insert_row_through': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Rows(4).Insert: s = Err.Number & ":" & ws.Range("D4").Formula
Err.Clear: ws.Columns("E").Insert: s = s & "|" & Err.Number & ":" & ws.Range("E3").Formula
On Error GoTo 0
Application.DisplayAlerts = True''',
    'insert_row_above': '''ws.Rows(1).Insert
s = ws.Range("D4").Formula & "|" & ws.Range("D4").FormulaR1C1 & "|" & CStr(ws.Range("D4").Value)
s = s & "|" & ws.Range("D4").CurrentArray.Address & "|" & ws.Range("O5").Formula
ws.Range("A4").Value = 2000: s = s & "|" & CStr(ws.Range("D4").Value)''',
    'insert_column_left': '''ws.Columns(1).Insert
s = ws.Range("E3").Formula & "|" & CStr(ws.Range("E3").Value) & "|" & ws.Range("I3").Formula & "|" & ws.Range("P4").Formula''',
    'delete_input_row': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Rows(1).Delete: s = Err.Number & ":" & ws.Range("D2").Formula & ":" & CStr(ws.Range("D2").Value)
s = s & ":" & ws.Range("H2").Formula & ":" & CStr(ws.Range("H2").Value)
On Error GoTo 0
Application.DisplayAlerts = True''',
    'whole_value_write': '''On Error Resume Next
Err.Clear: ws.Range("D3:E6").Value = 5
s = Err.Number & ":" & ws.Range("D3").Formula & ":" & CStr(ws.Range("E6").Value) & ":" & ws.Range("E6").HasArray
Err.Clear: ws.Range("N3:P5").Formula = "=1+1"
s = s & "|" & Err.Number & ":" & ws.Range("O4").Formula & ":" & CStr(ws.Range("O4").Value)''',
    'whole_table_clear': '''On Error Resume Next
Err.Clear: ws.Range("C2:E6").ClearContents
s = Err.Number & ":" & ws.Range("D2").Formula & ":" & ws.Range("D3").Formula & ":" & CStr(ws.Range("C3").Value)''',
    'refusal_messages': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Range("R2:S5").Table: s = Err.Number & ":" & Err.Description
Err.Clear: ws.Range("R2").Table ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Rows(4).Insert: s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("D3:D4").ClearContents: s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Rows(3).Delete: s = s & "|" & Err.Number & ":" & Err.Description & ":" & ws.Range("D3").Formula
On Error GoTo 0
Application.DisplayAlerts = True''',
    'more_messages': '''Dim other As Worksheet
Application.DisplayAlerts = False
On Error Resume Next
Set other = Worksheets.Add(After:=ws)
ws.Range("R2").Formula = "=A4": ws.Range("S2").Formula = "=A5"
Err.Clear: ws.Range("Q2:S2").Table ColumnInput:=ws.Range("A1"): s = Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:Q3").Table ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:S5").Table ColumnInput:=other.Range("A1"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:S5").Table ColumnInput:=ws.Range("A1:A2"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:S5").Table ColumnInput:=ws.Range("R3"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:S5").Table ColumnInput:=ws.Range("Q3"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:S5").Table RowInput:=ws.Range("A1"), ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("C2:D6").Table ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("D3").FormulaArray = "=1": s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("D3").Delete Shift:=xlUp: s = s & "|" & Err.Number & ":" & Err.Description
Err.Clear: ws.Range("Q2:S5").Table RowInput:=ws.Range("Q2"): s = s & "|" & Err.Number & ":" & Err.Description & ":" & ws.Range("R3").Formula
other.Delete
On Error GoTo 0
Application.DisplayAlerts = True''',
    'several_areas': '''ws.Range("R2").Formula = "=A4": ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("U2").Formula = "=A5": ws.Range("T3:T4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:R4,T2:U4").Table ColumnInput:=ws.Range("A1")
s = ws.Range("R3").Formula & "|" & ws.Range("U3").Formula & "|" & ws.Range("R3").CurrentArray.Address
s = s & "|" & Cells2(ws, "R3 U3")
ws.Range("D3:E6,H3:K3").ClearContents
s = s & "#" & ws.Range("D3").Formula & "|" & ws.Range("H3").Formula''',
    'thin_blocks': '''On Error Resume Next
ws.Range("R2").Formula = "=A4": ws.Range("S2").Formula = "=A5"
Err.Clear: ws.Range("Q2:S2").Table ColumnInput:=ws.Range("A1"): s = Err.Number & ":" & ws.Range("R2").Formula
ws.Range("Q3").Value = 0.01
Err.Clear: ws.Range("Q2:Q3").Table ColumnInput:=ws.Range("A1"): s = s & "|" & Err.Number & ":" & ws.Range("Q3").Formula''',
    'input_on_edge': '''On Error Resume Next
ws.Range("R2").Formula = "=A4": ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
Err.Clear: ws.Range("Q2:R4").Table ColumnInput:=ws.Range("Q2")
s = Err.Number & ":" & ws.Range("R3").Formula & ":" & CStr(ws.Range("R3").Value)
ws.Range("U2").Formula = "=A4": ws.Range("T3:T4").Value = Application.Transpose(Array(0.1, 0.2))
Err.Clear: ws.Range("T2:U4").Table ColumnInput:=ws.Range("T3")
s = s & "|" & Err.Number & ":" & ws.Range("U3").Formula & ":" & CStr(ws.Range("U3").Value)
ws.Range("X2").Formula = "=A4": ws.Range("W3:W4").Value = Application.Transpose(Array(0.1, 0.2))
Err.Clear: ws.Range("W2:X4").Table ColumnInput:=ws.Range("X2")
s = s & "|" & Err.Number & ":" & ws.Range("X3").Formula & ":" & CStr(ws.Range("X3").Value)''',
    'same_input_twice': '''On Error Resume Next
Err.Clear: ws.Range("M2:P5").Table RowInput:=ws.Range("A1"), ColumnInput:=ws.Range("A1")
s = Err.Number & ":" & ws.Range("N3").Formula & ":" & CStr(ws.Range("N3").Value)''',
    'semiautomatic_new': '''Application.Calculation = xlCalculationSemiautomatic
ws.Range("R2").Formula = "=A4": ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:R4").Table ColumnInput:=ws.Range("A1")
s = Cells2(ws, "R3 R4") & "#" & ws.Range("R3").Text
ws.Range("Q3").Value = 0.3: s = s & "#" & Cells2(ws, "R3 R4")
ws.Range("R2").Formula = "=A5": s = s & "#" & Cells2(ws, "R3 R4")''',
    'semiautomatic_reader': '''ws.Range("D8").Formula = "=SUM(D3:D6)"
Application.Calculation = xlCalculationSemiautomatic
ws.Range("A3").Value = 2000
s = Cells2(ws, "D8 D3")
ws.Range("D9").Formula = "=D3*1"
s = s & "#" & Cells2(ws, "D9")
ws.Calculate
s = s & "#" & Cells2(ws, "D8 D3 D9")''',
    'manual_new_table': '''Application.Calculation = xlCalculationManual
ws.Range("R2").Formula = "=A4": ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:R4").Table ColumnInput:=ws.Range("A1")
s = Cells2(ws, "R3 R4")''',
    'formula_row_blank_and_constant': '''ws.Range("R2").Value = 7
ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("Q2:S4").Table ColumnInput:=ws.Range("A1")
s = Cells2(ws, "R3 S3") & "#" & TypeName(ws.Range("S3").Value) & "#" & ws.Range("S3").Text & "#" & ws.Range("R4").Text
ws.Range("S2").Value = "abc": s = s & "#" & Cells2(ws, "S3 S4") & "#" & TypeName(ws.Range("S3").Value)''',
    'remake_and_overlap': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Range("C2:E6").Table ColumnInput:=ws.Range("A2")
s = Err.Number & ":" & ws.Range("D3").Formula & ":" & CStr(ws.Range("D3").Value)
Err.Clear: ws.Range("C2:D6").Table ColumnInput:=ws.Range("A1")
s = s & "|" & Err.Number & ":" & ws.Range("D3").Formula & ":" & ws.Range("E3").Formula
Err.Clear: ws.Range("C2:F7").Table ColumnInput:=ws.Range("A1")
s = s & "|" & Err.Number & ":" & ws.Range("D3").Formula & ":" & ws.Range("F7").Formula & ":" & ws.Range("D3").CurrentArray.Address
On Error GoTo 0
Application.DisplayAlerts = True''',
    'over_values_and_arrays': '''Application.DisplayAlerts = False
On Error Resume Next
ws.Range("R2").Formula = "=A4": ws.Range("Q3:Q4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("R3").Value = 99
Err.Clear: ws.Range("Q2:R4").Table ColumnInput:=ws.Range("A1")
s = Err.Number & ":" & ws.Range("R3").Formula & ":" & CStr(ws.Range("R3").Value)
ws.Range("U2").Formula = "=A4": ws.Range("T3:T4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("U3:U4").FormulaArray = "=A1*2"
Err.Clear: ws.Range("T2:U4").Table ColumnInput:=ws.Range("A1")
s = s & "|" & Err.Number & ":" & ws.Range("U3").Formula & ":" & ws.Range("U3").HasArray
ws.Range("X2").Formula = "=A4": ws.Range("W3:W4").Value = Application.Transpose(Array(0.1, 0.2))
ws.Range("X4:X5").FormulaArray = "=A1*2"
Err.Clear: ws.Range("W2:X4").Table ColumnInput:=ws.Range("A1")
s = s & "|" & Err.Number & ":" & ws.Range("X4").Formula & ":" & ws.Range("X5").Formula
On Error GoTo 0
Application.DisplayAlerts = True''',
    'copy_out': '''ws.Range("D3:D4").Copy ws.Range("R10")
s = ws.Range("R10").Formula & "|" & CStr(ws.Range("R10").Value) & "|" & ws.Range("R10").HasArray
ws.Range("C2:E6").Copy ws.Range("R20")
s = s & "#" & ws.Range("S21").Formula & "|" & CStr(ws.Range("S21").Value) & "|" & ws.Range("S21").HasArray
s = s & "|" & ws.Range("S20").Formula''',
    'formula_into_table': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Range("D3").Formula = "=1": s = Err.Number & ":" & ws.Range("D3").Formula
Err.Clear: ws.Range("D3").FormulaArray = "=1": s = s & "|" & Err.Number & ":" & ws.Range("D3").Formula
Err.Clear: ws.Range("D6:D7").Value = 1: s = s & "|" & Err.Number & ":" & ws.Range("D6").Formula & ":" & CStr(ws.Range("D7").Value)
Err.Clear: ws.Range("D2:D3").Value = 1: s = s & "|" & Err.Number & ":" & ws.Range("D2").Formula & ":" & ws.Range("D3").Formula
On Error GoTo 0
Application.DisplayAlerts = True''',
    'shift_inside': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Range("D3").Delete Shift:=xlUp: s = Err.Number & ":" & ws.Range("D3").Formula
Err.Clear: ws.Range("D2").Delete Shift:=xlUp: s = s & "|" & Err.Number & ":" & ws.Range("D2").Formula
Err.Clear: ws.Range("D1").Insert Shift:=xlDown: s = s & "|" & Err.Number & ":" & ws.Range("D2").Formula
On Error GoTo 0
Application.DisplayAlerts = True''',
    'shift_whole_block': '''Application.DisplayAlerts = False
On Error Resume Next
Err.Clear: ws.Range("C1:E1").Insert Shift:=xlDown
s = Err.Number & ":" & ws.Range("D4").Formula & ":" & ws.Range("D3").Formula & ":" & CStr(ws.Range("D4").Value)
s = s & ":" & ws.Range("D4").CurrentArray.Address
Err.Clear: ws.Range("C9:E9").Delete Shift:=xlUp: s = s & "|" & Err.Number & ":" & ws.Range("D4").Formula
On Error GoTo 0
Application.DisplayAlerts = True''',
    'sheet_copy': '''Dim copied As Worksheet
ws.Copy After:=ws
Set copied = Worksheets(ws.Index + 1)
copied.Range("A3").Value = 2000
s = copied.Range("D3").Formula & "|" & CStr(copied.Range("D3").Value) & "|" & CStr(ws.Range("D3").Value)
s = s & "|" & copied.Range("N3").CurrentArray.Address
Application.DisplayAlerts = False: copied.Delete: Application.DisplayAlerts = True''',
    'semiautomatic_range_calculate': '''Application.Calculation = xlCalculationSemiautomatic
ws.Range("A3").Value = 2000
ws.Range("D3:D4").Calculate
s = Cells2(ws, "D3 D4 D5 E3 H3")
ws.Range("A4").Calculate
s = s & "#" & Cells2(ws, "D5 H3")
Application.Calculate
s = s & "#" & Cells2(ws, "D5 H3")''',
    'paste_special': '''ws.Range("D3:D4").Copy
ws.Range("R10").PasteSpecial xlPasteFormulas
s = ws.Range("R10").Formula & "|" & ws.Range("R10").HasArray
ws.Range("S10").PasteSpecial xlPasteValues
s = s & "|" & ws.Range("S10").Formula
ws.Range("T10").PasteSpecial xlPasteAll
s = s & "|" & ws.Range("T10").Formula
Application.CutCopyMode = False''',
    'insert_row_at_top': '''ws.Rows(2).Insert
s = ws.Range("D4").Formula & "|" & ws.Range("D4").CurrentArray.Address & "|" & CStr(ws.Range("D4").Value)''',
    'delete_whole_table_columns': '''ws.Columns("G:K").Delete
s = ws.Range("I3").Formula & "|" & ws.Range("I3").CurrentArray.Address & "|" & CStr(ws.Range("I3").Value)
s = s & "|" & CStr(ws.Range("H3").Value)''',
}


def macro(body: str) -> str:
    return BUILD + f'''
Function Probe() As String
Dim ws As Worksheet, s As String
Set ws = Worksheets.Add
Build ws
{body}
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
Probe = s
End Function
'''


#: The files the last part saves, each after this is done to the workbook Build filled, calculation automatic.
STEPS = {
    'tables': '',
    'semiautomatic': 'Application.Calculation = xlCalculationSemiautomatic',
    'deleted_input': 'wb.Worksheets(1).Rows(1).Delete',
}


def file_probe(folder: Path) -> str:
    saves = '\n'.join(f'''Application.Calculation = xlCalculationAutomatic
Set wb = Workbooks.Add(xlWBATWorksheet)
Build wb.Worksheets(1)
{step}
wb.SaveAs Filename:="{str(folder / f'{name}.xlsx').replace('"', '""')}", FileFormat:=xlOpenXMLWorkbook
wb.Close False''' for name, step in STEPS.items())
    return BUILD + f'''
Function Probe() As String
Dim wb As Workbook
Application.DisplayAlerts = False
{saves}
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Probe = "saved"
End Function
'''


#: Opens each of the model's files in {folder}, beside the workbook already open, and gives every cell's formula and
#: value, then the same after A3 changes. Alerts stay on, so a file Excel has to repair stops the probe.
OPENED = '''Function Dump(ws As Worksheet) As String
Dim one As Range, s As String
For Each one In ws.UsedRange.Cells
    If one.HasFormula Or Not IsEmpty(one.Value) Then
        s = s & "|" & one.Address(False, False) & "=" & one.Formula & "~" & CStr(one.Value)
    End If
Next
Dump = Mid$(s, 2)
End Function

Function Probe() As String
Dim wb As Workbook, ws As Worksheet, s As String, name As Variant
For Each name In Array(NAMES)
    Set wb = Workbooks.Open(Filename:="{folder}" & name & ".xlsx")
    Set ws = wb.Worksheets(1)
    s = s & "#" & Dump(ws)
    ws.Range("A3").Value = 2000
    s = s & "@" & Dump(ws)
    wb.Close False
Next
Probe = Mid$(s, 2)
End Function
'''.replace('NAMES', ', '.join(f'"{name}"' for name in STEPS))


def model_saves(folder: Path) -> None:
    """The files as the model saves them, built as Excel's were."""
    from pyopenvba.apps.excel import ExcelApplication

    for name, step in STEPS.items():
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(BUILD + f'\nSub Save()\nDim wb As Workbook\nSet wb = ActiveWorkbook\nBuild wb.Worksheets(1)\n'
                       f'{step}\nEnd Sub\n', name='Harness')
        app.run('Harness.Save')
        app.save(folder / f'{name}.xlsx')


_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)


def main() -> None:
    record({name: macro(body) for name, body in BODIES.items()}, OUT)
    with tempfile.TemporaryDirectory() as excels, tempfile.TemporaryDirectory() as models, \
            ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        saved = excel.run_vba(file_probe(Path(excels)), 'Probe', timeout=120.0, line_numbers=False)
        assert saved.ok and not saved.dialogs, saved
        files: dict[str, dict[str, object]] = {}
        for name, step in STEPS.items():
            with zipfile.ZipFile(Path(excels) / f'{name}.xlsx') as package:
                sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                workbook = package.read('xl/workbook.xml').decode('utf-8')
            found = re.search(r'<calcPr\b[^>]*/?>', workbook)
            files[name] = {'step': step, 'cells': {match.group(1): match.group(0) for match in _CELL.finditer(sheet)},
                           'calcPr': found.group(0) if found else ''}
        model_saves(Path(models))
        folder = str(Path(models)) + '\\'
        opened = excel.run_vba(OPENED.replace('{folder}', folder.replace('"', '""')), 'Probe', timeout=120.0,
                               line_numbers=False)
        assert opened.ok and not opened.dialogs, opened
    fixture = json.loads(OUT.read_text(encoding='utf-8'))
    fixture['files'] = files
    fixture['opened'] = {'macro': OPENED, 'result': opened.value}
    OUT.write_text(json.dumps(fixture, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({name: {key: value for key, value in one['cells'].items() if key[0] in 'DEHIJKNOP'}  # type: ignore[union-attr]
                      for name, one in files.items()}, indent=1)[:2000])
    print(str(opened.value)[:2000])


if __name__ == '__main__':
    main()
