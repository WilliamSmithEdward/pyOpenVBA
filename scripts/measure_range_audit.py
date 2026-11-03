"""What keeps a Range across cells moving: the clipboard, formatting objects taken from a range, sort settings, remembered filter
criteria, links and notes; and the edits that move cells other than by inserting and deleting: table rows and columns, subtotals.

    python scripts/measure_range_audit.py

writes tests/fixtures/range_audit.json, which tests/test_excel_range_tracking.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_audit.json'
PRE = 'Dim r As Range, q As Range, c As Range, s As String, i As Long, v As Variant, o As Object, lo As Object'
TABLE = 'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E9"), , xlYes)\nSet r = ws.Range("B3")\nSet q = lo.DataBodyRange\n'
FORMULAS = 'ws.Range("F3").Formula = "=D3*2"\nws.Range("G3").Formula = "=SUM($D$2:D3)"\nws.Range("H1").Formula = "=SUM(D2:D9)"\nws.Range("I1").Formula = "=D3+D5"'
CASES = [
    # The clipboard.
    'ws.Range("B3").Copy\nws.Rows(1).Insert\ns = Application.CutCopyMode\nws.Paste ws.Range("H1")\ns = s & ws.Range("H1").Value',
    'ws.Range("B3").Cut\nws.Rows(1).Insert\ns = Application.CutCopyMode\nws.Paste ws.Range("H1")\ns = s & ws.Range("H1").Value',
    'ws.Range("B3").Copy\nws.Rows(20).Insert\ns = Application.CutCopyMode\nws.Paste ws.Range("H1")\ns = s & ws.Range("H1").Value',
    'ws.Range("B3").Copy\nws.Range("B3").Delete xlShiftUp\ns = Application.CutCopyMode\nws.Paste ws.Range("H1")\ns = s & ws.Range("H1").Value',
    'ws.Range("B3").Copy\nws.Range("Z1").Value = 1\ns = Application.CutCopyMode\nws.Paste ws.Range("H1")\ns = s & ws.Range("H1").Value',
    'ws.Range("B3").Cut\nws.Range("Z1").Value = 1\ns = Application.CutCopyMode\nws.Paste ws.Range("H1")\ns = s & ws.Range("H1").Value',
    # Objects taken from a range.
    'Set o = ws.Range("B3").Font\nws.Rows(1).Insert\no.Bold = True\ns = ws.Range("B3").Font.Bold & ws.Range("B4").Font.Bold',
    'Set o = ws.Range("B3").Interior\nws.Rows(1).Insert\no.Color = 255\ns = ws.Range("B3").Interior.Color & "|" & ws.Range("B4").Interior.Color',
    'Set o = ws.Range("B3").Borders\nws.Rows(1).Insert\no.LineStyle = xlContinuous\ns = ws.Range("B3").Borders(xlEdgeTop).LineStyle & "|" & ws.Range("B4").Borders(xlEdgeTop).LineStyle',
    'Set o = ws.Range("B3").Font\nws.Rows(3).Delete\ns = o.Bold',
    'Set o = ws.Range("B3").Validation\nws.Rows(1).Insert\no.Add xlValidateWholeNumber, , xlBetween, "1", "5"\nOn Error Resume Next\ns = ws.Range("B3").Validation.Type & "|" & Err.Number\nErr.Clear\ns = s & ws.Range("B4").Validation.Type & "|" & Err.Number',
    'Set o = ws.Range("B3").FormatConditions\nws.Rows(1).Insert\no.Add xlCellValue, xlGreater, "=1"\ns = ws.Range("B3").FormatConditions.Count & ws.Range("B4").FormatConditions.Count',
    'Set o = ws.Range("B3").FormatConditions.Add(xlCellValue, xlGreater, "=1")\nws.Rows(1).Insert\ns = o.AppliesTo.Address',
    'Set o = ws.Range("B3").Characters(1, 1)\nws.Rows(1).Insert\no.Font.Bold = True\ns = ws.Range("B3").Font.Bold & ws.Range("B4").Font.Bold & TypeName(o)',
    # Sort settings and filter criteria.
    'With ws.Sort\n.SetRange ws.Range("A1:E9")\n.SortFields.Add Key:=ws.Range("D2:D9"), Order:=xlDescending\n.Header = xlYes\nEnd With\nws.Rows(1).Insert\nws.Sort.Apply\ns = ws.Sort.Rng.Address & ws.Sort.SortFields(1).Key.Address & ws.Range("D3").Value & ws.Range("D2").Value',
    'ws.Sort.SetRange ws.Range("A1:E9")\nws.Rows(20).Insert\ns = ws.Sort.Rng.Address',
    'ws.Range("H1").Value = "Team"\nws.Range("H2").Value = "t3"\nws.Range("A1:E9").AdvancedFilter xlFilterInPlace, ws.Range("H1:H2")\nws.ShowAllData\nws.Rows(1).Insert\nws.Range("A2:E10").AdvancedFilter xlFilterCopy, , ws.Range("L1")\ns = ws.Range("L2").Value & ws.Cells(ws.Rows.Count, 12).End(xlUp).Row',
    'ws.Range("H1").Value = "Team"\nws.Range("H2").Value = "t3"\nws.Range("A1:E9").AdvancedFilter xlFilterInPlace, ws.Range("H1:H2")\nws.ShowAllData\ns = ThisWorkbook.Names.Count & ws.Names.Count\nOn Error Resume Next\ns = s & ws.Names("Criteria").RefersTo & ws.Names("_FilterDatabase").RefersTo\ns = s & Err.Number',
    # Links and notes.
    'Set o = ws.Hyperlinks.Add(ws.Range("B3"), "https://example.com")\nws.Rows(1).Insert\ns = o.Range.Address & ws.Hyperlinks(1).Range.Address',
    'Set o = ws.Range("B3").AddComment("x")\nws.Rows(1).Insert\ns = o.Parent.Address & ws.Range("B4").Comment.Text',
    'Set o = ws.Range("B3").AddComment("x")\nws.Rows(3).Delete\nOn Error Resume Next\ns = ws.Comments.Count & o.Parent.Address & Err.Number',
    'On Error Resume Next\ns = ws.Columns("8:9").Address\ns = s & "|" & Err.Number',
    # Other edits that move cells.
    TABLE + 'lo.ListRows.Add 1\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ListRows.Add\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ListRows(1).Delete\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ListRows(2).Delete\nOn Error Resume Next\ns = r.Address & Err.Number & q.Address & lo.Range.Address',
    TABLE + 'lo.ListColumns.Add 1\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ListColumns.Add\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ListColumns(1).Delete\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ListColumns(2).Delete\nOn Error Resume Next\ns = r.Address & Err.Number & q.Address & lo.Range.Address',
    TABLE + 'ws.Rows(1).Insert\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'ws.Rows(4).Insert\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.Resize ws.Range("A1:E12")\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.ShowTotals = True\ns = r.Address & q.Address & lo.Range.Address',
    TABLE + 'lo.Delete\nOn Error Resume Next\ns = r.Address & Err.Number',
    TABLE + 'lo.Unlist\ns = r.Address & q.Address',
    'Set r = ws.Range("B5")\nws.Range("A1:E9").Sort ws.Range("B1"), Header:=xlYes\nws.Range("A1:E9").Subtotal 2, xlSum, Array(4)\ns = r.Address & ws.UsedRange.Address',
    'ws.Range("A1:E9").Subtotal 2, xlSum, Array(4)\nSet r = ws.Range("B5")\nws.Range("A1").CurrentRegion.RemoveSubtotal\nOn Error Resume Next\ns = r.Address & Err.Number & ws.UsedRange.Address',
    'Set r = ws.Range("B5")\nws.Range("A1:E9").RemoveDuplicates 2, xlYes\ns = r.Address',
    'ws.Range("B4:B5").Value = "t3"\nSet r = ws.Range("B5")\nws.Range("A1:E9").RemoveDuplicates 2, xlYes\ns = r.Address & r.Value',
    'Set r = ws.Range("B5")\nws.Range("A1:E9").Sort ws.Range("D1"), xlDescending, Header:=xlYes\ns = r.Address & r.Value',
    'Set r = ws.Range("A1:B1")\nr.Merge\ns = r.Address\nr.UnMerge\ns = s & r.Address',
    'Set r = ws.Range("D3")\nws.Rows(3).Group\ns = r.Address\nws.Rows(3).Hidden = True\ns = s & r.Address & r.Height',
    'Set r = ws.Range("H1")\nws.Range("B3").Cut r\ns = r.Address & r.Value',
    'Set r = ws.Range("B3")\nr.Cut ws.Range("H1")\ns = r.Address & r.Value\nr.EntireRow.Delete\ns = s & ws.Range("B3").Value',
    'Set r = ws.Range("B3")\nr.Delete xlShiftUp\nOn Error Resume Next\ns = r.Address & Err.Number',
    'Set r = ws.Range("B3")\nr.Insert xlShiftDown\ns = r.Address',
    'Set r = ws.Rows(3)\nr.Insert\ns = r.Address',
    'Set r = ws.Rows(3)\nr.Delete\nOn Error Resume Next\ns = r.Address & Err.Number',
    'Set r = ws.Range("A1")\nws.Parent.Close False\nOn Error Resume Next\ns = r.Address & Err.Number',
] + [
    # An insert or delete with something on the clipboard: what it answers, and the sheet afterwards.
    f'{put}\nOn Error Resume Next\n{edit}\ns = Err.Number & "|" & Look(ws)'
    for put, edits in (
        ('ws.Range("B3").Copy', (
            'ws.Rows(1).Insert', 'ws.Rows(2).Insert', 'ws.Rows(3).Insert', 'ws.Rows(4).Insert', 'ws.Rows(5).Insert',
            'ws.Rows(9).Insert', 'ws.Rows(10).Insert', 'ws.Rows(20).Insert', 'ws.Rows("30:31").Insert',
            'ws.Columns(1).Insert', 'ws.Columns(2).Insert', 'ws.Columns(3).Insert', 'ws.Columns(10).Insert',
            'ws.Range("D1").Insert xlShiftDown', 'ws.Range("D1").Insert xlShiftToRight', 'ws.Range("A1").Insert xlShiftDown',
            'ws.Range("B1").Insert xlShiftDown', 'ws.Range("B5").Insert xlShiftDown', 'ws.Range("A3").Insert xlShiftToRight',
            'ws.Range("D1").Insert', 'ws.Range("D1:D2").Insert xlShiftDown', 'ws.Range("D20").Insert xlShiftDown',
            'ws.Rows(1).Delete', 'ws.Rows(3).Delete', 'ws.Rows(5).Delete', 'ws.Rows(20).Delete', 'ws.Columns(1).Delete',
            'ws.Columns(2).Delete', 'ws.Columns(5).Delete', 'ws.Range("D1").Delete xlShiftUp', 'ws.Range("B1").Delete xlShiftUp',
            'ws.Range("B5").Delete xlShiftUp', 'ws.Range("A3").Delete xlShiftToLeft',
        )),
        ('ws.Range("B3").Cut', (
            'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Rows(5).Insert', 'ws.Rows(20).Insert', 'ws.Columns(1).Insert',
            'ws.Columns(10).Insert', 'ws.Range("D1").Insert xlShiftDown', 'ws.Range("B1").Insert xlShiftDown',
            'ws.Range("D1").Insert', 'ws.Rows(1).Delete', 'ws.Rows(3).Delete', 'ws.Rows(5).Delete', 'ws.Rows(20).Delete',
            'ws.Range("D1").Delete xlShiftUp', 'ws.Range("B5").Delete xlShiftUp',
        )),
        ('ws.Rows(3).Copy', (
            'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Rows(6).Insert', 'ws.Rows(20).Insert', 'ws.Columns(1).Insert',
            'ws.Range("D1").Insert xlShiftDown', 'ws.Rows(1).Delete', 'ws.Rows(6).Delete',
        )),
        ('ws.Rows(3).Cut', (
            'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Rows(6).Insert', 'ws.Rows(20).Insert', 'ws.Rows(1).Delete', 'ws.Rows(6).Delete',
        )),
        ('ws.Columns(2).Copy', ('ws.Columns(1).Insert', 'ws.Columns(4).Insert', 'ws.Rows(1).Insert', 'ws.Rows(20).Insert')),
        ('ws.Range("B3:C4").Copy', (
            'ws.Rows(1).Insert', 'ws.Rows(20).Insert', 'ws.Range("D1").Insert xlShiftDown', 'ws.Range("D1").Insert xlShiftToRight',
            'ws.Range("G1").Insert', 'ws.Range("G1:H2").Insert',
        )),
        ('ws.Parent.Worksheets.Add(After:=ws).Range("A1:A2").Value = 7\nws.Parent.Worksheets(2).Range("A1").Copy\nws.Activate', (
            'ws.Rows(1).Insert', 'ws.Rows(20).Insert', 'ws.Range("D1").Insert xlShiftDown', 'ws.Rows(1).Delete',
        )),
    )
    for edit in edits
] + [
    # What the first round left open.
    f'{put}\nOn Error Resume Next\n{edit}\ns = s & "|" & Err.Number & "|" & Look(ws)'
    for put, edits in (
        ('ws.Range("B3").Copy', (
            'ws.Rows("1:2").Insert', 'ws.Rows("2:3").Insert', 'ws.Rows("5:6").Insert', 'ws.Columns("1:2").Insert',
            'ws.Columns("2:3").Insert', 'ws.Columns("8:9").Insert', 'ws.Rows(3).Insert\nws.Paste ws.Range("H1")',
            'ws.Rows(1).Insert\nws.Range("D1").Insert\nws.Paste ws.Range("H1")', 'ws.Range("B1").Insert xlShiftDown\nws.Paste ws.Range("H1")',
        )),
        ('ws.Range("B3:B4").Copy', (
            'ws.Range("D1").Insert', 'ws.Rows(20).Insert', 'ws.Columns(8).Insert', 'ws.Columns(1).Insert', 'ws.Columns(2).Insert',
        )),
        ('ws.Range("B3:C3").Copy', (
            'ws.Range("D1").Insert', 'ws.Rows(20).Insert', 'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Columns(8).Insert',
        )),
        ('ws.Range("A3:E3").Copy', ('ws.Rows(20).Insert', 'ws.Rows(1).Insert')),
        ('ws.Range("B3:C4").Copy', (
            'ws.Range("D1:D4").Insert xlShiftDown', 'ws.Range("D1:G4").Insert xlShiftDown', 'ws.Range("D1:E3").Insert xlShiftDown',
            'ws.Rows("20:21").Insert', 'ws.Columns("8:9").Insert',
        )),
        ('ws.Rows("3:4").Copy', ('ws.Rows(1).Insert', 'ws.Rows("1:4").Insert', 'ws.Rows("1:3").Insert', 'ws.Rows(4).Insert')),
        ('ws.Rows(3).Copy', ('ws.Rows("1:2").Insert',)),
        ('ws.Columns(2).Copy', ('ws.Columns("1:2").Insert',)),
        ('ws.Range("A1,C1").Copy', ('ws.Range("D1").Insert', 'ws.Rows(20).Insert')),
        ('ws.Range("B3").Cut', (
            'ws.Columns(2).Insert', 'ws.Columns(3).Insert', 'ws.Rows(2).Insert', 'ws.Rows(4).Insert', 'ws.Range("B5").Insert xlShiftDown',
            'ws.Range("B4").Insert xlShiftDown', 'ws.Range("B3").Insert xlShiftDown', 'ws.Range("B3").Insert', 'ws.Range("A3").Insert',
            'ws.Range("A3").Insert xlShiftDown', 'ws.Range("A3").Insert xlShiftToRight', 'ws.Range("C3").Insert',
            'ws.Range("C3").Insert xlShiftToRight', 'ws.Range("D3").Insert xlShiftToRight', 'ws.Range("B1").Insert',
            'ws.Range("B1").Insert xlShiftToRight', 'ws.Range("D5").Insert', 'ws.Range("D5").Insert xlShiftToRight',
            'ws.Range("D1:D2").Insert xlShiftDown', 'ws.Range("D1:E2").Insert',
        )),
        ('ws.Range("B3:C4").Cut', (
            'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Rows(20).Insert', 'ws.Range("G1").Insert', 'ws.Range("B1").Insert xlShiftDown',
            'ws.Range("C1").Insert xlShiftDown', 'ws.Range("B2").Insert', 'ws.Range("C3").Insert', 'ws.Range("A3").Insert xlShiftToRight',
            'ws.Range("D3").Insert xlShiftToRight', 'ws.Range("A4").Insert xlShiftToRight', 'ws.Columns(1).Insert', 'ws.Columns(8).Insert',
        )),
        ('ws.Rows(3).Cut', (
            'ws.Rows(2).Insert', 'ws.Rows(4).Insert', 'ws.Rows(5).Insert', 'ws.Rows("1:2").Insert', 'ws.Columns(1).Insert',
            'ws.Range("D1").Insert', 'ws.Range("A1").Insert',
        )),
        ('ws.Rows("3:4").Cut', (
            'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Rows(4).Insert', 'ws.Rows(5).Insert', 'ws.Rows(6).Insert', 'ws.Rows(20).Insert',
            'ws.Rows("6:9").Insert',
        )),
        ('ws.Columns(2).Cut', (
            'ws.Columns(1).Insert', 'ws.Columns(2).Insert', 'ws.Columns(3).Insert', 'ws.Columns(4).Insert', 'ws.Columns(8).Insert',
            'ws.Rows(1).Insert', 'ws.Range("D1").Insert',
        )),
        ('ws.Parent.Worksheets.Add(After:=ws).Range("A1:B2").Value = 7\nws.Parent.Worksheets(2).Range("A1").Cut\nws.Activate', (
            'ws.Rows(1).Insert', 'ws.Rows(20).Insert', 'ws.Range("D1").Insert xlShiftDown',
        )),
        ('ws.Parent.Worksheets.Add(After:=ws).Range("A1:B2").Value = 7\nws.Parent.Worksheets(2).Rows(1).Cut\nws.Activate', (
            'ws.Rows(1).Insert', 'ws.Rows(20).Insert',
        )),
        ('ws.Parent.Worksheets.Add(After:=ws).Range("A1:B2").Value = 7\nws.Parent.Worksheets(2).Rows(1).Copy\nws.Activate', (
            'ws.Rows(1).Insert', 'ws.Rows(20).Insert',
        )),
        # Formulas, formats and ranges held across it.
        *((f'{FORMULAS}\nSet r = ws.Range("B5")\nSet q = ws.Range("B3")\nws.Rows(3).{verb}', tuple(
            f'{edit}\ns = r.Address(False, False) & "," & Forms(ws)\nErr.Clear\ns = s & "," & q.Address(False, False) & Err.Number\nErr.Clear'
            for edit in ('ws.Rows(1).Insert', 'ws.Rows(6).Insert', 'ws.Rows(20).Insert')
        )) for verb in ('Copy', 'Cut')),
        (f'{FORMULAS}\nSet r = ws.Range("D5")\nSet q = ws.Range("D3")\nws.Range("D3").Cut', tuple(
            f'{edit}\ns = r.Address(False, False) & "," & Forms(ws) & "," & q.Address(False, False)'
            for edit in ('ws.Range("D1").Insert xlShiftDown', 'ws.Range("F1").Insert xlShiftDown', 'ws.Range("D5").Insert xlShiftDown')
        )),
        ('ws.Rows(3).RowHeight = 40\nws.Range("B3").Font.Bold = True\nws.Rows(3).Copy', (
            'ws.Rows(6).Insert\ns = (ws.Rows(6).RowHeight > 30) & ws.Range("B6").Font.Bold & (ws.Rows(7).RowHeight > 30)',
        )),
        ('ws.Range("B3").Font.Bold = True\nws.Rows(3).Cut', (
            'ws.Rows(6).Insert\ns = ws.Range("B5").Font.Bold & ws.Range("B3").Font.Bold',
        )),
        ('ws.Range("B3:D4").Copy', ('ws.Rows(20).Insert', 'ws.Rows("20:21").Insert', 'ws.Columns(8).Insert', 'ws.Columns(1).Insert')),
        ('ws.Range("A3:E4").Copy', ('ws.Rows(20).Insert', 'ws.Rows("20:21").Insert', 'ws.Rows("20:23").Insert', 'ws.Rows(1).Insert')),
        ('ws.Range("A3:E3").Copy', ('ws.Rows("20:21").Insert',)),
        ('ws.Range("B3:C4").Copy', ('ws.Rows("1:2").Insert', 'ws.Rows("3:4").Insert', 'ws.Rows("20:22").Insert', 'ws.Columns(8).Insert', 'ws.Columns(1).Insert')),
        ('ws.Range("B3:B4").Copy', ('ws.Rows(1).Insert', 'ws.Rows("20:21").Insert', 'ws.Rows("20:22").Insert', 'ws.Rows(3).Insert')),
        ('ws.Range("B3:B5").Copy', ('ws.Rows("20:21").Insert',)),
        ('ws.Range("B3:B4").Cut', ('ws.Range("D1").Insert', 'ws.Range("D3").Insert', 'ws.Range("D5").Insert', 'ws.Range("A3").Insert', 'ws.Rows(1).Insert')),
        ('ws.Range("B3:C3").Cut', ('ws.Range("D1").Insert', 'ws.Range("D3").Insert', 'ws.Range("A3").Insert', 'ws.Range("B1").Insert', 'ws.Range("C1").Insert')),
        ('ws.Rows("3:4").Copy', ('ws.Rows(3).Insert', 'ws.Rows("3:6").Insert')),
        ('ws.Rows(3).Cut', ('ws.Range("A5").Insert', 'ws.Range("A3").Insert')),
        ('ws.Columns(2).Cut', ('ws.Range("D2").Insert',)),
        ('ws.Range("B3").Copy', ('ws.Columns("H:I").Insert', 'ws.Columns("B:C").Insert', 'ws.Columns("A:B").Insert')),
        ('ws.Range("B3:C4").Copy', ('ws.Columns("H:I").Insert', 'ws.Columns("A:B").Insert')),
        ('ws.Columns(2).Copy', ('ws.Columns("A:B").Insert', 'ws.Columns("G:H").Insert')),
        ('ws.Columns("B:C").Copy', ('ws.Columns("C").Insert', 'ws.Columns("A:D").Insert', 'ws.Columns("G:I").Insert')),
        ('ws.Rows(3).Copy', ('ws.Rows("6:7").Insert',)),
        ('ws.Range("B3:B4").Cut', ('ws.Range("B1").Insert', 'ws.Range("B6").Insert')),
        ('ws.Range("B3:C4").Cut', ('ws.Range("B6").Insert', 'ws.Range("B4").Insert', 'ws.Range("E3").Insert', 'ws.Range("D3").Insert')),
        ('ws.Rows("3:4").Cut', ('ws.Rows(2).Insert',)),
        ('ws.Columns("B:C").Cut', ('ws.Columns("A").Insert', 'ws.Columns("C").Insert', 'ws.Columns("D").Insert', 'ws.Columns("G:H").Insert')),
        ('ws.Range("B3").Font.Bold = True\nws.Range("B3").Copy', (
            'ws.Rows(6).Insert\ns = ws.Range("Z6").Font.Bold & ws.Range("Z7").Font.Bold',
        )),
    )
    for edit in edits
]
PROBE = """Private Function Look(ByVal ws As Object) As String
    Dim i As Long, j As Long, t As String
    t = Application.CutCopyMode & "|" & ws.UsedRange.Address(False, False) & "|" & Selection.Address(False, False)
    If ws.Parent.Worksheets.Count > 1 Then t = t & "|" & ws.Parent.Worksheets(2).UsedRange.Address(False, False)
    For i = 1 To 7
        t = t & "/"
        For j = 1 To 8
            t = t & ws.Cells(i, j).Text & ","
        Next
    Next
    Look = t
End Function

Private Function Forms(ByVal ws As Object) As String
    Dim i As Long, j As Long, t As String
    For i = 1 To 22
        For j = 1 To 9
            If ws.Cells(i, j).HasFormula Then t = t & ws.Cells(i, j).Address(False, False) & ws.Cells(i, j).Formula & ";"
        Next
    Next
    Forms = t
End Function

Private Function Fill() As Object
    Dim ws As Object, i As Long
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    ws.Range("A1:E1").Value = Array("Name", "Team", "Unit", "Qty", "Cost")
    For i = 2 To 9
        ws.Cells(i, 1).Value = "n" & i: ws.Cells(i, 2).Value = "t" & i: ws.Cells(i, 3).Value = "u" & i
        ws.Cells(i, 4).Value = i: ws.Cells(i, 5).Value = i * 2
    Next
    Set Fill = ws
End Function

Public Function Audit(ByVal which As Long) As String
    Dim ws As Object, wb As Object
    @@PRE@@
    Application.DisplayAlerts = False
    Set ws = Fill()
    Set wb = ws.Parent
    On Error GoTo Failed
    Select Case which
@@CASES@@
    End Select
    s = "ok " & s
Answered:
    On Error Resume Next
    Application.CutCopyMode = False
    wb.Close False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Audit = s
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & s
    Resume Answered
End Function
"""


def source() -> str:
    cases = '\n'.join(f'    Case {index}\n        {case}' for index, case in enumerate(CASES))
    return PROBE.replace('@@PRE@@', PRE).replace('@@CASES@@', cases)


def main() -> None:
    text = source()
    cases: list[list[str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, case in enumerate(CASES):
            result = excel.run_vba(text, 'Audit', (index,), timeout=60.0)
            assert result.ok and not result.dialogs, (case, result)
            cases.append([case, str(result.value)])
            print(index, case.replace('\n', '; ')[:80], '=>', result.value, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
