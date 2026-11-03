"""Which ways of holding a Range follow its cells: a With block, an argument, a collection, an array, a For Each variable, a
function's result, a range made from a held one; and what a For Each does when its own cells are deleted or moved.

    python scripts/measure_range_holding.py

writes tests/fixtures/range_holding.json, which tests/test_excel_range_tracking.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_holding.json'
HELPERS = '''Private Function ByValue(ByVal r As Range, ByVal ws As Object) As String
    ws.Rows(1).Insert
    ByValue = r.Address
End Function

Private Function ByReference(ByRef r As Range, ByVal ws As Object) As String
    ws.Rows(1).Insert
    ByReference = r.Address
End Function

Private Function AsVariant(ByVal r As Variant, ByVal ws As Object) As String
    ws.Rows(1).Insert
    AsVariant = r.Address
End Function

Private Function AsObject(ByVal r As Object, ByVal ws As Object) As String
    ws.Rows(1).Insert
    AsObject = r.Address
End Function

Private Function Found(ByVal ws As Object) As Range
    Set Found = ws.Range("B3")
End Function

Private Function Deleted(ByVal r As Range, ByVal ws As Object) As String
    ws.Rows(3).Delete
    Deleted = TypeName(r) & r.Address
End Function
'''
PRE = 'Dim r As Range, q As Range, c As Range, s As String, i As Long, v As Variant, o As Object, a(1 To 2) As Variant, cc As New Collection'
CASES = [
    'With ws.Range("B3")\nws.Rows(1).Insert\ns = .Address\nEnd With',
    'With ws.Range("B3")\nws.Rows(3).Delete\ns = TypeName(.Parent) & .Address\nEnd With',
    'With ws.Range("B3:C5")\nws.Rows(4).Insert\ns = .Address & .Count\nEnd With',
    'Set r = ws.Range("B3")\nWith r\nws.Rows(1).Insert\ns = .Address\nEnd With',
    's = ByValue(ws.Range("B3"), ws)',
    's = ByReference(ws.Range("B3"), ws)',
    's = AsVariant(ws.Range("B3"), ws)',
    's = AsObject(ws.Range("B3"), ws)',
    's = Deleted(ws.Range("B3"), ws)',
    'Set r = ws.Range("B3")\ns = ByValue(r, ws) & r.Address',
    'Set r = ws.Range("B3")\ns = ByReference(r, ws) & r.Address',
    'Set r = Found(ws)\nws.Rows(1).Insert\ns = r.Address',
    's = Found(ws).Address\nws.Rows(1).Insert\ns = s & Found(ws).Address',
    'cc.Add ws.Range("B3")\nws.Rows(1).Insert\ns = cc(1).Address',
    'cc.Add ws.Range("B3"), "k"\nws.Rows(3).Delete\ns = TypeName(cc("k"))\nOn Error Resume Next\ns = s & cc("k").Address & Err.Number',
    'Set a(1) = ws.Range("B3")\nws.Rows(1).Insert\ns = a(1).Address',
    'Set v = ws.Range("B3")\nws.Rows(1).Insert\ns = v.Address',
    'Set o = ws.Range("B3")\nws.Rows(1).Insert\ns = o.Address',
    'Set o = CreateObject("Scripting.Dictionary")\no.Add "k", ws.Range("B3")\nws.Rows(1).Insert\ns = o("k").Address',
    'Set r = ws.Range("B3")\nSet q = r\nws.Rows(1).Insert\ns = r.Address & q.Address & (r Is q)',
    'Set r = ws.Range("B3")\nSet q = r.Offset(1)\nws.Rows(1).Insert\ns = r.Address & q.Address',
    'Set r = ws.Range("B3:B4").Cells(2)\nws.Rows(1).Insert\ns = r.Address',
    'Set r = ws.Range("B3,D5")\nws.Rows(1).Insert\ns = r.Address & r.Areas.Count',
    'Set r = ws.Range("B3,D5")\nws.Rows(3).Delete\ns = r.Address & r.Areas.Count',
    'Set r = ws.Range("B3,D5")\nSet q = r.Areas(2)\nws.Rows(1).Insert\ns = q.Address',
    'Set r = Union(ws.Range("B3"), ws.Range("B5"))\nws.Rows(4).Insert\ns = r.Address',
    'Set r = Application.Intersect(ws.Range("A1:C5"), ws.Range("B3:D9"))\nws.Rows(1).Insert\ns = r.Address',
    'Set r = ws.Cells(3, 2)\nws.Columns(1).Insert\ns = r.Address',
    'Set r = ws.UsedRange\nws.Rows(1).Insert\ns = r.Address',
    'Set r = ws.Range("A1").CurrentRegion\nws.Rows(5).Insert\ns = r.Address',
    'For Each c In ws.Range("B3:B4")\nws.Rows(1).Insert\ns = s & c.Address & "|"\nNext',
    'For Each c In ws.Range("A2:A9")\nIf c.Row Mod 2 = 0 Then c.EntireRow.Delete\ns = s & "|"\nNext\ns = s & ws.Range("A2").Value & ws.Range("A3").Value & ws.Range("A4").Value & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',
    'For Each c In ws.Range("A2:A9")\nIf c.Value = "n3" Or c.Value = "n4" Then c.EntireRow.Delete\nNext\nFor i = 2 To 8\ns = s & ws.Cells(i, 1).Value & ","\nNext',
    'For Each c In ws.Range("D2:D5")\ni = i + 1\nIf i > 6 Then Exit For\ns = s & c.Address & "|"\nc.EntireRow.Insert\nNext\ns = s & ws.UsedRange.Address',
    'For Each c In ws.Range("D2:D5")\ns = s & c.Address\nc.Delete xlShiftUp\nNext',
    'For Each c In ws.Range("D2:D5")\ns = s & c.Value & ","\nIf c.Value = 3 Then ws.Rows(4).Delete\nNext',
    'For Each c In ws.Range("A2:A9").Rows\nIf c.Row Mod 2 = 0 Then c.EntireRow.Delete\ns = s & c.Row & "|"\nNext',
    'For Each c In ws.Range("A:B").Columns\ns = s & c.Address & "|"\nc.Delete\nNext',
    'For Each c In ws.Range("A2:A4,C2:C3")\ns = s & c.Address & "|"\nws.Rows(1).Insert\nNext',
    'Set r = ws.Range("A2:A4")\nFor Each c In r\nr.Rows(1).EntireRow.Delete\ns = s & "|"\nNext\ns = s & r.Address',
    'ws.Activate\nws.Range("B3").Select\nSet r = Selection\nws.Rows(1).Insert\ns = r.Address & Selection.Address & ActiveCell.Address',
    'For Each c In ws.Range("A2:A4")\nc.Cut ws.Cells(c.Row + 20, 8)\ns = s & c.Address & "|"\nNext',
    'Set r = ws.Range("B3")\nws.Range("B3").Cut ws.Range("H1")\nws.Range("H1").Cut ws.Range("J2")\ns = r.Address',
    'Set r = ws.Range("B3")\nws.Rows(3).Delete\nOn Error Resume Next\ns = (r Is Nothing) & TypeName(r)\nSet q = r\ns = s & Err.Number\nErr.Clear\nSet q = r.Offset(1)\ns = s & Err.Number\nErr.Clear\nv = r.Value\ns = s & Err.Number\nErr.Clear\nr.Value = 1\ns = s & Err.Number\nErr.Clear\nFor Each c In r\ns = s & "in"\nNext\ns = s & Err.Number & IsObject(r)\nErr.Clear\ns = s & IsError(r) & IsEmpty(r) & (ws.Range("A1").Address = "x")',
    'Set r = ws.Range("B3")\nws.Rows(3).Delete\nOn Error Resume Next\nws.Range("A1").Value = r\ns = Err.Number & Err.Description',
    'Set r = ws.Range("B3")\nws.Rows(3).Delete\nws.Rows(3).Insert\nOn Error Resume Next\ns = r.Address & Err.Number',
    'Set r = ws.Range("B3")\nSet q = ws.Range("B3")\ns = (r Is q)\nws.Rows(1).Insert\ns = s & (r Is q) & r.Address & q.Address',
    'Dim wb2 As Object\nSet r = ws.Range("B3")\nSet wb2 = ws.Parent\nws.Copy After:=wb2.Worksheets(wb2.Worksheets.Count)\ns = r.Worksheet.Name & r.Address',
]
PROBE = """Private Function Fill() As Object
    Dim ws As Object, i As Long
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    ws.Range("A1:E1").Value = Array("Name", "Team", "Unit", "Qty", "Cost")
    For i = 2 To 9
        ws.Cells(i, 1).Value = "n" & i: ws.Cells(i, 2).Value = "t" & i: ws.Cells(i, 3).Value = "u" & i
        ws.Cells(i, 4).Value = i: ws.Cells(i, 5).Value = i * 2
    Next
    Set Fill = ws
End Function

@@HELPERS@@

Public Function Held(ByVal which As Long) As String
    Dim ws As Object
    @@PRE@@
    Application.DisplayAlerts = False
    Set ws = Fill()
    On Error GoTo Failed
    Select Case which
@@CASES@@
    End Select
    s = "ok " & s
Answered:
    On Error GoTo 0
    ws.Parent.Close False
    Application.DisplayAlerts = True
    Held = s
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & s
    Resume Answered
End Function
"""


def source() -> str:
    cases = '\n'.join(f'    Case {index}\n        {case}' for index, case in enumerate(CASES))
    return PROBE.replace('@@HELPERS@@', HELPERS).replace('@@PRE@@', PRE).replace('@@CASES@@', cases)


def main() -> None:
    text = source()
    cases: list[list[str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, case in enumerate(CASES):
            result = excel.run_vba(text, 'Held', (index,), timeout=60.0)
            assert result.ok and not result.dialogs, (case, result)
            cases.append([case, str(result.value)])
            print(index, case.replace('\n', '; ')[:90], '=>', result.value, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
