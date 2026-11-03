"""What a Range held in a variable does when the cells around it move, and what a chart makes of several areas.

Each case makes a sheet with values in A1:E9, holds ranges in variables, makes an edit, and reads each held range's Address, or
"E" and an error number where it has gone. The chart cases read each series' formula.

    python scripts/measure_range_tracking.py

writes tests/fixtures/range_tracking.json, which tests/test_excel_range_tracking.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_tracking.json'
EDITS = [
    'ws.Rows(1).Insert', 'ws.Rows(3).Insert', 'ws.Rows(5).Insert', 'ws.Rows("2:3").Insert', 'ws.Rows(9).Insert', 'ws.Rows(1).Delete', 'ws.Rows(3).Delete', 'ws.Rows(4).Delete', 'ws.Rows("3:4").Delete', 'ws.Rows("2:6").Delete', 'ws.Rows("1:2").Delete', 'ws.Rows(9).Delete',
    'ws.Columns(1).Insert', 'ws.Columns(2).Insert', 'ws.Columns(3).Insert', 'ws.Columns(1).Delete', 'ws.Columns(2).Delete', 'ws.Columns("B:C").Delete', 'ws.Columns("A:D").Delete',
    'ws.Range("B1").Insert xlShiftDown', 'ws.Range("B3").Insert xlShiftDown', 'ws.Range("A3").Insert xlShiftDown', 'ws.Range("B1").Insert xlShiftToRight', 'ws.Range("A3").Insert xlShiftToRight', 'ws.Range("B3").Insert xlShiftToRight', 'ws.Range("B1").Delete xlShiftUp', 'ws.Range("B3").Delete xlShiftUp', 'ws.Range("A1").Delete xlShiftToLeft', 'ws.Range("A3").Delete xlShiftToLeft', 'ws.Range("B3").Delete xlShiftToLeft', 'ws.Range("B3:C4").Delete xlShiftUp', 'ws.Range("B3:C4").Insert xlShiftDown', 'ws.Range("A1:E1").Delete xlShiftUp',
    'ws.Range("B3").Cut ws.Range("H1")', 'ws.Range("B3:C4").Cut ws.Range("H1")', 'ws.Range("B2:C5").Cut ws.Range("H1")', 'ws.Range("A1").Cut ws.Range("B3")', 'ws.Range("B3").Copy ws.Range("H1")', 'ws.Range("B3").Clear', 'ws.Range("B3:C4").ClearContents', 'ws.Sort.SortFields.Clear', 'ws.Range("A1:E9").Sort ws.Range("A1"), xlDescending', 'ws.Range("B3").Value = 1',
    'ws.Name = "Moved"', 'wo.Move', 'ws.Move After:=wb.Worksheets(wb.Worksheets.Count)', 'wb.Worksheets.Add Before:=ws', 'ws.Delete',
]
READS = ['p.Address', 'q.Address', 'e.Address', 'w.Address', 'p.Worksheet.Name', 'p.Value', 'q.Count', 'TypeName(p)', 'p Is Nothing', 'p.Row', 'q.Rows.Count', 'p.Offset(1).Address']
SOURCES = ['A1:A9,D1:D9', 'A1:A9,D1:E9', 'A2:A9,D2:D9', 'D1:D9,E1:E9', 'A1:B9,D1:D9', 'A1:E1,A3:E3', 'A1:E1,A3:E4', 'A1:A5,D1:D9', 'A1:A9,D3:D9', 'D1:D9,A1:A9', 'E1:E9,A1:A9,D1:D9', 'A1,D1', 'A1:A9,D1:D9,B1:B9', 'A1:A9,A1:A9', 'D2:D9,E2:E9', 'A:A,D:D', '1:1,3:3']
KINDS = ['ch.SetSourceData ws.Range("{}")', 'ws.Range("{}").Select: Set ch = ws.Shapes.AddChart2(201, xlColumnClustered).Chart']
PROBE = """Private Function Fill(ByVal wb As Object) As Object
    Dim ws As Object, i As Long
    Set ws = wb.Worksheets(1)
    ws.Range("A1:E1").Value = Array("Name", "Team", "Unit", "Qty", "Cost")
    For i = 2 To 9
        ws.Cells(i, 1).Value = "n" & i: ws.Cells(i, 2).Value = "t" & i: ws.Cells(i, 3).Value = "u" & i
        ws.Cells(i, 4).Value = i: ws.Cells(i, 5).Value = i * 2
    Next
    Set Fill = ws
End Function

Public Function Moved(ByVal edit As Long) As String
    Dim wb As Object, ws As Object, wo As Object, p As Object, q As Object, e As Object, w As Object, v As String, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets.Add After:=wb.Worksheets(1)
    Set ws = Fill(wb)
    Set wo = wb.Worksheets(2)
    Set p = ws.Range("B3")
    Set q = ws.Range("B3:C5")
    Set e = ws.Rows(3)
    Set w = ws.Columns(2)
    On Error GoTo Failed
    Select Case edit
@@EDITS@@
    End Select
    s = "ok "
Answered:
    On Error Resume Next
@@READS@@
    On Error GoTo 0
    wb.Close False
    Application.DisplayAlerts = True
    Moved = s
    Exit Function
Failed:
    s = "Err " & Err.Number & " "
    Resume Answered
End Function

Public Function Charted(ByVal which As Long) As String
    Dim wb As Object, ws As Object, ch As Object, s As String, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = Fill(wb)
    ws.Range("H20").Select
    If which < 100 Then Set ch = ws.ChartObjects.Add(10, 10, 200, 100).Chart
    On Error GoTo Failed
    Select Case which
@@SOURCES@@
    End Select
    s = "ok " & ch.SeriesCollection.Count & " " & ch.PlotBy
    For i = 1 To ch.SeriesCollection.Count
        s = s & "|" & ch.SeriesCollection(i).Formula
    Next
Answered:
    On Error GoTo 0
    wb.Close False
    Application.DisplayAlerts = True
    Charted = s
    Exit Function
Failed:
    s = "Err " & Err.Number
    Resume Answered
End Function
"""


def source() -> str:
    edits = '\n'.join(f'    Case {index}\n        {edit}' for index, edit in enumerate(EDITS))
    reads = '\n'.join(f'    Err.Clear\n    v = ""\n    v = {read}\n    If Err.Number Then v = "E" & Err.Number\n    s = s & "|" & v' for read in READS)
    sources = '\n'.join(f'    Case {offset + index}\n        {kind.format(area)}' for offset, kind in zip((0, 100), KINDS, strict=True)
                        for index, area in enumerate(SOURCES))
    return PROBE.replace('@@EDITS@@', edits).replace('@@READS@@', reads).replace('@@SOURCES@@', sources)


def main() -> None:
    text = source()
    moved: list[list[str]] = []
    charted: list[list[str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, edit in enumerate(EDITS):
            result = excel.run_vba(text, 'Moved', (index,), timeout=60.0)
            assert result.ok and not result.dialogs, (edit, result)
            moved.append([edit, str(result.value)])
            print(edit, '=>', result.value, flush=True)
        for offset, kind in zip((0, 100), KINDS, strict=True):
            for index, area in enumerate(SOURCES):
                result = excel.run_vba(text, 'Charted', (offset + index,), timeout=60.0)
                assert result.ok and not result.dialogs, (area, result)
                charted.append([kind.format(area), str(result.value)])
                print(kind.format(area)[:50], '=>', result.value, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': text, 'moved': moved, 'charted': charted}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
