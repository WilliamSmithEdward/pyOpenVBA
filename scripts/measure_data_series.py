"""Range.DataSeries: what Excel writes for each kind of series, to the bit.

Each case clears a sheet, writes its start cells, runs one DataSeries and
records, for every cell of the range, its type and value: a Double as the
hex of its eight bytes, copied with LSet since VBA's own comparisons run in
extended precision, and anything else as text. Errors are recorded as their
number and description.

    python scripts/measure_data_series.py

writes tests/fixtures/data_series.json, which tests/test_excel_data_series.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/data_series.json'

#: Each case: the cells set first (a statement each, ws the sheet), the range, and DataSeries' arguments.
CASES = [
    (['ws.Range("A1").Value = 1'], 'A1:A10', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=2'),
    (['ws.Range("A1").Value = 0'], 'A1:A40', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=0.1'),
    (['ws.Range("A1").Value = 0.3'], 'A1:A40', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=0.7'),
    (['ws.Range("A1").Value = 1 / 3'], 'A1:A30', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=1 / 7'),
    (['ws.Range("A1").Value = 1'], 'A1:A10', 'Rowcol:=xlColumns, Type:=xlLinear'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2:A10").Value = "old"'], 'A1:A10',
     'Rowcol:=xlColumns, Type:=xlLinear, Step:=3, Stop:=10'),
    (['ws.Range("A1").Value = 10'], 'A1:A10', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=-3, Stop:=0'),
    (['ws.Range("A1").Value = 1'], 'A1:A10', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=0'),
    (['ws.Range("A1").Value = 1'], 'A1:A12', 'Rowcol:=xlColumns, Type:=xlGrowth, Step:=2'),
    (['ws.Range("A1").Value = 1'], 'A1:A40', 'Rowcol:=xlColumns, Type:=xlGrowth, Step:=1.1'),
    (['ws.Range("A1").Value = 3'], 'A1:A10', 'Rowcol:=xlColumns, Type:=xlGrowth, Step:=2, Stop:=50'),
    (['ws.Range("A1").Value = 1'], 'A1:E1', ''),
    (['ws.Range("A1").Value = 1', 'ws.Range("B1").Value = 10'], 'A1:B5', 'Type:=xlLinear, Step:=5'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 100'], 'A1:C2', 'Type:=xlLinear, Step:=5'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 100'], 'A1:C3', 'Rowcol:=xlRows, Step:=5'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 5'], 'A1:A8', 'Rowcol:=xlColumns, Type:=xlLinear'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 3'], 'A1:A8', 'Rowcol:=xlColumns, Trend:=True'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 3'], 'A1:A8', 'Rowcol:=xlColumns, Type:=xlGrowth, Trend:=True'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 2', 'ws.Range("A3").Value = 4'], 'A1:A8',
     'Rowcol:=xlColumns, Trend:=True'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 3'], 'A1:A8', 'Rowcol:=xlColumns, Type:=xlAutoFill'),
    (['ws.Range("A1").Value = "Item 1"'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlAutoFill'),
    (['ws.Range("A1").Value = "Item 1"'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlLinear'),
    ([], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=1'),
    (['ws.Range("A1").Formula = "=1+1"'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlLinear, Step:=1'),
    (['ws.Range("A1").Value = 5', 'ws.Range("A1").NumberFormat = "0.00"'], 'A1:A4', 'Rowcol:=xlColumns, Step:=1'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 31)'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlDay'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 30)'], 'A1:A8', 'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlWeekday'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 31)'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlMonth'),
    (['ws.Range("A1").Value = DateSerial(2024, 2, 29)'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlYear'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 31)'], 'A1:A6',
     'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlMonth, Step:=2'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 1)'], 'A1:A6',
     'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlDay, Step:=7, Stop:=DateSerial(2026, 1, 25)'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 1) + TimeSerial(6, 0, 0)'], 'A1:A4',
     'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlDay, Step:=0.5'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlDay'),
    (['ws.Range("A1").Value = 1'], 'A1', 'Rowcol:=xlColumns, Step:=1'),
    (['ws.Range("A1").Value = 1', 'ws.Range("C1").Value = 1'], 'A1:A3,C1:C3', 'Rowcol:=xlColumns, Step:=1'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=3, Step:=1'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=xlColumns, Type:=9, Step:=1'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=xlColumns, Step:="x"'),
    (['ws.Range("A1").Value = True'], 'A1:A4', 'Rowcol:=xlColumns, Step:=1'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = "x"'], 'A1:A4', 'Rowcol:=xlColumns, Trend:=True'),
    (['ws.Range("A1").Value = 0.1', 'ws.Range("A2").Value = 0.3'], 'A1:A12', 'Rowcol:=xlColumns, Trend:=True'),
    (['ws.Range("A1").Value = 0.1', 'ws.Range("A2").Value = 0.3'], 'A1:A12', 'Rowcol:=xlColumns, Type:=xlAutoFill'),
    (['ws.Range("A1").Value = 0.1'], 'A1:A12', 'Rowcol:=xlColumns, Step:=0.2'),
    (['ws.Range("A1").Value = 0.1', 'ws.Range("A2").Value = 0.3', 'ws.Range("A3").Value = 0.5'], 'A1:A8',
     'Rowcol:=xlColumns, Trend:=True'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 31)'], 'A1:A6',
     'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlWeekday'),
    (['ws.Range("A1").Value = DateSerial(2026, 1, 31) + TimeSerial(9, 30, 0)'], 'A1:A4',
     'Rowcol:=xlColumns, Type:=xlChronological, Date:=xlMonth'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A2").Value = 10'], 'A1:B2', 'Step:=5'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlGrowth, Step:=-2'),
    (['ws.Range("A1").Value = 0'], 'A1:A4', 'Rowcol:=xlColumns, Type:=xlGrowth, Step:=3'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=xlColumns, Step:=2, Stop:=6'),
    (['ws.Range("A1").Value = 7'], 'A1:A6', 'Rowcol:=xlColumns, Step:=2, Stop:=3'),
    (['ws.Range("A1").Value = 1'], 'A1:A6', 'Rowcol:=xlColumns, Type:=xlGrowth, Step:=0.5, Stop:=0.2'),
    (['ws.Range("A1").Value = 1', 'ws.Range("A1").Font.Bold = True', 'ws.Range("A3").Interior.Color = 255'], 'A1:A4',
     'Rowcol:=xlColumns, Step:=1'),
]


def _probe() -> str:
    lines = ['Private Type DoubleBits', '    v As Double', 'End Type', 'Private Type ByteBits', '    b(7) As Byte',
             'End Type', '', 'Private Function Bits(ByVal value As Variant) As String',
             '    Dim d As DoubleBits, e As ByteBits, i As Long, s As String', '    d.v = value', '    LSet e = d',
             '    For i = 7 To 0 Step -1', '        s = s & Right$("0" & Hex$(e.b(i)), 2)', '    Next', '    Bits = s',
             'End Function', '', 'Private Function Cells(ByVal r As Range) As String', '    Dim c As Range, s As String',
             '    For Each c In r', '        If TypeName(c.Value) = "Double" Then',
             '            s = s & "," & Bits(c.Value)', '        Else',
             '            s = s & "," & TypeName(c.Value) & ":" & CStr(c.Value)', '        End If',
             '        s = s & ":" & c.HasFormula & ":" & c.Text & ":" & c.Font.Bold & ":" & c.Interior.Color',
             '    Next', '    Cells = Mid(s, 2)', 'End Function', '',
             'Private Function Tried(ByVal ws As Worksheet, ByVal which As Long) As String',
             '    Dim v As Variant, target As String', '    ws.Cells.Clear', '    Select Case which']
    for index, (setup, target, arguments) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {line}' for line in setup]
        lines.append(f'        target = "{target}"')
    lines += ['    End Select', '    On Error GoTo Failed', '    Select Case which']
    for index, (_, target, arguments) in enumerate(CASES):
        lines.append(f'    Case {index}: v = ws.Range("{target}").DataSeries({arguments})')
    lines += ['    End Select', '    Tried = TypeName(v) & ":" & CStr(v) & "=" & Cells(ws.Range(target).Areas(1).Resize(, 1).Parent'
              '.Range(target))', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description & "=" & Cells(ws.Range(target))', 'End Function',
              '', 'Public Function Probe() As String', '    Dim ws As Worksheet, s As String, i As Long',
              '    Set ws = ThisWorkbook.Worksheets.Add', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & vbLf & Tried(ws, i)', '    Next', '    Application.DisplayAlerts = False', '    ws.Delete',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


#: A sink for the Change events DataSeries raises; the harness runs with events off, so the macro turns them on and
#: puts the setting back.
SINK = '''Public WithEvents App As Application
Public Log As String

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "|change " & Target.Address(False, False)
End Sub
'''

EVENTS = '''Private sink As SeriesSink

Public Function Events() As String
    Dim ws As Worksheet, s As String, wasOn As Boolean
    wasOn = Application.EnableEvents
    Application.EnableEvents = True
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Range("A1").Value = 1
    ws.Range("C1").Value = 1
    ws.Range("E5").Select
    Set sink = New SeriesSink
    Set sink.App = Application
    ws.Range("A1:A5").DataSeries Rowcol:=xlColumns, Step:=1
    s = "column:" & sink.Log & "#selection " & Selection.Address(False, False)
    sink.Log = ""
    ws.Range("C1:C3,A7:A8").DataSeries Rowcol:=xlColumns, Step:=1
    s = s & "#areas:" & sink.Log
    sink.Log = ""
    ws.Range("G1:G3").DataSeries Rowcol:=xlColumns, Step:=1
    s = s & "#empty:" & sink.Log
    Set sink.App = Nothing
    Set sink = Nothing
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Application.EnableEvents = wasOn
    Events = s
End Function
'''


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('\n')
        excel.add_module('SeriesSink', SINK, kind='class')
        events = excel.run_vba(EVENTS, 'Events', timeout=120.0, module_name='SeriesEvents')
        assert events.ok and not events.dialogs, events
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'sink': SINK,
              'events_source': EVENTS, 'events': str(events.value),
              'cases': [[setup, target, arguments, answer]
                        for (setup, target, arguments), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for setup, target, arguments, answer in record['cases']:
        print(target, arguments, '->', answer[:300])


if __name__ == '__main__':
    main()
