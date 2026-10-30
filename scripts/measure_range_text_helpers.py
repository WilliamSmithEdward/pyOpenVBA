"""Range.ListNames, Range.AutoComplete and Range.Parse: what Excel writes or answers for each.

- ListNames: the names of a workbook pasted from a cell, hidden, sheet and workbook names among them.
- AutoComplete: what a cell's column offers for a typed beginning, above and below the cell.
- Parse: text split at the widths a bracket pattern marks, into the cells from Destination.

Each case records the cells it reads back, or the error raised.

    python scripts/measure_range_text_helpers.py

writes tests/fixtures/range_text_helpers.json, which tests/test_excel_range_text_helpers.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_text_helpers.json'

#: Each case: statements on a fresh sheet ws of a new workbook wb, and the cells or expression read back.
CASES = [
    (['wb.Names.Add "Zed", "=Sheet1!$A$1"', 'wb.Names.Add "alpha", "=10"', 'ws.Names.Add "Local", "=Sheet1!$B$2"',
      'wb.Names.Add "Hidden", "=5", Visible:=False', 'wb.Names.Add "Mid", "=SUM(Sheet1!$A$1:$A$3)"',
      'ws.Range("D2").ListNames'], 'D1:F7'),
    (['ws.Range("D2:E3").ListNames'], 'D1:E3'),
    (['wb.Names.Add "One", "=1"', 'ws.Range("D2:E3").ListNames'], 'D1:E3'),
    (['wb.Names.Add "One", "=1"', 'ws.Range("D2").Value = "keep"', 'ws.Range("E2").Value = "keep"',
      'ws.Range("D3").Value = "keep"', 'ws.Range("D2").ListNames'], 'D1:E3'),
    (['wb.Names.Add "One", "=1"', 'ws.Range("D2,G2").ListNames'], 'D1:H3'),
    (['ws.Range("A1").Value = "Apple"', 'ws.Range("A2").Value = "Apricot"', 'ws.Range("A3").Value = "Banana"',
      'ws.Range("A5").Value = "blueberry"', 'ws.Range("A6").Value = 123'],
     'ws.Range("A4").AutoComplete("b") & "|" & ws.Range("A4").AutoComplete("ap") & "|" & '
     'ws.Range("A4").AutoComplete("app") & "|" & ws.Range("A4").AutoComplete("Ban") & "|" & '
     'ws.Range("A4").AutoComplete("x") & "|" & ws.Range("A4").AutoComplete("1") & "|" & '
     'ws.Range("A4").AutoComplete("") & "|" & ws.Range("A4").AutoComplete("BLUE") & "|" & '
     'ws.Range("A4").AutoComplete("apple") & "|" & ws.Range("B4").AutoComplete("ap")'),
    (['ws.Range("A1").Value = "Apple"', 'ws.Range("A3").Value = "Apple pie"'],
     'ws.Range("A2").AutoComplete("app") & "|" & ws.Range("A10").AutoComplete("app") & "|" & '
     'ws.Range("A4").AutoComplete("Apple ")'),
    (['ws.Range("A1").Value = "Apple"', 'ws.Range("A2").Value = "apple"'], 'ws.Range("A3").AutoComplete("a")'),
    (['ws.Range("A1").Value = "Apple"', 'ws.Range("A2").Value = "Apricot"', 'Application.EnableAutoComplete = False'],
     'ws.Range("A3").AutoComplete("app")'),
    (['ws.Range("A1").Value = "Apple"', 'ws.Range("A2").Value = ""', 'ws.Range("A4").Value = "Apricot"'],
     'ws.Range("A3").AutoComplete("apr") & "|" & ws.Range("A3").AutoComplete("appl")'),
    (['ws.Range("A1").Value = "Apple"'], 'ws.Range("A1:A2").AutoComplete("ap")'),
    (['ws.Range("A1").Value = "20261029 Smith   0042"', 'ws.Range("A1").Parse "[xxxxxxxx] [xxxxxxxx][xxxx]"'],
     'A1:E1'),
    (['ws.Range("A1").Value = "20261029 Smith   0042"', 'ws.Range("A1").Parse "[xxxxxxxx][xxxxxxxxx][xxxx]"'],
     'A1:E1'),
    (['ws.Range("A1").Value = "abcdefgh"', 'ws.Range("A2").Value = "ijklmnop"', 'ws.Range("A1:A2").Parse "[xx][xxx]"'],
     'A1:D2'),
    (['ws.Range("A1").Value = "abcdefgh"', 'ws.Range("A1").Parse "[xx][xxx]", ws.Range("C5")'], 'A1:F5'),
    (['ws.Range("A1").Value = "abcdefgh"', 'ws.Range("A1").Parse'], 'A1:D1'),
    (['ws.Range("A1").Value = "abcdefgh"', 'ws.Range("A1").Parse "xx"'], 'A1:D1'),
    (['ws.Range("A1").Value = "abcdefgh"', 'ws.Range("A1:B1").Parse "[xx][xx]"'], 'A1:D1'),
    (['ws.Range("A1").Value = "0042abc"', 'ws.Range("A1").Parse "[xxxx][xxx]"'], 'A1:C1'),
    (['ws.Range("A1").Value = 12345678', 'ws.Range("A1").Parse "[xxxx][xxxx]"'], 'A1:C1'),
    (['ws.Range("A1").Value = "ab"', 'ws.Range("A1").Parse "[xxxx][xxxx]"'], 'A1:C1'),
    (['ws.Range("A1").Value = "abcdefgh"', 'ws.Range("A1").Parse "[xx]  [xx]"'], 'A1:C1'),
    (['wb.Worksheets.Add(After:=ws).Name = "Sheet2"', 'wb.Worksheets("Sheet2").Names.Add "Other", "=Sheet2!$A$1"',
      'wb.Names.Add "alpha", "=10"', 'ws.Activate', 'ws.Range("D2").ListNames'], 'D1:E4'),
    (['wb.Names.Add "Same", "=1"', 'ws.Names.Add "Same", "=2"', 'ws.Range("D2").ListNames'], 'D1:E4'),
    (['wb.Worksheets.Add(After:=ws).Name = "Sheet2"', 'wb.Worksheets("Sheet2").Names.Add "Two", "=2"',
      'ws.Names.Add "One", "=1"', 'ws.Activate', 'wb.Worksheets("Sheet2").Range("D2").ListNames',
      'ws.Range("A1").Value = wb.Worksheets("Sheet2").Range("D2").Formula & "/" & '
      'wb.Worksheets("Sheet2").Range("D3").Formula'], 'A1:A1'),
    (['ws.Range("A1").Value = "Apple"', 'ws.Range("A2").Formula = "=""Apricot"""', 'ws.Range("A3").Value = "Banana"'],
     'ws.Range("A4").AutoComplete("apr") & "|" & ws.Range("A4").AutoComplete("app") & "|" & '
     'ws.Range("A4").AutoComplete("ban")'),
    (['ws.Range("A1").Formula = "=1+1"', 'ws.Range("A2").Value = "Apple"'], 'ws.Range("A3").AutoComplete("app")'),
]


def _probe() -> str:
    lines = ['Private Function Cells(ByVal ws As Worksheet, ByVal address As String) As String', '    Dim c As Range, s As String',
             '    For Each c In ws.Range(address)',
             '        s = s & ";" & c.Address(False, False) & "=" & TypeName(c.Value) & ":" & c.Formula', '    Next',
             '    Cells = Mid(s, 2)', 'End Function', '', 'Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, answer As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Name = "Sheet1"', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if ':' in read and read.replace(':', '').replace('$', '').isalnum() and not read.startswith('ws.'):
            lines.append(f'        answer = Cells(ws, "{read}")')
        else:
            lines.append(f'        answer = {read}')
    lines += ['    End Select', '    Tried = answer', 'Done:', '    Application.EnableAutoComplete = True',
              '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim s As String, i As Long', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


#: A sink for the Change events ListNames raises; the harness runs with events off, so the macro turns them on and
#: puts the setting back.
SINK = '''Public WithEvents App As Application
Public Log As String

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "|change " & Sh.Name & "!" & Target.Address(False, False)
End Sub
'''

EVENTS = '''Private sink As NamesSink

Public Function Events() As String
    Dim wb As Workbook, wasOn As Boolean
    wasOn = Application.EnableEvents
    Application.EnableEvents = True
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "Sheet1"
    wb.Names.Add "One", "=1"
    wb.Names.Add "Two", "=2"
    Set sink = New NamesSink
    Set sink.App = Application
    wb.Worksheets(1).Range("D2").ListNames
    Events = Mid(sink.Log, 2)
    Set sink.App = Nothing
    Set sink = Nothing
    wb.Close False
    Application.EnableEvents = wasOn
End Function
'''


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        excel.add_module('NamesSink', SINK, kind='class')
        events = excel.run_vba(EVENTS, 'Events', timeout=120.0, module_name='NamesEvents')
        assert events.ok and not events.dialogs, events
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'sink': SINK,
              'events_source': EVENTS, 'events': str(events.value),
              'cases': [[statements, read, answer] for (statements, read), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, answer in record['cases']:
        print(statements[-1], '->', answer[:400])


if __name__ == '__main__':
    main()
