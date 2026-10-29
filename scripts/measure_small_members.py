"""Small members everyday macros read: Creator across Excel's objects, Range.InsertIndent, Workbook.HasVBProject and
CheckCompatibility, Name.Index, and Names' Parent and Application.

Each answer is recorded as text, an error as its number and description.

    python scripts/measure_small_members.py

writes tests/fixtures/small_members.json, which tests/test_excel_small_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/small_members.json'

#: Objects whose Creator is read, each as a VBA expression, the sheet ws and workbook wb set up first.
CREATORS = ['Application', 'wb', 'ws', 'ws.Range("A1")', 'ws.Range("A1").Font', 'ws.Range("A1").Interior',
            'ws.Range("A1").Borders', 'ws.Range("A1").Borders(xlEdgeTop)', 'wb.Names', 'wb.Names(1)',
            'Application.Windows', 'ActiveWindow', 'ws.Hyperlinks', 'wb.Sheets', 'wb.Worksheets', 'Application.Workbooks',
            'ws.Shapes', 'ws.Comments', 'ws.ListObjects', 'ws.Range("A1").Validation', 'ws.Range("A1").FormatConditions',
            'ws.PageSetup', 'ws.Range("A1:B2,D4").Areas', 'ws.Range("A1").Characters', 'wb.Styles', 'wb.Styles(1)',
            'Application.WorksheetFunction', 'ws.Sort', 'ws.Outline', 'ws.VPageBreaks', 'ws.Cells', 'ws.Rows',
            'ws.Range("A1").EntireColumn', 'wb.BuiltinDocumentProperties']
#: Each case: a statement setting it up, then what is read.
CASES = [
    ('ws.Range("A1").InsertIndent 2', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent 3', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent -10', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent 250', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent 1', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent 2.6', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent "1"', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").InsertIndent "x"', 'ws.Range("A1").IndentLevel'),
    ('ws.Range("A1").IndentLevel = 0', 'ws.Range("A1").HorizontalAlignment'),
    ('ws.Range("A1").InsertIndent 1', 'ws.Range("A1").HorizontalAlignment'),
    ('ws.Range("B1").HorizontalAlignment = xlCenter: ws.Range("B1").InsertIndent 1',
     'ws.Range("B1").HorizontalAlignment & " " & ws.Range("B1").IndentLevel'),
    ('ws.Range("C1").HorizontalAlignment = xlRight: ws.Range("C1").InsertIndent 1',
     'ws.Range("C1").HorizontalAlignment & " " & ws.Range("C1").IndentLevel'),
    ('ws.Range("D1").IndentLevel = 3: ws.Range("D2").IndentLevel = 5: ws.Range("D1:D2").InsertIndent 1',
     'ws.Range("D1").IndentLevel & " " & ws.Range("D2").IndentLevel & " " & IsNull(ws.Range("D1:D2").IndentLevel)'),
    ('ws.Range("E1").InsertIndent 0', 'ws.Range("E1").IndentLevel & " " & ws.Range("E1").HorizontalAlignment'),
    ('ws.Range("F1").IndentLevel = 249: ws.Range("F1").InsertIndent 2', 'ws.Range("F1").IndentLevel'),
    ('ws.Range("G1").InsertIndent 251', 'ws.Range("G1").IndentLevel'),
    ('', 'wb.HasVBProject'),
    ('', 'ThisWorkbook.HasVBProject'),
    ('', 'wb.CheckCompatibility'),
    ('wb.CheckCompatibility = True', 'wb.CheckCompatibility & " " & wb.Saved'),
    ('wb.CheckCompatibility = "x"', 'wb.CheckCompatibility'),
    ('wb.Names.Add "Zed", "=1": wb.Names.Add "Alpha", "=2": wb.Names.Add "mid", "=3"',
     'wb.Names("Zed").Index & " " & wb.Names("Alpha").Index & " " & wb.Names("mid").Index & " " & wb.Names(1).Name'),
    ('ws.Names.Add "Local", "=4"', 'wb.Names("Sheet1!Local").Index & " " & ws.Names("Local").Index'),
    ('', 'TypeName(wb.Names.Parent) & " " & TypeName(wb.Names.Application) & " " & TypeName(ws.Names.Parent)'),
    ('', 'TypeName(wb.Names(1).Parent) & " " & TypeName(wb.Names(1).Application)'),
]


def _probe() -> str:
    lines = ['Private wb As Workbook', 'Private ws As Worksheet', '',
             'Private Function Shown(ByVal v As Variant) As String',
             '    If IsNull(v) Then Shown = "Null" Else Shown = TypeName(v) & ":" & CStr(v)', 'End Function', '',
             'Public Function Creators() As String', '    Dim s As String',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    wb.Names.Add "First", "=1"',
             '    ws.Range("A1").Value = "x"', '    ws.Range("A1").AddComment "note"', '    On Error Resume Next']
    for expression in CREATORS:
        lines += ['    Err.Clear', f'    s = s & "|" & Shown({expression}.Creator)',
                  '    If Err.Number <> 0 Then s = s & "Err " & Err.Number']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Application.DisplayAlerts = True', '    Creators = Mid(s, 2)',
              'End Function', '']
    lines += ['Private Function Tried(ByVal which As Long) As String', '    On Error GoTo Failed', '    Select Case which']
    for index, (setup, read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        if setup:
            lines.append(f'        {setup}')
        lines.append(f'        Tried = Shown({read})')
    lines += ['    End Select', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description',
              'End Function', '', 'Public Function Cases() As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
              '    Set ws = wb.Worksheets(1)', '    ws.Range("A1:G1").Value = "x"',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "|" & Tried(i)', '    Next', '    wb.Close False',
              '    Application.DisplayAlerts = True', '    Cases = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        answers = {}
        for proc in ('Creators', 'Cases'):
            result = excel.run_vba(source, proc, timeout=120.0, module_name='SmallProbe')
            assert result.ok and not result.dialogs, result
            answers[proc] = str(result.value).split('|')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'creators': dict(zip(CREATORS, answers['Creators'], strict=True)),
              'cases': [[setup, read, answer] for (setup, read), answer in zip(CASES, answers['Cases'], strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for key, value in record['creators'].items():
        print(key, '->', value)
    for setup, read, answer in record['cases']:
        print(setup, '|', read, '->', answer)


if __name__ == '__main__':
    main()
