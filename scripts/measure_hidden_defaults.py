"""The hidden _Default members, reached through bracketed names as a macro reaches them: what each answers next to the
visible member it stands for, and the errors. Workbook._CodeName is left out: it reads as CodeName does, and a
workbook a macro has just made reads an empty code name until its project is compiled. Windows._Default is compared
by caption: Is answers False against ActiveWindow, COM handing back another pointer to the same window.

Each case makes a workbook whose sheet holds 4 in A1 and 5 in B1, a rectangle, a chart over A1:B1, a name Total for
A1, a list object over D1:E3 and a drop-down, runs the case's statement and answers what it set.

    python scripts/measure_hidden_defaults.py

writes tests/fixtures/hidden_defaults.json, which tests/test_excel_hidden_defaults.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/hidden_defaults.json'
CASES = [
    's = ws.Shapes.[_Default](1).Name & "," & ws.Shapes.[_Default]("Box").Name & "," & ws.Shapes(1).Name',
    's = TypeName(ws.ChartObjects.[_Default](1)) & "," & ws.ChartObjects.[_Default](1).Name',
    's = wb.Sheets.[_Default](1).Name & "," & wb.Worksheets.[_Default](1).Name',
    's = (Workbooks.[_Default](wb.Name) Is wb)',
    's = wb.Names.[_Default]("Total").RefersTo',
    's = (Windows.[_Default](1).Caption = ActiveWindow.Caption)',
    's = wb.Styles.[_Default]("Normal").Name',
    's = ch.SeriesCollection.[_Default](1).Name',
    's = ch.Axes.[_Default](xlValue).Type',
    's = ws.Range("A1:B1,D1").Areas.[_Default](2).Address',
    's = ws.Range("A1").Borders.[_Default](xlEdgeTop).LineStyle',
    's = (wb.PivotCaches.Count = 0)',
    's = ws.ListObjects.[_Default](1).Name',
    's = ws.Range("A1").[_Default] & "," & ws.Range("A1:B1").Cells(2).[_Default]',
    'ws.Range("C1").[_Default] = 7: s = ws.Range("C1").Value',
    's = ws.Range("A1:B1").[_Default](1, 2)',
    's = wb.Names("Total").[_Default] & "|" & wb.Names("Total").Name',
    's = Application.[_Default] & "|" & Application.Name',
    's = ws.Shapes("Drop").ControlFormat.[_Default] & "|" & ws.Shapes("Drop").ControlFormat.Value',
    'ws.Shapes("Drop").ControlFormat.[_Default] = 2: s = ws.Shapes("Drop").ControlFormat.Value',
    's = TypeName(ws.Shapes.[_Default](9))',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1").Value = 4: ws.Range("B1").Value = 5',
             '    ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 30).Name = "Box"',
             '    Set ch = ws.ChartObjects.Add(100, 10, 200, 100).Chart',
             '    ch.SetSourceData ws.Range("A1:B1")',
             '    wb.Names.Add "Total", "=Sheet1!$A$1"',
             '    ws.Range("D1:E3").Value = 1: ws.Range("D1").Value = "x": ws.Range("E1").Value = "y"',
             '    ws.ListObjects.Add xlSrcRange, ws.Range("D1:E3"), , xlYes',
             '    ws.Range("G1:G3").Value = Application.Transpose(Array("a", "b", "c"))',
             '    With ws.Shapes.AddFormControl(xlDropDown, 10, 150, 80, 20): .Name = "Drop": '
             '.ControlFormat.ListFillRange = "$G$1:$G$3": End With',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case[:80], '->', answer)


if __name__ == '__main__':
    main()
