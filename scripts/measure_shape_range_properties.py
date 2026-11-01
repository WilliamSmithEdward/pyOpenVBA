"""ShapeRange.Left, Top, Width, Height, Name, Visible, Rotation, AlternativeText, Title, LockAspectRatio,
HorizontalFlip and VerticalFlip: what a range of one or two shapes reads, what setting each does to its shapes, and
the errors either raises.

Each case makes a workbook with two rectangles, A (10, 10, 60 by 30) and B (100, 40, 80 by 20), takes a ShapeRange
of them as the case says, runs the case's statements and reads both shapes' own values, then saves it.

    python scripts/measure_shape_range_properties.py

writes tests/fixtures/shape_range_properties.json, which tests/test_excel_shape_range_properties.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_range_properties.json'
SHAPES = ('"|" & ws.Shapes("A").Left & "," & ws.Shapes("A").Top & "," & ws.Shapes("A").Width & "," & '
          'ws.Shapes("A").Height & "," & ws.Shapes("A").Rotation & "," & ws.Shapes("A").Visible & ";" & '
          'ws.Shapes(2).Name & "," & ws.Shapes(2).Left & "," & ws.Shapes(2).Top & "," & ws.Shapes(2).Width & "," & '
          'ws.Shapes(2).Height & "," & ws.Shapes(2).Rotation & "," & ws.Shapes(2).Visible')
ONE = 'Set r = ws.Shapes.Range(Array("A"))'
TWO = 'Set r = ws.Shapes.Range(Array("A", "B"))'
CASES = [
    ONE + ': s = s & "|" & r.Left & "," & r.Top & "," & r.Width & "," & r.Height & "," & r.Name & "," & r.Visible',
    TWO + ': s = s & "|" & r.Left & "," & r.Top & "," & r.Width & "," & r.Height',
    TWO + ': s = s & "|" & r.Name',
    TWO + ': s = s & "|" & r.Visible & "," & r.Rotation & "," & r.LockAspectRatio & "," & r.HorizontalFlip',
    TWO + ': s = s & "|" & r.AlternativeText & "," & r.Title',
    ONE + ': r.Left = 50: r.Top = 60',
    TWO + ': r.Left = 50',
    TWO + ': r.Top = 70',
    TWO + ': r.Width = 200',
    TWO + ': r.Height = 100',
    TWO + ': r.Rotation = 30',
    ONE + ': r.Name = "Z"',
    TWO + ': r.Name = "Z"',
    TWO + ': r.AlternativeText = "both": s = s & "|" & ws.Shapes(1).AlternativeText & "," & '
          'ws.Shapes(2).AlternativeText & "," & r.AlternativeText',
    TWO + ': ws.Shapes("A").Rotation = 10: s = s & "|" & r.Rotation',
    TWO + ': r.LockAspectRatio = msoTrue: s = s & "|" & r.LockAspectRatio & "," & ws.Shapes(2).LockAspectRatio',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, r As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 60, 30).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 40, 80, 20).Name = "B"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', f'    Tried = "ok" & s & {SHAPES}', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & s', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case[len(TWO) + 2:][:60], '->', answer)


if __name__ == '__main__':
    main()
