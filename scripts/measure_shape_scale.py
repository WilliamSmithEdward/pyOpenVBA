"""Shape.ScaleWidth and ScaleHeight, ConnectionSiteCount and the style reads of every kind of shape the model makes,
and what a ShapeRange answers for the shapes in it.

KINDS adds one shape of each kind and reads READS of it. SCALES adds a rectangle at 100, 50, 100 by 40 (or the
case's own shapes), runs the statements and reads the geometry of sh, and of sh2 where the case adds it. RANGES reads
ShapeRange members over one shape or two. Each answer is its value, TypeName for an object, or "E" and the error
number.

    python scripts/measure_shape_scale.py

writes tests/fixtures/shape_scale.json, which tests/test_excel_shape_scale.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_scale.json'
#: Each kind of shape the model makes: the 21 AutoShapes, a text box, a line, a chart, a button and a group.
KINDS = [f'ws.Shapes.AddShape({kind}, 10, 10, 80, 40)' for kind in [*range(1, 21), 92]] + [
    'ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 80, 40)', 'ws.Shapes.AddLine(10, 10, 90, 50)',
    'ws.ChartObjects.Add(10, 10, 200, 100).ShapeRange(1)',
    'ws.Shapes.AddFormControl(xlButtonControl, 10, 10, 80, 20)',
    'ws.Shapes.Range(Array(ws.Shapes.AddShape(1, 10, 10, 40, 40).Name, ws.Shapes.AddShape(1, 60, 10, 40, 40).Name))'
    '.Group',
]
READS = ['sh.ConnectionSiteCount', 'sh.ShapeStyle', 'sh.BackgroundStyle', 'sh.GraphicStyle']
SCALES = [
    'sh.ScaleWidth 1.5, msoFalse',
    'sh.ScaleWidth 1.5, msoFalse, msoScaleFromTopLeft',
    'sh.ScaleWidth 1.5, msoFalse, msoScaleFromMiddle',
    'sh.ScaleWidth 1.5, msoFalse, msoScaleFromBottomRight',
    'sh.ScaleHeight 0.5, msoFalse',
    'sh.ScaleHeight 0.5, msoFalse, msoScaleFromMiddle',
    'sh.ScaleHeight 2, msoFalse, msoScaleFromBottomRight',
    'sh.ScaleWidth 1.3333, msoFalse',
    'sh.ScaleWidth 1 / 3, msoFalse, msoScaleFromMiddle',
    'sh.ScaleWidth 1.5, msoTrue',
    'sh.ScaleWidth 0, msoFalse',
    'sh.ScaleWidth -1, msoFalse',
    'sh.ScaleWidth 5, msoFalse, msoScaleFromBottomRight',
    'sh.ScaleHeight 5, msoFalse, msoScaleFromBottomRight',
    'sh.ScaleWidth 2, msoFalse, 7',
    'sh.LockAspectRatio = msoTrue: sh.ScaleWidth 2, msoFalse',
    'sh.LockAspectRatio = msoTrue: sh.ScaleHeight 2, msoFalse, msoScaleFromMiddle',
    'sh.Rotation = 90: sh.ScaleWidth 2, msoFalse',
    'sh.Rotation = 90: sh.ScaleWidth 2, msoFalse, msoScaleFromMiddle',
    'sh.ScaleWidth 2, msoFalse: sh.ScaleWidth 2, msoTrue',
    'sh.ScaleWidth 1000, msoFalse',
    'Set sh2 = ws.Shapes.AddShape(1, 300, 100, 60, 30): '
    'ws.Shapes.Range(Array(sh.Name, sh2.Name)).ScaleWidth 2, msoFalse, msoScaleFromMiddle',
    'Set sh2 = ws.Shapes.AddShape(1, 300, 100, 60, 30): '
    'ws.Shapes.Range(Array(sh.Name, sh2.Name)).ScaleHeight 0.5, msoFalse, msoScaleFromBottomRight',
    'sh.Delete: Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 50, 100, 40): '
    'sh.ScaleWidth 1.5, msoFalse, msoScaleFromMiddle',
    'sh.Delete: Set sh = ws.Shapes.AddLine(100, 50, 200, 90): sh.ScaleWidth 1.5, msoFalse, msoScaleFromMiddle',
    'sh.Delete: Set sh = ws.ChartObjects.Add(100, 50, 100, 40).ShapeRange(1): sh.ScaleWidth 1.5, msoFalse',
]
GEOMETRY = 'sh.Left & "/" & sh.Top & "/" & sh.Width & "/" & sh.Height'
RANGES = [
    ('one', 'TypeName(sr.TextFrame)'), ('one', 'TypeName(sr.TextFrame2)'), ('one', 'sr.BlackWhiteMode'),
    ('one', 'sr.ConnectionSiteCount'), ('one', 'TypeName(sr.Chart)'), ('one', 'sr.ShapeStyle'),
    ('chart', 'TypeName(sr.Chart)'), ('chart', 'sr.Chart.Parent.Name'),
    ('two', 'TypeName(sr.TextFrame)'), ('two', 'TypeName(sr.TextFrame2)'), ('two', 'sr.BlackWhiteMode'),
    ('two', 'sr.ConnectionSiteCount'), ('two', 'TypeName(sr.Chart)'),
]


def _probe() -> str:
    lines = ['Private Function Kind(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which']
    for index, made in enumerate(KINDS):
        lines += [f'    Case {index}', f'        Set sh = {made}']
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    s = s & "," & {read}', '    If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Kind = Mid(s, 2)', 'End Function', '',
              'Private Function Scaled(ByVal which As Long) As String',
              '    Dim wb As Workbook, ws As Worksheet, sh As Object, sh2 As Object, s As String',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
              '    Set sh = ws.Shapes.AddShape(1, 100, 50, 100, 40)', '    On Error GoTo Failed',
              '    Select Case which']
    for index, statements in enumerate(SCALES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', f'    s = "ok " & {GEOMETRY}',
              '    If Not sh2 Is Nothing Then s = s & "|" & sh2.Left & "/" & sh2.Top & "/" & sh2.Width & "/" & '
              'sh2.Height',
              'Done:', '    wb.Close False', '    Scaled = s', '    Exit Function',
              'Failed:', f'    s = "Err " & Err.Number & " " & {GEOMETRY}', '    Resume Done', 'End Function', '',
              'Private Function Ranged(ByVal which As Long) As String',
              '    Dim wb As Workbook, ws As Worksheet, sr As Object, a As Object, b As Object, s As String',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
              '    Set a = ws.Shapes.AddShape(1, 10, 10, 80, 40): Set b = ws.Shapes.AddShape(9, 100, 10, 80, 40)',
              '    ws.ChartObjects.Add 10, 100, 200, 100', '    Select Case which']
    for index, (shapes, _read) in enumerate(RANGES):
        chosen = {'one': 'Array(a.Name)', 'two': 'Array(a.Name, b.Name)', 'chart': 'Array(3)'}[shapes]
        lines += [f'    Case {index}', f'        Set sr = ws.Shapes.Range({chosen})']
    lines += ['    End Select', '    On Error Resume Next', '    Select Case which']
    for index, (_shapes, read) in enumerate(RANGES):
        lines += [f'    Case {index}', f'        s = {read}']
    lines += ['    End Select', '    If Err.Number Then s = "E" & Err.Number', '    On Error GoTo 0',
              '    wb.Close False', '    Ranged = s', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(KINDS) - 1}', '        s = s & "@@@" & Kind(i)', '    Next',
              f'    For i = 0 To {len(SCALES) - 1}', '        s = s & "@@@" & Scaled(i)', '    Next',
              f'    For i = 0 To {len(RANGES) - 1}', '        s = s & "@@@" & Ranged(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    names = KINDS + SCALES + [f'{shapes}: {read}' for shapes, read in RANGES]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[name, answer] for name, answer in zip(names, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer in record['cases']:
        print(name, '->', answer)


if __name__ == '__main__':
    main()
