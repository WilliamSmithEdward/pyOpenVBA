"""What a huge number, an unallocated array or an awkward count answers on members that once let a Python error out or never returned, read
line by line from live Excel, bound late and typed. No file is saved and nothing shows.

    python scripts/measure_excel_huge_numbers.py

writes tests/fixtures/excel_huge_numbers.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_huge_numbers.json'
SEPARATOR = '|~|'
VALUES = ['410', '50000', '100000', '169093', '1000000', '2147483647#', '2147483648#', '1E+10', '1E+38', '1E+39', '1E+100', '1E+308', '-1E+308']
LINES: list[tuple[str, str]] = []
for value in VALUES:
    LINES += [
        (f'ws.Rows(3).RowHeight = {value}', 'ws.Rows(3).RowHeight'), (f'c.Top = {value}', 'c.Top'), (f'c.Left = {value}', 'c.Left'), (f'c.Width = {value}', 'c.Width'),
        (f'c.Height = {value}', 'c.Height'), (f'ws.Tab.TintAndShade = {value}', 'ws.Tab.TintAndShade'), (f's.Top = {value}', 's.Top'), (f's.Shadow.Blur = {value}', 's.Shadow.Blur'),
        (f's.Shadow.OffsetX = {value}', 's.Shadow.OffsetX'), (f's.Shadow.OffsetY = {value}', 's.Shadow.OffsetY'), (f's.Reflection.Offset = {value}', 's.Reflection.Offset'),
        (f's.Reflection.Blur = {value}', 's.Reflection.Blur'), (f's.TextFrame2.TextRange.Font.Spacing = {value}', 's.TextFrame2.TextRange.Font.Spacing'),
        (f's.TextFrame2.TextRange.Font.Kerning = {value}', 's.TextFrame2.TextRange.Font.Kerning'), (f'Set z = Nothing: Set z = ws.Shapes.AddShape(1, {value}, 1, 5, 5)', 'TypeName(z)'),
        (f'Set z = Nothing: Set z = ws.Shapes.AddShape(1, 1, {value}, 5, 5)', 'TypeName(z)'), (f'Set z = Nothing: Set z = ws.Shapes.AddShape(1, 1, 1, {value}, 5)', 'TypeName(z)'),
        (f'Set z = Nothing: Set z = ws.Shapes.AddShape(1, 1, 1, 5, {value})', 'TypeName(z)'), (f'Set d = Nothing: Set d = ws.ChartObjects.Add(1, 1, {value}, 50)', 'TypeName(d)'),
        (f'Set d = Nothing: Set d = ws.ChartObjects.Add(1, 1, 50, {value})', 'TypeName(d)'),
    ]
    if not value.startswith('-'):
        # Where a negative place puts a chart is not known, and the model says so.
        LINES += [(f'Set d = Nothing: Set d = ws.ChartObjects.Add({value}, 1, 50, 50)', 'TypeName(d)'), (f'Set d = Nothing: Set d = ws.ChartObjects.Add(1, {value}, 50, 50)', 'TypeName(d)')]
LINES += [(f'ws.Parent.Sheets.Add , , {count}', 'ws.Parent.Sheets.Count') for count in ['1E+308', '1099511627776#', '2147483648#', '-1', '0', '1.5', '"x"', '""', '"2"', 'True', 'Null', 'Empty', '0.5', '0.9', '2.9', '256', '300', '255', '255.9']]
LINES += [(f'ws.Parent.Worksheets.Add Count:={count}', 'ws.Parent.Sheets.Count') for count in ['1E+308', '2147483648#', '256', '2.9']]
LINES += [(f'{setter} = a', 'TypeName(ws.Range("Z1").Value)') for setter in ['ws.Range("Z1").Value', 'ws.Range("Z1").Value2', 'ws.Range("Z1:AA2").Value', 'ws.Range("Z1")', 'ws.Range("Z1:AA2").Formula']]
LINES += [(f'x = Empty: x = {function}', 'TypeName(x)') for function in [
    'WorksheetFunction.MUnit(1E+308)', 'WorksheetFunction.MUnit(1099511627776#)', 'WorksheetFunction.MUnit(2147483648#)', 'WorksheetFunction.MUnit(1000000)', 'WorksheetFunction.MUnit(100000)',
    'WorksheetFunction.MUnit(-1)', 'WorksheetFunction.MUnit(0)', 'WorksheetFunction.MUnit(1.5)', 'WorksheetFunction.MUnit(100)', 'UBound(WorksheetFunction.MUnit(3000), 1)', 'UBound(WorksheetFunction.MUnit(5000), 1)',
    'Application.MUnit(1E+308)', 'Application.MUnit(1099511627776#)', 'Application.MUnit(2147483648#)', 'Application.MUnit(1000000)', 'UBound(Application.MUnit(100), 1)',
]]
TEMPLATE = '''Private Function Shown(ByVal x As Variant) As String
    On Error Resume Next
    Shown = TypeName(x) & " " & CStr(x)
End Function
Function Probe() As String
    Dim ws As {W}, c As {C}, s As {S}, z As {S}, d As {C}, r As String, e As Long, x As Variant, a() As Long
    Set ws = Worksheets.Add
    Set c = ws.ChartObjects.Add(10, 100, 200, 100)
    Set s = ws.Shapes.AddShape(1, 10, 10, 50, 50)
    s.Shadow.Visible = True
    Application.DisplayAlerts = False
    On Error Resume Next
{BODY}
    Application.DisplayAlerts = True
    Probe = r
End Function
'''


def macro(*, typed: bool) -> str:
    body = ''.join(f'    {one.replace(": ", chr(10) + "    ")}\n    e = Err.Number: Err.Clear\n    r = r & e & "=" & Shown({two}) & "{SEPARATOR}"\n' for one, two in LINES)
    return (TEMPLATE.replace('{W}', 'Worksheet' if typed else 'Object').replace('{C}', 'ChartObject' if typed else 'Object').replace('{S}', 'Shape' if typed else 'Object').replace('{BODY}', body))


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, str] = {}
    for binding in ('late', 'typed'):
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(macro(typed=binding == 'typed'), 'Probe', timeout=240.0, line_numbers=False)
        if not result.ok or result.dialogs:
            raise SystemExit(f'{binding} stopped: {result.message}')
        answers[binding] = str(result.value)
        print(binding, 'done', flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
