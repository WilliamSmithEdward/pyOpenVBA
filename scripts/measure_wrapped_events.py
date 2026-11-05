"""When Excel works a wrapped row's height out, and what else counts in it; and how many lines a shape's text wraps into. Needs Excel at 96 DPI, where StandardHeight reads 15, and stops if
the display changes.

    python scripts/measure_wrapped_events.py

Writes tests/fixtures/wrapped_events.json. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/wrapped_events.json'
LONG = 'one two three four five six seven'
#: Each is run on a new sheet, with column A 8.43 wide unless it says, and the row's height read after each statement.
EVENTS: dict[str, list[str]] = {
    'value then wrap': [f'c.Value = "{LONG}"', 'c.WrapText = True'],
    'wrap then value': ['c.WrapText = True', f'c.Value = "{LONG}"', 'c.Value = "x"', f'c.Value = "{LONG}"', 'c.ClearContents'],
    'width after': ['c.WrapText = True', f'c.Value = "{LONG}"', 'ws.Columns(1).ColumnWidth = 30', 'ws.Rows(1).AutoFit', 'ws.Columns(1).ColumnWidth = 5', 'ws.Columns(1).AutoFit'],
    'wrap off': ['c.WrapText = True', f'c.Value = "{LONG}"', 'c.WrapText = False'],
    'font after': ['c.WrapText = True', f'c.Value = "{LONG}"', 'c.Font.Size = 16', 'c.Font.Size = 8', 'c.Font.Bold = True', 'c.Font.Italic = True', 'c.Font.Name = "Calibri"'],
    'fixed height': ['ws.Rows(1).RowHeight = 40', 'c.WrapText = True', f'c.Value = "{LONG}"', 'ws.Rows(1).AutoFit'],
    'newlines': ['c.WrapText = True', 'c.Value = "ab" & vbLf & "cd"', 'c.Value = "ab" & vbLf & vbLf & "cd"', 'c.Value = "ab" & vbCr & "cd"', 'c.Value = "ab" & vbLf', 'c.Value = vbLf & "ab"', f'c.Value = "{LONG}" & vbLf & "{LONG}"'],
    'number': ['c.WrapText = True', 'c.Value = 123456789012', 'c.Value = 1.23456789012345', 'c.NumberFormat = "#,##0.00"', 'c.Value = DateSerial(2020, 1, 1)'],
    'formula': ['c.WrapText = True', f'c.Formula = "=""{LONG}"""', 'ws.Range("B1").Value = "x"', 'c.Formula = "=B1"', 'ws.Range("B1").Value = "one two three four five six"', 'ws.Rows(1).AutoFit'],
    'two cells': ['c.WrapText = True', 'c.Value = "one two three"', 'd.WrapText = True', f'd.Value = "{LONG}"', 'd.ClearContents', 'ws.Rows(1).AutoFit'],
    'beside big': ['c.WrapText = True', 'c.Value = "one two three"', 'd.Font.Size = 30', 'd.Value = "x"', 'd.ClearContents', 'ws.Rows(1).AutoFit'],
    'merged': ['ws.Range("A1:B1").Merge', 'c.WrapText = True', f'c.Value = "{LONG}"', 'ws.Rows(1).AutoFit'],
    'indent': ['c.WrapText = True', f'c.Value = "{LONG}"', 'c.IndentLevel = 2', 'ws.Rows(1).AutoFit', 'c.IndentLevel = 0', 'c.HorizontalAlignment = xlRight', 'ws.Rows(1).AutoFit', 'c.HorizontalAlignment = xlCenter', 'ws.Rows(1).AutoFit'],
    'range set': ['ws.Range("A1:A3").WrapText = True', f'ws.Range("A1:A3").Value = "{LONG}"'],
    'justify': [f'c.Value = "{LONG}"', 'c.HorizontalAlignment = xlJustify', 'c.HorizontalAlignment = xlGeneral', 'c.VerticalAlignment = xlDistributed', 'c.VerticalAlignment = xlJustify'],
    'cap': ['c.WrapText = True', f'c.Value = String(3000, "a")', 'ws.Rows(1).AutoFit'],
}
_READ = '    r = r & "|" & ws.Rows(1).RowHeight & "/" & ws.Rows(2).RowHeight'
SHAPE_TEXTS = ['x', 'ab', 'abcd', 'abcdefgh', 'one two', 'ab cd', 'a-b', 'one two three four five six seven', 'abcdefghijklmnopqrstuvwxyz', 'The quick brown fox jumps over the lazy dog']
SHAPE_WIDTHS = list(range(10, 90, 2)) + list(range(90, 320, 10))


def event_source(name: str, statements: list[str]) -> str:
    lines = ['Public Function Go() As String', '    Dim ws As Object, c As Object, d As Object, r As String', '    Set ws = ActiveSheet', '    ws.Cells.Clear', '    ws.Rows("1:3").RowHeight = 15', '    ws.Rows("1:3").UseStandardHeight = True', '    ws.Columns("A:B").ColumnWidth = 8.43', '    Set c = ws.Range("A1")', '    Set d = ws.Range("B1")', _READ]
    for statement in statements:
        lines += [f'    {statement}', _READ]
    lines += ['    Go = ActiveSheet.StandardHeight & r', 'End Function']
    return '\n'.join(lines) + '\n'


SHAPES = r'''Public Function Go(ByVal text As String, ByVal widths As String) As String
    Dim ws As Object, s As Object, r As String, one As Variant
    Set ws = ActiveSheet
    Set s = ws.Shapes.AddShape(1, 10, 10, 100, 300)
    s.TextFrame2.TextRange.Text = text
    For Each one In Split(widths, ",")
        s.Width = Val(one)
        r = r & "|" & s.TextFrame2.TextRange.Lines.Count
    Next
    s.Delete
    Go = ActiveSheet.StandardHeight & r
End Function
'''


def main() -> None:
    events: dict[str, list[str]] = {}
    shapes: dict[str, list[str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, statements in EVENTS.items():
            result = excel.run_vba(event_source(name, statements), 'Go', (), timeout=60.0, line_numbers=False)
            assert result.ok and not result.dialogs, (name, result)
            height, *found = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            events[name] = found
        widths = ','.join(str(one) for one in SHAPE_WIDTHS)
        for text in SHAPE_TEXTS:
            result = excel.run_vba(SHAPES, 'Go', (text, widths), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, (text, result)
            height, *found = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            shapes[text] = found
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'events': EVENTS, 'rows': events, 'shape_widths': SHAPE_WIDTHS, 'shapes': shapes}, indent=0) + '\n', encoding='utf-8')
    for name, statements in EVENTS.items():
        print(name, events[name])
    print({text: found[:18] for text, found in list(shapes.items())[:3]})


if __name__ == '__main__':
    main()
