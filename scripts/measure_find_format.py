"""Application.FindFormat and ReplaceFormat: what a CellFormat reads before and after a set, which cells Range.Find
with SearchFormat finds, FindNext and FindPrevious after it, and what Range.Replace with SearchFormat and
ReplaceFormat changes. The formats last for Excel's session, so the probe clears both after every case.

    python scripts/measure_find_format.py

writes tests/fixtures/find_format.json, which tests/test_excel_find_format.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/find_format.json'
F = 'Application.FindFormat'
R = 'Application.ReplaceFormat'
FOUND = 'Got(ws.Range("A1:C4").Find("", ws.Range("C4"), xlFormulas, xlPart, xlByRows, xlNext, False, , True))'
#: Two sheets of formats over the same values in A1:C4.
SETUPS = {
    'a': ['ws.Range("B2").Font.Bold = True', 'ws.Range("C3").Font.Bold = True',
          'ws.Range("A2").Interior.Color = vbYellow', 'ws.Range("C1").Interior.Color = vbYellow',
          'ws.Range("A3").NumberFormat = "0.00"', 'ws.Range("B4").Font.Italic = True',
          'ws.Range("C2").Font.Color = vbRed', 'ws.Range("A4").Locked = False',
          'ws.Range("C3").Interior.Color = vbYellow'],
    'b': ['ws.Range("B2").Font.Bold = True', 'ws.Range("C3").Font.Bold = True', 'ws.Range("A2").Font.Name = "Arial"',
          'ws.Range("B3").Font.Size = 14', 'ws.Range("C2").Font.Underline = xlUnderlineStyleSingle',
          'ws.Range("A3").HorizontalAlignment = xlCenter', 'ws.Range("B1").WrapText = True',
          'ws.Range("C1").IndentLevel = 2', 'ws.Range("A1").FormulaHidden = True',
          'ws.Range("B2").Interior.Color = RGB(250, 250, 5)', 'ws.Range("C4").Interior.Color = vbYellow',
          'ws.Range("E1").Value = 1', 'ws.Range("F2").Font.Bold = True', 'ws.Range("G1:G3").Value = 1'],
}
#: Each case: its setup, statements after it, and what is read back.
CASES = [
    ('a', [], f'TypeName({F}) & " " & TypeName({R}) & " " & ({F} Is {F})'),
    ('a', [], f'Show({F}.Font.Bold) & " " & Show({F}.Interior.Color) & " " & Show({F}.NumberFormat) & " " & '
              f'Show({F}.Font.Name) & " " & Show({F}.Font.Size) & " " & Show({F}.Locked) & " " & '
              f'Show({F}.Interior.ColorIndex)'),
    ('a', [], f'Show({F}.HorizontalAlignment) & " " & Show({F}.WrapText) & " " & Show({F}.IndentLevel) & " " & '
              f'Show({F}.Font.Color) & " " & Show({F}.Font.Italic) & " " & Show({F}.Interior.Pattern) & " " & '
              f'Show({F}.FormulaHidden)'),
    ('a', [], f'Show({F}.Font.ColorIndex) & " " & Show({F}.Font.Strikethrough) & " " & Show({F}.VerticalAlignment) '
              f'& " " & Show({F}.Orientation) & " " & Show({F}.ShrinkToFit) & " " & Show({F}.MergeCells) & " " & '
              f'Show({F}.AddIndent)'),
    ('a', [f'{F}.Font.Bold = True'], f'Show({F}.Font.Bold) & " " & Show({F}.Font.Italic) & " " & ' + FOUND),
    ('a', [f'{F}.Font.Bold = False'], FOUND),
    ('a', [f'{F}.Font.Italic = True'], f'Show({F}.Font.Bold) & " " & Show({F}.Font.Italic)'),
    ('a', [f'{F}.Font.Bold = False'], f'Show({F}.Font.Bold) & " " & Show({F}.Font.Italic)'),
    ('a', [f'{F}.Font.Bold = True', f'{F}.Font.Italic = True'], f'Show({F}.Font.Bold) & " " & Show({F}.Font.Italic)'),
    ('a', [f'{F}.Font.Size = 11'], f'Show({F}.Font.Bold) & " " & Show({F}.Font.Italic) & " " & Show({F}.Font.Name)'),
    ('a', ['ws.Range("C3").Font.Italic = True', f'{F}.Font.Bold = True'],
     'Got(ws.Range("A1:C4").Find("", ws.Range("B2"), xlFormulas, xlPart, xlByRows, xlNext, False, , True))'),
    ('a', [f'{F}.Interior.Color = vbYellow'],
     f'Show({F}.Interior.Color) & " " & Show({F}.Interior.Pattern) & " " & Show({F}.Interior.ColorIndex) & " " & '
     + FOUND),
    ('a', [f'{F}.NumberFormat = "0.00"'], FOUND),
    ('a', [f'{F}.Font.Bold = True', f'{F}.Clear'], f'Show({F}.Font.Bold) & " " & ' + FOUND),
    ('a', [], FOUND),
    ('a', [f'{F}.Font.Bold = True'],
     'Got(ws.Range("A1:C4").Find("x", , xlFormulas, xlPart, xlByRows, xlNext, False, , True))'),
    ('a', [f'{F}.Font.Bold = True', 'Set c = ws.Range("A1:C4").Find("", ws.Range("C4"), , , , , , , True)'],
     'Got(c) & " " & Got(ws.Range("A1:C4").FindNext(c))'),
    ('a', [f'{F}.Font.Italic = True'], FOUND),
    ('a', [f'{F}.Font.Color = vbRed'], FOUND),
    ('a', [f'{F}.Locked = False'], FOUND),
    ('a', [f'{F}.Interior.ColorIndex = 6'], FOUND),
    ('a', [f'{F}.Font.Bold = True', f'{F}.Interior.Color = vbYellow'], FOUND),
    ('a', [f'{R}.Interior.Color = vbRed', 'ws.Range("A1:C4").Replace "x", "y", xlPart, xlByRows, False, , False, True'],
     'Grid(ws) & " " & Fills(ws)'),
    ('a', [f'{F}.Font.Bold = True', f'{R}.Font.Italic = True',
           'ws.Range("A1:C4").Replace "", "", xlPart, xlByRows, False, , True, True'], 'Grid(ws) & " " & Italics(ws)'),
    ('a', [f'{F}.Font.Bold = True', 'ws.Range("A1:C4").Replace "x", "z", xlPart, xlByRows, False, , True, False'],
     'Grid(ws)'),
    ('a', [f'{R}.Font.Bold = False', 'ws.Range("A1:C4").Replace "x", "y", xlPart, xlByRows, False, , False, True'],
     'Grid(ws) & " " & Bolds(ws)'),
    ('a', [f'{R}.NumberFormat = "0%"', 'ws.Range("A1:C4").Replace "7", "8", xlPart, xlByRows, False, , False, True'],
     'Grid(ws) & " " & ws.Range("B3").NumberFormat'),
    ('b', [f'{F}.Font.Name = "Arial"'], f'Show({F}.Font.Name) & " " & ' + FOUND),
    ('b', [f'{F}.Font.Size = 14'], f'Show({F}.Font.Size) & " " & ' + FOUND),
    ('b', [f'{F}.Font.Underline = xlUnderlineStyleSingle'], f'Show({F}.Font.Underline) & " " & ' + FOUND),
    ('b', [f'{F}.HorizontalAlignment = xlCenter'], f'Show({F}.HorizontalAlignment) & " " & ' + FOUND),
    ('b', [f'{F}.WrapText = True'], f'Show({F}.WrapText) & " " & ' + FOUND),
    ('b', [f'{F}.IndentLevel = 2'], f'Show({F}.IndentLevel) & " " & ' + FOUND),
    ('b', [f'{F}.FormulaHidden = True'], f'Show({F}.FormulaHidden) & " " & ' + FOUND),
    ('b', [f'{F}.Interior.ColorIndex = 6'], f'Show({F}.Interior.Color) & " " & Show({F}.Interior.Pattern)'),
    ('b', [f'{F}.Interior.Color = RGB(250, 250, 5)'], f'Show({F}.Interior.ColorIndex) & " " & ' + FOUND),
    ('b', [f'{F}.Interior.Pattern = xlSolid'], f'Show({F}.Interior.Pattern) & " " & ' + FOUND),
    ('b', [f'{F}.Font.Color = vbRed'], f'Show({F}.Font.Color) & " " & Show({F}.Font.ColorIndex)'),
    ('b', [f'{F}.Font.Bold = True'], 'Got(ws.Range("E1:F2").Find("", , , , , , , , True))'),
    ('b', [f'{F}.Font.Bold = True', 'Set c = ws.Range("A1:C4").Find("", ws.Range("C4"), , , , , , , True)'],
     'Got(ws.Range("A1:C4").Find(""))'),
    ('b', [f'{F}.Font.Bold = True', 'Set c = ws.Range("A1:C4").Find("", ws.Range("C4"), , , , , , , True)'],
     'Got(ws.Range("A1:C4").FindPrevious(c))'),
    ('b', [f'{F}.Font.Bold = True'], 'Got(ws.Range("A1:C4").Find("x", ws.Range("C4"), , , , , , , True))'),
    ('b', [f'{F}.Font.Bold = True'], 'Got(ws.Range("A1:C4").Find("x3", ws.Range("C4"), , xlWhole, , , , , True))'),
    ('b', [f'{F}.Font.Bold = True'], 'Got(ws.Range("A1:C4").Find("", ws.Range("C4"), , , , xlPrevious, , , True))'),
    ('b', [], 'Got(ws.Range("A1:C4").Find("", ws.Range("C4"), , , , , , , True))'),
    ('b', [f'{F}.Font.Bold = True'], 'Got(ws.Range("G1:G3").Find("", , , , , , , , True))'),
    ('b', [f'{R}.Font.Bold = True', f'{R}.Interior.Color = vbGreen',
           'ws.Range("A1:C4").Replace "q", "r", xlWhole, , , , , True'],
     'ws.Range("B4").Formula & " " & ws.Range("B4").Font.Bold & " " & ws.Range("B4").Interior.Color & " " & '
     'ws.Range("B4").Interior.Pattern'),
    ('b', [f'{R}.Font.Bold = True', 'ws.Range("A1:C4").Replace "zz", "r", xlWhole, , , , , True'], 'Bolds(ws)'),
    ('b', [f'{R}.Font.Bold = True', 'ws.Range("A1:C4").Replace "", "", xlPart, , , , , True'], 'Bolds(ws)'),
    ('b', [f'{R}.Font.Bold = True', 'ws.Range("A1:C4").Replace "x", "x", xlPart, , , , , True'], 'Bolds(ws)'),
    ('b', [f'{F}.Font.Bold = True', 'ws.Range("A1:C4").Replace "", "Q", xlPart, , , , True'], 'Grid(ws)'),
    ('b', [f'{R}.Font.Bold = True', 'v = ws.Range("A1:C4").Replace("x", "y", xlPart, , , , , True)'], 'v'),
    ('b', [f'{F}.Font.Bold = True', f'{R}.Interior.ColorIndex = 4',
           'ws.Range("A1:C4").Replace "", "", xlPart, , , , True, True'],
     'ws.Range("B2").Interior.Color & " " & ws.Range("C3").Interior.Color & " " & ws.Range("A1").Interior.Color'),
    ('b', [f'{R}.Font.Bold = True', 'ws.Range("A1:C4").Replace "x", "y", xlPart, , , , False, False'], 'Bolds(ws)'),
]


def _listing(name: str, expression: str) -> list[str]:
    return [f'Private Function {name}(ByVal ws As Worksheet) As String', '    Dim c As Range, s As String',
            '    For Each c In ws.Range("A1:C4")', f'        s = s & "|" & {expression}', '    Next',
            f'    {name} = Mid(s, 2)', 'End Function', '']


def _probe() -> str:
    lines = ['Private Function Show(ByVal v As Variant) As String', '    If IsNull(v) Then', '        Show = "Null"',
             '    Else', '        Show = TypeName(v) & ":" & CStr(v)', '    End If', 'End Function', '',
             'Private Function Got(ByVal c As Range) As String', '    If c Is Nothing Then', '        Got = "Nothing"',
             '    Else', '        Got = c.Address(False, False)', '    End If', 'End Function', '',
             *_listing('Grid', 'c.Formula'), *_listing('Fills', 'c.Interior.Color'),
             *_listing('Italics', 'c.Font.Italic'), *_listing('Bolds', 'c.Font.Bold'),
             'Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, c As Range, v As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:C4").Value = Application.Transpose(Application.Transpose(Array('
             'Array("x1", "a", 3), Array("b", "x2", 5), Array(1.5, 7, "x3"), Array("", "q", "x4"))))',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (setup, statements, read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in SETUPS[setup] + statements]
        lines.append(f'        Tried = {read}')
    lines += ['    End Select', 'Done:', f'    {F}.Clear', f'    {R}.Clear', '    wb.Close False', '    Exit Function',
              'Failed:', '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[setup, statements, read, answer]
                        for (setup, statements, read), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for _, statements, read, answer in record['cases']:
        print(statements[-1:] or [read[:60]], '->', answer)


if __name__ == '__main__':
    main()
