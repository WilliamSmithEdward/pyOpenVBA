"""Range.CreateNames: the names Excel makes from a range's labels, and how it cleans a label into a name.

Three parts, each case on sheet Data laid out again with every name cleared
first, alerts off as a macro runs it unattended:

- cleaning: one column of labels beside numbers, named from the left
  column alone, so each name shows what one label became;
- sides: a grid with labels on all four sides, named with every side given
  True or False;
- guessing: small layouts asked with sides left out, which Excel guesses.

Each case records what CreateNames returned and every name with what it
refers to, or the error raised.

    python scripts/measure_create_names.py

writes tests/fixtures/create_names.json, which tests/test_excel_create_names.py
replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/create_names.json'

#: Labels in A1 down, each beside a number in B.
LABELS = ['North', 'South East', 'R1C1', 'A1', '123', 'Total %', "O'Neil", 'TRUE', '', 'x' * 260, '_ok', '\\back',
          'r', 'C', 'XFD1', 'XFE1', '  lead', 'end  ', 'a%b', 'a  b', '%x', 'x%', '1.5', 'a?b', 'a$b', 'a.b-c',
          'Feb', 'feb', '3rd', '.dot', 'a,b', '"q"', 'a_', 'RC', 'R', 'rc1', 'C1R1', 'tab\tx', 'k&l', '#N/A',
          'a(b)', 'ab12', 'AB12', 'name ', '=x']
#: The grid for the sides: labels on the top row and left column, and on the bottom row and right column.
GRID = {'B1': 'TopA', 'C1': 'TopB', 'D1': 'TopC', 'A2': 'LeftA', 'A3': 'LeftB', 'A4': 'LeftC', 'B5': 'BottomA',
        'C5': 'BottomB', 'D5': 'BottomC', 'E2': 'RightA', 'E3': 'RightB', 'E4': 'RightC', 'A1': 'CornerTL',
        'E1': 'CornerTR', 'A5': 'CornerBL', 'E5': 'CornerBR'}
SIDES = [(top, left, bottom, right) for top in (False, True) for left in (False, True)
         for bottom in (False, True) for right in (False, True)]
#: Layouts for guessing, each cell to value: text, or a number written as one.
GUESSES = {
    'labels both ways': {'A1': '', 'B1': 'Jan', 'C1': 'Feb', 'A2': 'North', 'A3': 'South', 'B2': '1', 'C2': '2',
                         'B3': '3', 'C3': '4'},
    'labels on top': {'A1': 'Jan', 'B1': 'Feb', 'A2': '1', 'B2': '2', 'A3': '3', 'B3': '4'},
    'labels on the left': {'A1': 'North', 'A2': 'South', 'B1': '1', 'B2': '2', 'C1': '3', 'C2': '4'},
    'text everywhere': {'A1': 'a', 'B1': 'b', 'A2': 'c', 'B2': 'd'},
    'numbers everywhere': {'A1': '1', 'B1': '2', 'A2': '3', 'B2': '4'},
    'corner label': {'A1': 'Corner', 'B1': 'Jan', 'A2': 'North', 'B2': '1'},
    'top row text left numbers': {'A1': 'Jan', 'B1': 'Feb', 'A2': '5', 'B2': '6', 'A3': '7', 'B3': '8'},
}
#: Each guessing case: the layout, the range, and the arguments given.
GUESS_CASES = [(layout, 'A1:C3', '') for layout in GUESSES] + [
    ('labels on top', 'A1:C3', 'Top:=True'), ('labels on top', 'A1:C3', 'Left:=True'),
    ('labels on top', 'A1:C3', 'Top:=False'), ('labels both ways', 'A1:C3', 'Bottom:=True'),
    ('labels both ways', 'A1:C3', 'Top:=True, Left:=False'), ('labels on the left', 'A1:C3', 'Right:=False'),
    ('labels both ways', 'A1:C3', 'Top:="x"'), ('labels both ways', 'A1:C3', 'Top:=1'),
    ('labels both ways', 'A1:C3', 'Top:=Null'), ('labels both ways', 'B1', 'Top:=True'),
    ('labels both ways', 'A1:A3', 'Top:=True'), ('labels both ways', 'A1:C3', 'Top:=xlNo'),
    ('labels both ways', 'A1:C1,A3:C3', 'Top:=True'),
]


def _literal(text: str) -> str:
    return '"' + text.replace('"', '""').replace('\t', '" & vbTab & "') + '"'


def _put(cells: dict[str, str]) -> list[str]:
    out = ['    ws.Cells.Clear']
    for cell, value in cells.items():
        if not value:
            continue
        written = value if value.replace('.', '', 1).isdigit() else _literal(value)
        out.append(f'    ws.Range("{cell}").Value = {written}')
    return out


def _probe() -> str:
    lines = ['Private Function Named(ByVal wb As Workbook) As String', '    Dim n As Name, s As String',
             '    For Each n In wb.Names', '        s = s & "|" & n.Name & "=" & n.RefersTo', '    Next',
             '    Named = s', 'End Function', '', 'Private Sub Clear(ByVal ws As Worksheet)',
             '    Do While ws.Parent.Names.Count > 0', '        ws.Parent.Names(1).Delete', '    Loop', 'End Sub', '',
             'Public Function Cleaning() As String', '    Dim ws As Worksheet, v As Variant',
             '    Set ws = ActiveSheet', '    Clear ws', '    ws.Cells.Clear']
    for row, label in enumerate(LABELS, start=1):
        if label:
            lines.append(f'    ws.Cells({row}, 1).Value = {_literal(label)}')
        lines.append(f'    ws.Cells({row}, 2).Value = {row}')
    lines += [f'    v = ws.Range("A1:B{len(LABELS)}").CreateNames(Top:=False, Left:=True, Bottom:=False, Right:=False)',
              '    Cleaning = TypeName(v) & " " & CStr(v) & Named(ws.Parent)', 'End Function', '',
              'Public Function Sides(ByVal t As Boolean, ByVal l As Boolean, ByVal b As Boolean, '
              'ByVal r As Boolean) As String', '    Dim ws As Worksheet, v As Variant', '    Set ws = ActiveSheet',
              '    Clear ws', *_put(GRID), '    ws.Range("B2:D4").Value = 7',
              '    On Error GoTo Failed',
              '    v = ws.Range("A1:E5").CreateNames(Top:=t, Left:=l, Bottom:=b, Right:=r)',
              '    Sides = TypeName(v) & " " & CStr(v) & Named(ws.Parent)', '    Exit Function', 'Failed:',
              '    Sides = "Err " & Err.Number & " " & Split(Err.Description, vbLf)(0) & Named(ws.Parent)',
              'End Function', '']
    for index, (layout, target, arguments) in enumerate(GUESS_CASES):
        lines += [f'Public Function Guess{index}() As String', '    Dim ws As Worksheet, v As Variant',
                  '    Set ws = ActiveSheet', '    Clear ws', *_put(GUESSES[layout]), '    On Error GoTo Failed',
                  f'    v = ws.Range("{target}").CreateNames({arguments})',
                  f'    Guess{index} = TypeName(v) & " " & CStr(v) & Named(ws.Parent)', '    Exit Function', 'Failed:',
                  f'    Guess{index} = "Err " & Err.Number & " " & Split(Err.Description, vbLf)(0) & Named(ws.Parent)', 'End Function', '']
    lines += ['Public Function Kinds() As String', '    Dim ws As Worksheet, v As Variant', '    Set ws = ActiveSheet',
              '    Clear ws', '    ws.Cells.Clear', *KINDS, '    ws.Range("B1:B12").Value = 1',
              '    ActiveWorkbook.Names.Add Name:="Kept", RefersTo:="=99"',
              '    v = ws.Range("A1:B12").CreateNames(Top:=False, Left:=True, Bottom:=False, Right:=False)',
              '    Kinds = TypeName(v) & " " & CStr(v) & Named(ws.Parent)', 'End Function', '']
    return '\n'.join(lines)


#: Labels of other kinds, in A1 down: a formula that shows text, dates, an error, a number formatted as text, blank
#: text, a name the workbook has already, and one label twice.
KINDS = ['    ws.Range("A1").Formula = "=""Calc"" & ""ulated"""', '    ws.Range("A2").Value = DateSerial(2026, 1, 1)',
         '    ws.Range("A3").Value = DateSerial(2026, 2, 3)', '    ws.Range("A3").NumberFormat = "yyyy-mm-dd"',
         '    ws.Range("A4").Formula = "=NA()"', '    ws.Range("A5").NumberFormat = "@"', '    ws.Range("A5").Value = 5',
         '    ws.Range("A6").Value = "   "', '    ws.Range("A7").Value = "Kept"', '    ws.Range("A8").Value = "Twice"',
         '    ws.Range("A9").Value = "twice"', '    ws.Range("A10").Formula = "=1+1"',
         '    ws.Range("A11").Value = "kept"', '    ws.Range("A12").Value = "x" & String(250, "y") & " %%%%%z"']


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()

        def run(proc: str, *args: object) -> str:
            result = excel.run_vba(source, proc, args, timeout=120.0, module_name='NamesProbe')
            assert result.ok and not result.dialogs, result
            return str(result.value)

        record: dict[str, object] = {'measured_at': datetime.now(timezone.utc).isoformat(), 'labels': LABELS,
                                     'grid': GRID, 'guesses': GUESSES, 'source': source}
        record['sides'] = {'%d%d%d%d' % sides: run('Sides', *sides) for sides in SIDES}
        record['guess_cases'] = [[layout, target, arguments, run(f'Guess{index}')]
                                 for index, (layout, target, arguments) in enumerate(GUESS_CASES)]
        record['kinds'] = run('Kinds')
        # Last: Names(1).Delete then fails on one of the names left, so nothing is cleared after it.
        record['cleaning'] = run('Cleaning')
    OUT.write_text(json.dumps(record, indent=2, ensure_ascii=True) + '\n', encoding='utf-8')
    print(record['cleaning'].replace('|', '\n'))
    for key, answer in record['sides'].items():  # type: ignore[union-attr]
        print(key, answer)
    for case in record['guess_cases']:  # type: ignore[union-attr]
        print(case)


if __name__ == '__main__':
    main()
