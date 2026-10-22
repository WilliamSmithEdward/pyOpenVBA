"""Column AutoFit: how wide Excel makes a column for the text its cells show, on a 96-DPI display.

AutoFit makes a column wide enough for the widest text its cells show,
measured in the display's pixels with each character's width in the
cell's font. So this measures, for each of the two Normal fonts the model
knows (Aptos Narrow 11 and Calibri 11) and each style, the width AutoFit
gives a cell holding each character of Windows-1252 on its own; then runs
of one character up to the widest column, to find how the text's width
becomes the column's; random strings, to show the widths add up; numbers
in the formats that show them; and the cells, ranges and settings AutoFit
looks at. Every probe works in a workbook of its own.

    python scripts/measure_column_autofit.py

writes tests/fixtures/column_autofit.json, which
tests/test_excel_column_autofit.py replays.
"""
from __future__ import annotations

import json
import random
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/column_autofit.json'
NORMALS = ('Aptos Narrow', 'Calibri')
STYLES = {'regular': (False, False), 'bold': (True, False), 'italic': (False, True), 'bold italic': (True, True)}
#: Windows-1252's characters by their Unicode code points: ASCII, Latin-1 and the 27 in between.
CODES = [*range(0x20, 0x7F), *range(0xA0, 0x100),
         0x20AC, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017D,
         0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x017E,
         0x0178]

#: A workbook of its own whose Normal font is NORMAL, and a sheet ws in it; the probe body follows.
OPENING = '''Function Fitted(ws As Worksheet, column As Long) As Long
ws.Columns(column).AutoFit
Fitted = CLng(ws.Columns(column).Width * 4 / 3)
End Function

Function Probe() As String
Dim wb As Workbook, ws As Worksheet, s As String, i As Long, items As Variant
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
wb.Styles("Normal").Font.Name = "NORMAL"
wb.Styles("Normal").Font.Size = 11
Set ws = wb.Worksheets(1)
'''
CLOSING = '''
wb.Close False
Application.DisplayAlerts = True
Probe = Mid$(s, 2)
End Function
'''


def module(normal: str, body: str) -> str:
    return OPENING.replace('NORMAL', normal) + body + CLOSING


def glyphs(normal: str, font: str, bold: bool, italic: bool) -> str:
    """Each character c in ``font``: the width AutoFit gives column A for c alone, xc, cx and xcx, after x and xx.
    The differences separate what c adds between two characters from what its ink reaches past it at either end
    of a text."""
    extras = ', '.join(str(code) for code in CODES if code > 0xFF)
    return module(normal, f'''With ws.Range("A1").Font
    .Name = "{font}": .Size = 11: .Bold = {bold}: .Italic = {italic}
End With
ws.Range("A1").Value = "'x": s = " x:" & Fitted(ws, 1)
ws.Range("A1").Value = "'xx": s = s & " xx:" & Fitted(ws, 1)
For i = 32 To 126: s = s & One(ws, i): Next
For i = 160 To 255: s = s & One(ws, i): Next
items = Array({extras})
For i = LBound(items) To UBound(items): s = s & One(ws, CLng(items(i))): Next''') + '''
Function One(ws As Worksheet, code As Long) As String
Dim s As String
s = " " & code & ":"
ws.Range("A1").Value = "'" & ChrW(code): s = s & Fitted(ws, 1)
ws.Range("A1").Value = "'x" & ChrW(code): s = s & "," & Fitted(ws, 1)
ws.Range("A1").Value = "'" & ChrW(code) & "x": s = s & "," & Fitted(ws, 1)
ws.Range("A1").Value = "'x" & ChrW(code) & "x": s = s & "," & Fitted(ws, 1)
One = s
End Function
'''


def runs(normal: str) -> str:
    """Runs of 0, i and W, 1 to 260 long: AutoFit's width for each, to the widest column."""
    return module(normal, '''items = Array("0", "i", "W")
For i = LBound(items) To UBound(items)
    s = s & "#" & items(i) & ":"
    For n = 1 To 260
        ws.Range("A1").Value = "'" & String(n, items(i))
        s = s & " " & Fitted(ws, 1)
    Next
Next''').replace('Dim wb As Workbook,', 'Dim n As Long, wb As Workbook,')


def strings(seed: int) -> list[str]:
    """Strings of Windows-1252's printable characters, spaces among them, 1 to 40 long, and pairs that kern in
    typesetting."""
    chosen = random.Random(seed)
    pool = [chr(code) for code in CODES if code != 0x20] + [' '] * 8
    found = [''.join(chosen.choice(pool) for _ in range(chosen.randint(1, 40))).strip() or 'x' for _ in range(160)]
    return found + ['AV', 'AVAVAVAV', 'To', 'Wa', 'Ty', 'LT', 'Yo', 'fi', 'ff', 'Hello World', 'Total Revenue']


def latin(seed: int) -> list[str]:
    """Strings of ASCII and Latin-1 alone, 1 to 60 long."""
    chosen = random.Random(seed)
    pool = [chr(code) for code in CODES if code <= 0xFF and code != 0x20] + [' '] * 8
    return [''.join(chosen.choice(pool) for _ in range(chosen.randint(1, 60))).strip() or 'x' for _ in range(120)]


def styled(normal: str, bold: bool, italic: bool, texts: list[str]) -> str:
    """Each text in the Normal font in a style: AutoFit's width."""
    return module(normal, f'ws.Range("A1").Font.Bold = {bold}: ws.Range("A1").Font.Italic = {italic}\n' + '\n'.join(
        f'ws.Range("A1").Value = "\'" & {_vba(text)}: s = s & " " & Fitted(ws, 1)' for text in texts))


def _vba(text: str) -> str:
    """A VBA expression for ``text``, ChrW for anything past ASCII."""
    pieces: list[str] = []
    plain = ''
    for char in text:
        if 32 <= ord(char) < 127 and char != '"':
            plain += char
            continue
        if plain:
            pieces.append(f'"{plain}"')
            plain = ''
        pieces.append(f'ChrW({ord(char)})')
    if plain:
        pieces.append(f'"{plain}"')
    return ' & '.join(pieces)


def mixed(normal: str, texts: list[str]) -> str:
    lines = '\n'.join(f'ws.Range("A1").Value = "\'" & {_vba(text)}: s = s & " " & Fitted(ws, 1)' for text in texts)
    return module(normal, lines)


#: Numbers and what shows them: (formula, number format).
NUMBERS = [
    ('=1', 'General'), ('=12345', 'General'), ('=12345678901', 'General'), ('=123456789012', 'General'),
    ('=0.5', 'General'), ('=1104.6221254112', 'General'), ('=-1', 'General'), ('=-0.5', 'General'),
    ('=1/3', 'General'), ('=-1/3', 'General'), ('=1E+20', 'General'), ('=1.5E-20', 'General'), ('=0.000123', 'General'),
    ('=123456.789', 'General'), ('=1234.5', '#,##0.00'), ('=1234567.891', '#,##0'), ('=-1234.5', '#,##0.00'),
    ('=0.256', '0%'), ('=0.256', '0.00%'), ('=1234.5', '$#,##0.00'), ('=-1234.5', '$#,##0.00_);($#,##0.00)'),
    ('=-1234.5', '$#,##0.00_);[Red]($#,##0.00)'), ('=1234.5', '0.00E+00'), ('=0.5', '# ?/?'), ('=46317', 'm/d/yyyy'),
    ('=46317', 'mmmm d, yyyy'), ('=46317', 'dddd'), ('=0.75', 'h:mm AM/PM'), ('=0.75', 'h:mm:ss'),
    ('=46317.75', 'm/d/yyyy h:mm'), ('=1234.5', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'),
    ('=0', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'), ('=TRUE', 'General'), ('=FALSE', 'General'),
    ('=1/0', 'General'), ('=NA()', 'General'), ('=""', 'General'), ('="text"', 'General'), ('=1234.5', '@'),
    ('=1234.5', '0.0'), ('=1234.5', '000000'), ('=7', '"Qty: "0'), ('=12', '0" items"'),
    ('=0', 'General'), ('=0.00001', 'General'), ('=-1.5E-20', 'General'), ('=-1234567890123', 'General'),
    ('=1234567890123', 'General'), ('=12345.6789012345', 'General'), ('=-12345.6789012345', 'General'),
    ('=1E+100', 'General'), ('=-0.000001234', 'General'), ('=99999999999', 'General'), ('=999999999999', 'General'),
    ('=0.1+0.2', 'General'), ('=2/3', 'General'), ('=-2/3', 'General'), ('=1234.5', '#,##0.00;[Red]-#,##0.00'),
    ('=-1234.5', '#,##0.00;[Red]-#,##0.00'), ('=46317', 'General'), ('=46317', 'd-mmm-yy'), ('=0.5', '0.0%'),
]
#: Numbers in bold: (formula, number format).
BOLD_NUMBERS = [('=0.5', 'General'), ('=1104.6221254112', 'General'), ('=-1/3', 'General'), ('=1234.5', '#,##0.00'),
                ('=46317', 'm/d/yyyy'), ('=1.5E-20', 'General'), ('=TRUE', 'General')]


def numbers(normal: str, cases: list[tuple[str, str]], *, bold: bool = False) -> str:
    """Each number in its format: AutoFit's width and the text it shows then, a tab before each. A cell that shows
    nothing leaves the column as it was, so each case starts from a column 10 pixels wide."""
    lines = '\n'.join(f'ws.Columns(1).ColumnWidth = 0.83: ws.Range("A1").NumberFormat = "{code.replace(chr(34), chr(34) * 2)}": '
                      f'ws.Range("A1").Formula = "{formula.replace(chr(34), chr(34) * 2)}": '
                      f's = s & Chr(9) & Fitted(ws, 1) & "~" & ws.Range("A1").Text'
                      for formula, code in cases)
    return module(normal, f'ws.Range("A1").Font.Bold = {bold}\n' + lines)


#: What AutoFit looks at, each case on a fresh column: (name, setup).
CELLS = {
    'widest_cell': 'ws.Range("A1").Value = "\'abc": ws.Range("A2").Value = "\'abcdefghij": ws.Range("A3").Value = 5',
    'hidden_row': 'ws.Range("A1").Value = "\'abc": ws.Range("A2").Value = "\'abcdefghijklmnop": ws.Rows(2).Hidden = True',
    'wrapped': 'ws.Range("A1").Value = "\'abcdefghij klmnopqrst": ws.Range("A1").WrapText = True',
    'line_break': 'ws.Range("A1").Value = "abc" & Chr(10) & "abcdefghijkl"',
    'merged': 'ws.Range("A1").Value = "\'abcdefghijklmnopqrst": ws.Range("A1:B1").Merge: ws.Range("A2").Value = "\'abc"',
    'indent': 'ws.Range("A1").Value = "\'abc": ws.Range("A1").IndentLevel = 2',
    'right': 'ws.Range("A1").Value = "\'abcdef": ws.Range("A1").HorizontalAlignment = xlRight',
    'center': 'ws.Range("A1").Value = "\'abcdef": ws.Range("A1").HorizontalAlignment = xlCenter',
    'spaces': 'ws.Range("A1").Value = "\'   abc   "',
    'shrink': 'ws.Range("A1").Value = "\'abcdefghij": ws.Range("A1").ShrinkToFit = True',
    'turned': 'ws.Range("A1").Value = "\'abcdefghij": ws.Range("A1").Orientation = 90',
    'formula_text': 'ws.Range("A1").Formula = "=REPT(""ab"",5)"',
    'empty_text': 'ws.Range("A1").Formula = "=""""": ws.Range("A2").Value = "\'ab"',
    'nothing': 'ws.Columns(1).ColumnWidth = 20',
    'big_font': 'ws.Range("A1").Value = "\'abc": ws.Range("A1").Font.Size = 20',
    'mixed_runs': 'ws.Range("A1").Value = "\'abcdef": ws.Range("A1").Characters(1, 3).Font.Bold = True',
    'filtered_out': 'ws.Range("A1").Value = "\'h": ws.Range("A2").Value = "\'abc": ws.Range("A3").Value = "\'abcdefghijklmnop": '
                    'ws.Range("A1:A3").AutoFilter Field:=1, Criteria1:="abc"',
    'wrapped_beside': 'ws.Range("A1").Value = "\'abcdefghij klmnopqrst uvwxyz": ws.Range("A1").WrapText = True: '
                      'ws.Range("A2").Value = "\'abc"',
    'line_break_unwrapped': 'ws.Range("A1").Value = "abc" & Chr(10) & "abcdefghijkl": ws.Range("A1").WrapText = False',
    'line_break_wraps': 'ws.Range("A1").Value = "abc" & Chr(10) & "abcdefghijkl": s = s & "#wraps=" & ws.Range("A1").WrapText',
    'bold_header': 'ws.Range("A1").Value = "\'Total Revenue": ws.Range("A1").Font.Bold = True: ws.Range("A2").Value = 1234.5',
}
#: Sizes and texts measured beyond the Normal font's 11pt.
SIZES = (8, 9, 10, 12, 14, 16, 18, 20, 24, 36)
SIZED = ('abc', 'Hello World', '0123456789', 'WWWW', 'iiii')


def sizes(normal: str) -> str:
    """Each text at each size, regular, bold and italic: AutoFit's width."""
    lines = []
    for bold, italic in ((False, False), (True, False), (False, True)):
        lines.append(f'ws.Range("A1").Font.Bold = {bold}: ws.Range("A1").Font.Italic = {italic}')
        for size in SIZES:
            lines.append(f'ws.Range("A1").Font.Size = {size}')
            lines += [f'ws.Range("A1").Value = "\'{text}": s = s & " " & Fitted(ws, 1)' for text in SIZED]
    return module(normal, '\n'.join(lines))


def states(normal: str) -> str:
    """A hidden column AutoFit, and what ColumnWidth and Width read after AutoFit."""
    return module(normal, '''ws.Range("A1").Value = "'abcdefghij"
ws.Columns(1).Hidden = True
ws.Columns(1).AutoFit
s = s & "#hidden_column=" & ws.Columns(1).Hidden & ":" & ws.Columns(1).ColumnWidth & ":" & ws.Columns(1).Width
ws.Columns(1).Hidden = False
s = s & ":" & ws.Columns(1).ColumnWidth
ws.Range("B1").Value = "'Hello World": ws.Columns(2).AutoFit
s = s & "#readback=" & ws.Columns(2).ColumnWidth & ":" & ws.Columns(2).Width
ws.Range("C1").Value = "'" & String(300, "W"): ws.Columns(3).AutoFit
s = s & "#widest=" & ws.Columns(3).ColumnWidth & ":" & ws.Columns(3).Width''')


def cells(normal: str) -> str:
    lines = '\n'.join(f'''If ws.AutoFilterMode Then ws.AutoFilterMode = False
ws.Cells.UnMerge: ws.Cells.Clear: ws.Columns(1).ColumnWidth = 8.43: ws.Rows("1:3").Hidden = False
{setup}
s = s & "#{name}=" & Fitted(ws, 1)''' for name, setup in CELLS.items())
    return module(normal, lines)


#: Which cells AutoFit reads when it is asked of a range: (name, statement).
RANGES = {
    'columns_of_part': 'ws.Range("A2:A3").Columns.AutoFit',
    'entire_column': 'ws.Range("A2:A3").EntireColumn.AutoFit',
    'whole_columns': 'ws.Columns("A:B").AutoFit',
    'used_range': 'ws.UsedRange.Columns.AutoFit',
    'cells': 'ws.Cells.Columns.AutoFit',
    'one_cell_columns': 'ws.Range("B2").Columns.AutoFit',
}


def ranges(normal: str) -> str:
    lines = '\n'.join(f'''ws.Columns("A:B").ColumnWidth = 8.43
On Error Resume Next
Err.Clear
{statement}
s = s & "#{name}=" & Err.Number & ":" & CLng(ws.Columns(1).Width * 4 / 3) & ":" & CLng(ws.Columns(2).Width * 4 / 3)
On Error GoTo 0''' for name, statement in RANGES.items())
    setup = ('ws.Range("A1").Value = "\'abcdefghijklmnopqrstuvwxyz": ws.Range("A2").Value = "\'abc": '
             'ws.Range("A3").Value = "\'abcdefgh": ws.Range("B1").Value = "\'x": ws.Range("B2").Value = "\'abcdefghijkl"\n')
    return module(normal, setup + lines + '''
ws.Columns("A:B").ColumnWidth = 8.43
On Error Resume Next
Err.Clear
ws.Range("A2:A3").AutoFit
s = s & "#range_autofit=" & Err.Number & ":" & Err.Description
On Error GoTo 0''')


def probes() -> dict[str, str]:
    found: dict[str, str] = {}
    for normal in NORMALS:
        tag = normal.lower().replace(' ', '_')
        found['display_' + tag] = module(normal, 's = "|" & ws.StandardHeight & "|" & ws.StandardWidth & "|" & '
                                                 'Application.StandardFont & "|" & wb.Styles("Normal").Font.Name')
        for font in NORMALS:
            for style, (bold, italic) in STYLES.items():
                found[f'glyphs_{tag}_{font.lower().replace(" ", "_")}_{style.replace(" ", "_")}'] = \
                    glyphs(normal, font, bold, italic)
        found['runs_' + tag] = runs(normal)
        found['mixed_' + tag] = mixed(normal, strings(2026))
        for style, (bold, italic) in STYLES.items():
            found[f'latin_{tag}_{style.replace(" ", "_")}'] = styled(normal, bold, italic, latin(2027))
        found['numbers_' + tag] = numbers(normal, NUMBERS)
        found['bold_numbers_' + tag] = numbers(normal, BOLD_NUMBERS, bold=True)
        found['cells_' + tag] = cells(normal)
        found['ranges_' + tag] = ranges(normal)
        found['sizes_' + tag] = sizes(normal)
        found['states_' + tag] = states(normal)
    return found


#: A workbook whose columns A and B are AutoFit and C sized by hand, saved for the <cols> Excel writes.
FILE = '''Function Probe() As String
Dim wb As Workbook, ws As Worksheet
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Range("A1").Value = "'Hello World": ws.Range("B1").Value = 1104.6221254112: ws.Range("C1").Value = "'x"
ws.Range("D1").Value = "'abcdefghijklmnop": ws.Range("E1").Value = "'abc"
ws.Columns("A:B").AutoFit
ws.Columns("C").ColumnWidth = 20
ws.Columns("D:E").AutoFit
ws.Columns("E").ColumnWidth = 15
wb.SaveAs Filename:="{path}", FileFormat:=xlOpenXMLWorkbook
wb.Close False
Application.DisplayAlerts = True
Probe = "saved"
End Function
'''


def main() -> None:
    results: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\nEnd Function',
                                'V', line_numbers=False)
        for name, source in probes().items():
            result = excel.run_vba(source, 'Probe', timeout=300.0, line_numbers=False)
            assert result.ok and not result.dialogs, (name, result)
            results.append({'name': name, 'macro': source, 'result': result.value})
            print(name, str(result.value)[:200], flush=True)
        path = Path(folder) / 'fitted.xlsx'
        saved = excel.run_vba(FILE.replace('{path}', str(path).replace('"', '""')), 'Probe', timeout=120.0,
                              line_numbers=False)
        assert saved.ok and not saved.dialogs, saved
        with zipfile.ZipFile(path) as package:
            sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
        found = re.search(r'<cols>.*?</cols>', sheet)
        columns = found.group(0) if found else ''
        print(columns)
    OUT.write_text(json.dumps({'excel_version_build': version.value,
                               'measured_at': datetime.now(timezone.utc).isoformat(),
                               'strings': strings(2026), 'latin': latin(2027), 'numbers': NUMBERS,
                               'bold_numbers': BOLD_NUMBERS, 'cells': CELLS, 'ranges': RANGES, 'sizes': SIZES,
                               'sized': SIZED, 'file': {'macro': FILE, 'cols': columns},
                               'probes': results}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
