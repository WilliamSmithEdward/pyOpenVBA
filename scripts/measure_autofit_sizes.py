"""Column AutoFit and Range.Text at every font size, on a 96-DPI display.

AutoFit was measured at 11pt (scripts/measure_column_autofit.py). This
measures it at 35 sizes from 6 to 72pt, for each of the two Normal fonts
the model knows (Aptos Narrow 11 and Calibri 11) and each style: random
texts of ASCII and Latin-1 and a few set ones, each fitted alone. Then
Range.Text in columns of every width at ten sizes, for numbers in General
and other formats, a date, text, a logical and a fill; the column a
number format widens at those sizes; and a cell in the other font, which
shows whether its own font or the Normal font decides.

    python scripts/measure_autofit_sizes.py

writes tests/fixtures/autofit_sizes.json, which tests/test_excel_autofit_sizes.py
replays.
"""
from __future__ import annotations

import json
import random
from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/autofit_sizes.json'
NORMALS = ('Aptos Narrow', 'Calibri')
STYLES = {'regular': (False, False), 'bold': (True, False), 'italic': (False, True), 'bold italic': (True, True)}
#: The sizes AutoFit is measured at: every whole size to 24 and the font box's sizes past it, with three halves.
SIZES = (6, 7, 8, 9, 10, 10.5, 11, 11.5, 12, 13, 13.5, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 26, 28, 30, 32, 34,
         36, 40, 44, 48, 54, 60, 66, 72)
#: The sizes Range.Text is measured at.
NARROW_SIZES = (8, 9, 10, 12, 14, 16, 18, 20, 24, 36)
#: What Range.Text is read of: (formula, number format).
COMMA = r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)'
NARROW_VALUES = [
    ('=1104.6221254112', 'General'), ('=-1234.5678', 'General'), ('=1/3', 'General'), ('=1E+20', 'General'),
    ('=0.000123456789', 'General'), ('=1234.5', '#,##0.00'), ('=46317', 'm/d/yyyy'), ('="Hello World"', 'General'),
    ('=TRUE', 'General'), ('=1234.5', COMMA),
]
#: The sizes the widening a number format gives a column is measured at.
WIDENING_SIZES = (8, 10, 12, 14, 18, 24, 36)


def texts() -> list[str]:
    """Random texts of ASCII and Latin-1, spaces among them, 1 to 40 long, and a few set ones."""
    chosen = random.Random(2026_10_22)
    pool = [chr(code) for code in (*range(0x21, 0x7F), *range(0xA0, 0x100))] + [' '] * 8
    found = [''.join(chosen.choice(pool) for _ in range(chosen.randint(1, 40))).strip() or 'x' for _ in range(24)]
    return ['abc', 'Hello World', '0123456789', 'WWWW', 'iiii', 'Total Revenue', 'x', *found]


def _vba(text: str) -> str:
    """A VBA expression for ``text``, ChrW for anything past ASCII and the quote."""
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
    return ' & '.join(pieces) or '""'


def _numbers(values: tuple[float, ...]) -> str:
    return ', '.join(f'{value:g}' for value in values)


OPENING = '''Function Fitted(ws As Worksheet, column As Long) As Long
ws.Columns(column).AutoFit
Fitted = CLng(ws.Columns(column).Width * 4 / 3)
End Function

Function Probe() As String
Dim wb As Workbook, ws As Worksheet, s As String, i As Long, j As Long, px As Long, sizes As Variant
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


def fits(normal: str, font: str, bold: bool, italic: bool, sizes: tuple[float, ...]) -> str:
    """AutoFit's width for each text alone in ``font`` at each size: #size: then the widths."""
    chosen = texts()
    named = '\n'.join(f'texts({index}) = {_vba(text)}' for index, text in enumerate(chosen))
    return OPENING.replace('NORMAL', normal) + f'''Dim texts({len(chosen) - 1}) As String
{named}
sizes = Array({_numbers(sizes)})
With ws.Range("A1").Font
    .Name = "{font}": .Bold = {bold}: .Italic = {italic}
End With
For i = LBound(sizes) To UBound(sizes)
    ws.Range("A1").Font.Size = sizes(i)
    s = s & "#" & sizes(i) & ":"
    For j = 0 To {len(chosen) - 1}
        ws.Range("A1").Value = "'" & texts(j)
        s = s & " " & Fitted(ws, 1)
    Next
Next''' + CLOSING


def narrow(normal: str, italic: bool, size: float) -> str:
    """Range.Text of each value at ``size`` in columns from 2 pixels to wider than any of them needs: a line a
    width, the width in pixels and then each text, tab apart."""
    cases = '\n'.join(f'ws.Range("A{index}").NumberFormat = "{code.replace(chr(34), chr(34) * 2)}": '
                      f'ws.Range("A{index}").Formula = "{formula.replace(chr(34), chr(34) * 2)}"'
                      for index, (formula, code) in enumerate(NARROW_VALUES, 1))
    widest = round(size * 14) + 20
    return OPENING.replace('NORMAL', normal) + f'''{cases}
With ws.Range("A1:A{len(NARROW_VALUES)}").Font
    .Size = {size:g}: .Italic = {italic}
End With
For px = 2 To {widest}
    ws.Columns(1).ColumnWidth = IIf(px < 12, px / 12, (px - 5) / 7)
    s = s & Chr(10) & CLng(ws.Columns(1).Width * 4 / 3)
    For i = 1 To {len(NARROW_VALUES)}
        s = s & Chr(9) & ws.Range("A" & i).Text
    Next
Next''' + CLOSING


def widening(normal: str) -> str:
    """At each size, in a column of its own: 1234567.5 given #,##0.00, 1234567.891 in General, and 46317 given
    m/d/yyyy; the column's width in pixels and what the cell shows, after each."""
    return OPENING.replace('NORMAL', normal) + f'''Dim c As Range, k As Long, values As Variant, codes As Variant
sizes = Array({_numbers(WIDENING_SIZES)})
values = Array(1234567.5, 1234567.891, 46317)
codes = Array("#,##0.00", "General", "m/d/yyyy")
For i = LBound(sizes) To UBound(sizes)
    For k = 0 To 2
        Set c = ws.Cells(1, 3 * i + k + 1)
        c.Font.Size = sizes(i)
        c.Value = values(k)
        c.NumberFormat = codes(k)
        s = s & "#" & sizes(i) & "/" & k & "=" & CLng(ws.Columns(c.Column).Width * 4 / 3) & "~" & c.Text
    Next
Next''' + CLOSING


def probes() -> dict[str, str]:
    found: dict[str, str] = {}
    for normal in NORMALS:
        tag = normal.lower().replace(' ', '_')
        for style, (bold, italic) in STYLES.items():
            # In two halves, each well inside the time a probe may take.
            name = f'fits_{tag}_{style.replace(" ", "_")}'
            found[name + '_small'] = fits(normal, normal, bold, italic, SIZES[:19])
            found[name + '_large'] = fits(normal, normal, bold, italic, SIZES[19:])
        for size in NARROW_SIZES:
            for style, italic in (('regular', False), ('italic', True)):
                found[f'narrow_{tag}_{style}_{size:g}'] = narrow(normal, italic, size)
        found[f'widening_{tag}'] = widening(normal)
    # A cell in the other font, under each Normal font: its own font decides at 11pt; at other sizes too?
    found['fits_calibri_under_aptos_narrow'] = fits('Aptos Narrow', 'Calibri', False, False, (8, 14, 24, 36))
    found['fits_aptos_narrow_under_calibri'] = fits('Calibri', 'Aptos Narrow', False, False, (8, 14, 24, 36))
    return found


def main() -> None:
    record(probes(), OUT)
    fixture = json.loads(OUT.read_text(encoding='utf-8'))
    fixture.update({'sizes': list(SIZES), 'texts': texts(), 'narrow_sizes': list(NARROW_SIZES),
                    'narrow_values': NARROW_VALUES, 'widening_sizes': list(WIDENING_SIZES)})
    OUT.write_text(json.dumps(fixture, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
