"""Range.Text in a column too narrow for what a cell shows, on a 96-DPI display.

A number in General is shown with fewer decimals, then in scientific
notation, as the column narrows; any other number, a date among them,
fills the column with # once it no longer fits. This records Range.Text
for numbers, dates, text, booleans and errors, each in a column of every
width from 2 to 130 pixels, under each of the two Normal fonts the model
knows, and in bold and italic; the same for formats with a * fill, which
repeats a character across what the cell has to spare; then what else a
cell may have, and when Excel widens a column on its own for a number
format a macro sets, with the columns a saved workbook keeps after it
does.

    python scripts/measure_narrow_text.py

writes tests/fixtures/narrow_text.json, which tests/test_excel_narrow_text.py
replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/narrow_text.json'
#: What each cell holds: (formula, number format).
VALUES = [
    ('=1104.6221254112', 'General'), ('=123456789', 'General'), ('=1234567890123', 'General'),
    ('=0.000123456789', 'General'), ('=-1234.5678', 'General'), ('=1/3', 'General'), ('=-2/3', 'General'),
    ('=1E+20', 'General'), ('=12345678901', 'General'), ('=-0.5', 'General'), ('=7', 'General'),
    ('=99999.99999', 'General'), ('=1.5E-20', 'General'), ('=-123456789012', 'General'), ('=1234.5', '#,##0.00'),
    ('=46317', 'm/d/yyyy'), ('=0.256', '0.0%'), ('=1234.5', '$#,##0.00'), ('="Hello World"', 'General'),
    ('=TRUE', 'General'), ('=1/0', 'General'), ('=1234.5', '0.00E+00'), ('=-1234.5', '#,##0.00'),
    ('=-0.0001', 'General'), ('=-0.4', 'General'), ('=1.23456789E+100', 'General'), ('=-1.5E-100', 'General'),
    ('=0.15', 'General'), ('=2.5', 'General'), ('=1234567.5', 'General'), ('=0.00001234', 'General'),
]
#: The Comma style's format, whose * fills the cell with spaces between its _( and its number.
COMMA = r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)'
#: Formats with a * fill, which repeats its character across what the cell has to spare, or an _ pad, or General beside
#: other text: (formula, number format).
FILLS = [
    ('=1234.5', COMMA), ('=-1234.5', COMMA), ('=0', COMMA), ('="abc"', COMMA),
    ('=1234.5', r'_($* #,##0.00_);_($* \(#,##0.00\);_($* "-"??_);_(@_)'), ('=42', '*-0'), ('=42', '0*.'),
    ('abc', '@*-'), ('=7', '"x"*=0'), ('=42', '* 0'), ('=1', '*W0'), ('=42', '0*W'), ('=42', '*-0_)'),
    ('=42', '_W*-0'), ('=-12', '*-0;-0'), ('=TRUE', '*-General'), ('=1234.5678', '*-General'),
    ('=-42', '* 0'), ('=42', '0_W'), ('=1234.5', '#,##0.00_)'), ('=1234.5678', 'General" kg"'),
    ('=-1234.5678', '*-General'),
]
#: What else a cell may have beside a fill: (name, setup on A1, which holds 1234.5 in the Comma style in a column 100
#: pixels wide).
FILL_SETTINGS = {
    'fill_indented': 'ws.Range("A1").IndentLevel = 1',
    'fill_right_aligned': 'ws.Range("A1").HorizontalAlignment = xlRight',
    'fill_centered': 'ws.Range("A1").HorizontalAlignment = xlCenter',
    'fill_shrunk': 'ws.Range("A1").ShrinkToFit = True',
    'fill_wrapped': 'ws.Range("A1").WrapText = True',
    'fill_merged': 'ws.Range("A1:B1").Merge',
}
#: What else a cell may have, beside a number too wide for its column: (name, setup on A1, which holds 1104.6221254112
#: in a column 40 pixels wide).
SETTINGS = {
    'hidden_column': 'ws.Columns(1).Hidden = True',
    'merged_across': 'ws.Range("A1:B1").Merge',
    'indented': 'ws.Range("A1").IndentLevel = 1',
    'shrunk': 'ws.Range("A1").ShrinkToFit = True',
    'wrapped': 'ws.Range("A1").WrapText = True',
    'right_aligned': 'ws.Range("A1").HorizontalAlignment = xlRight',
    'big_font': 'ws.Range("A1").Font.Size = 14',
}
#: When Excel widens a column on its own: each case on a column of its own, its cells c, below it d, and the next
#: column e; what is read after it is the column's width in pixels and what c shows.
WIDENING = {
    'date_value': 'c.Value = DateSerial(2026, 10, 22)',
    'general_long': 'c.Value = 1234567.891',
    'format_after_short': 'c.Value = 1234.5: c.NumberFormat = "#,##0.00"',
    'format_after_long': 'c.Value = 1234567.5: c.NumberFormat = "#,##0.00"',
    'format_before_long': 'c.NumberFormat = "#,##0.00": c.Value = 1234567.5',
    'custom_width': 'c.ColumnWidth = 5: c.Value = 1234567.5: c.NumberFormat = "#,##0.00"',
    'formula_then_format': 'c.Formula = "=1234567.5": c.NumberFormat = "#,##0.00"',
    'formula_date': 'c.Formula = "=DATE(2026,10,22)"',
    'long_date_format': 'c.Value = DateSerial(2024, 3, 5): c.NumberFormat = "dddd, mmmm d, yyyy"',
    'value2_date': 'c.Value2 = 46317: c.NumberFormat = "m/d/yyyy"',
    'currency_value': 'c.Value = CCur(1234567.5)',
    'percent_format': 'c.Value = 12345.678: c.NumberFormat = "0.00%"',
    'back_to_general': 'c.Value = 1234567.5: c.NumberFormat = "#,##0.00": c.NumberFormat = "General"',
    'typed_date': 'c.Value = "10/22/2026"',
    'wider_again': 'c.Value = 1234567.5: c.NumberFormat = "#,##0.00": c.NumberFormat = "#,##0.0000"',
    'narrower_format': 'c.Value = 1234567.5: c.NumberFormat = "#,##0.0000": c.NumberFormat = "#,##0.00"',
    'two_cells': 'c.Value = 1234.5: d.Value = 1234567.5: Range(c, d).NumberFormat = "#,##0.00"',
    'other_cell_shown': 'c.Value = 1: d.Value = 1234567.5: c.NumberFormat = "#,##0.00"',
    'whole_column': 'c.Value = 1234567.5: c.EntireColumn.NumberFormat = "#,##0.00"',
    'format_local': 'c.Value = 1234567.5: c.NumberFormatLocal = "#,##0.00"',
    'cell_style': 'c.Value = 1234567.5: c.Style = "Comma"',
    'after_autofit': 'c.Value = 12: c.EntireColumn.AutoFit: c.Value = 1234567.5: c.NumberFormat = "#,##0.00"',
    'bold_cell': 'c.Value = 1234567.5: c.Font.Bold = True: c.NumberFormat = "#,##0.00"',
    'italic_cell': 'c.Value = 1234567.5: c.Font.Italic = True: c.NumberFormat = "#,##0.00"',
    'text_value': 'c.Value = "1234567.5 long text": c.NumberFormat = "#,##0.00"',
    'merged_cell': 'c.Value = 1234567.5: Range(c, e).Merge: c.NumberFormat = "#,##0.00"',
    'wrapped_cell': 'c.Value = 1234567.5: c.WrapText = True: c.NumberFormat = "#,##0.00"',
    'hidden_column': 'c.Value = 1234567.5: c.EntireColumn.Hidden = True: c.NumberFormat = "#,##0.00": '
                     'c.EntireColumn.Hidden = False',
    'big_number_format': 'c.Value = 1234567890123#: c.NumberFormat = "0"',
    'too_wide': 'c.Value = 1E+250: c.NumberFormat = "0"',
    # Excel widens wherever AutoFit would, the number shown whole or not: 12345678 and its space fit 64 pixels,
    # and an _W leaves only a space in Range.Text but a W's width for AutoFit.
    'shown_whole_wider': 'c.Value = 12345678: c.NumberFormat = "00000000"" """',
    'pad_wider': 'c.Value = 1234567: c.NumberFormat = "0000000_W"',
    'autofit_as_wide': 'c.Value = 123456: c.NumberFormat = "000000""((("""',
    'autofit_narrower': 'c.Value = 12345678: c.NumberFormat = "00000000"',
}


def widening() -> str:
    """Each case, two columns apart from the next: the column's width in pixels and what c shows."""
    cases = '\n'.join(f'''Set c = ws.Cells(1, {3 * index - 2}): Set d = c.Offset(1, 0): Set e = c.Offset(0, 1)
{setup}
s = s & "#{name}=" & CLng(ws.Columns({3 * index - 2}).Width * 4 / 3) & "~" & c.Text''' for index, (name, setup) in
                      enumerate(WIDENING.items(), 1))
    return f'''Function Probe() As String
Dim wb As Workbook, ws As Worksheet, c As Range, d As Range, e As Range, s As String
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
On Error Resume Next
{cases}
On Error GoTo 0
wb.Close False
Application.DisplayAlerts = True
Probe = Mid$(s, 2)
End Function
'''
#: The column widths tried, in pixels.
PIXELS = range(2, 131)


def probe(normal: str, bold: bool, italic: bool, values: list[tuple[str, str]], pixels: str = '') -> str:
    """Range.Text of each value at each width in pixels: every one of PIXELS, or those listed in ``pixels``."""
    cases = '\n'.join(f'ws.Range("A{index}").NumberFormat = "{code.replace(chr(34), chr(34) * 2)}": '
                      f'ws.Range("A{index}").Formula = "{formula.replace(chr(34), chr(34) * 2)}"'
                      for index, (formula, code) in enumerate(values, 1))
    loop = (f'For Each width In Array({pixels})\n    px = width' if pixels else
            f'For px = {PIXELS.start} To {PIXELS.stop - 1} Step {PIXELS.step}')
    return f'''Function Probe() As String
Dim wb As Workbook, ws As Worksheet, s As String, i As Long, px As Long, width As Variant
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
wb.Styles("Normal").Font.Name = "{normal}"
Set ws = wb.Worksheets(1)
{cases}
ws.Range("A1:A{len(values)}").Font.Bold = {bold}
ws.Range("A1:A{len(values)}").Font.Italic = {italic}
{loop}
    ws.Columns(1).ColumnWidth = IIf(px < 12, px / 12, (px - 5) / 7)
    s = s & Chr(10) & CLng(ws.Columns(1).Width * 4 / 3)
    For i = 1 To {len(values)}
        s = s & Chr(9) & ws.Range("A" & i).Text
    Next
Next
wb.Close False
Application.DisplayAlerts = True
Probe = Mid$(s, 2)
End Function
'''


def fill_setting(setup: str) -> str:
    """Range.Text of 1234.5 in the Comma style in a column 100 pixels wide, with something else set on its cell."""
    return f'''Function Probe() As String
Dim wb As Workbook, ws As Worksheet
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Range("A1").NumberFormat = "{COMMA.replace(chr(34), chr(34) * 2)}"
ws.Range("A1").Value = 1234.5
ws.Columns(1).ColumnWidth = 95 / 7
ws.Columns(2).ColumnWidth = 5
{setup}
Probe = ws.Range("A1").Text
wb.Close False
Application.DisplayAlerts = True
End Function
'''


def setting(setup: str) -> str:
    """Range.Text of 1104.6221254112 in a column 40 pixels wide, with something else set on its cell."""
    return f'''Function Probe() As String
Dim wb As Workbook, ws As Worksheet
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Range("A1").Value = 1104.6221254112
ws.Columns(1).ColumnWidth = 5
ws.Columns(2).ColumnWidth = 5
{setup}
Probe = ws.Range("A1").Text
wb.Close False
Application.DisplayAlerts = True
End Function
'''


def main() -> None:
    modules = {}
    for normal in ('Aptos Narrow', 'Calibri'):
        for style, bold, italic in (('regular', False, False), ('bold', True, False), ('italic', False, True)):
            modules[f'{normal.lower().replace(" ", "_")}_{style}'] = probe(normal, bold, italic, VALUES)
    for normal, style, bold, italic in (('Aptos Narrow', 'regular', False, False), ('Calibri', 'regular', False, False),
                                        ('Aptos Narrow', 'italic', False, True)):
        modules[f'fill_{normal.lower().replace(" ", "_")}_{style}'] = probe(normal, bold, italic, FILLS)
    modules['fill_wide'] = probe('Aptos Narrow', False, False, FILLS, '200, 300, 425, 600, 1000')
    modules.update({name: setting(setup) for name, setup in SETTINGS.items()})
    modules.update({name: fill_setting(setup) for name, setup in FILL_SETTINGS.items()})
    modules['widening'] = widening()
    record(modules, OUT)
    fixture = json.loads(OUT.read_text(encoding='utf-8'))
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        path = Path(folder) / 'widened.xlsx'
        saved = excel.run_vba(WIDENED.replace('{path}', str(path).replace('"', '""')), 'Probe', timeout=120.0,
                              line_numbers=False)
        assert saved.ok and not saved.dialogs, saved
        with zipfile.ZipFile(path) as package:
            sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<cols>.*?</cols>', sheet)
    fixture['file'] = {'macro': WIDENED, 'cols': found.group(0) if found else ''}
    OUT.write_text(json.dumps(fixture, indent=2) + '\n', encoding='utf-8')
    print(fixture['file']['cols'])


#: A workbook whose column A Excel widened for a format, and B AutoFit then widened, saved for the <cols> it writes.
WIDENED = '''Function Probe() As String
Dim wb As Workbook, ws As Worksheet
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Range("A1").Value = 1234567.5: ws.Range("A1").NumberFormat = "#,##0.00"
ws.Range("B1").Value = 12: ws.Columns(2).AutoFit: ws.Range("B1").Value = 1234567.5: ws.Range("B1").NumberFormat = "#,##0.00"
wb.SaveAs Filename:="{path}", FileFormat:=xlOpenXMLWorkbook
wb.Close False
Application.DisplayAlerts = True
Probe = "saved"
End Function
'''


if __name__ == '__main__':
    main()
