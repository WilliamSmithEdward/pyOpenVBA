"""Which number formats make a Text cell, which keeps what a macro writes to it as text.

A cell formatted @ keeps a formula, or a string that would read as a
number or a date, as the text written. So does a cell whose format has
an @ anywhere and no section that shows digits or a date: "x"@,
General;;;@ and *-General;;;@ make Text cells, and 0;@, General;0;0;@,
m/d/yyyy;;;@ and the Comma style do not. This sets each format on a cell
of its own and writes each value to it, reading back what the cell holds.

    python scripts/measure_text_cells.py

writes tests/fixtures/text_cells.json, which tests/test_excel_text_cells.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/text_cells.json'
#: Formats with an @ and others beside them, as written; Excel refuses some.
FORMATS = [
    '@', '@*-', '@" kg"', '"x"@', '0;@', ';;;@', 'General;General;General;@', '@;@', '@@', '[Red]@', '_(@_)',
    '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', '0.00', 'General', '"@"', '\\@', '@_)', ';@', '0;0;0;@',
    '[>5]@', 'mm/dd @', '0;0;0;"x"@', 'General;General;General;@*-', ';;;"x"@', ';;;@" kg"', '0;0;0;_(@_)', '0;0;@',
    '0;0;0;@_)', ';;;_(@', 'General;General;General;_(@_)', '@*-;@', '0;0;0;[Red]@', 'General;;;@',
]
#: Formats that tell apart a rule asking every number section to show General or nothing from one asking only the
#: first: General;0;0;@ is a Text cell by the second and not by the first.
SPLIT = [
    'General;0;0;@', ';0;;@', '0;;;@', ';;0;@', '"x";;;@', 'm/d/yyyy;;;@', 'General" kg";;;@', '[Red]General;;;@',
    '#;;;@', 'General;General;@', 'General;@', ';@" kg"', '0;0;General;@', 'General;General;0;@', '"x"General;;;@',
    '*-General;;;@', '_(General;;;@',
]
#: What each cell is written, after its format.
WRITES = [
    'c.Formula = "=1+1"', 'c.Value = "=1+1"', 'c.Value = "123"', 'c.Value = 123', 'c.Value2 = 45000',
    'c.Value = "1/2/2026"', 'c.Formula = "abc"', 'c.FormulaR1C1 = "=1+1"', 'c.Value = DateSerial(2026, 1, 2)',
    'c.Value = CCur(12.5)', 'c.Value = True',
]


def _vba(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def probe(formats: list[str]) -> str:
    """Each format on a row of its own, each write in a column of its own: F and the error where Excel refuses
    the format, E and the error where it refuses the write, else what the cell holds, whether it is a formula,
    its Formula and its NumberFormat."""
    named = '\n'.join(f'formats({index}) = {_vba(code)}' for index, code in enumerate(formats))
    writes = '\n'.join(f'                Case {index}: {write}' for index, write in enumerate(WRITES, 1))
    return f'''Function Probe() As String
Dim wb As Workbook, ws As Worksheet, c As Range, s As String, formats({len(formats) - 1}) As String
Dim i As Long, j As Long
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
{named}
On Error Resume Next
For i = 0 To {len(formats) - 1}
    For j = 1 To {len(WRITES)}
        Set c = ws.Cells(i + 1, j)
        Err.Clear
        c.NumberFormat = formats(i)
        If Err.Number <> 0 Then
            s = s & "|F" & Err.Number
        Else
            Err.Clear
            Select Case j
{writes}
            End Select
            If Err.Number <> 0 Then
                s = s & "|E" & Err.Number
            Else
                s = s & "|" & TypeName(c.Value) & "~" & c.HasFormula & "~" & c.Formula & "~" & c.NumberFormat
            End If
        End If
    Next
Next
On Error GoTo 0
wb.Close False
Application.DisplayAlerts = True
Probe = Mid$(s, 2)
End Function
'''


def main() -> None:
    record({'formats': probe(FORMATS), 'split': probe(SPLIT)}, OUT)


if __name__ == '__main__':
    main()
