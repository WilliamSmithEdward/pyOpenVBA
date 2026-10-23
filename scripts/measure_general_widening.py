"""When a number's column widens on its own: a number General shows, and any number when its font changes.

Setting a number format on a number widens its column wherever AutoFit
would make it wider (scripts/measure_narrow_text.py). A number General
shows widens too, when General is set on it after its value or its font
changes, but to a width that is not AutoFit's for its whole text; writing
the value alone never widens. This writes each value at each size in a
column of its own at the standard width, in each way, and records the
column's width in pixels and what the cell shows; then the same for a
number given #,##0.00 before its font grows, and a date.

    python scripts/measure_general_widening.py

writes tests/fixtures/general_widening.json, which
tests/test_excel_general_widening.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/general_widening.json'
SIZES = (11, 12, 13, 14, 18, 24)
VALUES = ('1234567.891', '123456789', '12345.6789', '0.123456789', '1234567890123', '-1234567.891', '1E+20',
          '12345678', '123456.7', '98765.4321', '7', '1234.5')
#: How each value goes in: by Value alone, by Value then General, by Formula, and its font's size set after it.
WAYS = {
    'value': '{size}c.Value = {value}',
    'value_general': '{size}c.Value = {value}: c.NumberFormat = "General"',
    'formula': '{size}c.Formula = "={value}"',
    'size_after': 'c.Value = {value}: c.Font.Size = {points}',
}
#: A number not in General whose font grows after its format, and a date: does a font's change widen them too?
FORMATTED = {
    'comma_size_after': 'c.Value = 1234.5: c.NumberFormat = "#,##0.00": c.Font.Size = {points}',
    'date_size_after': 'c.Value = DateSerial(2026, 10, 22): c.Font.Size = {points}',
    'comma_bold_after': 'c.Value = 1234.5: c.NumberFormat = "#,##0.00": c.Font.Size = {points}: c.Font.Bold = True',
    'general_bold_after': 'c.Value = {value}: c.Font.Bold = True',
}


def cases() -> list[tuple[str, str, str]]:
    """Each case's name, size and statements."""
    found: list[tuple[str, str, str]] = []
    for points in SIZES:
        for value in VALUES:
            for way, statements in WAYS.items():
                size = f'c.Font.Size = {points}: '
                found.append((f'{points}/{value}/{way}', str(points),
                              statements.format(size=size, value=value, points=points)))
        for way, statements in FORMATTED.items():
            found.append((f'{points}/1234567.891/{way}', str(points),
                          statements.format(points=points, value='1234567.891')))
    return found


def probe(chosen: list[tuple[str, str, str]]) -> str:
    """Each case in a column of its own: |name=the column's width in pixels~what the cell shows."""
    body = '\n'.join(f'Set c = ws.Cells(1, {column}): {statements}\n'
                     f's = s & "|{name}=" & CLng(ws.Columns({column}).Width * 4 / 3) & "~" & c.Text'
                     for column, (name, _, statements) in enumerate(chosen, 1))
    return f'''Function Probe() As String
Dim wb As Workbook, ws As Worksheet, c As Range, s As String
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
{body}
wb.Close False
Application.DisplayAlerts = True
Probe = Mid$(s, 2)
End Function
'''


def main() -> None:
    every = cases()
    # A probe a size, well inside what one procedure may hold.
    record({f'size_{points}': probe([case for case in every if case[1] == str(points)]) for points in SIZES}, OUT)


if __name__ == '__main__':
    main()
