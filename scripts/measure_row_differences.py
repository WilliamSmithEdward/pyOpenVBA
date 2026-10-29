"""Range.RowDifferences and ColumnDifferences: which cells Excel finds, over one grid and the ranges asked.

A grid of constants and formulas splits the rival readings: values against
formulas compared as R1C1, text by case or not, blanks against zero, a
number against text and logical values, a constant against a formula of
the same value, and errors. Each case asks a range with a comparison cell
and records the address found, or the error raised.

    python scripts/measure_row_differences.py

writes tests/fixtures/row_differences.json, which
tests/test_excel_row_differences.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/row_differences.json'

#: The grid, by cell: a constant or a formula as Formula writes it.
GRID = {
    'A1': '1', 'B1': '1', 'C1': '2', 'D1': '1',
    'A2': 'a', 'B2': 'a', 'C2': 'A', 'D2': 'b',
    'A3': '=A1', 'B3': '=B1', 'C3': '=C1', 'D3': '=A1',
    'B4': '1', 'D4': '0',
    'A5': '1', 'B5': "'1", 'C5': '1.0', 'D5': 'TRUE',
    'A6': '5', 'B6': '=2+3', 'C6': "'5", 'D6': '5',
    'A7': '=NA()', 'B7': '=NA()', 'C7': '#N/A', 'D7': '=1/0',
    'A8': '=$A$1', 'B8': '=$A$1', 'C8': '=A$1', 'D8': '=SUM(A1:A2)',
    'A9': '=A1*2', 'B9': '=a1*2', 'C9': '=A1 * 2', 'D9': '=(A1*2)',
    'B10': '0.333333333333333', 'D10': "'0.333333333333333",
    'B11': '46023', 'C11': "'1/1/2026",
    'A12': '=E12+1', 'B12': '=F12 + 1', 'C12': '=G12+1', 'D12': '=h12+1',
    'A13': '="a"', 'B13': '="A"', 'C13': "'=\"a\"", 'D13': '="a" ',
    'A14': '1', 'B14': "'1 ", 'C14': "'01", 'D14': '1E0',
}
#: Cells set through Value or given a number format after the grid: 1/3 exactly, a date, formats.
AFTER = [
    'ws.Range("A10").Value = 1 / 3', 'ws.Range("C10").Value = 1 / 3', 'ws.Range("C10").NumberFormat = "0.00"',
    'ws.Range("A11").Value = DateSerial(2026, 1, 1)', 'ws.Range("D11").Value = DateSerial(2026, 1, 1)',
    'ws.Range("D11").NumberFormat = "0"', 'ws.Range("D1").NumberFormat = "0.00"',
]

#: Each case: what is asked, as a VBA expression giving a Range.
CASES = [
    'Range("A1:D9").RowDifferences(Range("A1"))',
    'Range("A1:D9").RowDifferences(Range("C1"))',
    'Range("A1:D9").RowDifferences(Range("D3"))',
    'Range("A1:D9").ColumnDifferences(Range("A1"))',
    'Range("A1:D9").ColumnDifferences(Range("A2"))',
    'Range("A1:D9").ColumnDifferences(Range("B6"))',
    'Range("B2").RowDifferences(Range("B2"))',
    'Range("B2").ColumnDifferences(Range("B2"))',
    'Range("B1:D9").RowDifferences(Range("A1"))',
    'Range("A2:D3").ColumnDifferences(Range("A1"))',
    'Range("A1:B3,C5:D9").RowDifferences(Range("A1"))',
    'Range("A1:D9").RowDifferences(Range("B1:C1"))',
    'Range("A1:B1").RowDifferences(Range("A1"))',
    'Range("A:D").RowDifferences(Range("A1"))',
    'Range("1:9").ColumnDifferences(Range("A1"))',
    'Range("A1:D9").RowDifferences(Range("F1"))',
    'Range("A1:D9").RowDifferences(Range("A20"))',
    'Range("E1:F2").RowDifferences(Range("E1"))',
    'Range("A10:D14").RowDifferences(Range("A10"))',
    'Range("A10:D14").ColumnDifferences(Range("A10"))',
    'Range("A1:D9").ColumnDifferences(Range("A2:B3"))',
    'Range("A1:D9").ColumnDifferences(Range("A3:B2"))',
    'Range("A1:D9").RowDifferences(Range("B2:C3"))',
    'Range("A1:D9").RowDifferences(Range("A1,C1"))',
    'Range("A1:D9").RowDifferences(Range("C1,A1"))',
    'Range("A1:D9").RowDifferences(Range("B1:B2"))',
    'Range("A1:D9").RowDifferences(Range("A1:D1"))',
    'Range("A1:D9").RowDifferences(Range("D1:A1"))',
    'Range("A1:D9").ColumnDifferences(Range("A2:A3"))',
    'Range("A1:D9").ColumnDifferences(Range("A2:B2"))',
    'Range("A1:D9").ColumnDifferences(Range("B3:B2"))',
    'Range("A1:D9").RowDifferences(Range("C1:D1,A2"))',
    'Range("A1:D9").RowDifferences(Range("A1").Resize(1, 3))',
    'Range("A1:D9").ColumnDifferences(Range("C2:D9"))',
]


def _literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _probe() -> str:
    lines = ['Private Function Tried(ByVal ws As Worksheet, ByVal which As Long) As String',
             '    Dim r As Range', '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines.append(f'    Case {index}: Set r = ws.{case}')
    lines += ['    End Select', '    Tried = r.Address(False, False) & "~" & r.Areas.Count', '    Exit Function',
              'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description', 'End Function', '',
              'Public Function Probe() As String', '    Dim ws As Worksheet, s As String, i As Long',
              '    Set ws = ThisWorkbook.Worksheets.Add', '    ws.Activate']
    lines += [f'    ws.Range("{cell}").Formula = {_literal(value)}' for cell, value in GRID.items()]
    lines += [f'    {line}' for line in AFTER]
    # What each rival reads of the cells: FormulaR1C1, Text and Value.
    lines += ['    Dim c As Range', '    For Each c In ws.Range("A10:D14")',
              '        s = s & "|" & c.Address(False, False) & "=" & c.FormulaR1C1 & "~" & c.Text & "~" & CStr(c.Value)',
              '    Next', '    s = s & "#"']
    lines += [f'    For i = 0 To {len(CASES) - 1}', '        s = s & "|" & Tried(ws, i)', '    Next',
              '    Application.DisplayAlerts = False', '    ws.Delete', '    Application.DisplayAlerts = True',
              '    Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(_probe(), 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        readings, _, found = str(result.value).partition('#')
        answers = found.removeprefix('|').split('|')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'grid': GRID, 'after': AFTER,
              'readings': readings.strip('|').split('|'), 'cases': dict(zip(CASES, answers, strict=True))}
    OUT.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    for case, answer in record['cases'].items():
        print(f'{case}  ->  {answer}')


if __name__ == '__main__':
    main()
