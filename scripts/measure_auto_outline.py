"""Measure Range.AutoOutline: which rows and columns it groups from the formulas that sum the cells beside them, at
what levels, and what it leaves of the sheet's Outline settings.

Each case runs on a new workbook, writes cells, calls AutoOutline and answers the outline levels of rows 1 to 12
and columns A to H, then Outline.SummaryRow and SummaryColumn.

    python scripts/measure_auto_outline.py

writes tests/fixtures/auto_outline.json, which tests/test_excel_auto_outline.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/auto_outline.json'
#: Each case: cells written (address, formula or value), then the statement that outlines.
CASES: list[tuple[list[tuple[str, str]], str]] = [
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=SUM(A1:A2)'), ('A4', '4'), ('A5', '5'), ('A6', '=SUM(A4:A5)'),
      ('A7', '=A3+A6')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=SUM(A1:A2)'), ('A4', '4'), ('A5', '5'), ('A6', '=SUM(A4:A5)'),
      ('A7', '=SUM(A3,A6)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '=SUM(A2:A4)'), ('A2', '1'), ('A3', '2'), ('A4', '3')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('B1', '2'), ('C1', '3'), ('D1', '=SUM(A1:C1)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('B1', '2'), ('C1', '=SUM(A1:B1)'), ('A2', '3'), ('B2', '4'), ('C2', '=SUM(A2:B2)'),
      ('A3', '=SUM(A1:A2)'), ('B3', '=SUM(B1:B2)'), ('C3', '=SUM(C1:C2)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=A1+A2+A3')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A5', '=SUM(A1:A3)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A2:A3)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)'), ('A6', '1'), ('A7', '2'), ('A8', '=SUM(A6:A7)')],
     'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)'), ('A6', '1'), ('A7', '2'), ('A8', '=SUM(A6:A7)')],
     'ws.Range("A1:A4").AutoOutline'),
    ([('A1', '"x"'), ('B1', '1'), ('A2', '"y"'), ('B2', '2'), ('A3', '"Total"'), ('B3', '=SUM(B1:B2)')],
     'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=AVERAGE(A1:A2)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=A1*2')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=A2')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)')],
     'ws.Rows("1:2").Group: ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)')],
     'ws.Range("A1").AutoOutline: ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)')], 'ws.Range("C5").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)'), ('B1', '1'), ('B2', '=SUM(B1)')],
     'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=SUM(A1:A2)'), ('B1', '1'), ('B2', '2'), ('B3', '3'), ('B4', '=SUM(B1:B3)')],
     'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=SUM(A1:A2)'), ('A4', '=SUM(A1:A3)')], 'ws.Range("A1").AutoOutline'),
    ([('A2', '1'), ('A3', '2'), ('A4', '=SUM(A2:A3)')], 'ws.Range("A1").AutoOutline'),
    ([('B1', '1'), ('B2', '2'), ('B3', '=SUM(B1:B2)')], 'ws.Range("A1").AutoOutline'),
    ([], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '=SUBTOTAL(9,A1:A2)')], 'ws.Range("A1").AutoOutline'),
    ([('A1', '1'), ('A2', '2'), ('A3', '3'), ('A4', '=SUM(A1:A3)')],
     'ws.Outline.SummaryRow = xlSummaryAbove: ws.Range("A1").AutoOutline'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed',
             '    Select Case which']
    for index, (cells, call) in enumerate(CASES):
        lines.append(f'    Case {index}')
        for address, content in cells:
            if content.startswith('='):
                lines.append(f'        ws.Range("{address}").Formula = "{content}"')
            else:
                lines.append(f'        ws.Range("{address}").Value = {content.replace(chr(34), chr(34))}')
        lines.append(f'        {call}')
    lines += ['    End Select', '    s = "rows="', '    For i = 1 To 12', '        s = s & ws.Rows(i).OutlineLevel',
              '    Next', '    s = s & " cols="', '    For i = 1 To 8', '        s = s & ws.Columns(i).OutlineLevel',
              '    Next', '    s = s & " below=" & (ws.Outline.SummaryRow = xlSummaryBelow) & " right=" & '
              '(ws.Outline.SummaryColumn = xlSummaryOnRight)', '    Tried = s', 'Done:', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description',
              '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (0, len(CASES) - 1), timeout=300.0)
    assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[cells, call, answer] for (cells, call), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (cells, call, answer) in enumerate(record['cases']):
        print(index, ' '.join(f'{a}{c}' for a, c in cells)[:70], '|', call[:40], '->', answer)


if __name__ == '__main__':
    main()
