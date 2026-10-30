"""Application.Intersect of ranges with several areas (issue #257): the areas it answers, in what order, and whether
it joins, drops or repeats pieces.

Predictions written before measuring: every area of the first range against every area of the second, in that
nesting, each piece kept as it comes, overlaps repeated, nothing joined; Nothing where no piece is left.

    python scripts/measure_intersect_areas.py

writes tests/fixtures/intersect_areas.json, which tests/test_excel_intersect_areas.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/intersect_areas.json'
#: Each case: the ranges given to Application.Intersect, as addresses on one sheet.
CASES = [
    ['A1:A3,C1:C3', 'A1:C10'],
    ['A1:C10', 'A1:A3,C1:C3'],
    ['C1:C3,A1:A3', 'A1:C10'],
    ['A1:A3,C1:C3', 'C2:C5,A2:A5'],
    ['A1:A3,C1:C3', 'A2:C2,A3:C3'],
    ['A1:B2,B2:C3', 'A1:C3'],
    ['A1:C3', 'A1:B2,B2:C3'],
    ['A1:B5', 'A1:A2,A3:A5'],
    ['A1:A3,A5:A7', 'A2:A6'],
    ['A1:A3,C1:C3', 'E1:E3'],
    ['A1:A3,C1:C3', 'A1:C10', 'B1:C2'],
    ['A1:A3,C1:C3', 'A1:C10', 'A2:C2,C3'],
    ['A1:C3', 'B2'],
    ['A1:A3,C1:C3', 'A2,C2,A2'],
    ['A1,C1,B1', 'A1:C1'],
    ['A1:A3,A2', 'A1:A5'],
    ['A2,A1:A3', 'A1:A5'],
    ['A1:B1,A2:C2', 'A1:C2'],
    ['A1,B1,A2,B2', 'A1:B2'],
    ['A1,A3,A2', 'A1:A3'],
    ['A1,B2,B1,A2', 'A1:B2'],
    ['A1:A2,A2:A3', 'A1:A5'],
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal ws As Worksheet, ByVal which As Long) As String', '    Dim r As Range',
             '    On Error GoTo Failed', '    Select Case which']
    for index, addresses in enumerate(CASES):
        arguments = ', '.join(f'ws.Range("{address}")' for address in addresses)
        lines += [f'    Case {index}', f'        Set r = Application.Intersect({arguments})']
    lines += ['    End Select', '    If r Is Nothing Then',
              '        Tried = "Nothing"', '    Else',
              '        Tried = r.Address(False, False) & "~" & r.Areas.Count & "~" & r.Cells.CountLarge', '    End If',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description', 'End Function',
              '', 'Public Function Probe() As String', '    Dim s As String, i As Long, wb As Workbook',
              '    Set wb = Workbooks.Add(xlWBATWorksheet)', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "|" & Tried(wb.Worksheets(1), i)', '    Next', '    wb.Close False', '    Probe = Mid(s, 2)',
              'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[addresses, answer] for addresses, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for addresses, answer in record['cases']:
        print(' x '.join(addresses), '->', answer)


if __name__ == '__main__':
    main()
