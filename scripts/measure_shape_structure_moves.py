"""How shapes follow their cells when rows grow by themselves and when rows and columns are inserted or deleted: a
rectangle's Left, Top, Width, Height and the cells under its corners, by its Placement.

Each case makes a workbook, adds a rectangle over C2:E5, sets the case's Placement, runs the case's statements and
reads it.

    python scripts/measure_shape_structure_moves.py

writes tests/fixtures/shape_structure_moves.json, which tests/test_excel_shape_structure_moves.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_structure_moves.json'
CHANGES = [
    'ws.Range("A1").Font.Size = 30',
    'ws.Range("A3").Font.Size = 30',
    'ws.Range("A3").WrapText = True: ws.Range("A3").Value = "one two three four five six seven"',
    'ws.Rows(1).Insert',
    'ws.Rows(3).Insert',
    'ws.Rows("3:4").Insert',
    'ws.Rows(1).Delete',
    'ws.Rows(3).Delete',
    'ws.Rows("2:5").Delete',
    'ws.Columns(1).Insert',
    'ws.Columns(4).Insert',
    'ws.Columns(1).Delete',
    'ws.Columns(4).Delete',
    'ws.Columns("C:E").Delete',
    'ws.Range("A3").EntireRow.Insert',
    'ws.Range("D1").Insert xlShiftToRight',
    'ws.Range("D3").Insert xlShiftToRight',
    'ws.Range("D3").Insert xlShiftDown',
    'ws.Range("D3").Delete xlShiftUp',
    'ws.Range("A3:Z3").Insert xlShiftDown',
]
PLACEMENTS = ['xlMoveAndSize', 'xlMove', 'xlFreeFloating']
CASES = [(placement, change) for placement in PLACEMENTS for change in CHANGES]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Shape',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 110, 20, 120, 40)', '    On Error GoTo Failed',
             '    Select Case which']
    for index, (placement, change) in enumerate(CASES):
        lines += [f'    Case {index}', f'        sh.Placement = {placement}', f'        {change}']
    lines += ['    End Select', '    If ws.Shapes.Count = 0 Then',
              '        Tried = "ok gone"', '    Else',
              '        Tried = "ok " & sh.Left & "," & sh.Top & "," & sh.Width & "," & sh.Height & "," & '
              'sh.TopLeftCell.Address(0, 0) & "," & sh.BottomRightCell.Address(0, 0)', '    End If',
              'Done:', '    wb.Close False', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number',
              '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{placement}: {change}', answer] for (placement, change), answer in zip(CASES, answers,
                                                                                                  strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
