"""How shapes and chart objects follow the cells they are placed over when column widths and row heights change: each
one's Left, Top, Width, Height and the cells under its corners, by its Placement.

Each case makes a workbook, adds a rectangle and a chart object, sets both to the case's Placement, then runs the
case's statements and reads both.

    python scripts/measure_shape_cell_moves.py

writes tests/fixtures/shape_cell_moves.json, which tests/test_excel_shape_cell_moves.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_cell_moves.json'
CHANGES = [
    '',
    'ws.Columns("A:B").ColumnWidth = 20',
    'ws.Columns("C").ColumnWidth = 30',
    'ws.Columns("E").ColumnWidth = 30',
    'ws.Columns("C").ColumnWidth = 2',
    'ws.Columns("C").ColumnWidth = 0',
    'ws.Columns("A:Z").ColumnWidth = 3',
    'ws.Rows("1:2").RowHeight = 40',
    'ws.Rows("3").RowHeight = 50',
    'ws.Rows("2").RowHeight = 3',
    'ws.Columns("B").Hidden = True',
    'ws.Rows("2").Hidden = True',
    'ws.Columns("C").AutoFit',
    'ws.StandardWidth = 12',
]
PLACEMENTS = ['xlMoveAndSize', 'xlMove', 'xlFreeFloating']
CASES = [(placement, change) for placement in PLACEMENTS for change in CHANGES]


def _probe() -> str:
    lines = ['Private Function Where(ByVal sh As Object) As String',
             '    Where = sh.Left & "," & sh.Top & "," & sh.Width & "," & sh.Height & "," & sh.TopLeftCell.Address(0, 0)'
             ' & "," & sh.BottomRightCell.Address(0, 0)',
             'End Function', '',
             'Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Shape, co As ChartObject',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 110, 20, 120, 40)',
             '    Set co = ws.ChartObjects.Add(70, 50, 150, 60)', '    On Error GoTo Failed', '    Select Case which']
    for index, (placement, change) in enumerate(CASES):
        lines += [f'    Case {index}', f'        sh.Placement = {placement}: co.Placement = {placement}']
        lines += [f'        {change}'] if change else []
    lines += ['    End Select', '    Tried = "ok " & Where(sh) & "|" & Where(co)', 'Done:', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
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
