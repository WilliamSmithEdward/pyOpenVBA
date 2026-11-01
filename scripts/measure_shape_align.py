"""ShapeRange.Align and ShapeRange.Distribute: where each shape goes, relative to the shapes or to the sheet, and the
errors they raise.

Each case makes a workbook with three rectangles, A (100, 50, 120 by 60), B (250, 80, 60 by 40) and C (50, 200,
40 by 30), runs the case's statements and reads every shape's name and box.

    python scripts/measure_shape_align.py

writes tests/fixtures/shape_align.json, which tests/test_excel_shape_align.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_align.json'
ALL = 'ws.Shapes.Range(Array("A", "B", "C"))'
TWO = 'ws.Shapes.Range(Array("A", "B"))'
CASES = [
    *(f'{ALL}.Align {command}, msoFalse' for command in range(6)),
    *(f'{ALL}.Align {command}, msoTrue' for command in range(6)),
    f'{TWO}.Align msoAlignCenters, msoFalse',
    f'ws.Shapes.Range(Array("B")).Align msoAlignLefts, msoFalse',
    f'ws.Shapes.Range(Array("B")).Align msoAlignRights, msoTrue',
    f'{ALL}.Align 6, msoFalse',
    f'{ALL}.Align msoAlignLefts, 5',
    f'{ALL}.Align msoAlignRights, 1',
    f'{ALL}.Align msoAlignRights, -2',
    f'{ALL}.Distribute msoDistributeVertically, 1',
    f'{ALL}.Distribute msoDistributeHorizontally, msoFalse',
    f'{ALL}.Distribute msoDistributeVertically, msoFalse',
    f'{ALL}.Distribute msoDistributeHorizontally, msoTrue',
    f'{ALL}.Distribute msoDistributeVertically, msoTrue',
    f'{TWO}.Distribute msoDistributeHorizontally, msoFalse',
    f'{TWO}.Distribute msoDistributeHorizontally, msoTrue',
    f'ws.Shapes.Range(Array("B")).Distribute msoDistributeVertically, msoTrue',
    f'{ALL}.Distribute 2, msoFalse',
    'ws.Shapes("C").Left = 260: ' + f'{ALL}.Distribute msoDistributeHorizontally, msoFalse',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 250, 80, 60, 40).Name = "B"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 50, 200, 40, 30).Name = "C"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select',
              '    For i = 1 To ws.Shapes.Count: s = s & "|" & ws.Shapes(i).Name & "," & ws.Shapes(i).Left & "," & '
              'ws.Shapes(i).Top & "," & ws.Shapes(i).Width & "," & ws.Shapes(i).Height: Next',
              '    Tried = "ok" & s', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
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
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case[-48:], '->', answer)


if __name__ == '__main__':
    main()
