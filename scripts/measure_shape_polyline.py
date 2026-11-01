"""Shapes.AddPolyline and Shapes.AddCurve: what the new shape answers, the errors, and the anchor the saved
workbook holds, with creation ids and cell corners cut out as scripts/measure_shape_duplicate.py records them.

Each case makes a workbook, runs the case's statements, where pts is a Single array of points the case fills, and
reads every shape's name, id, type, AutoShapeType and box, then saves it.

    python scripts/measure_shape_polyline.py

writes tests/fixtures/shape_polyline.json, which tests/test_excel_shape_polyline.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from measure_shape_group import anchors
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_polyline.json'


def points(*pairs: tuple[float, float]) -> str:
    """Statements that fill pts with these points."""
    filled = ': '.join(f'pts({index}, 1) = {x}: pts({index}, 2) = {y}' for index, (x, y) in enumerate(pairs, start=1))
    return f'ReDim pts(1 To {len(pairs)}, 1 To 2): {filled}'


CASES = [
    f'{points((10, 10), (110, 10), (60, 80))}: ws.Shapes.AddPolyline pts',
    f'{points((10, 10), (110, 10), (60, 80), (10, 10))}: ws.Shapes.AddPolyline pts',
    f'{points((50, 20), (150, 60))}: ws.Shapes.AddPolyline pts',
    f'{points((10, 10))}: ws.Shapes.AddPolyline pts',
    f'{points((10, 10), (110, 10), (60, 80))}: s = TypeName(ws.Shapes.AddPolyline(pts))',
    f'{points((10, 10), (40, 60), (80, 60), (110, 10))}: ws.Shapes.AddCurve pts',
    f'{points((10, 10), (40, 60), (80, 60), (110, 10), (140, -20), (170, 40), (200, 10))}: ws.Shapes.AddCurve pts',
    f'{points((10, 10), (40, 60), (80, 60))}: ws.Shapes.AddCurve pts',
    f'{points((10, 10), (110, 10), (60, 80))}: ws.Shapes.AddPolyline(pts).Name = "Tri"',
    f'{points((10.5, 10.25), (110.75, 10), (60, 80.5))}: ws.Shapes.AddPolyline pts',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, pts() As Single',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select',
              '    For Each sh In ws.Shapes: s = s & "|" & sh.Name & "," & sh.ID & "," & sh.Type & "," & '
              'sh.AutoShapeType & "," & sh.Left & "," & sh.Top & "," & sh.Width & "," & sh.Height: Next',
              '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        drawings = [anchors(Path(folder) / f'c{index}.xlsx') if answers[index].count('|') else []
                    for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, drawings, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case[-40:], '->', answer)
        for anchor in found:
            print('   ', anchor[:900])


if __name__ == '__main__':
    main()
