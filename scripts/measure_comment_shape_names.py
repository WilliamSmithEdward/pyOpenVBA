"""How comments, form controls and shapes share a sheet's names and ids: each shape's name, id and type, comments'
boxes among them, after comments, buttons, rectangles and charts are added in different orders.

Each case makes a workbook, runs the case's additions on its sheet, then reads every shape the sheet's Shapes lists.

    python scripts/measure_comment_shape_names.py

writes tests/fixtures/comment_shape_names.json, which tests/test_excel_comment_shape_names.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/comment_shape_names.json'
NOTE = 'ws.Range("{cell}").AddComment "n"'
BOX = 'ws.Shapes.AddShape msoShapeRectangle, 10, 10, 40, 20'
CHART = 'ws.ChartObjects.Add 100, 10, 100, 60'
BUTTON = 'ws.Shapes.AddFormControl xlButtonControl, 10, 100, 60, 20'
CASES = [
    f'{NOTE.format(cell="A1")}: {BOX}',
    f'{BOX}: {NOTE.format(cell="A1")}',
    f'{NOTE.format(cell="A1")}: {CHART}',
    f'{NOTE.format(cell="A1")}: {NOTE.format(cell="B2")}: {BOX}: {CHART}',
    f'{BOX}: {NOTE.format(cell="A1")}: {BOX}: {NOTE.format(cell="B2")}',
    f'{NOTE.format(cell="A1")}: ws.Range("A1").Comment.Delete: {BOX}',
    f'{BOX}: {NOTE.format(cell="A1")}: ws.Range("A1").Comment.Delete: {NOTE.format(cell="B2")}: {BOX}',
    f'{NOTE.format(cell="A1")}: {BUTTON}: {BOX}: {NOTE.format(cell="B2")}',
    f'{BOX}: {BUTTON}: {NOTE.format(cell="A1")}: {BOX}',
    f'{BUTTON}: {BUTTON}: {NOTE.format(cell="A1")}',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Shape, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select',
              '    For Each sh In ws.Shapes: s = s & "|" & sh.Name & "," & sh.ID & "," & sh.Type: Next',
              '    Tried = "ok " & s', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & s', '    Resume Done', 'End Function', '',
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
        print(case[:90], '->', answer)


if __name__ == '__main__':
    main()
