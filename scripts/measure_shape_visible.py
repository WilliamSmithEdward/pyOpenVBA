"""Shape.Visible and ShapeRange.Visible: what a shape reads after each setting, the errors bad values raise, the
cNvPr the saved drawing holds, and what the shape reads once the workbook is opened again.

Each case makes a workbook with two rectangles, A and B, runs the case's statements, reads both shapes' Visible,
saves it, opens it again and reads them once more.

    python scripts/measure_shape_visible.py

writes tests/fixtures/shape_visible.json, which tests/test_excel_shape_visible.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_visible.json'
CASES = [
    '',
    'ws.Shapes("A").Visible = msoFalse',
    'ws.Shapes("A").Visible = False',
    'ws.Shapes("A").Visible = msoFalse: ws.Shapes("A").Visible = msoTrue',
    'ws.Shapes("A").Visible = 0',
    'ws.Shapes("A").Visible = 5',
    'ws.Shapes("A").Visible = "x"',
    'ws.Shapes.Range(Array("A", "B")).Visible = msoFalse',
    'ws.Shapes("A").Visible = msoFalse: s = s & "|" & ws.Shapes.Range(Array("A", "B")).Visible',
    'ws.Shapes("A").Visible = msoFalse: ws.Shapes("A").AlternativeText = "d"',
]


def _probe() -> str:
    read = '"|" & ws.Shapes("A").Visible & "," & ws.Shapes("B").Visible'
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String, path As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 60, 30).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 40, 80, 20).Name = "B"',
             '    path = folder & "\\c" & which & ".xlsx"', '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {case}'] if case else [])
    lines += ['    End Select', f'    s = "ok" & s & {read}', 'Saved:', '    wb.SaveAs path, xlOpenXMLWorkbook',
              '    wb.Close False', '    Set wb = Workbooks.Open(path)', '    Set ws = wb.Worksheets(1)',
              f'    s = s & {read}', '    wb.Close False', '    Tried = s', '    Exit Function', 'Failed:',
              '    s = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def identities(path: Path) -> list[str]:
    """Each shape's cNvPr opening tag, its creation id cut."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<xdr:cNvPr\b[^>]*>', xml)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        tags = [identities(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, tags, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case, '->', answer, '|', found)


if __name__ == '__main__':
    main()
