"""ShapeRange's IncrementLeft, IncrementTop, IncrementRotation, Flip, ZOrder, Decorative and AutoShapeType: what each
shape reads after each case, and the errors.

Each case makes a workbook with a rectangle A (100, 50, 120 by 60), an oval B (250, 80, 60 by 40) and a rectangle C
(50, 200, 40 by 30), runs the case's statements on r, the range of A and B, and reads every shape's name, box,
rotation, flips, decorative, AutoShapeType and z-order position.

    python scripts/measure_shape_range_more.py

writes tests/fixtures/shape_range_more.json, which tests/test_excel_shape_range_more.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_range_more.json'
CASES = [
    'r.IncrementLeft 10', 'r.IncrementLeft -120', 'r.IncrementTop 15.5', 'r.IncrementTop -70',
    'r.IncrementRotation 30', 'r.IncrementRotation -400', 'ws.Shapes("A").Rotation = 10: r.IncrementRotation 5',
    'r.Flip msoFlipHorizontal', 'r.Flip msoFlipVertical', 'r.Flip 2',
    'ws.Shapes("A").Rotation = 30: r.Flip msoFlipHorizontal',
    'r.ZOrder msoSendToBack', 'r.ZOrder msoBringToFront', 'ws.Shapes.Range(Array("B", "C")).ZOrder msoSendToBack',
    'r.ZOrder msoBringForward', 'r.ZOrder msoSendBackward', 'ws.Shapes.Range(Array("A", "C")).ZOrder msoBringForward',
    'r.ZOrder 6',
    's = r.Decorative', 'r.Decorative = msoTrue', 'ws.Shapes("A").Decorative = msoTrue: s = r.Decorative',
    's = r.AutoShapeType', 's = ws.Shapes.Range(Array("A", "C")).AutoShapeType',
    'r.AutoShapeType = msoShapeOval', 'r.AutoShapeType = msoShapeRoundedRectangle',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, r As Object, sh As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeOval, 250, 80, 60, 40).Name = "B"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 50, 200, 40, 30).Name = "C"',
             '    Set r = ws.Shapes.Range(Array("A", "B"))',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select',
              '    For Each sh In ws.Shapes: s = s & "|" & sh.Name & "," & sh.Left & "," & sh.Top & "," & sh.Width & '
              '"," & sh.Height & "," & sh.Rotation & "," & sh.HorizontalFlip & "," & sh.VerticalFlip & "," & '
              'sh.Decorative & "," & sh.AutoShapeType & "," & sh.ZOrderPosition: Next',
              '    Tried = "ok " & s', 'Done:', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook',
              '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def geometries(path: Path) -> list[str]:
    """Each saved shape's prstGeom, in drawing order."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<a:prstGeom\b.*?</a:prstGeom>', xml, re.S)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [geometries(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case[-44:], '->', answer, found)


if __name__ == '__main__':
    main()
