"""Shape.Flip, HorizontalFlip, VerticalFlip, ZOrder and ZOrderPosition: what the shapes read after each change, the
errors bad values raise, and the shapes' order and transforms in the saved drawing.

Each case makes a workbook with three rectangles, A, B and C, added in that order, runs the case's statements and
reads each shape's flips and ZOrderPosition, then saves it. The saved drawing is recorded as each anchor's shape
name and its xfrm element.

    python scripts/measure_shape_flip_order.py

writes tests/fixtures/shape_flip_order.json, which tests/test_excel_shape_flip_order.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_flip_order.json'
READ = ('"|" & ws.Shapes("A").HorizontalFlip & "," & ws.Shapes("A").VerticalFlip & "," & '
        'ws.Shapes("A").ZOrderPosition & "," & ws.Shapes("B").ZOrderPosition & "," & ws.Shapes("C").ZOrderPosition & '
        '"," & ws.Shapes(1).Name & ws.Shapes(2).Name & ws.Shapes(3).Name')
CASES = [
    '',
    'ws.Shapes("A").Flip msoFlipHorizontal',
    'ws.Shapes("A").Flip msoFlipVertical',
    'ws.Shapes("A").Flip msoFlipHorizontal: ws.Shapes("A").Flip msoFlipHorizontal',
    'ws.Shapes("A").Flip msoFlipHorizontal: ws.Shapes("A").Flip msoFlipVertical',
    'ws.Shapes("A").Flip 2',
    'ws.Shapes("A").ZOrder msoBringToFront',
    'ws.Shapes("C").ZOrder msoSendToBack',
    'ws.Shapes("A").ZOrder msoBringForward',
    'ws.Shapes("C").ZOrder msoSendBackward',
    'ws.Shapes("B").ZOrder msoBringToFront: ws.Shapes("A").ZOrder msoBringForward',
    'ws.Shapes("C").ZOrder msoBringForward',
    'ws.Shapes("A").ZOrder msoSendBackward',
    'ws.Shapes("A").ZOrder 9',
    'ws.Shapes("A").Flip msoFlipHorizontal: ws.Shapes("A").Rotation = 30',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 60, 30).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 40, 20, 60, 30).Name = "B"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 70, 30, 60, 30).Name = "C"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {case}'] if case else [])
    lines += ['    End Select', f'    Tried = "ok" & {READ}', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def drawing(path: Path) -> list[list[str]]:
    """Each anchor's shape name and xfrm element, in the drawing's order."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return [[re.search(r'<xdr:cNvPr\b[^>]*\bname="([^"]*)"', anchor).group(1),  # type: ignore[union-attr]
             re.search(r'<a:xfrm\b[^>]*>', anchor).group()]  # type: ignore[union-attr]
            for anchor in re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        drawings = [drawing(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, drawings, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case, '->', answer, '|', ' '.join(f'{name}{xfrm[7:]}' for name, xfrm in found))


if __name__ == '__main__':
    main()
