"""Shape.Duplicate and ShapeRange.Duplicate: what they answer, the copies' names, ids, places and settings, and the
anchors the saved drawing holds.

Each case makes a workbook with a rectangle, A (100, 50, 120 by 60), runs the case's statements and reads every
shape's name, id, left, top, width and height, then saves it. The drawing is recorded anchor by anchor, each
anchor's creation id cut out, as Excel makes it at random.

    python scripts/measure_shape_duplicate.py

writes tests/fixtures/shape_duplicate.json, which tests/test_excel_shape_duplicate.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_duplicate.json'
CASES = [
    's = TypeName(ws.Shapes("A").Duplicate)',
    's = ws.Shapes("A").Duplicate.Name',
    'ws.Shapes("A").Duplicate: ws.Shapes("A").Duplicate',
    'ws.Shapes("A").Duplicate: ws.Shapes(2).Duplicate',
    'ws.Shapes("A").AlternativeText = "d": ws.Shapes("A").Rotation = 30: ws.Shapes("A").Flip msoFlipHorizontal: '
    'ws.Shapes("A").Duplicate: s = ws.Shapes(2).AlternativeText & "," & ws.Shapes(2).Rotation & "," & '
    'ws.Shapes(2).HorizontalFlip',
    'ws.Shapes("A").Left = 5: ws.Shapes("A").Top = 3: ws.Shapes("A").Duplicate',
    'ws.Shapes.AddShape(msoShapeOval, 300, 50, 40, 40).Name = "B": '
    's = TypeName(ws.Shapes.Range(Array("A", "B")).Duplicate)',
    'ws.Shapes("A").Name = "Box": ws.Shapes("Box").Duplicate',
    'ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 200, 80, 30).TextFrame.Characters.Text = "hi": '
    'ws.Shapes(2).Duplicate',
    'ws.Shapes("A").Rotation = 30: ws.Shapes("A").Flip msoFlipVertical: s = ws.Shapes("A").Rotation',
    'ws.Shapes("A").Rotation = 30: ws.Shapes("A").Flip msoFlipHorizontal: ws.Shapes("A").Flip msoFlipHorizontal: '
    's = ws.Shapes("A").Rotation',
    'ws.Shapes("A").Flip msoFlipHorizontal: s = ws.Shapes("A").Rotation',
]
READ = ('    For i = 1 To ws.Shapes.Count: s = s & "|" & ws.Shapes(i).Name & "," & ws.Shapes(i).ID & "," & '
        'ws.Shapes(i).Left & "," & ws.Shapes(i).Top & "," & ws.Shapes(i).Width & "," & ws.Shapes(i).Height: Next')


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60).Name = "A"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', READ, '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def anchors(path: Path) -> list[str]:
    """Each anchor with its creation id and its cell corners cut out. The corners follow the columns and rows as
    the measuring machine's display draws them, which at 144 DPI are not the 96 DPI ones the model keeps; the
    shape's own transform, in points, stays."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S)
    return [re.sub(r'<xdr:from>.*?</xdr:to>', '',
                   re.sub(r'(<a16:creationId\b[^>]*\bid=")\{[0-9A-F-]{36}\}', r'\1{ID}', anchor), flags=re.S)
            for anchor in found]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        drawings = [anchors(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, drawings, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case[:60], '->', answer)


if __name__ == '__main__':
    main()
