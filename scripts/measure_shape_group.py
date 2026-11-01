"""ShapeRange.Group, Shape.Ungroup and GroupItems: what a sheet's shapes read after grouping and ungrouping, the
errors they raise, and the drawing the saved workbook holds.

Each case makes a workbook with three rectangles, A (100, 50, 120 by 60), B (250, 80, 60 by 40) and C (50, 200,
40 by 40), runs the case's statements and reads every top-level shape's name, id, type, box and, for a group, its
items' names, then saves it. The drawing is recorded with creation ids and anchor corners cut out, as
scripts/measure_shape_duplicate.py records it.

    python scripts/measure_shape_group.py

writes tests/fixtures/shape_group.json, which tests/test_excel_shape_group.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_group.json'
CASES = [
    's = TypeName(ws.Shapes.Range(Array("A", "B")).Group)',
    'ws.Shapes.Range(Array("A", "B")).Group',
    'ws.Shapes.Range(Array("B", "A")).Group',
    'ws.Shapes.Range(Array("A", "B", "C")).Group',
    'ws.Shapes.Range(Array("A")).Group',
    'ws.Shapes.Range(Array("A", "B")).Group.Name = "G"',
    'ws.Shapes.Range(Array("A", "B")).Group: s = TypeName(ws.Shapes("Group 4").Ungroup) & "," & '
    'ws.Shapes.Range(Array("A", "B")).Count',
    'ws.Shapes.Range(Array("A", "B")).Group: ws.Shapes("Group 4").Ungroup',
    'ws.Shapes.Range(Array("A", "B")).Group: s = ws.Shapes(1).GroupItems(1).Name & "," & '
    'ws.Shapes(1).GroupItems("B").Left & "," & ws.Shapes(1).GroupItems.Count',
    'ws.Shapes("A").Ungroup',
    's = ws.Shapes("C").GroupItems.Count',
    'ws.Shapes.Range(Array("A", "C")).Group',
    'ws.Shapes.Range(Array("B", "C")).Group',
    'ws.Shapes.Range(Array("A", "C")).Group: ws.Shapes("Group 4").Ungroup',
    'ws.Shapes.Range(Array("A", "B")).Group: ws.Shapes.AddShape(msoShapeOval, 10, 10, 20, 20).Name = "D"',
]
READ = ('    For i = 1 To ws.Shapes.Count: s = s & "|" & ws.Shapes(i).Name & "," & ws.Shapes(i).ID & "," & '
        'ws.Shapes(i).Type & "," & ws.Shapes(i).Left & "," & ws.Shapes(i).Top & "," & ws.Shapes(i).Width & "," & '
        'ws.Shapes(i).Height: If ws.Shapes(i).Type = msoGroup Then For j = 1 To ws.Shapes(i).GroupItems.Count: '
        's = s & ";" & ws.Shapes(i).GroupItems(j).Name: Next', '    Next')


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, s As String, i As Long, j As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 250, 80, 60, 40).Name = "B"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 50, 200, 40, 40).Name = "C"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', *READ, '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def anchors(path: Path) -> list[str]:
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
