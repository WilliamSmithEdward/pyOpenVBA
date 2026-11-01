"""ChartObject's Visible, Locked, PrintObject, Placement, BringToFront and SendToBack, and ChartObjects' Left, Top,
Width, Height, Visible, Locked, PrintObject and Placement: what they answer, the errors, the shapes' order, and each
saved anchor's opening tag, cNvPr and clientData.

Each case makes a workbook with a rectangle A (10, 10, 50 by 30), an empty chart (100, 50, 200 by 120) and a
rectangle B (20, 200, 50 by 30), runs the case's statements with co as the chart, and reads every shape's name and
the chart's settings, then saves it.

    python scripts/measure_chart_object_more.py

writes tests/fixtures/chart_object_more.json, which tests/test_excel_chart_object_more.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_object_more.json'
CASES = [
    '', 'co.Visible = False', 'co.Visible = False: co.Visible = True', 'co.Locked = False', 'co.PrintObject = False',
    'co.Placement = xlMove', 'co.Placement = xlFreeFloating', 'co.Placement = xlMoveAndSize', 'co.Placement = 5',
    's = TypeName(co.BringToFront)', 'co.SendToBack', 's = TypeName(co.SendToBack)', 'co.SendToBack: co.BringToFront',
    'ws.ChartObjects.Left = 10', 'ws.ChartObjects.Top = 15.5', 'ws.ChartObjects.Width = 300',
    'ws.ChartObjects.Height = 90', 'ws.ChartObjects.Visible = False', 'ws.ChartObjects.Locked = False',
    'ws.ChartObjects.PrintObject = False', 'ws.ChartObjects.Placement = xlFreeFloating',
    's = ws.ChartObjects.Left',
]
READ = ('    For Each sh In ws.Shapes: s = s & "|" & sh.Name: Next: s = s & "|" & co.Left & "," & co.Top & "," & '
        'co.Width & "," & co.Height & "," & co.Visible & "," & co.Locked & "," & co.PrintObject & "," & co.Placement')


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, co As Object, sh As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 30).Name = "A"',
             '    Set co = ws.ChartObjects.Add(100, 50, 200, 120)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 20, 200, 50, 30).Name = "B"',
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


def anchor_settings(path: Path) -> list[str]:
    """Each anchor's opening tag, its cNvPr's opening tag and its clientData, in drawing order."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found: list[str] = []
    for anchor in re.findall(r'<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>', xml, re.S):
        parts = [re.search(pattern, anchor) for pattern in
                 (r'<xdr:twoCellAnchor\b[^>]*>', r'<xdr:cNvPr\b[^>]*>', r'<xdr:clientData\b[^>]*/>')]
        found.append(''.join(part.group(0) for part in parts if part))
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [anchor_settings(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case[-40:], '->', answer, [re.sub(r'<xdr:cNvPr id="(\d+)" name="([^"]*)"', r'#\1 \2', one)[:160]
                                         for one in found])


if __name__ == '__main__':
    main()
