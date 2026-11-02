"""Shape.Fill and Shape.Line: what an AutoShape, a text box and a line read before and after each change, the
errors, and the spPr and style each saved shape has.

Each case makes a workbook, adds the case's shape as sh, runs its statements, reads READS, each as its value or "E"
and the error number, and saves the workbook. The shape's spPr and style are recorded from the drawing part.

    python scripts/measure_shape_fill_line.py

writes tests/fixtures/shape_fill_line.json, which tests/test_excel_shape_fill_line.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_fill_line.json'
SHAPES = {'rect': 'ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 40)',
          'text': 'ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 100, 40)',
          'line': 'ws.Shapes.AddLine(10, 10, 110, 50)'}
CASES = [
    ('rect', ''), ('text', ''), ('line', ''),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.Visible = msoFalse'),
    ('rect', 'sh.Fill.Visible = msoFalse: sh.Fill.Visible = msoTrue'),
    ('rect', 'sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(0, 128, 255): sh.Fill.Transparency = 0.25'),
    ('rect', 'sh.Fill.Solid'),
    ('rect', 'sh.Line.Visible = msoFalse'),
    ('rect', 'sh.Line.ForeColor.RGB = RGB(0, 255, 0)'),
    ('rect', 'sh.Line.Weight = 3'),
    ('rect', 'sh.Line.Weight = 2.25: sh.Line.ForeColor.RGB = RGB(1, 2, 3)'),
    ('rect', 'sh.Line.DashStyle = msoLineDash'),
    ('rect', 'sh.Line.Transparency = 0.5'),
    ('text', 'sh.Fill.ForeColor.RGB = RGB(255, 255, 0)'),
    ('text', 'sh.Fill.Visible = msoFalse'),
    ('text', 'sh.Line.Visible = msoFalse'),
    ('text', 'sh.Line.Weight = 2'),
    ('line', 'sh.Line.ForeColor.RGB = RGB(255, 0, 0)'),
    ('line', 'sh.Line.Weight = 4'),
    ('line', 'sh.Line.DashStyle = msoLineRoundDot'),
    ('line', 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.ForeColor.RGB = -1'),
    ('rect', 'sh.Line.Weight = -1'),
    ('rect', 'sh.Fill.Transparency = 2'),
]
READS = [
    'sh.Fill.Visible', 'sh.Fill.ForeColor.RGB', 'sh.Fill.Transparency', 'sh.Fill.Type', 'sh.Fill.BackColor.RGB',
    'sh.Fill.ForeColor.ObjectThemeColor', 'sh.Fill.ForeColor.Type', 'sh.Line.Visible', 'sh.Line.ForeColor.RGB',
    'sh.Line.Weight', 'sh.Line.DashStyle', 'sh.Line.Style', 'sh.Line.Transparency',
    'sh.Line.ForeColor.ObjectThemeColor',
    'TypeName(sh.Fill)', 'TypeName(sh.Line)', 'TypeName(sh.Fill.ForeColor)', 'TypeName(sh.Fill.Parent)',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (shape, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set sh = {SHAPES[shape]}']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo Failed', '    Tried = "ok " & Mid(s, 2)', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function',
              '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def drawn(path: Path) -> str:
    """The first shape's spPr and style, as the drawing part holds them."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part]
                        for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '->', answer)
        print('    ', re.sub(r'<a:xfrm>.*?</a:xfrm>|<a:prstGeom.*?</a:prstGeom>', '', part)[:600])


if __name__ == '__main__':
    main()
