"""ShapeRange.Fill and ShapeRange.Line: what a range reads when its shapes agree and differ, what a change does to each
shape, and what a change one of them cannot take leaves behind.

Each case adds a, b and c and takes the range the case names as sr, runs its statements, reads READS of the range and
EACH of each shape, each as its value (TypeName for an object) or "E" and the error number, and saves the workbook.
Each shape's spPr is recorded from the drawing part, without its place and geometry.

    python scripts/measure_shape_range_fill_line.py

writes tests/fixtures/shape_range_fill_line.json, which tests/test_excel_shape_range_fill_line.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_range_fill_line.json'
#: Which shapes the range holds: two rectangles, a rectangle and a text box, a rectangle and a line, one rectangle.
RANGES = {'rects': 'Array(a.Name, b.Name)', 'text': 'Array(a.Name, c.Name)', 'line': 'Array(a.Name, d.Name)',
          'one': 'Array(a.Name)', 'lineb': 'Array(d.Name, a.Name)', 'three': 'Array(a.Name, b.Name, c.Name)'}
CASES = [
    ('rects', ''), ('text', ''), ('line', ''), ('one', ''),
    ('rects', 'a.Fill.ForeColor.RGB = 255'),
    ('rects', 'a.Fill.ForeColor.RGB = 255: b.Fill.ForeColor.RGB = 255'),
    ('rects', 'a.Fill.Transparency = 0.5'),
    ('rects', 'a.Fill.Visible = msoFalse'),
    ('rects', 'a.Line.ForeColor.RGB = 255'),
    ('rects', 'a.Line.Weight = 3'),
    ('rects', 'a.Line.DashStyle = msoLineDash'),
    ('rects', 'a.Line.Visible = msoFalse'),
    ('rects', 'a.Line.Transparency = 0.5'),
    ('rects', 'sr.Fill.ForeColor.RGB = 255'),
    ('rects', 'sr.Fill.Transparency = 0.5'),
    ('rects', 'sr.Fill.Visible = msoFalse'),
    ('rects', 'sr.Fill.Solid'),
    ('rects', 'sr.Line.ForeColor.RGB = 255'),
    ('rects', 'sr.Line.Weight = 3'),
    ('rects', 'sr.Line.DashStyle = msoLineDash'),
    ('rects', 'sr.Line.Visible = msoFalse'),
    ('rects', 'sr.Line.Transparency = 0.5'),
    ('text', 'sr.Fill.ForeColor.RGB = 255'),
    ('text', 'sr.Fill.Visible = msoFalse'),
    ('text', 'sr.Line.Weight = 3'),
    ('line', 'sr.Fill.ForeColor.RGB = 255'),
    ('lineb', 'sr.Fill.ForeColor.RGB = 255'),
    ('line', 'sr.Fill.Visible = msoFalse'),
    ('line', 'sr.Fill.Visible = msoTrue'),
    ('line', 'sr.Line.Weight = 3'),
    ('line', 'sr.Line.ForeColor.RGB = 255'),
    ('one', 'sr.Fill.ForeColor.RGB = 255'),
    ('rects', 'Set kept = sr.Fill: a.Fill.ForeColor.RGB = 255: s0 = kept.ForeColor.RGB'),
    ('rects', 'Set kept = sr.Fill: a.Delete: s0 = kept.ForeColor.RGB'),
    ('rects', 'sr.Line.Weight = -1'),
    ('rects', 'sr.Fill.Transparency = 2'),
    # Hiding through a range leaves every shape after the first a bare noFill: what that shape then reads and does.
    ('rects', 'sr.Fill.Visible = msoFalse: sr.Fill.Visible = msoTrue'),
    ('rects', 'sr.Fill.Visible = msoFalse: b.Fill.Visible = msoTrue'),
    ('rects', 'sr.Fill.Visible = msoFalse: a.Fill.Visible = msoTrue'),
    ('rects', 'sr.Fill.Visible = msoFalse: b.Fill.ForeColor.RGB = 255'),
    ('rects', 'sr.Fill.Visible = msoFalse: b.Fill.Transparency = 0.5'),
    ('rects', 'sr.Fill.Visible = msoFalse: sr.Fill.ForeColor.RGB = 255'),
    ('rects', 'a.Fill.Visible = msoFalse: b.Fill.Visible = msoFalse'),
    ('rects', 'a.Fill.Visible = msoFalse: sr.Fill.Visible = msoFalse'),
    ('rects', 'b.Fill.Visible = msoFalse: sr.Fill.Visible = msoFalse'),
    ('rects', 'a.Fill.ForeColor.RGB = 255: sr.Fill.Visible = msoFalse'),
    ('rects', 'b.Fill.ForeColor.RGB = 255: sr.Fill.Visible = msoFalse'),
    ('rects', 'sr.Fill.ForeColor.RGB = 255: sr.Fill.Visible = msoFalse'),
    ('rects', 'sr.Fill.ForeColor.RGB = 255: sr.Fill.Visible = msoFalse: sr.Fill.Visible = msoTrue'),
    ('rects', 'sr.Fill.Transparency = 0.5: sr.Fill.Visible = msoFalse'),
    ('rects', 'sr.Line.Weight = 3: sr.Fill.Visible = msoFalse'),
    ('lineb', 'sr.Fill.Visible = msoFalse'),
    ('three', 'sr.Fill.Visible = msoFalse'),
    ('rects', 'sr.Line.Visible = msoFalse: sr.Line.Visible = msoTrue'),
    ('rects', 'a.Line.Visible = msoFalse: b.Line.Visible = msoFalse'),
    ('rects', 'a.Line.Visible = msoFalse: sr.Line.Visible = msoFalse'),
    ('rects', 'sr.Line.Visible = msoFalse: sr.Line.Weight = 3'),
    ('rects', 'sr.Line.Visible = msoFalse: sr.Fill.ForeColor.RGB = 255'),
    ('rects', 'sr.Fill.BackColor.RGB = 255'),
]
READS = [
    'sr.Fill.Visible', 'sr.Fill.ForeColor.RGB', 'sr.Fill.Transparency', 'sr.Fill.Type', 'sr.Fill.BackColor.RGB',
    'sr.Fill.ForeColor.ObjectThemeColor', 'sr.Fill.ForeColor.Type', 'sr.Line.Visible', 'sr.Line.ForeColor.RGB',
    'sr.Line.Weight', 'sr.Line.DashStyle', 'sr.Line.Style', 'sr.Line.Transparency',
    'sr.Line.ForeColor.ObjectThemeColor', 'sr.Line.ForeColor.Type', 'TypeName(sr.Fill)', 'TypeName(sr.Line)',
    'TypeName(sr.Fill.ForeColor)', 'TypeName(sr.Fill.Parent)', 'TypeName(sr.Line.Parent)',
    'TypeName(sr.Fill.Application)', 's0',
]
EACH = ['a.Fill.ForeColor.RGB', 'b.Fill.ForeColor.RGB', 'a.Line.Weight', 'b.Line.Weight', 'd.Line.Weight',
        'a.Fill.Visible', 'a.Fill.Transparency', 'a.Fill.Type', 'a.Fill.BackColor.RGB',
        'a.Fill.ForeColor.ObjectThemeColor', 'a.Fill.ForeColor.Type',
        'b.Fill.Visible', 'b.Fill.Transparency', 'b.Fill.Type', 'b.Fill.BackColor.RGB',
        'b.Fill.ForeColor.ObjectThemeColor', 'b.Fill.ForeColor.Type',
        'c.Fill.Visible', 'c.Fill.ForeColor.RGB', 'd.Fill.Visible', 'd.Fill.ForeColor.RGB']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, a As Object, b As Object, c As Object, d As Object',
             '    Dim sr As Object, kept As Object, s As String, v As String, s0 As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    Set a = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 40)',
             '    Set b = ws.Shapes.AddShape(msoShapeRectangle, 10, 60, 100, 40)',
             '    Set c = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 110, 100, 40)',
             '    Set d = ws.Shapes.AddLine(10, 160, 110, 200)', '    On Error GoTo Failed', '    Select Case which']
    for index, (shapes, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set sr = ws.Shapes.Range({RANGES[shapes]})']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS + EACH:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def drawn(path: Path) -> str:
    """Each shape's spPr, without its place and geometry, joined."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.findall(r'<xdr:spPr\b.*?</xdr:spPr>', xml, re.DOTALL)
    return '|'.join(re.sub(r'<a:xfrm\b.*?</a:xfrm>|<a:prstGeom\b.*?</a:prstGeom>|</?xdr:spPr[^>]*>', '', one)
                    for one in found)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS + EACH, 'source': source,
              'cases': [[f'{shapes}: {statements}', answer, part]
                        for (shapes, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '\n  =', answer, '\n  ', part[:420])


if __name__ == '__main__':
    main()
