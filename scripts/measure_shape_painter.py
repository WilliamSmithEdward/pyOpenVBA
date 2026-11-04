"""Shape.PickUp and Apply, and a freeform's Vertices and Nodes: what each copies or reads, the errors, and each saved
shape's spPr and style.

    python scripts/measure_shape_painter.py

writes tests/fixtures/shape_painter.json, which tests/test_excel_shape_painter.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_painter.json'
RECTS = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddShape(1, 20, 70, 100, 40)')
TO_ROUND = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddShape(5, 20, 70, 100, 40)')
TO_LINE = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddLine(20, 70, 120, 110)')
FROM_LINE = ('ws.Shapes.AddLine(10, 10, 110, 50)', 'ws.Shapes.AddShape(1, 20, 70, 100, 40)')
TO_TEXT = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddTextbox(1, 20, 70, 100, 40)')
FROM_TEXT = ('ws.Shapes.AddTextbox(1, 10, 10, 100, 40)', 'ws.Shapes.AddShape(1, 20, 70, 100, 40)')
GO = 'a.PickUp: b.Apply'
LOOKS = [
    'a.Fill.ForeColor.RGB = 255', 'a.Fill.ForeColor.ObjectThemeColor = 6', 'a.Fill.ForeColor.Brightness = 0.4',
    'a.Fill.Transparency = 0.5', 'a.Fill.Visible = msoFalse', 'a.Fill.Patterned 22', 'a.Fill.TwoColorGradient 1, 1',
    'a.Line.ForeColor.RGB = 255', 'a.Line.Weight = 3', 'a.Line.DashStyle = msoLineDash', 'a.Line.Visible = msoFalse',
    'a.Line.Transparency = 0.5', 'a.Line.Pattern = 22', 'a.Line.InsetPen = msoTrue',
    'a.Shadow.Visible = msoTrue', 'a.Glow.Radius = 5', 'a.SoftEdge.Radius = 5', 'a.Width = 200',
    'a.Rotation = 30', 'a.Name = "x"', 'a.TextFrame.Characters.Text = "Hi"', 'a.TextFrame.MarginLeft = 3',
    'a.TextFrame.Characters.Font.Bold = True', 'a.Fill.ForeColor.RGB = 255: a.Line.Weight = 3',
]
POLY = ('Dim pts(1 To 3, 1 To 2) As Single: pts(1, 1) = 10: pts(1, 2) = 10: pts(2, 1) = 110: pts(2, 2) = 10: pts(3, 1) = 60: '
        'pts(3, 2) = 80: Set a = ws.Shapes.AddPolyline(pts)')
LINES = 'Set bld = ws.Shapes.BuildFreeform(1, 10, 10): bld.AddNodes 0, 0, 110, 10: bld.AddNodes 0, 0, 60, 80: Set a = bld.ConvertToShape'
CLOSED = LINES.replace('Set a = bld', 'bld.AddNodes 0, 0, 10, 10: Set a = bld')
CURVE = 'Set bld = ws.Shapes.BuildFreeform(1, 10, 10): bld.AddNodes 1, 1, 40, 60, 80, 60, 110, 10: Set a = bld.ConvertToShape'
CASES = [
    *[(RECTS, f'{look}: {GO}') for look in LOOKS],
    (RECTS, f'a.Fill.ForeColor.RGB = 255: b.Line.Weight = 3: {GO}'),
    (RECTS, f'b.Fill.ForeColor.RGB = 255: {GO}'), (RECTS, f'b.Line.Weight = 3: {GO}'),
    (RECTS, 'b.Apply'), (RECTS, GO),
    (RECTS, f'a.Fill.ForeColor.RGB = 255: a.PickUp: a.Fill.ForeColor.RGB = 65280: b.Apply'),
    (TO_ROUND, f'a.Fill.ForeColor.RGB = 255: {GO}'), (TO_LINE, f'a.Fill.ForeColor.RGB = 255: a.Line.Weight = 3: {GO}'),
    (FROM_LINE, f'a.Line.Weight = 3: {GO}'), (TO_TEXT, f'a.Fill.ForeColor.RGB = 255: a.Line.Weight = 3: {GO}'),
    (TO_TEXT, GO), (FROM_TEXT, f'a.Fill.ForeColor.RGB = 255: {GO}'), (FROM_TEXT, GO),
    (RECTS, 'a.Fill.ForeColor.RGB = 255: a.PickUp: ws.Shapes.Range(Array(1, 2)).Apply'),
    (RECTS, 'ws.Shapes.Range(Array(1)).PickUp: b.Apply'),
    (RECTS, 'a.Fill.ForeColor.RGB = 255: ws.Shapes.Range(Array(1, 2)).PickUp: b.Apply'),
]
FREEFORMS = [POLY, LINES, CLOSED, CURVE, 'Set a = ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'Set a = ws.Shapes.AddLine(10, 10, 110, 50)',
             'Set a = ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'Set a = ws.Shapes.AddTextbox(1, 10, 10, 100, 40)']
READS = ['a.Fill.ForeColor.RGB', 'b.Fill.ForeColor.RGB', 'b.Fill.Type', 'b.Fill.Visible', 'b.Fill.Transparency', 'b.Line.ForeColor.RGB',
         'b.Line.Weight', 'b.Line.Visible', 'b.Line.DashStyle', 'b.Width', 'b.Rotation', 'b.Name', 'b.AutoShapeType',
         'b.Shadow.Visible', 'b.Glow.Radius', 'b.SoftEdge.Radius', 'b.TextFrame.MarginLeft', 'b.TextFrame.Characters.Text']
NODE_READS = ['TypeName(a.Nodes)', 'a.Nodes.Count', 'TypeName(a.Nodes(1))', 'a.Nodes(1).SegmentType', 'a.Nodes(1).EditingType',
              'Pt(a.Nodes(1).Points)', 'Pt(a.Nodes(2).Points)', 'a.Nodes(2).SegmentType', 'a.Nodes(2).EditingType',
              'Pt(a.Nodes(a.Nodes.Count).Points)', 'a.Nodes(a.Nodes.Count + 1).SegmentType', 'a.Nodes(0).SegmentType',
              'TypeName(a.Vertices)', 'Pt(a.Vertices)', 'a.Width', 'a.Height']


def _probe() -> str:
    lines = ['Private Function Pt(p As Variant) As String',
             '    Dim i As Long, s As String', '    s = LBound(p, 1) & "-" & UBound(p, 1) & ":"',
             '    For i = LBound(p, 1) To UBound(p, 1): s = s & p(i, 1) & "/" & p(i, 2) & " ": Next', '    Pt = s', 'End Function', '',
             'Private Function Painted(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, a As Object, b As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which']
    for index, ((first, second), _) in enumerate(CASES):
        lines += [f'    Case {index}: Set a = {first}: Set b = {second}']
    lines += ['    End Select', '    On Error GoTo Failed', '    Select Case which']
    for index, (_, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Painted = s',
              '    wb.SaveAs folder & "\\p" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function', '',
              'Private Function Noded(ByVal which As Long) As String',
              '    Dim wb As Workbook, ws As Worksheet, a As Object, bld As Object, s As String, v As String',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which']
    for index, statements in enumerate(FREEFORMS):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    s = "ok"', '    On Error Resume Next']
    for read in NODE_READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Noded = s', '    wb.Close False', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Painted(i, folder)', '    Next',
              f'    For i = 0 To {len(FREEFORMS) - 1}', '        s = s & "@@@" & Noded(i)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def drawn(path: Path) -> list[str]:
    """Each shape's spPr, style and text, without its names and place."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?(?:<xdr:txBody>.*?</xdr:txBody>)?', xml, re.DOTALL)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'p{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'node_reads': NODE_READS, 'source': source,
              'cases': [[statements, answer, part] for (_, statements), answer, part in zip(CASES, answers[:len(CASES)], parts, strict=True)],
              'freeforms': [[statements, answer] for statements, answer in zip(FREEFORMS, answers[len(CASES):], strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
