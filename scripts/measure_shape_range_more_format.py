"""What a ShapeRange does with LineFormat's arrowheads, pattern and InsetPen, FillFormat's patterns and gradients,
ColorFormat's theme colour, brightness and tint, and with Adjustments and ConnectorFormat: reads where its shapes agree
and differ, sets, the errors, and each saved shape's spPr.

Each case makes two shapes on a sheet, ranges them as r, runs its statements, reads READS, and saves.

    python scripts/measure_shape_range_more_format.py

writes tests/fixtures/shape_range_more_format.json, which tests/test_excel_shape_range_more_format.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_range_more_format.json'
LINES = ('ws.Shapes.AddLine(10, 10, 110, 50)', 'ws.Shapes.AddLine(20, 70, 120, 110)')
RECTS = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddShape(1, 20, 70, 100, 40)')
ROUND = ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'ws.Shapes.AddShape(5, 20, 70, 100, 40)')
MIXED = ('ws.Shapes.AddShape(5, 10, 10, 100, 40)', 'ws.Shapes.AddLine(20, 70, 120, 110)')
ARROW = ('ws.Shapes.AddShape(33, 10, 10, 100, 40)', 'ws.Shapes.AddShape(5, 20, 70, 100, 40)')
CASES = [
    (LINES, ''), (RECTS, ''), (ROUND, ''), (MIXED, ''), (ARROW, ''),
    (LINES, 'r.Line.EndArrowheadStyle = 2'), (LINES, 'r.Line.BeginArrowheadLength = 3'),
    (LINES, 'r.Line.BeginArrowheadWidth = 1'),
    (LINES, 'a.Line.EndArrowheadStyle = 2'), (LINES, 'b.Line.EndArrowheadStyle = 3'),
    (MIXED, 'r.Line.EndArrowheadStyle = 2'), (RECTS, 'r.Line.EndArrowheadStyle = 2'),
    (LINES, 'r.Line.Pattern = 22'), (RECTS, 'r.Line.Pattern = 22'), (LINES, 'a.Line.Pattern = 22'),
    (LINES, 'r.Line.Pattern = 22: r.Line.BackColor.RGB = 255'), (LINES, 'r.Line.BackColor.RGB = 255'),
    (RECTS, 'r.Line.InsetPen = msoTrue'), (RECTS, 'a.Line.InsetPen = msoTrue'),
    (RECTS, 'r.Fill.ForeColor.ObjectThemeColor = 6'), (RECTS, 'r.Line.ForeColor.ObjectThemeColor = 6'),
    (RECTS, 'a.Fill.ForeColor.ObjectThemeColor = 6'),
    (RECTS, 'r.Fill.ForeColor.Brightness = 0.4'), (RECTS, 'a.Fill.ForeColor.Brightness = 0.4'),
    (RECTS, 'r.Fill.ForeColor.TintAndShade = 0.4'), (RECTS, 'a.Fill.ForeColor.TintAndShade = 0.4'),
    (RECTS, 'r.Line.ForeColor.Brightness = 0.4'), (RECTS, 'r.Fill.ForeColor.SchemeColor = 3'),
    (RECTS, 'r.Fill.Patterned 22'), (RECTS, 'a.Fill.Patterned 22'), (RECTS, 'r.Fill.Patterned 22: r.Fill.BackColor.RGB = 255'),
    (MIXED, 'r.Fill.Patterned 22'),
    (RECTS, 'r.Fill.TwoColorGradient 1, 1'), (RECTS, 'a.Fill.TwoColorGradient 1, 1'),
    (RECTS, 'r.Fill.OneColorGradient 1, 1, 0.25'), (RECTS, 'r.Fill.PresetGradient 1, 1, 7'),
    (RECTS, 'a.Fill.TwoColorGradient 1, 1: b.Fill.TwoColorGradient 2, 1'),
    (RECTS, 'r.Fill.TwoColorGradient 1, 1: r.Fill.GradientAngle = 30'),
    (RECTS, 'r.Fill.TwoColorGradient 1, 1: r.Fill.RotateWithObject = msoFalse'),
    (RECTS, 'r.Fill.TwoColorGradient 1, 1: r.Fill.GradientStops.Insert 255, 0.5'),
    (MIXED, 'r.Fill.TwoColorGradient 1, 1'),
    (ROUND, 'r.Adjustments(1) = 0.3'), (ROUND, 'a.Adjustments(1) = 0.3'), (ARROW, 'r.Adjustments(1) = 0.3'),
    (MIXED, 'r.Adjustments(1) = 0.3'), (RECTS, 'r.Adjustments(1) = 0.3'), (ROUND, 'r.Adjustments.Item(1) = 0.3'),
    (LINES, 'r.ConnectorFormat.Type = 2'), (RECTS, 'r.ConnectorFormat.Type = 2'),
]
READS = [
    'r.Line.BeginArrowheadStyle', 'r.Line.BeginArrowheadLength', 'r.Line.BeginArrowheadWidth', 'r.Line.EndArrowheadStyle',
    'r.Line.EndArrowheadLength', 'r.Line.EndArrowheadWidth', 'r.Line.Pattern', 'r.Line.InsetPen', 'r.Line.BackColor.RGB',
    'r.Line.BackColor.ObjectThemeColor', 'r.Line.BackColor.Type', 'r.Line.ForeColor.Brightness', 'r.Line.ForeColor.TintAndShade',
    'r.Line.ForeColor.ObjectThemeColor',
    'r.Fill.Type', 'r.Fill.Pattern', 'r.Fill.ForeColor.Brightness', 'r.Fill.ForeColor.TintAndShade', 'r.Fill.ForeColor.RGB',
    'r.Fill.BackColor.RGB', 'r.Fill.GradientColorType', 'r.Fill.GradientStyle', 'r.Fill.GradientVariant', 'r.Fill.GradientDegree',
    'r.Fill.GradientAngle', 'r.Fill.PresetGradientType', 'r.Fill.RotateWithObject', 'r.Fill.GradientStops.Count',
    'TypeName(r.Adjustments)', 'r.Adjustments.Count', 'r.Adjustments(1)', 'r.Adjustments(2)',
    'TypeName(r.ConnectorFormat)', 'r.ConnectorFormat.Type', 'r.ConnectorFormat.BeginConnected',
    'a.Line.EndArrowheadStyle', 'b.Line.EndArrowheadStyle', 'b.Fill.Type', 'b.Fill.ForeColor.Brightness', 'b.Adjustments(1)',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, a As Object, b As Object, r As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which']
    for index, ((first, second), _) in enumerate(CASES):
        lines += [f'    Case {index}: Set a = {first}: Set b = {second}']
    lines += ['    End Select', '    a.Name = "a": b.Name = "b"', '    Set r = ws.Shapes.Range(Array("a", "b"))',
              '    On Error GoTo Failed', '    Select Case which']
    for index, (_, statements) in enumerate(CASES):
        if statements:
            lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function',
              '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def drawn(path: Path) -> list[str]:
    """Each shape's spPr, as the drawing part holds it."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<xdr:spPr\b.*?</xdr:spPr>', xml, re.DOTALL)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[statements, answer, part] for (_, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
