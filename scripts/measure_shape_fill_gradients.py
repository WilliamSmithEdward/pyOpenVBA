"""FillFormat's gradients: OneColorGradient, TwoColorGradient, PresetGradient, GradientStops and what each reads, on an AutoShape, a text box and a line, with the errors and the spPr each saved shape has.

Each case makes a workbook, adds the case's shape as sh, runs its statements, reads READS, each as its value or "E"
and the error number, and saves the workbook. The shape's spPr and style are recorded from the drawing part.

    python scripts/measure_shape_fill_gradients.py

writes tests/fixtures/shape_fill_gradients.json, which tests/test_excel_shape_fill_gradients.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_fill_gradients.json'
SHAPES = {'rect': 'ws.Shapes.AddShape(1, 10, 10, 100, 40)',
          'text': 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)',
          'line': 'ws.Shapes.AddLine(10, 10, 110, 50)'}
RED = 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'
BLUE = 'sh.Fill.BackColor.RGB = RGB(0, 0, 255)'
CASES = [
    ('rect', ''), ('text', ''), ('line', ''),
    # Two colours.
    *[('rect', f'sh.Fill.TwoColorGradient {style}, {variant}') for style in (1, 2, 3, 4, 5, 6, 7)
      for variant in (1, 2, 3, 4)],
    ('rect', 'sh.Fill.TwoColorGradient 0, 1'), ('rect', 'sh.Fill.TwoColorGradient 8, 1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 0'), ('rect', 'sh.Fill.TwoColorGradient 1, 5'),
    ('rect', 'sh.Fill.TwoColorGradient -2, 1'),
    ('text', 'sh.Fill.TwoColorGradient 1, 1'), ('line', 'sh.Fill.TwoColorGradient 1, 1'),
    ('rect', f'{RED}: sh.Fill.TwoColorGradient 1, 1'), ('rect', f'{BLUE}: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', f'{RED}: {BLUE}: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', f'sh.Fill.TwoColorGradient 1, 1: {RED}'), ('rect', f'sh.Fill.TwoColorGradient 1, 1: {BLUE}'),
    ('rect', f'sh.Fill.TwoColorGradient 1, 1: {RED}: {BLUE}'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientAngle = 30'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Visible = msoFalse'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Solid'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Patterned 22'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.TwoColorGradient 1, 1'),
    # One colour, with its degree.
    *[('rect', f'sh.Fill.OneColorGradient {style}, 1, {degree}') for style in (1, 2, 3, 4, 5, 6, 7)
      for degree in (0.25, 1)],
    *[('rect', f'sh.Fill.OneColorGradient 1, {variant}, 0.5') for variant in (1, 2, 3, 4)],
    ('rect', 'sh.Fill.OneColorGradient 1, 1, 0'), ('rect', 'sh.Fill.OneColorGradient 1, 1, 2'),
    ('rect', 'sh.Fill.OneColorGradient 1, 1, -1'), ('rect', f'{RED}: sh.Fill.OneColorGradient 1, 1, 0.5'),
    # Preset gradients.
    *[('rect', f'sh.Fill.PresetGradient 1, 1, {preset}') for preset in range(1, 26)],
    ('rect', 'sh.Fill.PresetGradient 1, 1, 0'), ('rect', 'sh.Fill.PresetGradient 1, 1, 99'),
    *[('rect', f'sh.Fill.PresetGradient {style}, {variant}, 7') for style, variant in
      ((2, 1), (3, 1), (4, 1), (5, 1), (6, 1), (7, 1), (1, 2), (1, 3), (1, 4))],
    # Stops.
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Delete 1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Color.RGB = RGB(1, 2, 3)'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Position = 0.25'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Transparency = 0.5'),
    ('rect', 'sh.Fill.RotateWithObject = msoFalse'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.RotateWithObject = msoFalse'),
    # Degrees, and how a one-colour gradient's second colour follows.
    *[('rect', f'sh.Fill.OneColorGradient 1, 1, {degree}') for degree in (0.1, 0.4, 0.6, 0.75, 0.9)],
    ('rect', 'sh.Fill.OneColorGradient 1, 1, 0.25: sh.Fill.GradientDegree = 0.75'),
    ('rect', 'sh.Fill.OneColorGradient 1, 1, 0.25: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.OneColorGradient 1, 1, 0.25: sh.Fill.BackColor.RGB = RGB(0, 0, 255)'),
    # Where the colours come from.
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = 6: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.OneColorGradient 1, 1, 0.25'),
    ('rect', f'sh.Fill.Patterned 22: {BLUE}: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Fill.Transparency = 0.5: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Fill.Visible = msoFalse: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Line.Visible = msoFalse: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.TwoColorGradient 2, 1'),
    ('rect', f'sh.Fill.TwoColorGradient 1, 1: {BLUE}: sh.Fill.TwoColorGradient 2, 1'),
    ('rect', f'sh.Fill.TwoColorGradient 1, 3: {RED}'), ('rect', f'sh.Fill.TwoColorGradient 1, 3: {BLUE}'),
    ('rect', f'sh.Fill.TwoColorGradient 1, 4: {RED}'), ('rect', f'sh.Fill.TwoColorGradient 1, 4: {BLUE}'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.ForeColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.BackColor.ObjectThemeColor = 6'),
    # Leaving a gradient.
    ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.Solid'), ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.Patterned 22'),
    ('rect', 'sh.Fill.OneColorGradient 1, 1, 0.25: sh.Fill.Solid'),
    ('rect', 'sh.Fill.OneColorGradient 1, 1, 0.25: sh.Fill.Patterned 22'),
    ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.TwoColorGradient 1, 1'),
    ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.OneColorGradient 1, 1, 0.25'),
    ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.BackColor.RGB = RGB(0, 0, 255)'),
    # Presets across variants and centre styles, odd positions included.
    *[('rect', f'sh.Fill.PresetGradient 1, {variant}, {preset}') for preset in range(1, 25) for variant in (3, 4) if preset != 7],
    *[('rect', f'sh.Fill.PresetGradient 5, {variant}, 7') for variant in (2, 3, 4)],
    *[('rect', f'sh.Fill.PresetGradient 7, {variant}, 7') for variant in (2, 3, 4)],
    ('rect', 'sh.Fill.PresetGradient 2, 3, 7'), ('rect', 'sh.Fill.PresetGradient 1, 0, 7'),
    ('rect', 'sh.Fill.PresetGradient 1, 5, 7'), ('rect', 'sh.Fill.PresetGradient 0, 1, 7'),
    ('text', 'sh.Fill.PresetGradient 1, 1, 7'), ('line', 'sh.Fill.PresetGradient 1, 1, 7'),
    ('line', 'sh.Fill.OneColorGradient 1, 1, 0.5'),
    # The angle.
    ('rect', 'sh.Fill.GradientAngle = 30'),
    ('rect', 'sh.Fill.TwoColorGradient 5, 1: sh.Fill.GradientAngle = 30'),
    ('rect', 'sh.Fill.TwoColorGradient 7, 1: sh.Fill.GradientAngle = 30'),
    *[('rect', f'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientAngle = {angle}') for angle in (0, 359.9, 360, 400, -30, 30.5)],
    ('rect', 'sh.Fill.TwoColorGradient 1, 2: sh.Fill.GradientAngle = 30'),
    ('rect', 'sh.Fill.PresetGradient 1, 1, 7: sh.Fill.GradientAngle = 30'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStyle = 2'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientAngle = 30: sh.Fill.TwoColorGradient 2, 1'),
    # Stops.
    ('rect', 'sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5, 0.25'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5, 0, 1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5, 0, 2'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5, 0, 3'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5, 0, 9'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5, 0, -1, 0.4'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 2'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), -1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5: sh.Fill.GradientStops.Delete 1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5: sh.Fill.GradientStops.Delete 3'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5: sh.Fill.GradientStops.Delete 4'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5: sh.Fill.GradientStops.Delete'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops.Insert RGB(0, 255, 0), 0.5: sh.Fill.Solid'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(3).Position = 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(0).Position = 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Position = 2'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Transparency = 2'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(2).Color.RGB = RGB(1, 2, 3)'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Color.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(2).Color.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Color.Brightness = 0.4'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Color.TintAndShade = 0.4'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Color.RGB = RGB(1, 2, 3): sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(2).Position = 0.75: sh.Fill.TwoColorGradient 1, 3'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Position = 0.25: sh.Fill.GradientStops(2).Position = 0.1'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Visible = msoFalse: sh.Fill.Visible = msoTrue'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Visible = msoFalse: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Visible = msoFalse: sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Transparency = 0.5: sh.Fill.GradientStops(1).Transparency = 0.25'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.GradientStops(1).Transparency = 0.25: sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.Transparency = 0.5: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.RotateWithObject = msoTrue'), ('line', 'sh.Fill.RotateWithObject = msoFalse'),
    ('rect', 'sh.Fill.TwoColorGradient 1, 1: sh.Fill.RotateWithObject = msoFalse: sh.Fill.TwoColorGradient 2, 1'),
    ('rect', 'sh.Fill.PresetGradientType = 3'),
    *[('rect', f'sh.Fill.OneColorGradient 1, 1, {degree}') for degree in (0.01, 0.05, 0.15, 0.2, 0.3, 0.35, 0.45, 0.49, 0.51, 0.55, 0.65, 0.7, 0.8, 0.85, 0.95, 0.99)],
]
READS = [
    'sh.Fill.Type', 'sh.Fill.Visible', 'sh.Fill.Pattern', 'sh.Fill.Transparency', 'sh.Fill.ForeColor.RGB',
    'sh.Fill.ForeColor.ObjectThemeColor', 'sh.Fill.ForeColor.Type', 'sh.Fill.ForeColor.Brightness',
    'sh.Fill.BackColor.RGB', 'sh.Fill.BackColor.ObjectThemeColor', 'sh.Fill.BackColor.Type',
    'sh.Fill.GradientColorType', 'sh.Fill.GradientStyle', 'sh.Fill.GradientVariant', 'sh.Fill.GradientDegree',
    'sh.Fill.GradientAngle', 'sh.Fill.PresetGradientType', 'sh.Fill.RotateWithObject', 'sh.Fill.GradientStops.Count',
    'sh.Fill.GradientStops(1).Position', 'sh.Fill.GradientStops(1).Color.RGB',
    'sh.Fill.GradientStops(1).Color.ObjectThemeColor', 'sh.Fill.GradientStops(1).Color.Type',
    'sh.Fill.GradientStops(1).Color.Brightness', 'sh.Fill.GradientStops(1).Transparency',
    'sh.Fill.GradientStops(2).Position', 'sh.Fill.GradientStops(2).Color.RGB',
    'sh.Fill.GradientStops(2).Transparency', 'sh.Fill.GradientStops(3).Position',
    'TypeName(sh.Fill.GradientStops)', 'TypeName(sh.Fill.GradientStops(1))',
    'TypeName(sh.Fill.GradientStops(1).Color)', 'TypeName(sh.Fill.GradientStops.Parent)',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    Select Case which']
    for index, (shape, _) in enumerate(CASES):
        lines += [f'    Case {index}: Set sh = {SHAPES[shape]}']
    lines += ['    End Select', '    On Error GoTo Failed', '    Select Case which']
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
        result = excel.run_vba(source, 'Probe', (folder,), timeout=500.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part]
                        for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    strip = re.compile(r'<a:xfrm\b.*?</a:xfrm>|<a:prstGeom\b.*?</a:prstGeom>|<xdr:spPr[^>]*>|</xdr:spPr>|<xdr:style>.*')
    for case, answer, part in record['cases']:
        print(case, '->', answer)
        print('    ', strip.sub('', part)[:900])


if __name__ == '__main__':
    main()
