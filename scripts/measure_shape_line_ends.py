"""LineFormat's arrowheads, back colour, pattern and InsetPen, and ColorFormat's Brightness, TintAndShade, SchemeColor
and ObjectThemeColor: what an AutoShape, a text box and a line read before and after each change, the errors, and the
spPr each saved shape has.

Each case makes a workbook, adds the case's shape as sh, runs its statements, reads READS, each as its value or "E"
and the error number, and saves the workbook. The shape's spPr and style are recorded from the drawing part.

    python scripts/measure_shape_line_ends.py

writes tests/fixtures/shape_line_ends.json, which tests/test_excel_shape_line_ends.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_line_ends.json'
SHAPES = {'rect': 'ws.Shapes.AddShape(1, 10, 10, 100, 40)',
          'text': 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)',
          'line': 'ws.Shapes.AddLine(10, 10, 110, 50)'}
ENDS = ('End', 'Begin')
CASES = [
    ('rect', ''), ('text', ''), ('line', ''),
    # Each style at each end, on a line.
    *[('line', f'sh.Line.{end}ArrowheadStyle = {style}') for end in ENDS for style in range(1, 7)],
    # Length and width, alone and after a style, in both orders.
    *[('line', f'sh.Line.{end}Arrowhead{part} = {size}') for end in ENDS for part in ('Length', 'Width')
      for size in (1, 2, 3)],
    *[('line', f'sh.Line.EndArrowheadStyle = 2: sh.Line.EndArrowhead{part} = {size}') for part in ('Length', 'Width')
      for size in (1, 3)],
    ('line', 'sh.Line.EndArrowheadLength = 3: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.EndArrowheadWidth = 1: sh.Line.EndArrowheadLength = 3: sh.Line.EndArrowheadStyle = 4'),
    ('line', 'sh.Line.BeginArrowheadStyle = 3: sh.Line.EndArrowheadStyle = 5'),
    ('line', 'sh.Line.BeginArrowheadStyle = 3: sh.Line.BeginArrowheadLength = 1: sh.Line.BeginArrowheadWidth = 3:'
             ' sh.Line.EndArrowheadStyle = 6: sh.Line.EndArrowheadLength = 3: sh.Line.EndArrowheadWidth = 1'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.EndArrowheadStyle = 1'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.EndArrowheadLength = 3: sh.Line.EndArrowheadStyle = 1'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.EndArrowheadLength = 3: sh.Line.EndArrowheadStyle = 1:'
             ' sh.Line.EndArrowheadStyle = 2'),
    # Odd values.
    *[('line', f'sh.Line.EndArrowheadStyle = {value}') for value in (0, 7, -2, 2.5)],
    *[('line', f'sh.Line.EndArrowheadLength = {value}') for value in (0, 4, -2)],
    *[('line', f'sh.Line.EndArrowheadWidth = {value}') for value in (0, 4, -2)],
    # With everything else a line takes, in both orders.
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.Weight = 3'),
    ('line', 'sh.Line.Weight = 3: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.ForeColor.RGB = RGB(255, 0, 0)'),
    ('line', 'sh.Line.ForeColor.RGB = RGB(255, 0, 0): sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.DashStyle = msoLineDash'),
    ('line', 'sh.Line.DashStyle = msoLineDash: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.DashStyle = msoLineRoundDot: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.Transparency = 0.5'),
    ('line', 'sh.Line.Transparency = 0.5: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.Visible = msoFalse'),
    ('line', 'sh.Line.EndArrowheadStyle = 2: sh.Line.Visible = msoFalse: sh.Line.Visible = msoTrue'),
    ('line', 'sh.Line.Visible = msoFalse: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.Visible = msoFalse: sh.Line.Visible = msoTrue: sh.Line.EndArrowheadStyle = 2'),
    # On shapes with a closed outline.
    ('rect', 'sh.Line.EndArrowheadStyle = 2'), ('rect', 'sh.Line.BeginArrowheadLength = 3'),
    ('text', 'sh.Line.EndArrowheadStyle = 2'), ('text', 'sh.Line.BeginArrowheadWidth = 3'),
    # Back colour and pattern.
    ('line', 'sh.Line.BackColor.RGB = RGB(1, 2, 3)'), ('rect', 'sh.Line.BackColor.RGB = RGB(1, 2, 3)'),
    ('text', 'sh.Line.BackColor.RGB = RGB(1, 2, 3)'),
    ('line', 'sh.Line.Pattern = 1'), ('line', 'sh.Line.Pattern = 22'), ('line', 'sh.Line.Pattern = 48'),
    ('rect', 'sh.Line.Pattern = 22'), ('text', 'sh.Line.Pattern = 22'),
    ('line', 'sh.Line.Pattern = 0'), ('line', 'sh.Line.Pattern = 99'), ('line', 'sh.Line.Pattern = -2'),
    ('line', 'sh.Line.Pattern = 22: sh.Line.BackColor.RGB = RGB(1, 2, 3)'),
    ('line', 'sh.Line.BackColor.RGB = RGB(1, 2, 3): sh.Line.Pattern = 22'),
    ('line', 'sh.Line.Pattern = 22: sh.Line.ForeColor.RGB = RGB(255, 0, 0): sh.Line.BackColor.RGB = RGB(1, 2, 3)'),
    ('line', 'sh.Line.Pattern = 22: sh.Line.ForeColor.RGB = RGB(255, 0, 0): sh.Line.Weight = 3'),
    ('line', 'sh.Line.Pattern = 22: sh.Line.EndArrowheadStyle = 2'),
    ('line', 'sh.Line.Pattern = 22: sh.Line.Visible = msoFalse'),
    # InsetPen.
    ('line', 'sh.Line.InsetPen = msoFalse'), ('line', 'sh.Line.InsetPen = msoTrue'),
    ('rect', 'sh.Line.InsetPen = msoFalse'), ('rect', 'sh.Line.InsetPen = msoTrue'),
    ('rect', 'sh.Line.InsetPen = msoTrue: sh.Line.InsetPen = msoFalse'),
    ('text', 'sh.Line.InsetPen = msoTrue'),
    # Colours: theme colour, then its brightness and tint, in both orders.
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('rect', 'sh.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('line', 'sh.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('text', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('text', 'sh.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('rect', 'sh.Fill.BackColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('rect', 'sh.Line.BackColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorHyperlink'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = 0'), ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = 99'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0): sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4'), ('rect', 'sh.Fill.ForeColor.Brightness = -0.25'),
    ('rect', 'sh.Line.ForeColor.Brightness = 0.4'), ('line', 'sh.Line.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 2'), ('rect', 'sh.Fill.ForeColor.Brightness = 0'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2: sh.Fill.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0): sh.Fill.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Fill.ForeColor.TintAndShade = 0.4'), ('rect', 'sh.Fill.ForeColor.TintAndShade = -0.25'),
    ('rect', 'sh.Line.ForeColor.TintAndShade = 0.4'), ('line', 'sh.Line.ForeColor.TintAndShade = 0.4'),
    ('rect', 'sh.Fill.ForeColor.TintAndShade = 2'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0): sh.Fill.ForeColor.TintAndShade = 0.4'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.ForeColor.TintAndShade = 0.2'),
    ('rect', 'sh.Fill.ForeColor.SchemeColor = 3'), ('rect', 'sh.Line.ForeColor.SchemeColor = 3'),
    ('line', 'sh.Line.ForeColor.SchemeColor = 3'), ('rect', 'sh.Fill.BackColor.SchemeColor = 3'),
    ('rect', 'sh.Fill.ForeColor.SchemeColor = 0'), ('rect', 'sh.Fill.ForeColor.SchemeColor = 99'),
    ('rect', 'sh.Fill.ForeColor.SchemeColor = 3: sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'),
    *[('rect', f'sh.Fill.ForeColor.SchemeColor = {number}') for number in (8, 20, 63, 64, 65, 80, 81, -1)],
    ('rect', 'sh.Fill.ForeColor.SchemeColor = 3: sh.Fill.ForeColor.Brightness = 0.4'),
    # Every pattern number and theme colour.
    *[('line', f'sh.Line.Pattern = {number}') for number in range(1, 64)],
    *[('rect', f'sh.Fill.ForeColor.ObjectThemeColor = {number}') for number in range(1, 18)],
    # The back colour of a fill, on each kind.
    ('rect', 'sh.Fill.BackColor.RGB = RGB(1, 2, 3)'), ('text', 'sh.Fill.BackColor.RGB = RGB(1, 2, 3)'),
    ('line', 'sh.Fill.BackColor.RGB = RGB(1, 2, 3)'), ('text', 'sh.Fill.BackColor.ObjectThemeColor = 6'),
    # What the odd writes do next.
    ('rect', 'sh.Line.BackColor.RGB = RGB(1, 2, 3): sh.Line.ForeColor.RGB = RGB(255, 0, 0)'),
    ('rect', 'sh.Line.BackColor.RGB = RGB(1, 2, 3): sh.Line.Weight = 3'),
    ('text', 'sh.Line.InsetPen = msoFalse'),
    ('rect', 'sh.Line.Pattern = 22: sh.Line.ForeColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Line.Pattern = 22: sh.Line.BackColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Line.ForeColor.RGB = RGB(255, 0, 0): sh.Line.Pattern = 22'),
    ('rect', 'sh.Line.ForeColor.ObjectThemeColor = 6: sh.Line.Pattern = 22'),
    ('rect', 'sh.Line.Pattern = 22: sh.Line.Pattern = 48'),
    ('rect', 'sh.Line.Pattern = 22: sh.Line.ForeColor.RGB = RGB(255, 0, 0): sh.Line.Pattern = 48'),
    ('rect', 'sh.Line.Pattern = 22: sh.Line.DashStyle = msoLineDash'),
    ('rect', 'sh.Line.Pattern = 22: sh.Line.Transparency = 0.5'),
    # Brightness and tint: both orders, again, and on the colours the style shades.
    ('rect', 'sh.Fill.ForeColor.TintAndShade = 0.2: sh.Fill.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.ForeColor.Brightness = -0.25'),
    ('rect', 'sh.Fill.ForeColor.TintAndShade = 0.4: sh.Fill.ForeColor.TintAndShade = -0.25'),
    ('rect', 'sh.Fill.ForeColor.TintAndShade = 0'), ('rect', 'sh.Fill.ForeColor.Brightness = 1'),
    ('rect', 'sh.Fill.ForeColor.Brightness = -1'), ('rect', 'sh.Fill.ForeColor.TintAndShade = 1'),
    ('rect', 'sh.Fill.ForeColor.TintAndShade = -1'),
    ('rect', 'sh.Line.ForeColor.Brightness = -0.25'), ('rect', 'sh.Line.ForeColor.TintAndShade = -0.25'),
    ('rect', 'sh.Line.ForeColor.TintAndShade = 0'), ('rect', 'sh.Line.ForeColor.Brightness = 0'),
    ('text', 'sh.Fill.ForeColor.Brightness = -0.25'), ('text', 'sh.Fill.ForeColor.TintAndShade = -0.25'),
    ('text', 'sh.Line.ForeColor.Brightness = 0.4'), ('text', 'sh.Line.ForeColor.TintAndShade = 0.4'),
    ('rect', 'sh.Fill.BackColor.Brightness = 0.4'), ('rect', 'sh.Fill.BackColor.TintAndShade = 0.4'),
    ('rect', 'sh.Line.BackColor.Brightness = 0.4'), ('rect', 'sh.Line.BackColor.TintAndShade = 0.4'),
    ('line', 'sh.Fill.ForeColor.Brightness = 0.4'), ('line', 'sh.Fill.ForeColor.TintAndShade = 0.4'),
    ('line', 'sh.Fill.ForeColor.ObjectThemeColor = 6'), ('line', 'sh.Fill.ForeColor.SchemeColor = 3'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(18, 52, 86): sh.Fill.ForeColor.Brightness = 0.25'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(18, 52, 86): sh.Fill.ForeColor.Brightness = -0.1'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(200, 180, 90): sh.Fill.ForeColor.Brightness = 0.6'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(200, 180, 90): sh.Fill.ForeColor.Brightness = -0.5'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(120, 120, 120): sh.Fill.ForeColor.Brightness = 0.3'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(18, 52, 86): sh.Fill.ForeColor.TintAndShade = 0.25'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(18, 52, 86): sh.Fill.ForeColor.TintAndShade = -0.1'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(200, 180, 90): sh.Fill.ForeColor.TintAndShade = 0.6'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(200, 180, 90): sh.Fill.ForeColor.TintAndShade = -0.5'),
    ('rect', 'sh.Fill.ForeColor.RGB = RGB(120, 120, 120): sh.Fill.ForeColor.TintAndShade = 0.3'),
    # With a hidden fill and line.
    ('rect', 'sh.Fill.Visible = msoFalse: sh.Fill.ForeColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.Visible = msoFalse: sh.Fill.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Line.Visible = msoFalse: sh.Line.ForeColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Line.Visible = msoFalse: sh.Line.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Line.ForeColor.ObjectThemeColor = 6: sh.Line.Visible = msoFalse'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = 6: sh.Fill.Visible = msoFalse'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.Visible = msoFalse'),
    ('rect', 'sh.Line.ForeColor.Brightness = 0.4: sh.Line.Visible = msoFalse'),
    ('rect', 'sh.Line.ForeColor.ObjectThemeColor = 6: sh.Line.Transparency = 0.5'),
    ('rect', 'sh.Fill.ForeColor.ObjectThemeColor = 6: sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.ForeColor.Brightness = 0.4: sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.Transparency = 0.5: sh.Fill.ForeColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.Transparency = 0.5: sh.Fill.ForeColor.Brightness = 0.4'),
]
READS = [
    'sh.Line.BeginArrowheadStyle', 'sh.Line.BeginArrowheadLength', 'sh.Line.BeginArrowheadWidth',
    'sh.Line.EndArrowheadStyle', 'sh.Line.EndArrowheadLength', 'sh.Line.EndArrowheadWidth',
    'sh.Line.Pattern', 'sh.Line.InsetPen', 'sh.Line.Style', 'sh.Line.DashStyle', 'sh.Line.Weight', 'sh.Line.Visible',
    'sh.Line.Transparency', 'sh.Line.ForeColor.RGB', 'sh.Line.ForeColor.ObjectThemeColor', 'sh.Line.ForeColor.Type',
    'sh.Line.ForeColor.Brightness', 'sh.Line.ForeColor.TintAndShade', 'sh.Line.ForeColor.SchemeColor',
    'sh.Line.BackColor.RGB', 'sh.Line.BackColor.ObjectThemeColor', 'sh.Line.BackColor.Type',
    'sh.Line.BackColor.Brightness', 'sh.Line.BackColor.TintAndShade', 'sh.Line.BackColor.SchemeColor',
    'sh.Fill.ForeColor.RGB', 'sh.Fill.ForeColor.ObjectThemeColor', 'sh.Fill.ForeColor.Type',
    'sh.Fill.ForeColor.Brightness', 'sh.Fill.ForeColor.TintAndShade', 'sh.Fill.ForeColor.SchemeColor',
    'sh.Fill.BackColor.RGB', 'sh.Fill.BackColor.ObjectThemeColor', 'sh.Fill.BackColor.Type',
    'sh.Fill.BackColor.Brightness', 'sh.Fill.BackColor.TintAndShade', 'sh.Fill.BackColor.SchemeColor',
    'sh.Fill.Type', 'sh.Fill.Visible',
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
        result = excel.run_vba(source, 'Probe', (folder,), timeout=400.0)
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
        print('    ', strip.sub('', part)[:700])


if __name__ == '__main__':
    main()
