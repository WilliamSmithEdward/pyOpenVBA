"""Shape.Shadow, Glow, Reflection and SoftEdge: what an AutoShape, a text box and a line read before and after each change,
the errors, and the spPr each saved shape has.

Each case makes a workbook, adds the case's shape as sh, runs its statements, reads READS, each as its value or "E" and
the error number, and saves the workbook. The shape's spPr and style are recorded from the drawing part.

    python scripts/measure_shape_effects.py

writes tests/fixtures/shape_effects.json, which tests/test_excel_shape_effects.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_shape_fill_line import SHAPES, drawn

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_effects.json'
STATEMENTS = [
    '',
    'sh.Shadow.Visible = msoTrue', 'sh.Shadow.Visible = msoFalse', 'sh.Shadow.Visible = msoTrue: sh.Shadow.Visible = msoFalse',
    'sh.Shadow.ForeColor.RGB = RGB(255, 0, 0)', 'sh.Shadow.Visible = msoTrue: sh.Shadow.ForeColor.RGB = RGB(255, 0, 0)',
    'sh.Shadow.OffsetX = 5', 'sh.Shadow.OffsetY = 5', 'sh.Shadow.OffsetX = 5: sh.Shadow.OffsetY = 7', 'sh.Shadow.Visible = msoTrue: sh.Shadow.OffsetX = -3',
    'sh.Shadow.Blur = 4', 'sh.Shadow.Visible = msoTrue: sh.Shadow.Blur = 4', 'sh.Shadow.Transparency = 0.5', 'sh.Shadow.Visible = msoTrue: sh.Shadow.Transparency = 0.5',
    'sh.Shadow.Size = 120', 'sh.Shadow.Style = msoShadowStyleOuterShadow', 'sh.Shadow.Style = msoShadowStyleInnerShadow', 'sh.Shadow.Type = msoShadow21', 'sh.Shadow.Type = msoShadow1',
    'sh.Shadow.Obscured = msoTrue', 'sh.Shadow.RotateWithShape = msoFalse', 'sh.Shadow.IncrementOffsetX 2', 'sh.Shadow.IncrementOffsetY 2',
    'sh.Shadow.Blur = -1', 'sh.Shadow.Transparency = 2', 'sh.Shadow.Size = 0', 'sh.Shadow.Visible = 5',
    'sh.Glow.Radius = 5', 'sh.Glow.Color.RGB = RGB(255, 0, 0)', 'sh.Glow.Radius = 5: sh.Glow.Color.RGB = RGB(255, 0, 0)', 'sh.Glow.Radius = 5: sh.Glow.Transparency = 0.5',
    'sh.Glow.Transparency = 0.5', 'sh.Glow.Radius = 5: sh.Glow.Radius = 0', 'sh.Glow.Radius = -1', 'sh.Glow.Radius = 200', 'sh.Glow.Color.ObjectThemeColor = msoThemeColorAccent2',
    'sh.SoftEdge.Radius = 5', 'sh.SoftEdge.Type = msoSoftEdgeType1', 'sh.SoftEdge.Type = msoSoftEdgeType3', 'sh.SoftEdge.Type = msoSoftEdgeTypeNone', 'sh.SoftEdge.Radius = 5: sh.SoftEdge.Radius = 0', 'sh.SoftEdge.Radius = -1', 'sh.SoftEdge.Radius = 200',
    'sh.Reflection.Type = msoReflectionType1', 'sh.Reflection.Type = msoReflectionType5', 'sh.Reflection.Type = msoReflectionTypeNone', 'sh.Reflection.Transparency = 0.5', 'sh.Reflection.Size = 50', 'sh.Reflection.Offset = 3', 'sh.Reflection.Blur = 2',
    'sh.Reflection.Type = msoReflectionType1: sh.Reflection.Transparency = 0.2', 'sh.Reflection.Type = msoReflectionType1: sh.Reflection.Type = msoReflectionTypeNone',
    'sh.Shadow.Visible = msoTrue: sh.Glow.Radius = 5: sh.SoftEdge.Radius = 3: sh.Reflection.Type = msoReflectionType1',
    'sh.Glow.Radius = 5: sh.Shadow.Visible = msoTrue', 'sh.Fill.ForeColor.RGB = RGB(0, 255, 0): sh.Shadow.Visible = msoTrue', 'sh.Shadow.Visible = msoTrue: sh.Fill.ForeColor.RGB = RGB(0, 255, 0)',
    'sh.Line.Visible = msoFalse: sh.Glow.Radius = 5', 'sh.Glow.Radius = 5: sh.Line.Visible = msoFalse', 'sh.Fill.Visible = msoFalse: sh.Shadow.Visible = msoTrue',
]
STATEMENTS += [f'sh.Reflection.Type = {index}' for index in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, -1, -2)]
STATEMENTS += ['sh.Reflection.Transparency = 2', 'sh.Reflection.Size = 200', 'sh.Reflection.Offset = -1', 'sh.Reflection.Blur = -1', 'sh.Glow.Transparency = 2', 'sh.Shadow.Size = 300',
               'sh.Reflection.Size = 50: sh.Reflection.Blur = 2: sh.Reflection.Offset = 3: sh.Reflection.Transparency = 0.5', 'sh.SoftEdge.Radius = 2.5', 'sh.SoftEdge.Radius = 100', 'sh.Glow.Radius = 150', 'sh.Glow.Radius = 2.5',
               'sh.SoftEdge.Type = 2', 'sh.SoftEdge.Type = 4', 'sh.SoftEdge.Type = 5', 'sh.SoftEdge.Type = 6', 'sh.SoftEdge.Type = 7', 'sh.SoftEdge.Radius = 2.5: sh.SoftEdge.Radius = 5',
               'sh.Glow.Color.RGB = RGB(255, 0, 0): sh.Glow.Radius = 0', 'sh.Glow.Radius = 5: sh.Glow.Radius = 0: sh.Glow.Radius = 6', 'sh.Glow.Color.RGB = RGB(255, 0, 0): sh.Glow.Transparency = 0.25: sh.Glow.Radius = 5',
               'sh.Shadow.Visible = msoTrue: sh.Shadow.OffsetX = 5: sh.Shadow.OffsetY = 7: sh.Shadow.Blur = 4: sh.Shadow.Size = 120: sh.Shadow.Transparency = 0.5: sh.Shadow.ForeColor.RGB = RGB(255, 0, 0)',
               'sh.Shadow.Visible = msoTrue: sh.Shadow.OffsetX = 0: sh.Shadow.OffsetY = 0', 'sh.Shadow.Visible = msoTrue: sh.Shadow.Visible = msoFalse: sh.Shadow.Visible = msoTrue',
               'sh.Shadow.ForeColor.RGB = RGB(255, 0, 0): sh.Shadow.Visible = msoTrue', 'sh.Shadow.Blur = 4: sh.Shadow.Visible = msoTrue', 'sh.Glow.Radius = 5: sh.Shadow.Blur = 4', 'sh.Shadow.Blur = 4: sh.Glow.Radius = 5']
CASES = [('rect', statement) for statement in STATEMENTS]
CASES += [(shape, statement) for shape in ('text', 'line') for statement in ('', 'sh.Shadow.Visible = msoTrue', 'sh.Glow.Radius = 5', 'sh.SoftEdge.Radius = 5', 'sh.Reflection.Type = msoReflectionType1')]
READS = ['TypeName(sh.Shadow)', 'sh.Shadow.Visible', 'sh.Shadow.ForeColor.RGB', 'sh.Shadow.OffsetX', 'sh.Shadow.OffsetY', 'sh.Shadow.Blur', 'sh.Shadow.Transparency', 'sh.Shadow.Size', 'sh.Shadow.Style', 'sh.Shadow.Type',
         'sh.Shadow.Obscured', 'sh.Shadow.RotateWithShape', 'TypeName(sh.Shadow.Parent)', 'TypeName(sh.Glow)', 'sh.Glow.Radius', 'sh.Glow.Color.RGB', 'sh.Glow.Transparency', 'TypeName(sh.Glow.Parent)', 'TypeName(sh.SoftEdge)', 'sh.SoftEdge.Radius',
         'sh.SoftEdge.Type', 'TypeName(sh.Reflection)', 'sh.Reflection.Type', 'sh.Reflection.Transparency', 'sh.Reflection.Size', 'sh.Reflection.Offset', 'sh.Reflection.Blur']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String, head As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    head = "ok "',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (shape, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set sh = {SHAPES[shape]}']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = head & s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    head = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part] for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '->', answer)
        print('    ', re.sub(r'<a:xfrm>.*?</a:xfrm>|<a:prstGeom.*?</a:prstGeom>', '', part)[:500])


if __name__ == '__main__':
    main()
