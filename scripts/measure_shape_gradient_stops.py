"""GradientStops.Insert2, walking the stops with For Each, and the stops' Application and Creator: what each reads, the errors,
and the gradient each saved shape has.

    python scripts/measure_shape_gradient_stops.py

writes tests/fixtures/shape_gradient_stops.json, which tests/test_excel_shape_gradient_stops.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_gradient_stops.json'
MAKE = 'Set sh = ws.Shapes.AddShape(1, 10, 10, 100, 40): sh.Fill.ForeColor.RGB = 255: sh.Fill.BackColor.RGB = 65280: sh.Fill.TwoColorGradient 1, 1: Set st = sh.Fill.GradientStops: '
CASES = [
    's2 = st.Count', *[f'st.Insert2 {args}' for args in (
        '16711680, 0.5', '16711680, 0.5, 0.5', '16711680, 0.5, , 1', '16711680, 0.5, , 3', '16711680, 0.5, , 2', '16711680, 0.5, , 4', '16711680, 0.5, , 0', '16711680, 0.5, , , 0.4', '16711680, 0.5, , , -0.4',
        '16711680, 0.5, , , 0', '16711680, 0.5, , , 1', '16711680, 0.5, , , -1', '16711680, 0.5, , , 2', '16711680, 0.5, 0.5, 2, 0.4', '16711680, 2', '16711680, -1', '16711680, 0.5, 2', '16711680', '16711680, 0', '16711680, 1', 'RGB:=16711680, Position:=0.5, Brightness:=0.4')],
    'st.Insert 16711680, 0.5: st.Insert2 255, 0.25, , , 0.4', 'st.Insert2 16711680, 0.5, , , 0.4: st.Insert 255, 0.25', 'st.Insert2 16711680, 0.5, , , 0.4: s2 = st(3).Color.Brightness & "|" & st(3).Color.RGB & "|" & st(3).Color.Type & "|" & st(3).Color.TintAndShade',
    'st.Insert2 16711680, 0.5, , , 0.4: st.Delete 3', 'st.Insert2 16711680, 0.5, , , 0.4: st(3).Position = 0.75', 'st.Insert2 16711680, 0.5, , , 0.4: st(3).Color.RGB = 255', 'st.Insert2 16711680, 0.5, , , 0.4: st(3).Transparency = 0.5',
    'For Each o In st: s2 = s2 & o.Position & "/" & o.Color.RGB & ",": Next', 'For Each o In st: o.Transparency = 0.5: Next', 'st.Insert 16711680, 0.5: For Each o In st: s2 = s2 & TypeName(o) & o.Position & ",": Next',
    's2 = TypeName(st.Application) & "|" & st.Application.Name', 's2 = st.Creator', 's2 = TypeName(st(1).Application) & "|" & st(1).Application.Name', 's2 = st(1).Creator', 's2 = TypeName(st)',
    'Set o = ws.Shapes.AddShape(1, 1, 1, 5, 5).Fill.GradientStops: s2 = TypeName(o.Application)', 'Set o = ws.Shapes.AddShape(1, 1, 1, 5, 5).Fill.GradientStops: s2 = o.Creator',
    'Set o = ws.Shapes.AddShape(1, 1, 1, 5, 5).Fill.GradientStops: For Each v2 In o: s2 = "x": Next', 'Set o = ws.Shapes.AddShape(1, 1, 1, 5, 5).Fill.GradientStops: o.Insert2 255, 0.5',
]
READS = ['st.Count', *[f'st({place}).{one}' for place in (1, 2, 3, 4) for one in ('Position', 'Transparency', 'Color.RGB', 'Color.Brightness')], 's2']


def _probe() -> str:
    lines: list[str] = ['Private Sub Act(ByVal which As Long, ByVal ws As Worksheet, sh As Object, st As Object, o As Object, s2 As Variant, failure As Long)', '    Dim v2 As Variant',
                        '    On Error GoTo Failed', f'    {MAKE.rstrip(": ")}', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    Exit Sub', 'Failed:', '    failure = Err.Number', 'End Sub', '',
              'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, sh As Object, st As Object, o As Object, s As String, v As String, s2 As Variant, failure As Long',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Act which, ws, sh, st, o, s2, failure',
              '    s = IIf(failure = 0, "ok", "Err " & failure)', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', 'End Function',
              '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = []
        for index in range(len(CASES)):
            with zipfile.ZipFile(Path(folder) / f'c{index}.xlsx') as package:
                xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
            parts.append(re.findall(r'<a:gradFill\b.*?</a:gradFill>', xml, re.DOTALL)[0])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'reads': READS, 'source': source,
              'cases': [[statements, answer, part] for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
