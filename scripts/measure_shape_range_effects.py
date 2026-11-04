"""A ShapeRange's Shadow, Glow, SoftEdge and Reflection: reads where its shapes agree and differ, sets, and each saved
shape's spPr.

    python scripts/measure_shape_range_effects.py

writes tests/fixtures/shape_range_effects.json, which tests/test_excel_shape_range_effects.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_range_effects.json'
RECTS = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddShape(1, 20, 70, 100, 40)')
MIXED = ('ws.Shapes.AddShape(1, 10, 10, 100, 40)', 'ws.Shapes.AddLine(20, 70, 120, 110)')
SETS = [
    'Shadow.Visible = msoTrue', 'Shadow.OffsetX = 5', 'Shadow.OffsetY = 5', 'Shadow.Blur = 3', 'Shadow.Transparency = 0.5',
    'Shadow.Size = 110', 'Shadow.ForeColor.RGB = 255', 'Shadow.Style = 2', 'Shadow.Type = 21',
    'Glow.Radius = 5', 'Glow.Color.RGB = 255', 'Glow.Transparency = 0.5', 'SoftEdge.Radius = 5', 'SoftEdge.Type = 1',
    'Reflection.Type = 1', 'Reflection.Transparency = 0.5', 'Reflection.Size = 50', 'Reflection.Offset = 3', 'Reflection.Blur = 2',
]
CASES = [(RECTS, ''), (MIXED, ''),
         *[(RECTS, f'r.{one}') for one in SETS], *[(RECTS, f'a.{one}') for one in SETS],
         (MIXED, 'r.Shadow.Visible = msoTrue'), (MIXED, 'r.Glow.Radius = 5'), (MIXED, 'r.SoftEdge.Radius = 5'),
         (MIXED, 'r.Reflection.Type = 1')]
READS = ['TypeName(r.Shadow)', 'r.Shadow.Visible', 'r.Shadow.OffsetX', 'r.Shadow.OffsetY', 'r.Shadow.Blur', 'r.Shadow.Transparency',
         'r.Shadow.Size', 'r.Shadow.ForeColor.RGB', 'r.Shadow.ForeColor.ObjectThemeColor', 'r.Shadow.Style', 'r.Shadow.Type',
         'r.Shadow.Obscured', 'r.Shadow.RotateWithShape',
         'TypeName(r.Glow)', 'r.Glow.Radius', 'r.Glow.Color.RGB', 'r.Glow.Transparency',
         'TypeName(r.SoftEdge)', 'r.SoftEdge.Radius', 'r.SoftEdge.Type',
         'TypeName(r.Reflection)', 'r.Reflection.Type', 'r.Reflection.Transparency', 'r.Reflection.Size', 'r.Reflection.Offset',
         'r.Reflection.Blur', 'TypeName(r.Shadow.Parent)']


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
            parts.append(re.findall(r'<xdr:spPr\b.*?</xdr:spPr>', xml, re.DOTALL))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[statements, answer, part] for (_, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
