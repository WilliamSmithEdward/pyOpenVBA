"""TextFrame2.Column, Shape.Callout and Shape.SetShapesDefaultProperties: what each reads, the errors, and each saved
shape's spPr, style and text body.

    python scripts/measure_shape_columns.py

writes tests/fixtures/shape_columns.json, which tests/test_excel_shape_columns.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_columns.json'
RECT = 'ws.Shapes.AddShape(1, 10, 10, 100, 40)'
TEXT = 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)'
LINE = 'ws.Shapes.AddLine(10, 10, 110, 50)'
CALL = 'ws.Shapes.AddCallout(2, 10, 10, 100, 40)'
NEXT = 'Set c = ws.Shapes.AddShape(1, 20, 70, 100, 40)'
CASES = [
    (RECT, ''), (TEXT, ''), (LINE, ''), (CALL, ''),
    (RECT, 'sh.TextFrame2.Column.Number = 2'), (RECT, 'sh.TextFrame2.Column.Number = 16'), (RECT, 'sh.TextFrame2.Column.Number = 17'),
    (RECT, 'sh.TextFrame2.Column.Number = 0'), (RECT, 'sh.TextFrame2.Column.Number = 1'), (RECT, 'sh.TextFrame2.Column.Spacing = 5'),
    (RECT, 'sh.TextFrame2.Column.Spacing = -1'), (RECT, 'sh.TextFrame2.Column.Number = 2: sh.TextFrame2.Column.Spacing = 5'),
    (RECT, 'sh.TextFrame2.Column.Spacing = 5: sh.TextFrame2.Column.Number = 2'), (RECT, 'sh.TextFrame2.Column.Spacing = 0'),
    (TEXT, 'sh.TextFrame2.Column.Number = 2'), (TEXT, 'sh.TextFrame2.Column.Spacing = 5'),
    (LINE, 'sh.TextFrame2.Column.Number = 2'),
    (RECT, 'sh.TextFrame2.Column.Number = 2: sh.TextFrame2.MarginLeft = 3'),
    (RECT, 'sh.TextFrame2.MarginLeft = 3: sh.TextFrame2.Column.Number = 2'),
    (CALL, 'sh.Callout.Type = 3'), (CALL, 'sh.Callout.Gap = 5'), (CALL, 'sh.Callout.Border = msoFalse'), (CALL, 'sh.Callout.Accent = msoTrue'),
    (CALL, 'sh.Callout.Angle = 2'), (CALL, 'sh.Callout.AutoLength = msoFalse'), (CALL, 'sh.Callout.Drop = 5'),
    (CALL, 'sh.Callout.PresetDrop 1'), (CALL, 'sh.Callout.CustomLength 5'), (CALL, 'sh.Callout.AutomaticLength'),
    (CALL, 'sh.Callout.CustomDrop 5'), (RECT, 'sh.Callout.Type = 3'),
    ('ws.Shapes.AddShape(109, 10, 10, 100, 40)', ''), ('ws.Shapes.AddShape(61, 10, 10, 100, 40)', ''),
    *[(f'ws.Shapes.AddShape({kind}, 10, 10, 100, 40)', '') for kind in range(109, 125)],
    *[(CALL, f'sh.Callout.Type = {kind}') for kind in (1, 2, 3, 4, 0, 5)],
    *[('ws.Shapes.AddShape(110, 10, 10, 100, 40)', f'sh.Callout.Type = {kind}') for kind in (1, 3, 4)],
    ('ws.Shapes.AddShape(109, 10, 10, 100, 40)', 'sh.Callout.Border = msoFalse'),
    ('ws.Shapes.AddShape(109, 10, 10, 100, 40)', 'sh.Callout.Accent = msoTrue'),
    (CALL, 'sh.Callout.Border = msoTrue'), (CALL, 'sh.Callout.Accent = msoTrue: sh.Callout.Accent = msoFalse'),
    (CALL, 'sh.Callout.Border = msoTrue: sh.Callout.Accent = msoTrue'),
    *[(CALL, f'sh.Callout.Angle = {angle}') for angle in (1, 3, 4, 5, 6, 0)],
    *[(CALL, f'sh.Callout.Gap = {gap}') for gap in (0, 20, 200, -5)],
    *[(CALL, f'sh.Callout.PresetDrop {drop}') for drop in (1, 2, 3, 4, 0, 5)],
    *[(CALL, f'sh.Callout.CustomDrop {drop}') for drop in (0, 20, 100, -5)],
    *[(CALL, f'sh.Callout.CustomLength {length}') for length in (0, 20, 100, -5)],
    (CALL, 'sh.Callout.CustomLength 20: sh.Callout.AutomaticLength'), (CALL, 'sh.Callout.AutoAttach = msoTrue'),
    (CALL, 'sh.Callout.AutoAttach = msoFalse'), (CALL, 'sh.Width = 200'), (CALL, 'sh.Height = 80'),
    ('ws.Shapes.AddCallout(3, 10, 10, 100, 40)', 'sh.Callout.Gap = 5'), ('ws.Shapes.AddCallout(4, 10, 10, 100, 40)', 'sh.Callout.CustomDrop 5'),
    ('ws.Shapes.AddCallout(3, 10, 10, 100, 40)', 'sh.Callout.CustomLength 5'), (CALL, 'sh.Adjustments(1) = 0.5'),
    (RECT, f'sh.SetShapesDefaultProperties: {NEXT}'),
    (RECT, f'sh.Fill.ForeColor.RGB = 255: sh.SetShapesDefaultProperties: {NEXT}'),
    (RECT, f'sh.Line.Weight = 3: sh.SetShapesDefaultProperties: {NEXT}'),
    (RECT, f'sh.Fill.ForeColor.RGB = 255: sh.SetShapesDefaultProperties: Set c = ws.Shapes.AddShape(5, 20, 70, 100, 40)'),
    (RECT, f'sh.Fill.ForeColor.RGB = 255: sh.SetShapesDefaultProperties: Set c = ws.Shapes.AddTextbox(1, 20, 70, 100, 40)'),
    (RECT, f'sh.Fill.ForeColor.RGB = 255: sh.SetShapesDefaultProperties: Set c = ws.Shapes.AddLine(20, 70, 120, 110)'),
    (RECT, f'sh.Fill.ForeColor.RGB = 255: sh.SetShapesDefaultProperties: Set c = wb.Worksheets.Add.Shapes.AddShape(1, 20, 70, 100, 40)'),
    (RECT, f'sh.TextFrame.Characters.Text = "Hi": sh.TextFrame.Characters.Font.Bold = True: sh.SetShapesDefaultProperties: {NEXT}'),
    (RECT, f'sh.TextFrame.MarginLeft = 3: sh.SetShapesDefaultProperties: {NEXT}'),
]
READS = ['sh.TextFrame2.Column.Number', 'sh.TextFrame2.Column.Spacing', 'TypeName(sh.TextFrame2.Column)', 'TypeName(sh.TextFrame2.Column.Parent)',
         'TypeName(sh.Callout)', 'sh.Callout.Type', 'sh.Callout.Gap', 'sh.Callout.Border', 'sh.Callout.Accent', 'sh.Callout.Angle', 'sh.Callout.AutoAttach',
         'sh.Callout.AutoLength', 'sh.Callout.Drop', 'sh.Callout.DropType', 'sh.Callout.Length', 'sh.AutoShapeType',
         'sh.Type', 'c.Fill.ForeColor.RGB', 'c.Line.Weight', 'c.Name', 'c.TextFrame.MarginLeft', 'c.TextFrame.Characters.Font.Bold']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, c As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which']
    for index, (shape, _) in enumerate(CASES):
        lines += [f'    Case {index}: Set sh = {shape}']
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
                theme = package.read('xl/theme/theme1.xml').decode('utf-8')
            found = re.search(r'<a:objectDefaults>.*?</a:objectDefaults>|<a:objectDefaults/>', theme, re.DOTALL)
            parts.append([re.findall(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?(?:<xdr:txBody>.*?</xdr:txBody>)?', xml, re.DOTALL),
                          found.group(0) if found else ''])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part] for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
