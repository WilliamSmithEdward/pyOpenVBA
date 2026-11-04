"""TextRange2.Font.Fill and Line, over all of a shape's text and part of it: what each reads, the errors, and the text body
each saved shape has.

    python scripts/measure_shape_font2_paint.py

writes tests/fixtures/shape_font2_paint.json, which tests/test_excel_shape_font2_paint.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_font2_paint.json'
RECT = 'ws.Shapes.AddShape(1, 10, 10, 100, 40)'
TEXT = 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)'
SET = 'sh.TextFrame.Characters.Text = "Hello World": '
W = 'sh.TextFrame2.TextRange.Font.'
P = 'sh.TextFrame2.TextRange.Characters(1, 5).Font.'
L = 'sh.TextFrame2.TextRange.Characters(7, 5).Font.'
C = 'sh.TextFrame.Characters'
FILLS = ('Visible = msoFalse', 'Visible = msoTrue', 'ForeColor.RGB = 255', 'ForeColor.ObjectThemeColor = 5', 'ForeColor.ObjectThemeColor = 13', 'ForeColor.SchemeColor = 3',
         'ForeColor.TintAndShade = 0.5', 'ForeColor.Brightness = 0.4', 'Transparency = 0.5', 'Transparency = 0', 'Transparency = 2', 'Solid', 'BackColor.RGB = 255', 'TwoColorGradient 1, 1',
         'Patterned 5', 'OneColorGradient 1, 1, 0.5', 'PresetGradient 1, 1, 1', 'PresetTextured 1')
LINES = ('Visible = msoTrue', 'Visible = msoFalse', 'ForeColor.RGB = 255', 'ForeColor.ObjectThemeColor = 5', 'ForeColor.SchemeColor = 3', 'ForeColor.TintAndShade = 0.5', 'ForeColor.Brightness = 0.4',
         'Weight = 2', 'Weight = 0', 'Weight = -1', 'DashStyle = 4', 'DashStyle = 1', 'Style = 2', 'Transparency = 0.5', 'BackColor.RGB = 255', 'Pattern = 5')
CASES = [
    (RECT, SET.rstrip(': ')), (TEXT, SET.rstrip(': ')),
    *[(RECT, f'{SET}{W}Fill.{one}') for one in FILLS], *[(RECT, f'{SET}{W}Line.{one}') for one in LINES],
    *[(TEXT, f'{SET}{W}Fill.{one}') for one in ('Visible = msoFalse', 'ForeColor.RGB = 255', 'Transparency = 0.5')],
    *[(TEXT, f'{SET}{W}Line.{one}') for one in ('Visible = msoTrue', 'ForeColor.RGB = 255', 'Weight = 2')],
    *[(RECT, f'{SET}{P}Fill.{one}') for one in FILLS[:4] + FILLS[6:9]], *[(RECT, f'{SET}{L}Fill.{one}') for one in FILLS[:3]],
    *[(RECT, f'{SET}{P}Line.{one}') for one in LINES[:4] + LINES[7:8] + LINES[10:11] + LINES[13:14]], *[(RECT, f'{SET}{L}Line.{one}') for one in LINES[:3]],
    # Together, and against Font.Color and Fill.Visible.
    (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {W}Fill.Transparency = 0.5'), (RECT, f'{SET}{W}Fill.Transparency = 0.5: {W}Fill.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {W}Fill.Visible = msoFalse'), (RECT, f'{SET}{W}Fill.Visible = msoFalse: {W}Fill.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Fill.Visible = msoFalse: {W}Fill.Visible = msoTrue'), (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {W}Fill.Visible = msoFalse: {W}Fill.Visible = msoTrue'),
    (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {W}Fill.ForeColor.ObjectThemeColor = 5'), (RECT, f'{SET}{W}Fill.ForeColor.ObjectThemeColor = 5: {W}Fill.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Fill.ForeColor.ObjectThemeColor = 5: {W}Fill.ForeColor.TintAndShade = 0.5'), (RECT, f'{SET}{W}Fill.ForeColor.ObjectThemeColor = 5: {W}Fill.ForeColor.Brightness = 0.4'),
    (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {W}Fill.ForeColor.Brightness = 0.4'), (RECT, f'{SET}{W}Fill.ForeColor.ObjectThemeColor = 5: {W}Fill.Transparency = 0.5'),
    (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {C}.Font.Color = 65280'), (RECT, f'{SET}{C}.Font.Color = 65280: {W}Fill.ForeColor.RGB = 255'),
    (RECT, f'{SET}{C}.Font.Color = 65280: {W}Fill.Transparency = 0.5'), (RECT, f'{SET}{C}.Font.Color = 65280: {W}Fill.Visible = msoFalse'),
    (RECT, f'{SET}{W}Fill.Visible = msoFalse: {C}.Font.Color = 65280'), (RECT, f'{SET}{W}Fill.Transparency = 0.5: {C}.Font.Color = 65280'),
    (RECT, f'{SET}{W}Line.ForeColor.RGB = 255: {W}Line.Weight = 2'), (RECT, f'{SET}{W}Line.Weight = 2: {W}Line.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Line.ForeColor.RGB = 255: {W}Line.Visible = msoFalse'), (RECT, f'{SET}{W}Line.Visible = msoFalse: {W}Line.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Line.Visible = msoTrue: {W}Line.Visible = msoFalse'), (RECT, f'{SET}{W}Line.Visible = msoFalse: {W}Line.Visible = msoTrue'),
    (RECT, f'{SET}{W}Line.Visible = msoTrue: {W}Line.Weight = 2: {W}Line.DashStyle = 4: {W}Line.ForeColor.RGB = 255: {W}Line.Transparency = 0.5'),
    (RECT, f'{SET}{W}Line.ForeColor.ObjectThemeColor = 5: {W}Line.Transparency = 0.5'), (RECT, f'{SET}{W}Line.DashStyle = 4: {W}Line.Style = 2'),
    (RECT, f'{SET}{W}Line.ForeColor.RGB = 255: {W}Fill.ForeColor.RGB = 65280'), (RECT, f'{SET}{W}Fill.ForeColor.RGB = 65280: {W}Line.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Bold = msoTrue: {W}Line.ForeColor.RGB = 255: {W}Fill.ForeColor.RGB = 65280: {W}Size = 14'),
    (RECT, f'{SET}{W}Shadow.Visible = msoFalse: {W}Line.ForeColor.RGB = 255'),
    # Part over whole and whole over part.
    (RECT, f'{SET}{W}Fill.ForeColor.RGB = 255: {P}Fill.ForeColor.RGB = 65280'), (RECT, f'{SET}{P}Fill.ForeColor.RGB = 65280: {W}Fill.ForeColor.RGB = 255'),
    (RECT, f'{SET}{W}Line.ForeColor.RGB = 255: {P}Line.Weight = 2'), (RECT, f'{SET}{P}Line.ForeColor.RGB = 65280: {W}Line.ForeColor.RGB = 255'),
    (RECT, f'{SET}{P}Fill.ForeColor.RGB = 255: {L}Fill.ForeColor.RGB = 255'), (RECT, f'{SET}{P}Line.ForeColor.RGB = 255: {L}Line.ForeColor.RGB = 65280'),
    (RECT, f'{SET}{P}Fill.Transparency = 0.5: {C}(1, 5).Font.Color = 255'), (RECT, f'{SET}{P}Fill.ForeColor.RGB = 255: {C}(1, 5).Font.Bold = True'),
    (RECT, f'{SET}{P}Line.ForeColor.RGB = 255: {C}(1, 5).Text = "Bye"'), (RECT, f'{SET}{P}Line.ForeColor.RGB = 255: {C}.Text = "New"'),
    # On a line already shown, and what shows a hidden one.
    *[(RECT, f'{SET}{W}Line.{start}: {W}Line.{one}') for start in ('Visible = msoTrue', 'Weight = 2') for one in (
        'ForeColor.RGB = 255', 'ForeColor.ObjectThemeColor = 5', 'ForeColor.TintAndShade = 0.5', 'ForeColor.Brightness = 0.4', 'Weight = 3', 'DashStyle = 4', 'Style = 2', 'Transparency = 0.5', 'Transparency = 0')],
    (RECT, f'{SET}{W}Line.ForeColor.RGB = 255: {W}Line.Visible = msoTrue'), (RECT, f'{SET}{W}Line.Weight = 2: {W}Line.ForeColor.RGB = 255: {W}Line.Transparency = 0.5: {W}Line.Weight = 3'),
    (RECT, f'{SET}{W}Line.Weight = 2: {W}Line.Transparency = 0.5: {W}Line.ForeColor.RGB = 255'), (RECT, f'{SET}{W}Line.Weight = 2: {W}Line.ForeColor.ObjectThemeColor = 5: {W}Line.ForeColor.Brightness = 0.4'),
    (RECT, f'{SET}{W}Line.Weight = 2: {W}Line.ForeColor.RGB = 255: {W}Line.ForeColor.ObjectThemeColor = 5'), (RECT, f'{SET}{W}Line.Weight = 2: {W}Line.Transparency = 2'),
    (RECT, f'{SET}{P}Line.Visible = msoTrue: {P}Line.ForeColor.RGB = 255'), (RECT, f'{SET}{P}Line.Weight = 2: {P}Line.Transparency = 0.5'), (RECT, f'{SET}{P}Line.Weight = 2: {L}Line.Weight = 3'),
    (RECT, f'{SET}{P}Line.Weight = 2: {W}Line.Weight = 3'), (RECT, f'{SET}{P}Line.Weight = 2: {W}Line.ForeColor.RGB = 255'),
    (TEXT, f'{SET}{W}Fill.Transparency = 0'), (TEXT, f'{SET}{W}Fill.Solid'), (TEXT, f'{SET}{W}Fill.ForeColor.TintAndShade = 0.5'), (TEXT, f'{SET}{W}Fill.ForeColor.Brightness = 0.4'), (TEXT, f'{SET}{W}Fill.Visible = msoFalse: {W}Fill.Visible = msoTrue'),
    (TEXT, f'{SET}{C}.Font.TintAndShade = 0.5'),
    # No text.
    (RECT, f'{W}Fill.ForeColor.RGB = 255'), (RECT, f'{W}Line.ForeColor.RGB = 255'), (RECT, f'{W}Fill.ForeColor.RGB = 255: {C}.Text = "Hi"'), (RECT, f'{W}Line.ForeColor.RGB = 255: {C}.Text = "Hi"'),
]
READS = [*[f'{W}Fill.{one}' for one in ('Visible', 'Type', 'ForeColor.RGB', 'ForeColor.ObjectThemeColor', 'ForeColor.Type', 'ForeColor.TintAndShade', 'ForeColor.Brightness', 'ForeColor.SchemeColor', 'Transparency', 'BackColor.RGB', 'Pattern')],
         *[f'{W}Line.{one}' for one in ('Visible', 'ForeColor.RGB', 'ForeColor.ObjectThemeColor', 'ForeColor.Type', 'ForeColor.TintAndShade', 'ForeColor.Brightness', 'Weight', 'DashStyle', 'Style', 'Transparency', 'BackColor.RGB', 'Pattern')],
         *[f'{P}Fill.{one}' for one in ('Visible', 'ForeColor.RGB', 'Transparency')], *[f'{P}Line.{one}' for one in ('Visible', 'ForeColor.RGB', 'Weight')],
         f'{C}.Font.Color', f'{C}(1, 5).Font.Color', 'TypeName(' + W + 'Fill)', 'TypeName(' + W + 'Line)', 'TypeName(' + W + 'Fill.Parent)', 'TypeName(' + W + 'Line.Parent)']


def _probe() -> str:
    lines: list[str] = []
    chunks = [range(at, min(at + 60, len(CASES))) for at in range(0, len(CASES), 60)]
    # In procedures of their own, since one holding them all is too large to compile. Each takes its own error: passed up
    # through a call, Excel's -2147024809 turns into VBA's 5.
    for number, chunk in enumerate(chunks):
        lines += [f'Private Sub Act{number}(ByVal which As Long, ByVal sh As Object, failure As Long)', '    On Error GoTo Failed', '    Select Case which']
        for index in chunk:
            lines += [f'    Case {index}', f'        {CASES[index][1]}']
        lines += ['    End Select', '    Exit Sub', 'Failed:', '    failure = Err.Number', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String, failure As Long',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)']
    boxes = ', '.join(str(index) for index, (shape, _) in enumerate(CASES) if shape == TEXT)
    lines += ['    Select Case which', f'    Case {boxes}: Set sh = {TEXT}', f'    Case Else: Set sh = {RECT}', '    End Select', '    Select Case which \\ 60']
    for number in range(len(chunks)):
        lines += [f'    Case {number}: Act{number} which, sh, failure']
    lines += ['    End Select', '    s = IIf(failure = 0, "ok", "Err " & failure)', '    On Error Resume Next']
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
        result = excel.run_vba(source, 'Probe', (folder,), timeout=560.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = []
        for index in range(len(CASES)):
            with zipfile.ZipFile(Path(folder) / f'c{index}.xlsx') as package:
                xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
            parts.append(re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part] for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
