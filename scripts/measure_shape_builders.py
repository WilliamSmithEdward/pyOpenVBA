"""Shapes.BuildFreeform with FreeformBuilder.AddNodes and ConvertToShape, and Shapes.AddCallout: what each makes, the
errors, and the anchor each saved shape has.

    python scripts/measure_shape_builders.py

writes tests/fixtures/shape_builders.json, which tests/test_excel_shape_builders.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_builders.json'
BUILD = 'Set b = ws.Shapes.BuildFreeform(msoEditingCorner, 10, 10)'
LINE = 'b.AddNodes msoSegmentLine, msoEditingAuto, '
CURVE = 'b.AddNodes msoSegmentCurve, msoEditingCorner, '
SMOOTH = 'b.AddNodes msoSegmentCurve, msoEditingAuto, '
END = 'Set sh = b.ConvertToShape'
CASES = [
    f'{BUILD}: {LINE}110, 10: {LINE}60, 80: {END}',
    f'{BUILD}: {LINE}110, 10: {LINE}60, 80: {LINE}10, 10: {END}',
    f'{BUILD}: {LINE}110, 10: {END}',
    f'{BUILD}: {END}',
    f'Set b = ws.Shapes.BuildFreeform(msoEditingAuto, 10, 10): {LINE}110, 10: {LINE}60, 80: {END}',
    'Set b = ws.Shapes.BuildFreeform(3, 10, 10)',
    'Set b = ws.Shapes.BuildFreeform(-1, 10, 10)',
    f'{BUILD}: {CURVE}40, 60, 80, 60, 110, 10: {END}',
    f'{BUILD}: {CURVE}40, 60, 80, 60, 110, 10: {CURVE}140, -20, 170, 40, 200, 10: {END}',
    f'{BUILD}: {CURVE}40, 60, 80, 60, 110, 10: {LINE}200, 50: {END}',
    f'{BUILD}: {LINE}200, 50: {CURVE}40, 60, 80, 60, 110, 10: {END}',
    f'{BUILD}: {CURVE}40, 60, 80, 60, 10, 10: {END}',
    f'{BUILD}: {CURVE}40, 60: {END}',
    f'{BUILD}: {SMOOTH}40, 60: {SMOOTH}80, 60: {END}',
    f'{BUILD}: {SMOOTH}40, 60, 80, 60, 110, 10: {END}',
    f'{BUILD}: {LINE}40, 60, 80, 60, 110, 10: {END}',
    f'{BUILD}: b.AddNodes msoSegmentLine, msoEditingCorner, 110, 10: {END}',
    f'{BUILD}: b.AddNodes 9, msoEditingAuto, 110, 10: {END}',
    f'{BUILD}: b.AddNodes msoSegmentLine, 9, 110, 10: {END}',
    f'{BUILD}: {LINE}110.5, 10.25: {LINE}60.75, 80.5: {END}',
    f'{BUILD}: {LINE}110, 10: {LINE}60, 80: {END}: sh.Name = "Mine"',
    f'{BUILD}: {LINE}110, 10: {LINE}60, 80: {END}: {END}',
    f'{BUILD}: {LINE}110, 10: {END}: {LINE}60, 80: {END}',
    f'Set b = ws.Shapes.BuildFreeform(msoEditingCorner, 50, 40): {LINE}10, 10: {LINE}90, 5: {END}',
    *[f'Set sh = ws.Shapes.AddCallout({kind}, 10, 10, 100, 40)' for kind in (1, 2, 3, 4)],
    *[f'Set sh = ws.Shapes.AddCallout({kind}, 10, 10, 100, 40)' for kind in (0, 5, -2)],
    'Set sh = ws.Shapes.AddCallout(2, 10, 10, 100, 40): sh.Adjustments(1) = 0.3',
    'Set sh = ws.Shapes.AddCallout(2, 10, 10, 100, 40): sh.Fill.ForeColor.RGB = 255',
    'Set sh = ws.Shapes.AddCallout(2, 10, 10, 100, 40): sh.Line.Weight = 3',
    'Set sh = ws.Shapes.AddCallout(2, 10, 10, 100, 40): sh.AutoShapeType = 1',
    'Set sh = ws.Shapes.AddCallout(2, 10, 10, 100, 40): sh.TextFrame.Characters.Text = "Hi"',
]
READS = ['TypeName(b)', 'TypeName(sh)', 'sh.Name', 'sh.Type', 'sh.AutoShapeType', 'sh.Left', 'sh.Top', 'sh.Width', 'sh.Height',
         'ws.Shapes.Count', 'sh.Fill.Visible', 'sh.Line.Visible', 'sh.Adjustments.Count', 'sh.Adjustments(1)',
         'sh.Adjustments(2)', 'sh.Adjustments(3)', 'sh.Adjustments(4)', 'sh.HorizontalFlip', 'sh.VerticalFlip']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, b As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
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
    """The first shape as the drawing part holds it, or '' with no drawing."""
    with zipfile.ZipFile(path) as package:
        try:
            xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        except KeyError:
            return ''
    found = re.search(r'<xdr:(sp|cxnSp)\b.*?</xdr:\1>', xml, re.DOTALL)
    return found.group(0) if found else ''


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [drawn(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[statements, answer, part] for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    strip = re.compile(r'<xdr:nv\w+>.*?</xdr:nv\w+>|<xdr:style>.*?</xdr:style>|<xdr:txBody>.*?</xdr:txBody>')
    for statements, answer, part in record['cases']:
        print(statements, '->', answer)
        print('    ', strip.sub('', part)[:760])


if __name__ == '__main__':
    main()
