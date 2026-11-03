"""FillFormat's patterns: Patterned and Pattern and what each reads, on an AutoShape, a text box and a line, with the errors and the spPr each saved shape has.

Each case makes a workbook, adds the case's shape as sh, runs its statements, reads READS, each as its value or "E"
and the error number, and saves the workbook. The shape's spPr and style are recorded from the drawing part.

    python scripts/measure_shape_fill_patterns.py

writes tests/fixtures/shape_fill_patterns.json, which tests/test_excel_shape_fill_patterns.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_fill_patterns.json'
SHAPES = {'rect': 'ws.Shapes.AddShape(1, 10, 10, 100, 40)',
          'text': 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)',
          'line': 'ws.Shapes.AddLine(10, 10, 110, 50)'}
RED = 'sh.Fill.ForeColor.RGB = RGB(255, 0, 0)'
BLUE = 'sh.Fill.BackColor.RGB = RGB(0, 0, 255)'
CASES = [
    ('rect', ''), ('text', ''), ('line', ''),
    # Patterns.
    ('rect', 'sh.Fill.Patterned 22'), ('rect', 'sh.Fill.Patterned 1'), ('rect', 'sh.Fill.Patterned 54'),
    ('rect', 'sh.Fill.Patterned 0'), ('rect', 'sh.Fill.Patterned 55'), ('rect', 'sh.Fill.Patterned -2'),
    ('text', 'sh.Fill.Patterned 22'), ('line', 'sh.Fill.Patterned 22'),
    ('rect', f'sh.Fill.Patterned 22: {RED}'), ('rect', f'sh.Fill.Patterned 22: {BLUE}'),
    ('rect', f'sh.Fill.Patterned 22: {RED}: {BLUE}'), ('rect', f'{RED}: sh.Fill.Patterned 22'),
    ('rect', f'{BLUE}: sh.Fill.Patterned 22'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.Patterned 48'),
    ('rect', f'sh.Fill.Patterned 22: {RED}: sh.Fill.Patterned 48'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.ForeColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.BackColor.ObjectThemeColor = 6'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.ForeColor.Brightness = 0.4'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.Transparency = 0.5'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.Visible = msoFalse'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.Visible = msoFalse: sh.Fill.Visible = msoTrue'),
    ('rect', 'sh.Fill.Visible = msoFalse: sh.Fill.Patterned 22'),
    ('rect', 'sh.Fill.Patterned 22: sh.Fill.Solid'),
    ('rect', f'sh.Fill.Patterned 22: {RED}: sh.Fill.Solid'),
    ('rect', 'sh.Fill.Pattern = 22'),
]
READS = [
    'sh.Fill.Type', 'sh.Fill.Visible', 'sh.Fill.Pattern', 'sh.Fill.Transparency', 'sh.Fill.ForeColor.RGB',
    'sh.Fill.ForeColor.ObjectThemeColor', 'sh.Fill.ForeColor.Type', 'sh.Fill.ForeColor.Brightness',
    'sh.Fill.BackColor.RGB', 'sh.Fill.BackColor.ObjectThemeColor', 'sh.Fill.BackColor.Type',
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
