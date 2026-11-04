"""TextFrame.Characters.Font on a new AutoShape and text box: what each member reads, the errors, and the text body each
saved shape has.

    python scripts/measure_shape_font.py

writes tests/fixtures/shape_font.json, which tests/test_excel_shape_font.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_font.json'
RECT = 'ws.Shapes.AddShape(1, 10, 10, 100, 40)'
TEXT = 'ws.Shapes.AddTextbox(1, 10, 10, 100, 40)'
F = 'sh.TextFrame.Characters.Font.'
HI = 'sh.TextFrame.Characters.Text = "Hi": '
SETS = ['Bold = True', 'Bold = False', 'Italic = True', 'Size = 14', 'Size = 10.5', 'Size = 0', 'Size = 500', 'Name = "Arial"', 'Name = ""',
        'Color = 255', 'ColorIndex = 3', 'ColorIndex = xlAutomatic', 'Underline = xlUnderlineStyleSingle',
        'Underline = xlUnderlineStyleDouble', 'Underline = xlUnderlineStyleNone', 'Strikethrough = True', 'Superscript = True',
        'Subscript = True', 'FontStyle = "Bold"', 'FontStyle = "Bold Italic"', 'FontStyle = "Regular"', 'ThemeColor = 2', 'TintAndShade = 0.5',
        'ThemeFont = 1', 'OutlineFont = True', 'Shadow = True']
CASES = [(RECT, ''), (TEXT, ''), (RECT, HI.rstrip(': ')), (TEXT, HI.rstrip(': ')),
         *[(RECT, F + one) for one in SETS], *[(RECT, HI + F + one) for one in SETS], *[(TEXT, HI + F + one) for one in SETS],
         (RECT, f'{F}Bold = True: {HI.rstrip(": ")}'), (RECT, f'{HI}{F}Bold = True: {F}Size = 14'),
         (RECT, f'{HI}{F}Bold = True: {F}Bold = False'), (RECT, f'{HI}{F}Size = 14: {F}Name = "Arial"'),
         (RECT, f'{HI}{F}Bold = True: sh.TextFrame.Characters.Text = "Bye"'),
         (RECT, 'sh.TextFrame.Characters.Text = "ab" & vbLf & "cd": ' + F + 'Bold = True'),
         (RECT, f'{HI}sh.TextFrame2.TextRange.Font.Bold = msoTrue'), (RECT, f'{HI}sh.TextFrame2.TextRange.Font.Size = 14')]
READS = [F + one for one in ('Bold', 'Italic', 'Size', 'Name', 'Color', 'ColorIndex', 'Underline', 'Strikethrough', 'Superscript', 'Subscript',
                             'FontStyle', 'ThemeColor', 'TintAndShade', 'ThemeFont', 'OutlineFont', 'Shadow', 'Background')] + [
    'TypeName(sh.TextFrame.Characters.Font)', 'sh.TextFrame.Characters.Text', 'TypeName(sh.TextFrame.Characters.Font.Parent)']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String',
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
                  '    s = s & "|" & v']
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
            parts.append(re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)[0])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{shape}: {statements}', answer, part] for (shape, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
