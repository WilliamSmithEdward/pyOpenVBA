"""TextFrame's margins, alignment, orientation, overflow and reading order: what a rectangle and a text box answer
before and after each setting, the errors, and the saved bodyPr and first paragraph's pPr.

Each case makes a workbook with a rectangle R (100, 50, 120 by 60) and a text box T (250, 50, 120 by 60), both
holding "Hi", runs the case's statements on the shape the case names through tf, reads every setting of both
shapes, then saves it.

    python scripts/measure_text_frame.py

writes tests/fixtures/text_frame.json, which tests/test_excel_text_frame.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/text_frame.json'
SETTINGS = ['tf.MarginLeft = 10', 'tf.MarginTop = 3.5', 'tf.MarginRight = 0', 'tf.MarginBottom = 20',
            'tf.MarginLeft = -1', 'tf.MarginLeft = 10: tf.AutoMargins = True', 'tf.AutoMargins = False',
            'tf.AutoMargins = True: tf.MarginTop = 2', 'tf.HorizontalAlignment = xlHAlignCenter',
            'tf.HorizontalAlignment = xlHAlignRight', 'tf.HorizontalAlignment = xlHAlignLeft',
            'tf.HorizontalAlignment = xlHAlignJustify', 'tf.HorizontalAlignment = xlHAlignDistributed',
            'tf.HorizontalAlignment = xlHAlignGeneral', 'tf.HorizontalAlignment = 5',
            'tf.VerticalAlignment = xlVAlignCenter', 'tf.VerticalAlignment = xlVAlignBottom',
            'tf.VerticalAlignment = xlVAlignTop', 'tf.VerticalAlignment = xlVAlignJustify',
            'tf.VerticalAlignment = xlVAlignDistributed', 'tf.VerticalAlignment = 5',
            # msoTextOrientation 1 to 6 and one past, then xlUpward, xlDownward, xlVertical and xlHorizontal.
            *(f'tf.Orientation = {value}' for value in (1, 2, 3, 4, 5, 6, 7, -4171, -4170, -4166, -4128)),
            'tf.HorizontalOverflow = xlOartHorizontalOverflowClip',
            'tf.HorizontalOverflow = xlOartHorizontalOverflowOverflow', 'tf.HorizontalOverflow = 2',
            'tf.VerticalOverflow = xlOartVerticalOverflowClip', 'tf.VerticalOverflow = xlOartVerticalOverflowEllipsis',
            'tf.VerticalOverflow = xlOartVerticalOverflowOverflow', 'tf.VerticalOverflow = 3',
            'tf.ReadingOrder = xlRTL', 'tf.ReadingOrder = xlLTR', 'tf.ReadingOrder = xlContext',
            'tf.ReadingOrder = 5', 's = tf.AutoMargins', 'tf.AutoMargins = True: s = tf.AutoMargins',
            'tf.HorizontalAlignment = 7', 'tf.HorizontalAlignment = xlHAlignCenter: tf.ReadingOrder = xlRTL',
            'tf.Orientation = 2: tf.Orientation = 1', 'tf.ReadingOrder = xlRTL: tf.ReadingOrder = xlLTR',
            'tf.VerticalOverflow = xlOartVerticalOverflowOverflow: tf.MarginLeft = 10: tf.Orientation = 2: '
            'tf.VerticalAlignment = xlVAlignBottom: tf.HorizontalAlignment = xlHAlignCenter',
            'tf.VerticalOverflow = xlOartVerticalOverflowOverflow: tf.VerticalOverflow = xlOartVerticalOverflowClip']
CASES = ['', *(f'Set tf = ws.Shapes("{name}").TextFrame: {setting}' for name in 'RT' for setting in SETTINGS)]
READ = ('    For Each sh In ws.Shapes: Set tf = sh.TextFrame: s = s & "|" & sh.Name & "," & tf.MarginLeft & "," & '
        'tf.MarginTop & "," & tf.MarginRight & "," & tf.MarginBottom & "," & '
        'tf.HorizontalAlignment & "," & tf.VerticalAlignment & "," & tf.Orientation & "," & tf.HorizontalOverflow & '
        '"," & tf.VerticalOverflow & "," & tf.ReadingOrder: Next')


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, tf As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    With ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60): .Name = "R": '
             '.TextFrame.Characters.Text = "Hi": End With',
             '    With ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 250, 50, 120, 60): .Name = "T": '
             '.TextFrame.Characters.Text = "Hi": End With',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', READ, '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def text_settings(path: Path) -> list[str]:
    """Each shape's bodyPr and its first paragraph's pPr, as saved."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found: list[str] = []
    for body in re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.S):
        head = re.search(r'<a:bodyPr\b[^>]*?(?:/>|>.*?</a:bodyPr>)', body, re.S)
        paragraph = re.search(r'<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)', body, re.S)
        found.append((head.group(0) if head else '') + (paragraph.group(0) if paragraph else ''))
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=290.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [text_settings(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case[-44:], '->', answer, found)


if __name__ == '__main__':
    main()
