"""TextFrame2's margins, AutoSize, HorizontalAnchor, VerticalAnchor, NoTextRotation, Orientation, WordWrap and
DeleteText: what a rectangle and a text box answer before and after each setting, the errors, each shape's box, and
the saved txBody.

Each case makes a workbook with a rectangle R (100, 50, 120 by 60) and a text box T (250, 50, 120 by 60), both
holding "Hi", runs the case's statements on the shape the case names through tf, reads every setting of both
shapes, then saves it.

    python scripts/measure_text_frame2.py

writes tests/fixtures/text_frame2.json, which tests/test_excel_text_frame2.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/text_frame2.json'
SETTINGS = ['tf.MarginLeft = 10', 'tf.MarginBottom = 0', 'tf.MarginTop = -1', 'tf.AutoSize = msoAutoSizeNone',
            'tf.AutoSize = msoAutoSizeTextToFitShape', 'tf.AutoSize = 3',
            'tf.HorizontalAnchor = msoAnchorCenter', 'tf.HorizontalAnchor = msoAnchorNone',
            'tf.HorizontalAnchor = 3', 'tf.VerticalAnchor = msoAnchorTop', 'tf.VerticalAnchor = msoAnchorTopBaseline',
            'tf.VerticalAnchor = msoAnchorMiddle', 'tf.VerticalAnchor = msoAnchorBottom',
            'tf.VerticalAnchor = msoAnchorBottomBaseLine', 'tf.VerticalAnchor = 6', 'tf.NoTextRotation = msoTrue',
            'tf.NoTextRotation = msoFalse', 'tf.NoTextRotation = 5', 'tf.Orientation = msoTextOrientationUpward',
            'tf.Orientation = 7', 'tf.WordWrap = msoFalse', 'tf.WordWrap = msoTrue', 'tf.WordWrap = 5',
            'tf.DeleteText', 'tf.MarginLeft = 10: tf.VerticalAnchor = msoAnchorMiddle: tf.WordWrap = msoFalse']
CASES = ['', *(f'Set tf = ws.Shapes("{name}").TextFrame2: {setting}' for name in 'RT' for setting in SETTINGS)]
READ = ('    For Each sh In ws.Shapes: Set tf = sh.TextFrame2: s = s & "|" & sh.Name & "," & tf.MarginLeft & "," & '
        'tf.MarginTop & "," & tf.MarginRight & "," & tf.MarginBottom & "," & tf.AutoSize & "," & '
        'tf.HorizontalAnchor & "," & tf.VerticalAnchor & "," & tf.NoTextRotation & "," & tf.Orientation & "," & '
        'tf.WordWrap & "," & tf.HasText & "," & sh.TextFrame.VerticalAlignment & "," & sh.Width & "," & '
        'sh.Height: Next')


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


def text_bodies(path: Path) -> list[str]:
    """Each shape's txBody, as saved."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return re.findall(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.S)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=290.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [text_bodies(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        heads = [re.sub(r'<a:lstStyle/>.*', '', one.replace('<xdr:txBody>', ''))[:260] for one in found]
        print(case[-44:], '->', answer, heads)


if __name__ == '__main__':
    main()
