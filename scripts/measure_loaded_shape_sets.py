"""Settings a macro changes on the text, connections and picture of a shape in a file: what it reads after, and what the file holds.

    python scripts/measure_loaded_shape_sets.py

Each case opens tests/fixtures/loaded_shape_text.xlsx (scripts/measure_loaded_shape_text.py), changes one shape and saves.

Writes tests/fixtures/loaded_shape_sets.json. tests/test_excel_loaded_shape_text.py replays it.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / 'tests/fixtures/loaded_shape_text.xlsx'
OUT = ROOT / 'tests/fixtures/loaded_shape_sets.json'

#: Statements on the shape, s. CASES holds (shape, what it is, statement).
_FRAME = ['s.TextFrame.MarginLeft = 20', 's.TextFrame2.MarginTop = 11', 's.TextFrame.HorizontalAlignment = xlHAlignCenter', 's.TextFrame.VerticalAlignment = xlVAlignBottom', 's.TextFrame2.WordWrap = False',
          's.TextFrame2.AutoSize = 2', 's.TextFrame2.AutoSize = 0', 's.TextFrame2.Orientation = 3', 's.TextFrame2.VerticalAnchor = 3', 's.TextFrame2.HorizontalAnchor = 2', 's.TextFrame2.NoTextRotation = True', 's.TextFrame.HorizontalOverflow = 0',
          's.TextFrame.VerticalOverflow = 2', 's.TextFrame.ReadingOrder = xlRTL', 's.TextFrame2.Column.Number = 2', 's.TextFrame2.Column.Spacing = 5']
_FONT = ['s.TextFrame.Characters.Font.Bold = True', 's.TextFrame.Characters.Font.Italic = True', 's.TextFrame.Characters.Font.Size = 16', 's.TextFrame.Characters.Font.Name = "Arial"', 's.TextFrame.Characters.Font.Color = RGB(1, 2, 3)',
         's.TextFrame.Characters.Font.Underline = xlUnderlineStyleSingle', 's.TextFrame.Characters.Font.Strikethrough = True', 's.TextFrame.Characters.Font.Superscript = True',
         's.TextFrame.Characters(1, 3).Font.Bold = True', 's.TextFrame.Characters(2, 3).Font.Size = 9', 's.TextFrame.Characters(1, 3).Font.Color = RGB(9, 8, 7)',
         's.TextFrame2.TextRange.Font.Bold = True', 's.TextFrame2.TextRange.Font.Size = 13', 's.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = 6', 's.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(4, 5, 6)',
         's.TextFrame2.TextRange.Font.Fill.Transparency = 0.25', 's.TextFrame2.TextRange.Font.Line.Visible = True', 's.TextFrame2.TextRange.Font.Line.Weight = 1.5', 's.TextFrame2.TextRange.Characters(2, 2).Font.Italic = True',
         's.TextFrame2.TextRange.Font.Spacing = 2', 's.TextFrame2.TextRange.Font.Allcaps = True', 's.TextFrame2.TextRange.Font.UnderlineStyle = 3', 's.TextFrame2.TextRange.Font.Strike = 2']
_WORDS = ['s.TextFrame.Characters.Text = "New words"', 's.TextFrame2.TextRange.Text = "One" & vbCr & "Two"', 's.TextFrame2.TextRange.InsertAfter "!"', 's.TextFrame2.TextRange.InsertBefore "?"', 's.TextFrame.Characters(1, 2).Delete', 's.TextFrame.Characters(2, 1).Text = "xyz"', 's.TextFrame2.DeleteText']
_PARAGRAPH = ['s.TextFrame2.TextRange.ParagraphFormat.Alignment = 2', 's.TextFrame2.TextRange.Paragraphs(1).ParagraphFormat.Alignment = 3', 's.TextFrame2.TextRange.ParagraphFormat.LeftIndent = 10', 's.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = 5', 's.TextFrame2.TextRange.ParagraphFormat.IndentLevel = 2', 's.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = 6', 's.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 4', 's.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 2', 's.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2', 's.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible = True']
_LINE = ['s.ConnectorFormat.BeginDisconnect', 's.ConnectorFormat.EndDisconnect', 's.ConnectorFormat.BeginConnect ws.Shapes("End2"), 2', 's.ConnectorFormat.EndConnect ws.Shapes("End1"), 4', 's.RerouteConnections']
_MOVED = ['ws.Shapes("End1").Left = 500', 'ws.Shapes("End2").Top = 200']
_PICTURE = ['s.PictureFormat.Brightness = 0.7', 's.PictureFormat.Contrast = 0.3', 's.PictureFormat.ColorType = 2', 's.PictureFormat.ColorType = 3', 's.PictureFormat.ColorType = 4', 's.PictureFormat.ColorType = 1', 's.PictureFormat.TransparentBackground = False', 's.PictureFormat.TransparencyColor = RGB(0, 255, 0)', 's.PictureFormat.IncrementBrightness 0.1']
CASES: list[tuple[str, str, str]] = []
for name in ('Plain', 'Box', 'Clipped'):
    CASES += [(name, 'text', one) for one in _FRAME]
for name in ('Plain', 'Box', 'Runs', 'Paint', 'NoText'):
    CASES += [(name, 'text', one) for one in _FONT]
for name in ('Plain', 'Runs', 'Paras', 'NoText'):
    CASES += [(name, 'text', one) for one in _WORDS]
for name in ('Plain', 'Box', 'Paras', 'Laid'):
    CASES += [(name, 'text', one) for one in _PARAGRAPH]
for name in ('Joined', 'Elbow', 'Curve', 'Line'):
    CASES += [(name, 'line', one) for one in _LINE]
CASES += [(name, 'line', one) for name in ('Joined', 'Elbow') for one in _MOVED]
for name in ('Pic', 'Tuned', 'Clear'):
    CASES += [(name, 'picture', one) for one in _PICTURE]

READS = {
    'text': ('s.TextFrame.MarginLeft', 's.TextFrame.MarginTop', 's.TextFrame.HorizontalAlignment', 's.TextFrame.VerticalAlignment', 's.TextFrame.AutoSize', 's.TextFrame2.WordWrap', 's.TextFrame2.AutoSize', 's.TextFrame2.Orientation',
             's.TextFrame2.HorizontalAnchor', 's.TextFrame2.NoTextRotation', 's.TextFrame.HorizontalOverflow', 's.TextFrame.VerticalOverflow', 's.TextFrame.ReadingOrder', 's.TextFrame2.Column.Number', 's.TextFrame2.Column.Spacing',
             's.TextFrame.Characters.Text', 's.TextFrame2.TextRange.Runs.Count', 's.TextFrame2.TextRange.Paragraphs.Count',
             's.TextFrame.Characters.Font.Bold', 's.TextFrame.Characters.Font.Italic', 's.TextFrame.Characters.Font.Size', 's.TextFrame.Characters.Font.Name', 's.TextFrame.Characters.Font.Color', 's.TextFrame.Characters.Font.Underline',
             's.TextFrame.Characters(1, 2).Font.Bold', 's.TextFrame.Characters(1, 2).Font.Size', 's.TextFrame2.TextRange.Font.Bold', 's.TextFrame2.TextRange.Font.Fill.ForeColor.RGB', 's.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor', 's.TextFrame2.TextRange.Font.Fill.Transparency',
             's.TextFrame2.TextRange.Font.Line.Visible', 's.TextFrame2.TextRange.Font.Line.Weight', 's.TextFrame2.TextRange.Font.Spacing', 's.TextFrame2.TextRange.Font.Caps', 's.TextFrame2.TextRange.Font.UnderlineStyle', 's.TextFrame2.TextRange.Font.Strike',
             's.TextFrame2.TextRange.ParagraphFormat.Alignment', 's.TextFrame2.TextRange.ParagraphFormat.LeftIndent', 's.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent', 's.TextFrame2.TextRange.ParagraphFormat.IndentLevel', 's.TextFrame2.TextRange.ParagraphFormat.SpaceBefore', 's.TextFrame2.TextRange.ParagraphFormat.SpaceAfter', 's.TextFrame2.TextRange.ParagraphFormat.SpaceWithin', 's.TextFrame2.TextRange.ParagraphFormat.TextDirection', 's.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible'),
    'line': ('s.ConnectorFormat.BeginConnected', 's.ConnectorFormat.EndConnected', 's.ConnectorFormat.BeginConnectedShape.Name', 's.ConnectorFormat.EndConnectedShape.Name', 's.ConnectorFormat.BeginConnectionSite', 's.ConnectorFormat.EndConnectionSite', 's.Left', 's.Top', 's.Width', 's.Height', 's.HorizontalFlip', 's.VerticalFlip'),
    'picture': ('s.PictureFormat.Brightness', 's.PictureFormat.Contrast', 's.PictureFormat.ColorType', 's.PictureFormat.TransparentBackground', 's.PictureFormat.TransparencyColor'),
}
_SHOWN = '''Private Function Shown(v As Variant) As String
    If IsObject(v) Then
        Shown = "object"
    ElseIf IsArray(v) Or IsNull(v) Then
        Shown = TypeName(v)
    Else
        Shown = CStr(v)
    End If
End Function
'''


def _reader(kind: str) -> str:
    lines = [f'Private Function Read{kind}(s As Object) As String', '    Dim v As Variant, r As String', '    On Error Resume Next']
    for read in READS[kind]:
        lines += ['    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        r = r & "|E" & Err.Number', '    Else', '        r = r & "|" & Shown(v)', '    End If']
    return '\n'.join([*lines, f'    Read{kind} = r', 'End Function', ''])


def case_source(number: int, statement: str) -> str:
    """One case, on the workbook already open."""
    return '\n'.join([f'Private Sub Act{number}(ws As Object, s As Object)', f'    {statement}', 'End Sub', ''])


def source() -> str:
    lines = [_SHOWN, _reader('text'), _reader('line'), _reader('picture')]
    for number, (_, _, statement) in enumerate(CASES):
        lines.append(case_source(number, statement))
    lines += ['Public Function Case1(ByVal which As Long, ByVal book As String, ByVal saved As String) As String',
              '    Dim wb As Workbook, ws As Object, s As Object, r As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(book)', '    Set ws = wb.Worksheets(1)', '    Select Case which']
    for number, (name, _, _) in enumerate(CASES):
        lines += [f'    Case {number}', f'        Set s = ws.Shapes("{name}")']
    lines += ['    End Select', '    On Error Resume Next', '    Select Case which']
    for number in range(len(CASES)):
        lines += [f'    Case {number}', f'        Act{number} ws, s']
    lines += ['    End Select', '    If Err.Number Then r = "E" & Err.Number & "!"', '    Err.Clear', '    On Error GoTo 0', '    Select Case which']
    for number, (_, kind, _) in enumerate(CASES):
        lines += [f'    Case {number}', f'        r = r & Read{kind}(s)']
    lines += ['    End Select', '    wb.SaveAs saved, xlOpenXMLWorkbook', '    wb.Close False', '    Application.DisplayAlerts = True', '    Case1 = r', 'End Function']
    return '\n'.join(lines) + '\n'


def anchor_of(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    for found in re.finditer(r'<xdr:(?:twoCell|oneCell|absolute)Anchor\b.*?</xdr:(?:twoCell|oneCell|absolute)Anchor>', xml, re.S):
        if f'name="{name}"' in found.group():
            return found.group()
    return ''


def main() -> None:
    code = source()
    records = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        book = Path(folder) / 'book.xlsx'
        shutil.copyfile(BOOK, book)
        for number, (name, kind, statement) in enumerate(CASES):
            saved = Path(folder) / f'c{number}.xlsx'
            result = excel.run_vba(code, 'Case1', (number, str(book), str(saved)), timeout=60.0, line_numbers=False)
            assert result.ok and not result.dialogs, (name, statement, result)
            records.append({'shape': name, 'kind': kind, 'statement': statement, 'answer': str(result.value), 'anchor': anchor_of(saved, name), 'end1': anchor_of(saved, 'End1') if 'End1' in statement else ''})
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': code, 'cases': records}, indent=0) + '\n', encoding='utf-8')
    print(len(records), 'cases;', sum(one['answer'].startswith('E') for one in records), 'raised')


if __name__ == '__main__':
    main()
