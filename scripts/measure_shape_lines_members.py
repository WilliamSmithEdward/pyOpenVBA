"""TextRange2.Lines as a member, and what else moves its breaks, into tests/fixtures/shape_lines_members.json. Each case is statements on a rectangle, then a read, with its type and any error."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'shape_lines_members.json'
TEXT = 'The quick brown fox jumps over the lazy dog'
LAYOUT = 'LAYOUT'
T = 's.TextFrame2.TextRange'
SET = f'{T}.Text = '
CASES: list[tuple[str, str]] = [
    ('', f'{T}.Lines.Count'), ('', f'TypeName({T}.Lines)'), ('', f'{T}.Lines.Text'), ('', f'{T}.Lines.Start'), ('', f'{T}.Lines.Length'), ('', f'{T}.Lines(1).Text'), ('', f'{T}.Lines(2).Text'), ('', f'{T}.Lines(2).Start'), ('', f'{T}.Lines(2).Length'),
    ('', f'{T}.Lines(2).Count'), ('', f'{T}.Lines(1, 2).Text'), ('', f'{T}.Lines(1, 2).Count'), ('', f'{T}.Lines(2, 1).Text'), ('', f'{T}.Lines(99).Text'), ('', f'{T}.Lines(0).Text'), ('', f'{T}.Lines(-1).Text'), ('', f'{T}.Lines(1, 0).Text'), ('', f'{T}.Lines(1, 99).Text'), ('', f'{T}.Lines(2, 99).Length'),
    ('', f'{T}.Lines(2, -1).Text'), ('', f'{T}.Lines(-1, 1).Text'), ('', f'{T}.Lines(-1, 2).Text'), ('', f'{T}.Lines(3).Text'), ('', f'{T}.Lines(3, 1).Text'), ('', f'{T}.Lines(2, 0).Text'), ('', f'{T}.Lines(2, 0).Start'), ('', f'{T}.Lines(2, 0).Length'), ('', f'{T}.Lines(1, 0).Count'), ('', f'{T}.Lines(3, 0).Text'),
    ('x = 0: For Each t In ' + T + '.Lines: x = x + t.Length: Next', 'x'), ('', f'{T}.Lines(1, -1).Text'), ('', f'{T}.Lines.Item(2).Text'), ('', f'{T}.Lines.Lines.Count'), ('', f'{T}.Lines(2).Lines.Count'), ('', f'{T}.Lines(2).Lines(1).Text'), ('', f'{T}.Lines(2).Words.Count'), ('', f'{T}.Lines(2).Characters.Count'), ('', f'{T}.Lines(2).Paragraphs.Count'),
    ('', f'{T}.Characters(1, 10).Lines.Count'), ('', f'{T}.Characters(1, 10).Lines(1).Text'), ('', f'{T}.Characters(5, 30).Lines.Count'), ('', f'{T}.Characters(5, 30).Lines(1).Text'), ('', f'{T}.Characters(5, 30).Lines(2).Text'), ('', f'{T}.Characters(15, 30).Lines(1).Start'),
    ('', f'{T}.Characters(5, 30).Lines(1).Length'), ('', f'{T}.Characters(5, 0).Lines.Count'), ('', f'{T}.Words(2).Lines.Count'), ('', f'{T}.Words(2).Lines(1).Text'), ('', f'{T}.Paragraphs(1).Lines.Count'), ('', f'{T}.Paragraphs(1).Lines(2).Text'), ('', f'{T}.Runs(1).Lines.Count'), ('', f'{T}.Sentences(1).Lines.Count'),
    ('Set t = ' + T + '.Lines: s.Width = 400', 't.Count'), ('Set t = ' + T + '.Lines: s.Width = 60', 't.Count'), ('Set t = ' + T + '.Lines(2): s.Width = 400', 't.Text'), ('Set t = ' + T + '.Lines(2): s.Width = 60', 't.Text'), ('Set t = ' + T + '.Lines: ' + SET + '"a"', 't.Count'), ('Set t = ' + T + '.Lines(2): ' + SET + '"a"', 't.Text'),
    ('Set t = ' + T + '.Lines(2): ' + T + '.InsertBefore "abc "', 't.Text'), ('Set t = ' + T + '.Lines(2): ' + T + '.InsertAfter " abc"', 't.Text'), ('Set t = ' + T + '.Lines: ' + T + '.InsertAfter " abc def ghi jkl mno pqr"', 't.Count'), ('Set t = ' + T + '.Characters(5, 30).Lines: s.Width = 60', 't.Count'),
    (SET + '""', f'{T}.Lines.Count'), (SET + '""', f'{T}.Lines.Length'), (SET + '""', f'{T}.Lines(1).Length'), (SET + '""', LAYOUT), (SET + '" "', LAYOUT), (SET + '"     "', LAYOUT), (SET + 'vbLf', LAYOUT), (SET + 'vbLf & vbLf', LAYOUT), (SET + '"a" & vbLf', LAYOUT), (SET + '"a" & vbLf & vbLf', LAYOUT),
    (SET + '"one two three four five" & vbCr & "six"', LAYOUT), (SET + '"one two three four five" & vbCrLf & "six"', LAYOUT), (SET + '"one" & vbTab & "two three four five six seven"', LAYOUT), (SET + '"one two three four" & Chr(11) & "five"', LAYOUT), (SET + '"one two three four" & Chr(12) & "five"', LAYOUT),
    (SET + '"one two" & ChrW(8212) & "three four five six seven eight"', LAYOUT), (SET + 'ChrW(20013) & ChrW(25991) & ChrW(20013) & ChrW(25991) & ChrW(20013) & ChrW(25991) & ChrW(20013) & ChrW(25991) & ChrW(20013) & ChrW(25991)', LAYOUT), (SET + '"one two three" & ChrW(160) & "four five six seven eight"', LAYOUT), (SET + '"one two three four five" & ChrW(173) & "six seven eight"', LAYOUT),
    (SET + '"one/two/three/four/five/six/seven/eight/nine"', LAYOUT), (SET + '"one,two,three,four,five,six,seven,eight,nine"', LAYOUT), (SET + '"one_two_three_four_five_six_seven_eight_nine"', LAYOUT),
    ('', LAYOUT), ('s.Rotation = 90', LAYOUT), ('s.Rotation = 33', LAYOUT), ('s.Flip 0', LAYOUT), ('s.Height = 10', LAYOUT), ('s.TextFrame2.AutoSize = 1', LAYOUT), ('s.TextFrame2.AutoSize = 2', LAYOUT), ('s.TextFrame2.Orientation = 2', LAYOUT), ('s.TextFrame2.Orientation = 3', LAYOUT), ('s.TextFrame2.Orientation = 6', LAYOUT),
    ('s.TextFrame2.Column.Number = 2', LAYOUT), ('s.TextFrame2.VerticalAnchor = 3', LAYOUT), ('s.TextFrame2.HorizontalAnchor = 2', LAYOUT), ('s.TextFrame2.WordWrap = msoFalse', LAYOUT), ('s.TextFrame.AutoMargins = False', LAYOUT), ('s.TextFrame.MarginLeft = 20', LAYOUT), ('s.TextFrame.Orientation = 2', LAYOUT),
    (f'{T}.ParagraphFormat.Alignment = 2', LAYOUT), (f'{T}.ParagraphFormat.Alignment = 4', LAYOUT), (f'{T}.ParagraphFormat.LeftIndent = 20', LAYOUT), (f'{T}.ParagraphFormat.FirstLineIndent = 20', LAYOUT), (f'{T}.ParagraphFormat.RightIndent = 20', LAYOUT), (f'{T}.ParagraphFormat.IndentLevel = 2', LAYOUT),
    (f'{T}.ParagraphFormat.Bullet.Visible = msoTrue', LAYOUT), (f'{T}.Font.Spacing = 2', LAYOUT), (f'{T}.Font.Allcaps = msoTrue', LAYOUT), (f'{T}.Font.Smallcaps = msoTrue', LAYOUT), (f'{T}.Font.Superscript = msoTrue', LAYOUT), (f'{T}.Font.Size = 20', LAYOUT), (f'{T}.Font.Bold = msoTrue', LAYOUT),
    (f'{T}.Characters(1, 9).Font.Size = 20', LAYOUT), (f'{T}.Characters(11, 5).Font.Bold = msoTrue', LAYOUT), (f'{T}.Characters(11, 5).Font.Name = "Calibri"', LAYOUT), (f'{T}.Font.Name = "Arial"', LAYOUT), (f'{T}.Font.Name = "No Such Font"', LAYOUT), ('s.TextFrame.Characters.Font.Size = 20', LAYOUT), ('s.TextFrame.Characters(1, 9).Font.Size = 20', LAYOUT),
    ('s.DrawingObject.Font.Size = 20', LAYOUT), (f'{T}.Font.Name = "+mj-lt"', LAYOUT), (f'{T}.Font.Name = "+mn-lt"', LAYOUT),
    ('s.Width = 1', LAYOUT), ('s.Width = 14.4', LAYOUT), ('s.Width = 15', LAYOUT), ('s.TextFrame2.MarginLeft = 80: s.TextFrame2.MarginRight = 80', LAYOUT), ('s.AutoShapeType = 9', LAYOUT), ('s.AutoShapeType = 5', LAYOUT), ('s.AutoShapeType = 4', LAYOUT), ('s.AutoShapeType = 7', LAYOUT), ('s.AutoShapeType = 61', LAYOUT),
    ('s.Delete: Set s = ws.Shapes.AddLabel(1, 10, 10, 160, 100): ' + T + '.Text = "' + TEXT + '"', LAYOUT), ('s.Delete: Set s = ws.Shapes.AddTextbox(1, 10, 10, 160, 100): ' + T + '.Text = "' + TEXT + '"', LAYOUT), ('s.Delete: Set s = ws.Shapes.AddFormControl(xlButtonControl, 10, 10, 160, 100): s.TextFrame2.TextRange.Text = "' + TEXT + '"', LAYOUT),
    ('s.Delete: Set s = ws.Shapes.AddFormControl(xlLabel, 10, 10, 160, 100): s.TextFrame.Characters.Text = "' + TEXT + '"', LAYOUT), ('s.Delete: Set s = ws.Shapes.AddLine(10, 10, 160, 100)', f'{T}.Lines.Count'), ('s.Delete: Set s = ws.Shapes.AddConnector(1, 10, 10, 160, 100)', f'{T}.Lines.Count'),
]
SOURCE = r'''Public Function Layout(ByVal s As Object) As String
    Dim i As Long, r As String
    r = s.TextFrame2.TextRange.Lines.Count & ":"
    For i = 1 To s.TextFrame2.TextRange.Lines.Count
        r = r & s.TextFrame2.TextRange.Lines(i).Start & "+" & s.TextFrame2.TextRange.Lines(i).Length & ","
    Next
    Layout = "Width " & s.Width & " " & r
End Function

Public Function Go() As String
    Dim ws As Object, s As Object, t As Object, x As Variant, kind As String, code As Long, before As String
    Set ws = ActiveSheet
    before = ws.StandardHeight
    Set s = ws.Shapes.AddShape(1, 10, 10, 160, 100)
    s.TextFrame2.TextRange.Text = "%TEXT%"
    On Error Resume Next
    %SETUP%
    code = Err.Number
    Err.Clear
    x = %READ%
    If Err.Number <> 0 Then
        Go = before & "~" & ws.StandardHeight & "~" & code & "~" & Err.Number & "~~"
    Else
        Go = before & "~" & ws.StandardHeight & "~" & code & "~0~" & TypeName(x) & "~" & CStr(x)
    End If
    Err.Clear
    s.Delete
End Function
'''


#: Every shape in a file Excel saved: where its text breaks, read-only and with nothing saved.
LOADED = r'''Public Function Opened(ByVal path As String) As String
    Dim wb As Workbook, s As Object, r As String, i As Long, one As String, before As String
    before = ActiveSheet.StandardHeight
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, ReadOnly:=True)
    For Each s In wb.Worksheets(1).Shapes
        one = ""
        On Error Resume Next
        one = s.TextFrame2.TextRange.Lines.Count & ":"
        If Err.Number <> 0 Then
            one = "error " & Err.Number
        Else
            For i = 1 To s.TextFrame2.TextRange.Lines.Count
                one = one & s.TextFrame2.TextRange.Lines(i).Start & "+" & s.TextFrame2.TextRange.Lines(i).Length & ","
            Next
        End If
        Err.Clear
        On Error GoTo 0
        r = r & "|" & s.Name & "=" & s.Width & "=" & one
    Next
    wb.Close False
    Application.DisplayAlerts = True
    Opened = before & "~" & ActiveSheet.StandardHeight & r
End Function
'''


def source(setup: str, read: str) -> str:
    return SOURCE.replace('%TEXT%', TEXT).replace('%SETUP%', setup or "'").replace('%READ%', 'Layout(s)' if read == LAYOUT else read)


def main() -> None:
    answers = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for setup, read in CASES:
            result = excel.run_vba(source(setup, read), 'Go', (), timeout=60.0, line_numbers=False)
            assert result.ok and not result.dialogs, (setup, read, result)
            first, last, *found = str(result.value).split('~')
            assert (first, last) == ('15', '15'), f'Excel is not at 96 DPI: {first}, {last}'
            answers.append({'setup': setup, 'read': read, 'answer': '~'.join(found)})
        book = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'loaded_shape_text.xlsx'
        result = excel.run_vba(LOADED, 'Opened', (str(book),), timeout=120.0, line_numbers=False)
        assert result.ok and not result.dialogs, result
        heights, *rows = str(result.value).split('|')
        assert heights == '15~15', f'Excel is not at 96 DPI: {heights}'
        loaded = {row.split('=')[0]: row.split('=')[1:] for row in rows}
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'text': TEXT, 'source': SOURCE, 'cases': answers, 'loaded': loaded}, indent=0), encoding='utf-8', newline='\n')
    for one in answers[19:32]:
        print(one['setup'][:70], '|', one['read'][-34:], '=>', one['answer'][:90])
    for name, row in loaded.items():
        print(name, row)


if __name__ == '__main__':
    main()
