"""What a shape loaded from a file answers of its text, its connections and its picture.

    python scripts/measure_loaded_shape_text.py

Excel makes a workbook of shapes with settings beyond the defaults and saves it (tests/fixtures/loaded_shape_text.xlsx), then opens that file and reads every member below of every shape. The
model opens the same file and reads them one by one, so that what it refuses ends no other.

Writes tests/fixtures/loaded_shape_text.json. tests/test_excel_loaded_shape_text.py replays it.
"""
from __future__ import annotations

import json
import struct
import sys
import tempfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))
from fixture_workbook import copy_saved  # noqa: E402
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAUnsupportedError  # noqa: E402

OUT = ROOT / 'tests/fixtures/loaded_shape_text.json'
BOOK = ROOT / 'tests/fixtures/loaded_shape_text.xlsx'


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * 4 for _ in range(4))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 4, 4, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b'')


MAKE = r'''Public Function Make(ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet, s As Shape, top As Double
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Shapes.AddShape(1, 10, 10, 200, 60).Name = "NoText"
    Set s = ws.Shapes.AddShape(1, 10, 80, 200, 60): s.Name = "Plain": s.TextFrame.Characters.Text = "Hello world"
    Set s = ws.Shapes.AddTextbox(1, 10, 150, 200, 60): s.Name = "Box": s.TextFrame.Characters.Text = "Hello world"
    Set s = ws.Shapes.AddShape(1, 10, 220, 200, 60): s.Name = "Laid": s.TextFrame.Characters.Text = "Hello world"
    s.TextFrame2.MarginLeft = 12: s.TextFrame2.MarginTop = 6: s.TextFrame2.MarginRight = 3: s.TextFrame2.MarginBottom = 9
    s.TextFrame.HorizontalAlignment = xlHAlignCenter: s.TextFrame.VerticalAlignment = xlVAlignCenter
    s.TextFrame2.WordWrap = False: s.TextFrame2.AutoSize = 0
    Set s = ws.Shapes.AddShape(1, 10, 290, 200, 60): s.Name = "Grown": s.TextFrame.Characters.Text = "Hello": s.TextFrame2.AutoSize = 1
    Set s = ws.Shapes.AddShape(1, 10, 360, 200, 60): s.Name = "Shrunk": s.TextFrame.Characters.Text = "Hello": s.TextFrame2.AutoSize = 2
    Set s = ws.Shapes.AddShape(1, 10, 430, 200, 60): s.Name = "Bottom": s.TextFrame.Characters.Text = "Hello": s.TextFrame.VerticalAlignment = xlVAlignBottom: s.TextFrame.HorizontalAlignment = xlHAlignRight
    Set s = ws.Shapes.AddShape(1, 10, 500, 200, 60): s.Name = "Upward": s.TextFrame.Characters.Text = "Hello": s.TextFrame2.Orientation = 2
    Set s = ws.Shapes.AddShape(1, 10, 570, 200, 60): s.Name = "Clipped": s.TextFrame.Characters.Text = "Hello": s.TextFrame2.VerticalAnchor = 3: s.TextFrame2.HorizontalAnchor = 2: s.TextFrame2.NoTextRotation = True
    s.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowClip: s.TextFrame.VerticalOverflow = xlOartVerticalOverflowEllipsis
    Set s = ws.Shapes.AddShape(1, 220, 10, 200, 60): s.Name = "Runs": s.TextFrame.Characters.Text = "Hello world again"
    s.TextFrame.Characters(1, 5).Font.Bold = True
    s.TextFrame.Characters(7, 5).Font.Italic = True
    s.TextFrame.Characters(13, 5).Font.Color = RGB(10, 20, 30)
    Set s = ws.Shapes.AddTextbox(1, 220, 80, 200, 60): s.Name = "Fonts": s.TextFrame.Characters.Text = "Hello world again more"
    s.TextFrame.Characters(1, 5).Font.Name = "Arial": s.TextFrame.Characters(1, 5).Font.Size = 20
    s.TextFrame.Characters(7, 5).Font.Underline = xlUnderlineStyleSingle: s.TextFrame.Characters(7, 5).Font.Strikethrough = True
    s.TextFrame.Characters(13, 5).Font.Superscript = True
    s.TextFrame.Characters(19, 4).Font.Subscript = True
    Set s = ws.Shapes.AddShape(1, 220, 150, 200, 60): s.Name = "Whole": s.TextFrame.Characters.Text = "All one look"
    s.TextFrame.Characters.Font.Bold = True: s.TextFrame.Characters.Font.Size = 14: s.TextFrame.Characters.Font.Color = RGB(200, 100, 50): s.TextFrame.Characters.Font.Name = "Arial"
    Set s = ws.Shapes.AddShape(1, 220, 220, 200, 60): s.Name = "Paint": s.TextFrame.Characters.Text = "Painted text"
    With s.TextFrame2.TextRange.Characters(1, 7).Font
        .Fill.ForeColor.ObjectThemeColor = 5
        .Line.Visible = True: .Line.ForeColor.RGB = RGB(1, 2, 3): .Line.Weight = 2
        .Spacing = 3: .Kerning = 12: .Allcaps = True: .UnderlineStyle = 3
    End With
    s.TextFrame2.TextRange.Characters(8, 5).Font.Fill.Transparency = 0.5
    Set s = ws.Shapes.AddShape(1, 220, 290, 200, 60): s.Name = "Paras": s.TextFrame.Characters.Text = "First para" & vbCr & "Second one" & vbCr & "Third"
    s.TextFrame2.TextRange.Paragraphs(2).ParagraphFormat.Alignment = 2
    s.TextFrame2.TextRange.Paragraphs(2).Font.Bold = True
    Set s = ws.Shapes.AddShape(1, 220, 360, 200, 60): s.Name = "Broken": s.TextFrame.Characters.Text = "One" & vbLf & "Two": s.TextFrame2.TextRange.Characters(1, 3).Font.Size = 8
    Set s = ws.Shapes.AddShape(9, 220, 430, 100, 60): s.Name = "Oval": s.TextFrame.Characters.Text = "Round"
    Dim g As Shape, a As Shape, b As Shape
    Set a = ws.Shapes.AddShape(1, 450, 10, 60, 30): a.Name = "End1": a.TextFrame.Characters.Text = "One"
    Set b = ws.Shapes.AddShape(1, 450, 100, 60, 30): b.Name = "End2"
    Set s = ws.Shapes.AddConnector(1, 0, 0, 10, 10): s.Name = "Joined": s.ConnectorFormat.BeginConnect a, 3: s.ConnectorFormat.EndConnect b, 1
    Set s = ws.Shapes.AddConnector(2, 450, 150, 100, 40): s.Name = "Elbow": s.ConnectorFormat.BeginConnect a, 2
    Set s = ws.Shapes.AddConnector(3, 450, 200, 100, 40): s.Name = "Curve"
    Set s = ws.Shapes.AddLine(450, 250, 550, 260): s.Name = "Line"
    Set s = ws.Shapes.AddPicture(folder & "\p.png", False, True, 600, 10, 60, 60): s.Name = "Pic"
    Set s = ws.Shapes.AddPicture(folder & "\p.png", False, True, 600, 80, 60, 60): s.Name = "Tuned"
    s.PictureFormat.Brightness = 0.7: s.PictureFormat.Contrast = 0.3: s.PictureFormat.ColorType = 2
    s.PictureFormat.CropLeft = 5: s.PictureFormat.CropTop = 6: s.PictureFormat.CropRight = 7: s.PictureFormat.CropBottom = 8
    Set s = ws.Shapes.AddPicture(folder & "\p.png", False, True, 600, 150, 60, 60): s.Name = "Clear"
    s.PictureFormat.TransparentBackground = True: s.PictureFormat.TransparencyColor = RGB(255, 0, 0)
    Set s = ws.Shapes.AddShape(1, 700, 10, 60, 30): s.Name = "Inner1": s.TextFrame.Characters.Text = "Grouped text": s.TextFrame.Characters(1, 7).Font.Bold = True
    Set s = ws.Shapes.AddShape(1, 700, 50, 60, 30): s.Name = "Inner2"
    ws.Shapes.Range(Array("Inner1", "Inner2")).Group.Name = "Grp"
    wb.SaveAs folder & "\book.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
TEXT_SHAPES = ('NoText', 'Plain', 'Box', 'Laid', 'Grown', 'Shrunk', 'Bottom', 'Upward', 'Clipped', 'Runs', 'Fonts', 'Whole', 'Paint', 'Paras', 'Broken', 'Oval', 'End1', 'Grp.GroupItems("Inner1")')
FRAME = ('HorizontalAlignment', 'VerticalAlignment', 'MarginLeft', 'MarginTop', 'MarginRight', 'MarginBottom', 'Orientation', 'AutoSize', 'AutoMargins', 'HorizontalOverflow', 'VerticalOverflow', 'ReadingOrder', 'Characters.Text', 'Characters.Count')
FRAME2 = ('HasText', 'WordWrap', 'AutoSize', 'MarginLeft', 'MarginTop', 'MarginRight', 'MarginBottom', 'Orientation', 'VerticalAnchor', 'HorizontalAnchor', 'NoTextRotation', 'PathFormat', 'WarpFormat',
          'TextRange.Text', 'TextRange.Length', 'TextRange.Start', 'TextRange.Count', 'TextRange.Runs.Count', 'TextRange.Paragraphs.Count', 'TextRange.Words.Count', 'TextRange.Sentences.Count', 'TextRange.Characters.Count', 'TextRange.Lines.Count',
          'TextRange.Runs(1).Text', 'TextRange.Runs(2).Text', 'TextRange.Runs(2).Start', 'TextRange.Paragraphs(1).Text', 'TextRange.Paragraphs(2).Text', 'TextRange.Words(2).Text', 'TextRange.LanguageID')
PARAGRAPH = ('Alignment', 'BaselineAlignment', 'FirstLineIndent', 'LeftIndent', 'RightIndent', 'SpaceAfter', 'SpaceBefore', 'SpaceWithin', 'LineRuleAfter', 'LineRuleBefore', 'LineRuleWithin', 'WordWrap', 'FarEastLineBreakLevel', 'HangingPunctuation', 'TextDirection', 'IndentLevel', 'Bullet.Visible')
FONT = ('Name', 'Size', 'Bold', 'Italic', 'Underline', 'Strikethrough', 'Superscript', 'Subscript', 'Color', 'ColorIndex', 'FontStyle', 'Shadow', 'OutlineFont', 'TintAndShade', 'ThemeColor', 'ThemeFont')
FONT2 = ('Name', 'Size', 'Bold', 'Italic', 'UnderlineStyle', 'UnderlineColor.RGB', 'Strike', 'StrikeThrough', 'DoubleStrikeThrough', 'Superscript', 'Subscript', 'BaselineOffset', 'Caps', 'Allcaps', 'Smallcaps', 'Kerning', 'Spacing', 'Equalize', 'Embossed', 'Engrave', 'Shadow.Visible', 'Glow.Radius', 'Reflection.Type', 'Highlight.RGB',
         'NameAscii', 'NameComplexScript', 'NameFarEast', 'NameOther', 'AutorotateNumbers', 'Embeddable', 'Embedded', 'SoftEdgeFormat', 'WordArtformat', 'Fill.Visible', 'Fill.Type', 'Fill.ForeColor.RGB', 'Fill.ForeColor.ObjectThemeColor', 'Fill.ForeColor.Type', 'Fill.Transparency', 'Line.Visible', 'Line.ForeColor.RGB', 'Line.Weight', 'Line.DashStyle')
CONNECTOR = ('Connector', 'ConnectorFormat.Type', 'ConnectorFormat.BeginConnected', 'ConnectorFormat.EndConnected', 'ConnectorFormat.BeginConnectedShape.Name', 'ConnectorFormat.EndConnectedShape.Name', 'ConnectorFormat.BeginConnectionSite', 'ConnectorFormat.EndConnectionSite', 'ConnectionSiteCount', 'HorizontalFlip', 'VerticalFlip', 'Left', 'Top', 'Width', 'Height')
PICTURE = ('PictureFormat.Brightness', 'PictureFormat.Contrast', 'PictureFormat.ColorType', 'PictureFormat.TransparentBackground', 'PictureFormat.TransparencyColor', 'PictureFormat.CropLeft', 'PictureFormat.CropTop', 'PictureFormat.CropRight', 'PictureFormat.CropBottom',
           'PictureFormat.Crop.PictureWidth', 'PictureFormat.Crop.PictureHeight', 'PictureFormat.Crop.PictureOffsetX', 'PictureFormat.Crop.PictureOffsetY', 'PictureFormat.Crop.ShapeWidth', 'PictureFormat.Crop.ShapeHeight', 'PictureFormat.Crop.ShapeLeft', 'PictureFormat.Crop.ShapeTop', 'LockAspectRatio', 'Width', 'Height', 'Left', 'Top')


def reads() -> list[str]:
    out: list[str] = []
    for shape in TEXT_SHAPES:
        head = f'ws.Shapes("{shape}")' if '.' not in shape else f'ws.Shapes("Grp").GroupItems("Inner1")'
        out += [f'{head}.TextFrame.{member}' for member in FRAME]
        out += [f'{head}.TextFrame2.{member}' for member in FRAME2]
        out += [f'{head}.TextFrame2.TextRange.ParagraphFormat.{member}' for member in PARAGRAPH]
        out += [f'{head}.TextFrame.Characters.Font.{member}' for member in FONT]
        out += [f'{head}.TextFrame2.TextRange.Font.{member}' for member in FONT2]
    for shape, spans in (('Runs', ((1, 5), (7, 5), (13, 5), (1, 17), (5, 3))), ('Fonts', ((1, 5), (7, 5), (13, 5), (19, 4), (1, 22))), ('Paint', ((1, 7), (8, 5), (1, 12))), ('Paras', ((1, 10), (12, 10), (23, 5))), ('Broken', ((1, 3), (5, 3)))):
        for start, length in spans:
            out += [f'ws.Shapes("{shape}").TextFrame.Characters({start}, {length}).Text', f'ws.Shapes("{shape}").TextFrame.Characters({start}, {length}).Count']
            out += [f'ws.Shapes("{shape}").TextFrame.Characters({start}, {length}).Font.{member}' for member in FONT]
            out += [f'ws.Shapes("{shape}").TextFrame2.TextRange.Characters({start}, {length}).Font.{member}' for member in FONT2]
    out += [f'ws.Shapes("Paras").TextFrame2.TextRange.Paragraphs({n}).ParagraphFormat.{member}' for n in (1, 2, 3) for member in PARAGRAPH]
    for shape in ('Joined', 'Elbow', 'Curve', 'Line'):
        out += [f'ws.Shapes("{shape}").{member}' for member in CONNECTOR]
    for shape in ('Pic', 'Tuned', 'Clear'):
        out += [f'ws.Shapes("{shape}").{member}' for member in PICTURE]
    return out


READS = reads()
SHOWN = r'''Private Function Shown(v As Variant) As String
    If IsObject(v) Then
        Shown = "object"
    ElseIf IsArray(v) Or IsNull(v) Then
        Shown = TypeName(v)
    Else
        Shown = CStr(v)
    End If
End Function
'''


def reader(reads_: list[str]) -> str:
    lines = [SHOWN]
    names: list[str] = []
    for start in range(0, len(reads_), 60):
        name = f'Part{start}'
        names.append(name)
        lines += [f'Private Function {name}(ws As Object) As String', '    Dim v As Variant, s As String', '    On Error Resume Next']
        for read in reads_[start:start + 60]:
            lines += ['    s = s & "@@"', '    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        s = s & "E" & Err.Number: Err.Clear', '    Else', '        s = s & Shown(v)', '    End If']
        lines += [f'    {name} = s', 'End Function']
    lines += ['Public Function Opened(ByVal path As String) As String', '    Dim wb As Workbook, s As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)']
    lines += [f'    s = s & {name}(wb.Worksheets(1))' for name in names]
    lines += ['    wb.Close False', '    Application.DisplayAlerts = True', '    Opened = s', 'End Function']
    return '\n'.join(lines) + '\n'


def one_read(number: int, read: str) -> str:
    return '\n'.join([f'Public Function R{number}() As String', '    Dim ws As Object, v As Variant', '    Set ws = ActiveWorkbook.Worksheets(1)', '    On Error Resume Next', f'    v = {read}', '    If Err.Number Then', f'        R{number} = "E" & Err.Number',
                      '    ElseIf IsObject(v) Then', f'        R{number} = "object"', '    ElseIf IsArray(v) Or IsNull(v) Then', f'        R{number} = TypeName(v)', '    Else', f'        R{number} = CStr(v)', '    End If', 'End Function', ''])


def in_model() -> list[str]:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(''.join(one_read(number, read) for number, read in enumerate(READS)), name='Reader')
    answers: list[str] = []
    for number in range(len(READS)):
        try:
            answers.append(str(app.run(f'Reader.R{number}')))
        except VBAUnsupportedError as refused:
            answers.append(f'UNSUPPORTED {str(refused)[:80]}')
    return answers


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        (base / 'p.png').write_bytes(png())
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (folder,), timeout=120.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(base / 'book.xlsx', BOOK)
            read = excel.run_vba(reader(READS), 'Opened', (str(BOOK),), timeout=280.0, line_numbers=False)
            assert read.ok, read
            truth = str(read.value).split('@@')[1:]
    assert len(truth) == len(READS), (len(truth), len(READS))
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'reads': READS, 'excel': truth}, indent=0) + '\n', encoding='utf-8')
    model = in_model()
    wrong = sum(1 for a, b in zip(truth, model) if a != b and not b.startswith('UNSUPPORTED'))
    print(len(READS), 'reads; refused', sum(b.startswith('UNSUPPORTED') for b in model), 'wrong', wrong)


if __name__ == '__main__':
    main()
