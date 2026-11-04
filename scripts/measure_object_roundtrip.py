"""Whether the settings of a sheet's objects survive the model. Excel makes a workbook whose shapes, tables, comments, hyperlinks, names, validation
and conditional formats use settings beyond the defaults, and saves it; the model opens that file and reads every setting, makes an unrelated edit,
and saves; and Excel opens what the model saved and reads them again.

    python scripts/measure_object_roundtrip.py

writes tests/fixtures/object_roundtrip.json, with what Excel read from its own file and from the model's for each read, and copies Excel's file to
tests/fixtures/object_roundtrip.xlsx. tests/test_excel_object_roundtrip.py replays them. Excel makes each edit to its own file too: a difference between
what it then reads from its save and from the model's is something the model lost or changed, which only reopening shows.
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
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError  # noqa: E402

OUT = ROOT / 'tests/fixtures/object_roundtrip.json'
BOOK = ROOT / 'tests/fixtures/object_roundtrip.xlsx'
MAKE = r'''
Private Problems As String

Private Sub MakeShapes(ws As Object, folder As String)
    Dim s As Object
    Set s = ws.Shapes.AddShape(1, 10, 10, 100, 40): s.Name = "Hidden": s.Visible = msoFalse
    Set s = ws.Shapes.AddShape(1, 10, 60, 100, 40): s.Name = "Unlocked": s.Locked = msoFalse
    Set s = ws.Shapes.AddShape(1, 10, 110, 100, 40): s.Name = "NoPrint": s.DrawingObject.PrintObject = False
    Set s = ws.Shapes.AddShape(1, 10, 160, 100, 40): s.Name = "Moves": s.Placement = xlMove
    Set s = ws.Shapes.AddShape(1, 10, 210, 100, 40): s.Name = "Floats": s.Placement = xlFreeFloating
    Set s = ws.Shapes.AddShape(1, 10, 260, 100, 40): s.Name = "Described": s.AlternativeText = "Alt words": s.Title = "A title"
    Set s = ws.Shapes.AddShape(1, 150, 10, 100, 40): s.Name = "Turned": s.Rotation = 30: s.Flip msoFlipHorizontal: s.LockAspectRatio = msoTrue
    Set s = ws.Shapes.AddShape(1, 150, 60, 100, 40): s.Name = "Macro": s.OnAction = "Nothing"
    Set s = ws.Shapes.AddShape(1, 150, 110, 100, 40): s.Name = "Painted": s.Fill.ForeColor.RGB = 255: s.Fill.Transparency = 0.5: s.Line.ForeColor.RGB = 65280: s.Line.Weight = 3: s.Line.DashStyle = 4: s.Shadow.Visible = msoTrue
    Set s = ws.Shapes.AddShape(1, 150, 160, 100, 40): s.Name = "Texted": s.TextFrame.Characters.Text = "Hello World": s.TextFrame.Characters(1, 5).Font.Bold = True: s.TextFrame.HorizontalAlignment = xlHAlignCenter: s.TextFrame.VerticalAlignment = xlVAlignCenter: s.TextFrame.MarginLeft = 12: s.TextFrame2.WordWrap = msoFalse
    Set s = ws.Shapes.AddShape(1, 150, 210, 100, 40): s.Name = "Linked": ws.Hyperlinks.Add s, "http://example.com/", , "A tip"
    Set s = ws.Shapes.AddTextbox(1, 150, 260, 100, 40): s.Name = "Box": s.TextFrame.Characters.Text = "Text": s.Locked = msoFalse: s.TextFrame2.AutoSize = 0
    Dim a As Object, b As Object, c As Object
    Set a = ws.Shapes.AddShape(1, 300, 10, 50, 50): a.Name = "End1"
    Set b = ws.Shapes.AddShape(1, 400, 100, 50, 50): b.Name = "End2"
    Set c = ws.Shapes.AddConnector(1, 1, 1, 2, 2): c.Name = "Joined": c.ConnectorFormat.BeginConnect a, 4: c.ConnectorFormat.EndConnect b, 2: c.Line.EndArrowheadStyle = 2
    Set a = ws.Shapes.AddShape(1, 300, 200, 50, 50): a.Name = "G1"
    Set b = ws.Shapes.AddShape(9, 360, 200, 50, 50): b.Name = "G2": b.Locked = msoFalse
    Set c = ws.Shapes.Range(Array("G1", "G2")).Group: c.Name = "Grouped": c.Placement = xlMove
    Set s = ws.Shapes.AddPicture(folder & "\p.png", msoFalse, msoTrue, 500, 10, 40, 40): s.Name = "Pic": s.PictureFormat.Brightness = 0.7: s.Locked = msoFalse: s.Placement = xlFreeFloating
    ws.Buttons.Add(500, 100, 60, 20).Name = "Btn"
End Sub

Private Sub MakeTables(ws As Object)
    Dim t As Object
    ws.Range("A1:C1").Value = Array("a", "b", "c"): ws.Range("A2:C2").Value = Array(1, 2, 3): ws.Range("A3:C3").Value = Array(4, 5, 6)
    Set t = ws.ListObjects.Add(1, ws.Range("A1:C3"), , 1): t.Name = "Styled"
    t.TableStyle = "TableStyleMedium9": t.ShowTableStyleRowStripes = False: t.ShowTableStyleColumnStripes = True: t.ShowTableStyleFirstColumn = True: t.ShowTableStyleLastColumn = True
    t.ShowTotals = True: t.ListColumns(2).TotalsCalculation = xlTotalsCalculationAverage: t.ListColumns(1).Total.Value = "Label"
    t.AlternativeText = "Table alt": t.Summary = "Table summary": t.Comment = "Table comment"
    ws.Range("E1:F1").Value = Array("x", "y"): ws.Range("E2:F2").Value = Array(1, 2)
    Set t = ws.ListObjects.Add(1, ws.Range("E1:F2"), , 1): t.Name = "Plain": t.TableStyle = "": t.ShowAutoFilterDropDown = False
    ws.Range("H1:I1").Value = Array("p", "q"): ws.Range("H2:I2").Value = Array(1, 2)
    Set t = ws.ListObjects.Add(1, ws.Range("H1:I2"), , 1): t.Name = "NoHeader": t.ShowHeaders = False
    ws.Range("K1:L1").Value = Array("m", "n"): ws.Range("K2:L2").Value = Array(1, 2)
    Set t = ws.ListObjects.Add(1, ws.Range("K1:L2"), , 1): t.Name = "NoFilter": t.ShowAutoFilter = False
End Sub

Private Sub MakeComments(ws As Object)
    Dim n As Object
    Set n = ws.Range("A10").AddComment("Shown note"): n.Visible = True: n.Shape.Width = 200: n.Shape.Height = 80
    Set n = ws.Range("A11").AddComment("Hidden note"): n.Shape.Fill.ForeColor.RGB = 65535
    Set n = ws.Range("A12").AddComment("Locked off"): n.Shape.Locked = msoFalse: n.Shape.Placement = xlFreeFloating
    ws.Range("A13").AddComment("Bold words"): ws.Range("A13").Comment.Shape.TextFrame.Characters(1, 4).Font.Bold = True
End Sub

Private Sub MakeLinks(ws As Object)
    ws.Hyperlinks.Add ws.Range("B10"), "http://example.com/a", "section", "A tip", "Shown"
    ws.Hyperlinks.Add ws.Range("B11"), "", "Data!A1", "Inside", "Inner"
    ws.Hyperlinks.Add ws.Range("B12"), "mailto:someone@example.com?subject=Hi"
End Sub

Private Sub MakeNames(wb As Object, ws As Object)
    Dim n As Object
    Set n = wb.Names.Add("HiddenName", "=Data!$A$1"): n.Visible = False
    Set n = wb.Names.Add("Commented", "=Data!$A$2"): n.Comment = "A name's comment"
    Set n = ws.Names.Add("LocalName", "=Data!$A$3"): n.Visible = False
    Set n = wb.Names.Add("Constant", "=42")
    Set n = wb.Names.Add("Macroish", "=Data!$A$1"): n.MacroType = xlFunction: n.Category = "Financial"
    Set n = wb.Names.Add("Keyed", "=Data!$A$1"): n.MacroType = xlCommand: n.ShortcutKey = "k"
End Sub

Private Sub MakeValidation(ws As Object)
    With ws.Range("C10").Validation
        .Add xlValidateList, xlValidAlertWarning, , "a,b,c"
        .ErrorTitle = "Err title": .ErrorMessage = "Err words": .InputTitle = "In title": .InputMessage = "In words": .IgnoreBlank = False: .InCellDropdown = False: .ShowInput = False: .ShowError = False
    End With
    With ws.Range("C11").Validation
        .Add xlValidateWholeNumber, xlValidAlertInformation, xlBetween, "1", "9"
        .ErrorMessage = "One to nine"
    End With
    ws.Range("C12").Validation.Add xlValidateDecimal, xlValidAlertStop, xlNotEqual, "1.5"
    ws.Range("C13").Validation.Add xlValidateDate, xlValidAlertStop, xlGreater, "1/2/2020"
    ws.Range("C14").Validation.Add xlValidateTextLength, xlValidAlertStop, xlLessEqual, "5"
    ws.Range("C15").Validation.Add xlValidateCustom, xlValidAlertStop, , "=A1>0"
    ws.Range("C16").Validation.Add xlValidateList, xlValidAlertStop, , "=$A$1:$A$3"
    ws.Range("C17").Validation.Add xlValidateTime, xlValidAlertStop, xlLess, "0.5"
    ws.Range("C18").Validation.Add xlValidateInputOnly
    ws.Range("C18").Validation.InputMessage = "Just a note"
    ws.Range("C18").Validation.IMEMode = xlIMEModeOn
End Sub

Private Sub MakeFormats(ws As Object)
    Dim f As Object
    On Error Resume Next
    Set f = ws.Range("D10:D12").FormatConditions.Add(xlCellValue, xlGreater, "=5"): f.Font.Bold = True: f.Font.Color = 255: f.Interior.Color = 65535: f.StopIfTrue = True: f.NumberFormat = "0.00"
    If Err.Number Then Problems = Problems & "|line3:" & Err.Number & " " & Err.Description: Err.Clear
    f.Borders(xlBottom).LineStyle = xlContinuous
    If Err.Number Then Problems = Problems & "|line4:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("D10:D12").FormatConditions.Add(xlExpression, , "=$A$1>0"): f.Font.Italic = True: f.StopIfTrue = False
    If Err.Number Then Problems = Problems & "|line5:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("D13:D15").FormatConditions.AddColorScale(3): f.ColorScaleCriteria(1).FormatColor.Color = 255: f.ColorScaleCriteria(2).Type = xlConditionValuePercentile: f.ColorScaleCriteria(2).Value = 40
    If Err.Number Then Problems = Problems & "|line6:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("D16:D18").FormatConditions.AddDatabar: f.ShowValue = False: f.BarColor.Color = 255: f.BarFillType = xlDataBarFillSolid: f.Direction = xlRTL: f.BarBorder.Type = xlDataBarBorderSolid: f.MinPoint.Modify xlConditionValueNumber, 1: f.MaxPoint.Modify xlConditionValueNumber, 9
    If Err.Number Then Problems = Problems & "|line7:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("D19:D21").FormatConditions.AddIconSetCondition: f.IconSet = ws.Parent.IconSets(xl5Arrows): f.ShowIconOnly = True: f.ReverseOrder = True
    If Err.Number Then Problems = Problems & "|line8:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("E10:E12").FormatConditions.AddTop10: f.TopBottom = xlTop10Bottom: f.Percent = True: f.Rank = 15: f.Interior.Color = 255
    If Err.Number Then Problems = Problems & "|line9:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("E13:E15").FormatConditions.AddAboveAverage: f.AboveBelow = xlAboveStdDev: f.NumStdDev = 2: f.Font.Bold = True
    If Err.Number Then Problems = Problems & "|line10:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("E16:E18").FormatConditions.AddUniqueValues: f.DupeUnique = xlDuplicate: f.Interior.Color = 65280
    If Err.Number Then Problems = Problems & "|line11:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("E19:E21").FormatConditions.Add(Type:=xlTextString, String:="abc", TextOperator:=xlContains): f.Font.Underline = xlUnderlineStyleSingle
    If Err.Number Then Problems = Problems & "|line12:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("F10:F12").FormatConditions.Add(Type:=xlTimePeriod, DateOperator:=xlLast7Days): f.Font.Strikethrough = True
    If Err.Number Then Problems = Problems & "|line13:" & Err.Number & " " & Err.Description: Err.Clear
    Set f = ws.Range("F13:F15").FormatConditions.Add(xlBlanksCondition): f.Interior.Pattern = xlGray25: f.Interior.PatternColor = 255
    If Err.Number Then Problems = Problems & "|line14:" & Err.Number & " " & Err.Description: Err.Clear
End Sub

Public Function Make(ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Data"
    On Error GoTo Failed
    Make = "shapes": MakeShapes ws, folder
    Make = "tables": MakeTables ws
    Make = "comments": MakeComments ws
    Make = "links": MakeLinks ws
    Make = "names": MakeNames wb, ws
    Make = "validation": MakeValidation ws
    Make = "formats": MakeFormats ws
    wb.SaveAs folder & "\excel.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok" & Problems
    Exit Function
Failed:
    Make = Make & " failed " & Err.Number & " " & Err.Description
    Application.DisplayAlerts = True
End Function
'''
SHAPE = ['Visible', 'Locked', 'Placement', 'AlternativeText', 'Title', 'Rotation', 'HorizontalFlip', 'VerticalFlip', 'LockAspectRatio', 'OnAction', 'Left', 'Top', 'Width', 'Height', 'Type', 'AutoShapeType', 'Connector',
         'DrawingObject.PrintObject', 'ZOrderPosition', 'Fill.Visible', 'Fill.ForeColor.RGB', 'Fill.Transparency', 'Line.Visible', 'Line.ForeColor.RGB', 'Line.Weight', 'Line.DashStyle', 'Shadow.Visible']
TEXT = ['TextFrame.Characters.Text', 'TextFrame.HorizontalAlignment', 'TextFrame.VerticalAlignment', 'TextFrame.MarginLeft', 'TextFrame2.WordWrap', 'TextFrame2.AutoSize', 'TextFrame.Characters(1, 5).Font.Bold', 'TextFrame.Characters(7, 5).Font.Bold']
SHAPES = ['Hidden', 'Unlocked', 'NoPrint', 'Moves', 'Floats', 'Described', 'Turned', 'Macro', 'Painted', 'Texted', 'Linked', 'Box', 'Joined', 'Grouped', 'Pic', 'Btn']
READS: list[str] = [f'ws.Shapes("{name}").{one}' for name in SHAPES for one in SHAPE]
READS += [f'ws.Shapes("{name}").{one}' for name in ('Texted', 'Box') for one in TEXT]
READS += ['ws.Shapes("Linked").Hyperlink.Address', 'ws.Shapes("Linked").Hyperlink.ScreenTip', 'ws.Shapes.Count', 'ws.Shapes("Joined").ConnectorFormat.BeginConnected', 'ws.Shapes("Joined").ConnectorFormat.BeginConnectedShape.Name', 'ws.Shapes("Joined").ConnectorFormat.BeginConnectionSite',
          'ws.Shapes("Joined").ConnectorFormat.EndConnected', 'ws.Shapes("Joined").ConnectorFormat.EndConnectedShape.Name', 'ws.Shapes("Joined").ConnectorFormat.EndConnectionSite', 'ws.Shapes("Joined").Line.EndArrowheadStyle',
          'ws.Shapes("Grouped").GroupItems.Count', 'ws.Shapes("Grouped").GroupItems(2).Name', 'ws.Shapes("Grouped").GroupItems(2).Locked', 'ws.Shapes("Pic").PictureFormat.Brightness', 'ws.Buttons("Btn").Caption', 'ws.Buttons("Btn").PrintObject', 'ws.Buttons("Btn").Enabled']
TABLE = ['Name', 'DisplayName', 'Range.Address', 'TableStyle', 'ShowTableStyleRowStripes', 'ShowTableStyleColumnStripes', 'ShowTableStyleFirstColumn', 'ShowTableStyleLastColumn', 'ShowTotals', 'ShowHeaders', 'ShowAutoFilter', 'ShowAutoFilterDropDown', 'AlternativeText', 'Summary', 'Comment']
READS += [f'ws.ListObjects("{name}").{one}' for name in ('Styled', 'Plain', 'NoHeader', 'NoFilter') for one in TABLE]
READS += ['ws.ListObjects("Styled").ListColumns(2).TotalsCalculation', 'ws.ListObjects("Styled").ListColumns(1).TotalsCalculation', 'ws.ListObjects("Styled").ListColumns(1).Total.Value', 'ws.ListObjects("Styled").ListColumns(2).Total.Formula']
NOTE = ['Comment.Text', 'Comment.Visible', 'Comment.Shape.Width', 'Comment.Shape.Height', 'Comment.Shape.Locked', 'Comment.Shape.Placement', 'Comment.Shape.Fill.ForeColor.RGB', 'Comment.Shape.Visible', 'Comment.Shape.Name']
READS += [f'ws.Range("{cell}").{one}' for cell in ('A10', 'A11', 'A12', 'A13') for one in NOTE]
READS += ['ws.Range("A13").Comment.Shape.TextFrame.Characters(1, 4).Font.Bold', 'ws.Comments.Count']
LINK = ['Address', 'SubAddress', 'ScreenTip', 'TextToDisplay', 'Name', 'Range.Address']
READS += [f'ws.Hyperlinks({at}).{one}' for at in (1, 2, 3, 4) for one in LINK] + ['ws.Hyperlinks.Count', 'ws.Range("B10").Value', 'ws.Range("B10").Style', 'ws.Range("B11").Value']
NAME = ['Name', 'RefersTo', 'Visible', 'Comment', 'MacroType', 'Category', 'ShortcutKey', 'Value']
READS += [f'wb.Names("{name}").{one}' for name in ('HiddenName', 'Commented', 'Data!LocalName', 'Constant', 'Macroish', 'Keyed') for one in NAME] + ['wb.Names.Count', 'ws.Names.Count']
VALID = ['Type', 'AlertStyle', 'Operator', 'Formula1', 'Formula2', 'ErrorTitle', 'ErrorMessage', 'InputTitle', 'InputMessage', 'IgnoreBlank', 'InCellDropdown', 'ShowInput', 'ShowError', 'IMEMode']
READS += [f'ws.Range("C{row}").Validation.{one}' for row in range(10, 19) for one in VALID]
COMMON = ['Type', 'Priority', 'StopIfTrue', 'AppliesTo.Address']
STYLE = ['Font.Bold', 'Font.Italic', 'Font.Color', 'Font.Underline', 'Font.Strikethrough', 'Interior.Color', 'Interior.Pattern', 'Interior.PatternColor', 'NumberFormat', 'Borders(xlBottom).LineStyle', 'Operator', 'Formula1']
RULES = {'D10': (1, ''), 'E10': (1, 'TopBottom,Percent,Rank'), 'E13': (1, 'AboveBelow,NumStdDev'), 'E16': (1, 'DupeUnique'), 'E19': (1, 'Text,TextOperator'), 'F10': (1, 'DateOperator'), 'F13': (1, '')}
READS += [f'ws.Range("{cell}").FormatConditions.Count' for cell in ('D10', 'D13', 'D16', 'D19', 'E10', 'E13', 'E16', 'E19', 'F10', 'F13')]
READS += [f'ws.Range("D10").FormatConditions(2).{one}' for one in (*COMMON, 'Font.Italic', 'Formula1')]
for cell, (at, extra) in RULES.items():
    READS += [f'ws.Range("{cell}").FormatConditions({at}).{one}' for one in (*COMMON, *STYLE, *(extra.split(',') if extra else []))]
READS += [f'ws.Range("D13").FormatConditions(1).{one}' for one in (*COMMON, 'ColorScaleCriteria.Count', 'ColorScaleCriteria(1).FormatColor.Color', 'ColorScaleCriteria(1).Type', 'ColorScaleCriteria(2).Type', 'ColorScaleCriteria(2).Value', 'ColorScaleCriteria(3).Type')]
READS += [f'ws.Range("D16").FormatConditions(1).{one}' for one in (*COMMON, 'ShowValue', 'BarColor.Color', 'BarFillType', 'Direction', 'BarBorder.Type', 'MinPoint.Type', 'MinPoint.Value', 'MaxPoint.Type', 'MaxPoint.Value', 'PercentMin', 'PercentMax', 'AxisPosition', 'NegativeBarFormat.ColorType')]
READS += [f'ws.Range("D19").FormatConditions(1).{one}' for one in (*COMMON, 'ShowIconOnly', 'ReverseOrder', 'IconSet.ID', 'PercentileValues', 'IconCriteria.Count', 'IconCriteria(2).Type', 'IconCriteria(2).Value', 'IconCriteria(2).Operator')]
#: The edits a save is tried after: one no object has anything to do with, then ones that could reach them.
EDITS = {'value': 'ws.Range("Z100").Value = "edited"', 'sheet': 'wb.Worksheets.Add After:=ws', 'row_far': 'ws.Rows(200).Insert', 'row_near': 'ws.Rows(1).Insert', 'column_near': 'ws.Columns(1).Insert',
         'rename': 'ws.Name = "Renamed"', 'twice': 'ws.Range("Z100").Value = 1: ws.Range("Z100").ClearContents'}
EDITOR = '''Public Function Edited(ByVal path As String, ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path): Set ws = wb.Worksheets(1)
    {edit}
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Edited = "ok"
End Function
'''


def _reader() -> str:
    lines: list[str] = []
    parts = [READS[at:at + 60] for at in range(0, len(READS), 60)]
    for number, part in enumerate(parts):
        lines += [f'Private Function Read{number}(wb As Object, ws As Object) As String', '    Dim s As String, v As Variant', '    On Error Resume Next']
        for read in part:
            lines += ['    Err.Clear', '    v = Empty', f'    v = {read}', '    If Err.Number Then', '        s = s & "@@E" & Err.Number', '    ElseIf IsObject(v) Or IsArray(v) Or IsError(v) Or IsNull(v) Then', '        s = s & "@@" & TypeName(v)', '    Else', '        s = s & "@@" & CStr(v)', '    End If']
        lines += [f'    Read{number} = s', 'End Function', '']
    lines += ['Public Function ReadAll(wb As Object) As String', '    Dim ws As Object, s As String', '    Set ws = wb.Worksheets(1)']
    lines += [f'    s = s & Read{number}(wb, ws)' for number in range(len(parts))]
    lines += ['    ReadAll = Mid(s, 3)', 'End Function', '']
    return '\n'.join(lines)


OPENER = '''Public Function Opened(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    Opened = ReadAll(wb)
    wb.Close False
    Application.DisplayAlerts = True
End Function
'''


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\0\xff\0\0\xff\0\0' * 2)) + chunk(b'IEND', b'')


def _one_by_one() -> str:
    """Each read a function of its own, since what the model refuses ends a macro, where a trappable error would not."""
    lines: list[str] = []
    for number, read in enumerate(READS):
        lines += [f'Public Function R{number}() As String', '    Dim wb As Object, ws As Object, v As Variant', '    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1)', '    On Error Resume Next', f'    v = {read}', '    If Err.Number Then',
                  f'        R{number} = "E" & Err.Number', '    ElseIf IsObject(v) Or IsArray(v) Or IsError(v) Or IsNull(v) Then', f'        R{number} = TypeName(v)', '    Else', f'        R{number} = CStr(v)', '    End If', 'End Function', '']
    return '\n'.join(lines)


def in_model(source: Path, folder: Path) -> tuple[list[str], dict[str, str]]:
    """What the model reads from Excel's file, or that it refuses; and for each edit, the path of its save, or what it refused."""
    app = ExcelApplication.open(source, with_vba=False)
    app.add_module(_one_by_one(), name='Reader')
    answers: list[str] = []
    for number in range(len(READS)):
        try:
            answers.append(str(app.run(f'Reader.R{number}')))
        except VBAUnsupportedError as refused:
            answers.append(f'UNSUPPORTED {str(refused)[:90]}')
    saved: dict[str, str] = {}
    for name, edit in EDITS.items():
        fresh = ExcelApplication.open(source, with_vba=False)
        fresh.add_module(f'''Public Sub Edit()
    Dim wb As Object, ws As Object
    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1)
    {edit}
End Sub
''', name='Editor')
        path = folder / f'model_{name}.xlsx'
        try:
            fresh.run('Editor.Edit')
            fresh.save(path)
            saved[name] = str(path)
        except (VBAUnsupportedError, VBARuntimeError) as refused:
            saved[name] = f'UNSUPPORTED {str(refused)[:200]}'
    return answers, saved


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        (Path(folder) / 'p.png').write_bytes(png())
        excel_file = Path(folder) / 'excel.xlsx'
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (folder,), timeout=120.0, line_numbers=False)
            print('MADE', made.value)
            assert made.ok and str(made.value).startswith('ok'), made
            read = excel.run_vba(_reader() + OPENER, 'Opened', (str(excel_file),), timeout=200.0, line_numbers=False)
            assert read.ok, read
            truth = str(read.value).split('@@')
        copy_saved(excel_file, BOOK)
        model, after_edits = in_model(BOOK, Path(folder))
        variants: dict[str, dict[str, object]] = {}
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            for name, edit in EDITS.items():
                saved = after_edits[name]
                # Excel makes the same edit to its own file, saves, and reads what it reopens: what the model's save is held to.
                own = Path(folder) / f'own_{name}.xlsx'
                source = _reader() + OPENER + EDITOR.format(edit=edit)
                done = excel.run_vba(source, 'Edited', (str(excel_file), str(own)), timeout=60.0, line_numbers=False)
                assert done.ok and done.value == 'ok', (name, done)
                held = excel.run_vba(source, 'Opened', (str(own),), timeout=200.0, line_numbers=False)
                assert held.ok, (name, held)
                if saved.startswith('UNSUPPORTED'):
                    variants[name] = {'edit': edit, 'excel': str(held.value).split('@@'), 'refused': saved, 'model': []}
                    print(name, saved[:140])
                    continue
                again = excel.run_vba(source, 'Opened', (saved,), timeout=200.0, line_numbers=False)
                print(name, again.outcome, [one.message[:160] for one in again.dialogs])
                variants[name] = {'edit': edit, 'excel': str(held.value).split('@@'), 'opened': again.outcome, 'model': str(again.value).split('@@') if again.ok else []}
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'reads': READS, 'excel': truth, 'model': model, 'variants': variants}, indent=0) + '\n', encoding='utf-8')
    print(len(READS), 'reads')


if __name__ == '__main__':
    main()
