"""What the model loses of a file's other parts: charts, pivots, the VBA project, the macro formats, properties, theme,
images, links, names, and settings on a sheet other than the first. Live Excel saves a workbook and says what it reads
there; the model reads the same file, sets a cell nothing reads and saves, opened by a macro and opened from Python; Excel
reads each result, against what it reads after making the same edit itself, since an edit can change an answer
(a far cell adds automatic page breaks). The packages are compared too, part by part and for the workbook and sheets
element by element, so a loss no macro can read still shows. SEQUENCES are one macro run in fresh Excel and in the model.
No row reads a size that follows the display, so this can be measured on any; see display_default for the one element that
follows it anyway.

    python scripts/measure_parts_loss.py [name ...]

writes tests/fixtures/parts_loss/, which tests/test_excel_parts_loss.py replays.
"""
from __future__ import annotations

import json
import re
import shutil
import struct
import sys
import tempfile
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "scripts"))

from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAUnsupportedError  # noqa: E402

OUT = ROOT / "tests/fixtures/parts_loss"
FORMATS = {"xlsx": 51, "xlsm": 52, "xltm": 53, "xlam": 55, "xlsb": 50, "xls": 56}
#: The zip packages, which the comparison and strip_save_path can look into.
PACKAGED = {"xlsx", "xlsm", "xltm", "xlam"}
#: Parts every save rewrites on purpose: the save's author and time, the stylesheet in use only, and the shared strings, whose
#: entries no cell reads any more are emptied as a changed sheet's text goes inline (Excel's reread is what checks that).
REWRITTEN = {"docProps/core.xml", "xl/styles.xml", "xl/sharedStrings.xml"}
#: Elements the edit itself changes.
EDITED = {"sheetData", "dimension"}

MODULE = 'Set o = wb.VBProject.VBComponents.Add(1): o.Name = "Mod1": o.CodeModule.AddFromString "Function F()" & vbLf & "F = 42" & vbLf & "End Function"'
MODULE_READS = ["wb.VBProject.VBComponents.Count", 'wb.VBProject.VBComponents("Mod1").CodeModule.Lines(1, 3)', "wb.HasVBProject"]
PIVOT = '''Set p = wb.PivotCaches.Create(1, ws.Range("A1:C5")).CreatePivotTable(s2.Range("A3"), "P")
    p.PivotFields("k").Orientation = 1
    p.AddDataField p.PivotFields("v"), "Sum v", -4157'''
PIVOT_READS = ['wb.Worksheets(2).PivotTables(1).Name', 'wb.Worksheets(2).PivotTables(1).TableRange1.Address', 'wb.Worksheets(2).Range("B4").Value',
               'wb.Worksheets(2).PivotTables(1).PivotCache.SourceData']

#: Name: the format saved in; what is done before saving; what is read.
ROWS: dict[str, tuple[str, str, list[str]]] = {
    "chart_embedded": ("xlsx", '''Set o = ws.ChartObjects.Add(10, 10, 300, 200)
    o.Chart.SetSourceData ws.Range("A1:B5")
    o.Chart.ChartType = 51
    o.Chart.HasTitle = True
    o.Chart.ChartTitle.Text = "T"''', ["ws.ChartObjects.Count", "ws.ChartObjects(1).Chart.ChartType", "ws.ChartObjects(1).Chart.ChartTitle.Text",
                                       "ws.ChartObjects(1).Chart.SeriesCollection(1).Formula"]),
    "chart_styled": ("xlsx", '''Set o = ws.Shapes.AddChart2(227, 4)
    o.Chart.SetSourceData ws.Range("A1:B5")''', ["ws.Shapes.Count", "ws.ChartObjects(1).Chart.ChartType", "ws.ChartObjects(1).Chart.ChartStyle",
                                                  "ws.ChartObjects(1).Chart.SeriesCollection.Count"]),
    "chart_sheet": ("xlsx", '''Set o = wb.Charts.Add
    o.SetSourceData ws.Range("A1:B5")
    o.ChartType = 4
    o.Name = "C"''', ["wb.Charts.Count", 'wb.Charts("C").ChartType', "wb.Sheets.Count", 'wb.Charts("C").SeriesCollection(1).Formula']),
    "pivot": ("xlsx", PIVOT, PIVOT_READS),
    "pivot_slicer": ("xlsx", PIVOT + '''
    wb.SlicerCaches.Add2(p, "k").Slicers.Add s2, , "S", "S", 10, 300, 100, 100''', PIVOT_READS + ["wb.SlicerCaches.Count", "wb.SlicerCaches(1).Slicers(1).Name"]),
    "vba_module": ("xlsm", MODULE, MODULE_READS),
    "vba_kinds": ("xlsm", MODULE + '''
    Set o = wb.VBProject.VBComponents.Add(2): o.Name = "Cls1": o.CodeModule.AddFromString "Public V As Long"
    wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString "Private Sub Workbook_Open()" & vbLf & "End Sub"
    wb.VBProject.VBComponents(ws.CodeName).CodeModule.AddFromString "Private Sub Worksheet_Activate()" & vbLf & "End Sub"''',
                  MODULE_READS + ['wb.VBProject.VBComponents("Cls1").Type', 'wb.VBProject.VBComponents("ThisWorkbook").CodeModule.CountOfLines',
                                  "wb.VBProject.VBComponents(ws.CodeName).CodeModule.CountOfLines"]),
    "vba_project": ("xlsm", MODULE + '''
    wb.VBProject.Name = "Named"
    wb.VBProject.Description = "About"
    wb.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0''',
                    MODULE_READS + ["wb.VBProject.Name", "wb.VBProject.Description", "wb.VBProject.References.Count"]),
    "vba_form": ("xlsm", 'Set o = wb.VBProject.VBComponents.Add(3): o.Name = "Form1": o.Designer.Controls.Add "Forms.Label.1", "Lab"',
                 ['wb.VBProject.VBComponents("Form1").Type', 'wb.VBProject.VBComponents("Form1").Designer.Controls.Count']),
    "format_xlsm_empty": ("xlsm", "", ["wb.HasVBProject", "wb.FileFormat", 'ws.Range("A2").Value']),
    "format_xltm": ("xltm", MODULE, MODULE_READS + ["wb.FileFormat", 'ws.Range("A2").Value']),
    "format_xlam": ("xlam", MODULE, MODULE_READS + ["wb.FileFormat", "wb.IsAddin", 'ws.Range("A2").Value']),
    "format_xlsb": ("xlsb", "", ["wb.FileFormat", 'ws.Range("A2").Value']),
    "format_xls": ("xls", "", ["wb.FileFormat", 'ws.Range("A2").Value']),
    "properties": ("xlsx", '''wb.BuiltinDocumentProperties("Title") = "T"
    wb.BuiltinDocumentProperties("Subject") = "S"
    wb.BuiltinDocumentProperties("Keywords") = "K"
    wb.BuiltinDocumentProperties("Comments") = "C"
    wb.BuiltinDocumentProperties("Category") = "G"
    wb.BuiltinDocumentProperties("Manager") = "M"
    wb.BuiltinDocumentProperties("Company") = "Y"
    wb.BuiltinDocumentProperties("Hyperlink base") = "http://example.com/"''',
                   [f'wb.BuiltinDocumentProperties("{name}").Value' for name in
                    ("Title", "Subject", "Keywords", "Comments", "Category", "Manager", "Company", "Hyperlink base")]),
    "properties_custom": ("xlsx", '''wb.Names.Add "Linked", "=Sheet1!$A$2"
    wb.CustomDocumentProperties.Add "Text", False, 4, "x"
    wb.CustomDocumentProperties.Add "Whole", False, 1, 7
    wb.CustomDocumentProperties.Add "Real", False, 5, 1.5
    wb.CustomDocumentProperties.Add "Flag", False, 2, True
    wb.CustomDocumentProperties.Add "When", False, 3, DateSerial(2020, 2, 3)
    wb.CustomDocumentProperties.Add Name:="Link", LinkToContent:=True, Type:=4, LinkSource:="Linked"''',
                          ["wb.CustomDocumentProperties.Count"] + [f'wb.CustomDocumentProperties("{name}").Value' for name in ("Text", "Whole", "Real", "Flag", "When", "Link")]
                          + ['wb.CustomDocumentProperties("Link").LinkToContent']),
    "theme": ("xlsx", '''wb.Theme.ThemeColorScheme.Colors(5).RGB = RGB(1, 2, 3)
    wb.Theme.ThemeFontScheme.MajorFont(1).Name = "Arial"
    ws.Range("A1").Interior.ThemeColor = 5''', ["wb.Theme.ThemeColorScheme.Colors(5).RGB", "wb.Theme.ThemeFontScheme.MajorFont(1).Name", 'ws.Range("A1").Interior.Color',
                                                'ws.Range("A1").Interior.ThemeColor']),
    "image": ("xlsx", 'ws.Shapes.AddPicture folder & "\\p.png", False, True, 10, 10, 50, 40',
              ["ws.Shapes.Count", "ws.Shapes(1).Type", "ws.Shapes(1).Width", "ws.Shapes(1).Left"]),
    "image_background": ("xlsx", 'ws.SetBackgroundPicture folder & "\\p.png"', ['ws.Range("A2").Value']),
    "image_header": ("xlsx", '''ws.PageSetup.LeftHeaderPicture.Filename = folder & "\\p.png"
    ws.PageSetup.LeftHeader = "&G"''', ["ws.PageSetup.LeftHeader", "ws.PageSetup.LeftHeaderPicture.Width"]),
    "link_cell": ("xlsx", '''Set o = Workbooks.Add(1)
    o.Worksheets(1).Name = "Data": o.Worksheets(1).Range("A1").Value = 11
    o.SaveAs folder & "\\source.xlsx", 51
    ws.Range("E7").Formula = "=[source.xlsx]Data!$A$1"
    o.Close False''', ['ws.Range("E7").Value', 'Mid$(ws.Range("E7").Formula, InStrRev(ws.Range("E7").Formula, "\\") + 1)', "UBound(wb.LinkSources(1))",
                       'Mid$(wb.LinkSources(1)(1), InStrRev(wb.LinkSources(1)(1), "\\") + 1)']),
    "link_name": ("xlsx", '''Set o = Workbooks.Add(1)
    o.Worksheets(1).Name = "Data": o.Worksheets(1).Range("A1").Value = 11
    o.SaveAs folder & "\\source.xlsx", 51
    wb.Names.Add "Far", "=[source.xlsx]Data!$A$1"
    o.Close False''', ["UBound(wb.LinkSources(1))", "wb.Names.Count", 'Mid$(wb.Names("Far").RefersTo, InStrRev(wb.Names("Far").RefersTo, "\\") + 1)']),
    "link_used": ("xlsx", '''Set o = Workbooks.Add(1)
    o.Worksheets(1).Name = "Data": o.Worksheets(1).Range("A1").Value = 11
    o.SaveAs folder & "\\source.xlsx", 51
    wb.Names.Add "Far", "=[source.xlsx]Data!$A$1"
    ws.Range("E7").Formula = "=Far"
    o.Close False''', ["UBound(wb.LinkSources(1))", 'ws.Range("E7").Value']),
    "link_both": ("xlsx", '''Set o = Workbooks.Add(1)
    o.Worksheets(1).Name = "Data": o.Worksheets(1).Range("A1").Value = 11
    o.SaveAs folder & "\\source.xlsx", 51
    wb.Names.Add "Far", "=[source.xlsx]Data!$A$1"
    ws.Range("E8").Formula = "=[source.xlsx]Data!$A$1"
    o.Close False''', ["UBound(wb.LinkSources(1))", 'ws.Range("E8").Value']),
    "names_scopes": ("xlsx", '''wb.Names.Add "G", "=Sheet1!$A$1"
    ws.Names.Add "L", "=Sheet1!$B$2"
    wb.Names.Add "Both", "=Sheet1!$A$3"
    s2.Names.Add "Both", "=Sheet2!$A$4"
    wb.Names.Add Name:="H", RefersTo:="=Sheet1!$A$5", Visible:=False
    wb.Names("G").Comment = "note"
    wb.Names.Add "Num", "=42"
    wb.Names.Add "Txt", "=""a"""
    ws.Activate
    ws.Range("B3").Select
    wb.Names.Add "Rel", "=Sheet1!A1"
    wb.Names.Add "Calc", "=SUM(Sheet1!$A$2:$A$5)"''',
                     ["wb.Names.Count", "ws.Names.Count", "wb.Worksheets(2).Names.Count", 'wb.Names("G").Comment', 'wb.Names("H").Visible', 'ws.Names("L").RefersTo',
                      'wb.Worksheets(2).Names("Both").RefersTo', 'wb.Names("Both").RefersTo', 'wb.Names("Num").RefersTo', 'wb.Names("Txt").RefersTo',
                      'wb.Names("Rel").RefersToR1C1', 'wb.Names("Calc").RefersTo']),
    "names_builtin": ("xlsx", '''ws.PageSetup.PrintArea = "$A$1:$C$5"
    ws.PageSetup.PrintTitleRows = "$1:$1"
    ws.PageSetup.PrintTitleColumns = "$A:$A"
    ws.Range("A1:C5").AutoFilter 1, "a"''', ["ws.PageSetup.PrintArea", "ws.PageSetup.PrintTitleRows", "ws.PageSetup.PrintTitleColumns", "ws.AutoFilterMode", "ws.Names.Count",
                                              'ws.Names(1).Name', "ws.Rows(2).Hidden"]),
    "later_sheets": ("xlsx", '''s2.Range("A1").Value = 1
    s2.Tab.Color = RGB(9, 8, 7)
    s2.Rows(3).RowHeight = 33
    s2.Columns(2).ColumnWidth = 22
    s2.Rows(4).Hidden = True
    s2.Rows(6).OutlineLevel = 2
    s2.Activate
    ActiveWindow.Zoom = 130
    ActiveWindow.DisplayGridlines = False
    s2.Range("D5:E6").Select
    s2.Protect "pw"
    Set o = wb.Worksheets.Add(After:=s2): o.Name = "Third": o.Range("A1").Value = 3
    Set o = wb.Worksheets.Add(After:=o): o.Name = "Hid": o.Tab.Color = RGB(1, 1, 1)
    o.Visible = 0
    Set o = wb.Worksheets.Add(After:=wb.Worksheets("Third")): o.Name = "Very": o.Visible = 2
    wb.Worksheets("Third").Activate''', ["wb.ActiveSheet.Name", "wb.Worksheets(2).Tab.Color", "wb.Worksheets(2).Rows(3).RowHeight",
                                         "wb.Worksheets(2).Rows(4).Hidden", "wb.Worksheets(2).Rows(6).OutlineLevel", "wb.Worksheets(2).ProtectContents",
                                         'wb.Worksheets("Hid").Visible', 'wb.Worksheets("Very").Visible', 'wb.Worksheets("Hid").Tab.Color', "wb.Windows(1).SheetViews(2).DisplayGridlines"]),
    "scenarios": ("xlsx", 'ws.Scenarios.Add "S", ws.Range("A2:A3"), Array(5, 6), "note"',
                  ["ws.Scenarios.Count", "ws.Scenarios(1).Name", "ws.Scenarios(1).Comment", "ws.Scenarios(1).ChangingCells.Address"]),
    "custom_views": ("xlsx", 'wb.CustomViews.Add "V", True, True', ["wb.CustomViews.Count", "wb.CustomViews(1).Name", "wb.CustomViews(1).PrintSettings"]),
    "edit_ranges": ("xlsx", 'ws.Protection.AllowEditRanges.Add "R", ws.Range("A2:A3")',
                    ["ws.Protection.AllowEditRanges.Count", "ws.Protection.AllowEditRanges(1).Title", "ws.Protection.AllowEditRanges(1).Range.Address"]),
    "ignored_errors": ("xlsx", 'ws.Range("D1").NumberFormat = "@": ws.Range("D1").Value = "12": ws.Range("D1").Errors(3).Ignore = True',
                       ['ws.Range("D1").Errors(3).Ignore', 'ws.Range("D1").Errors(3).Value']),
    "page_breaks": ("xlsx", 'ws.HPageBreaks.Add ws.Range("A5"): ws.VPageBreaks.Add ws.Range("C1")', ["ws.HPageBreaks.Count", "ws.VPageBreaks.Count", "ws.HPageBreaks(1).Location.Address"]),
    "custom_xml": ("xlsx", 'wb.CustomXMLParts.Add "<a xmlns=""urn:p""><b>x</b></a>"', ['wb.CustomXMLParts.SelectByNamespace("urn:p").Count', 'wb.CustomXMLParts.SelectByNamespace("urn:p")(1).XML']),
    "cell_style": ("xlsx", 'Set o = wb.Styles.Add("Mine"): o.Font.Bold = True: ws.Range("A1").Style = "Mine"', ['wb.Styles("Mine").Font.Bold', 'ws.Range("A1").Style.Name']),
    "sparklines": ("xlsx", 'ws.Range("E2").SparklineGroups.Add 1, "B2:C2"', ['ws.Range("E2").SparklineGroups.Count', 'ws.Range("E2").SparklineGroups(1).SourceData']),
    "form_button": ("xlsx", 'Set o = ws.Buttons.Add(10, 10, 80, 20): o.Caption = "Go"', ["ws.Buttons.Count", "ws.Buttons(1).Caption", "ws.Shapes(1).Type"]),
}

HELP = '''Private Function Tried(ByVal n As Long, wb As Object, ws As Object) As String
    Dim v As Variant
    On Error GoTo Failed
    Select Case n
{cases}
    End Select
    Tried = CStr(v)
    Exit Function
Failed:
    Tried = "E" & Err.Number
End Function
'''
MAKE = HELP + '''
Function Make(ByVal file As String, ByVal folder As String, ByVal kind As Long) As String
    Dim wb As Workbook, ws As Worksheet, s2 As Worksheet, o As Object, p As Object, i As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(1): Set ws = wb.Worksheets(1)
    Set s2 = wb.Worksheets.Add(After:=ws)
    ws.Activate
    ws.Range("A1:C1").Value = Array("k", "v", "w")
    For i = 2 To 5
        ws.Cells(i, 1).Value = Chr$(95 + i): ws.Cells(i, 2).Value = i * 2: ws.Cells(i, 3).Value = i * 3
    Next
    {action}
    wb.SaveAs file, kind
    wb.Close False
    Make = "ok"
End Function
'''
READ = HELP + '''
Function Read1(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, i As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    wb.Activate
    On Error GoTo 0
    For i = 0 To {last}
        s = s & "|" & Tried(i, wb, ws)
    Next
    wb.Close False
    Read1 = Mid$(s, 2)
End Function
'''
MODEL_READ = '''Function Read1(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, i As Long, s As String
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1)
    For i = 0 To {last}
        s = s & "|" & Tried(i, wb, ws)
    Next
    wb.Close False
    Read1 = Mid$(s, 2)
End Function
'''
MODEL_EDIT = '''Sub Edit1(ByVal file As String)
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file, UpdateLinks:=0, Editable:=True)
    wb.Worksheets(1).Range("Z99").Value = 7
    wb.Save
    wb.Close False
End Sub
'''


#: What LinkSources answers once a macro has touched a workbook's names, for a link only a name holds, one a cell reads
#: through a name, and one a cell names itself.
LINK_SEQUENCE = '''Function Go(ByVal folder As String) As String
    Dim wb As Workbook, s As String, x As Variant, i As Long, n As Object
    Application.DisplayAlerts = False
''' + """    Set wb = Book(folder): s = s & "|" & Src(wb): wb.Close False
    Set wb = Book(folder): i = wb.Names.Count: s = s & "|count " & Src(wb): wb.Close False
    Set wb = Book(folder): i = wb.Worksheets(1).Names.Count: s = s & "|sheet count " & Src(wb): wb.Close False
    Set wb = Book(folder): x = wb.Names(1).Visible: s = s & "|item " & Src(wb): wb.Close False
    Set wb = Book(folder): x = wb.Names("Far").Name: s = s & "|named " & Src(wb): s = s & " " & Src(wb): wb.Close False
    Set wb = Book(folder): x = wb.Names("Far").RefersTo: s = s & "|refers " & Src(wb): wb.Close False
    Set wb = Book(folder)
    For Each n In wb.Names
        i = i + 1
    Next
    s = s & "|each " & Src(wb): wb.Close False
    Set wb = Book(folder): x = wb.Worksheets(1).Range("A2").Value: s = s & "|cell " & Src(wb): wb.Close False
    Set wb = Book(folder): Set n = wb.Names(1): s = s & "|set " & Src(wb): wb.Worksheets(1).Range("Z9").Value = 1: s = s & " " & Src(wb): wb.Close False
""" + '''
    Go = Mid$(s, 2)
End Function
Private Function Src(wb As Object) As String
    Dim v As Variant
    v = wb.LinkSources(1)
    If IsEmpty(v) Then Src = "Empty" Else Src = CStr(UBound(v))
End Function
Private Function Book(ByVal folder As String) As Workbook
    Set Book = Workbooks.Open(folder & "\\l.xlsx", UpdateLinks:=0)
End Function
'''
#: Name: the fixture it opens, under the name it has here; the macro, which takes the folder and answers in a string.
SEQUENCES: dict[str, tuple[tuple[str, str], str]] = {
    "link_only_a_name": (("link_name.xlsx", "l.xlsx"), LINK_SEQUENCE),
    "link_through_a_name": (("link_used.xlsx", "l.xlsx"), LINK_SEQUENCE),
    "link_in_a_cell": (("link_both.xlsx", "l.xlsx"), LINK_SEQUENCE),
    "template_numbers": (("format_xltm.xltm", "m.xltm"), '''Function Go(ByVal folder As String) As String
    Dim a As Workbook, b As Workbook, c As Workbook, s As String
    Application.DisplayAlerts = False
    Set a = Workbooks.Open(folder & "\\m.xltm"): s = a.Name & " " & a.FileFormat
    Set b = Workbooks.Open(folder & "\\m.xltm"): s = s & "|" & b.Name
    Set c = Workbooks.Add(folder & "\\m.xltm"): s = s & "|" & c.Name
    a.Close False: b.Close False: c.Close False
    Set a = Workbooks.Open(folder & "\\m.xltm"): s = s & "|" & a.Name: a.Close False
    Set a = Workbooks.Add(folder & "\\m.xltm"): s = s & "|" & a.Name: a.Close False
    Set a = Workbooks.Open(folder & "\\m.xltm", Editable:=True): s = s & "|" & a.Name & " " & a.FileFormat: a.Close False
    Set a = Workbooks.Open(folder & "\\m.xltm"): s = s & "|" & a.Name: a.Close False
    Go = s
End Function
'''),
    "template_add_first": (("format_xltm.xltm", "q.xltm"), '''Function Go(ByVal folder As String) As String
    Dim a As Workbook, b As Workbook, c As Workbook, s As String
    Application.DisplayAlerts = False
    Set a = Workbooks.Add(folder & "\\q.xltm"): s = a.Name
    Set b = Workbooks.Open(folder & "\\q.xltm"): s = s & "|" & b.Name
    Set c = Workbooks.Open(folder & "\\q.xltm"): s = s & "|" & c.Name
    a.Close False: b.Close False: c.Close False
    Go = s
End Function
'''),
    "addin": (("format_xlam.xlam", "a.xlam"), '''Function Go(ByVal folder As String) As String
    Dim wb As Workbook, n As Workbook, w As Workbook, s As String, keep As String, count As Long
    Application.DisplayAlerts = False
    keep = ActiveWorkbook.Name
    Set wb = Workbooks.Open(folder & "\\a.xlam")
    s = Workbooks.Count & " " & (ActiveWorkbook.Name = keep) & " " & wb.IsAddin & " " & wb.Windows.Count & " " & Application.Windows.Count & " " & wb.Worksheets.Count
    s = s & "|" & Workbooks("a.xlam").Name
    For Each w In Workbooks
        If w.Name = "a.xlam" Then count = count + 1
    Next
    s = s & "|" & count
    wb.Activate
    s = s & "|" & (ActiveWorkbook.Name = keep)
    wb.IsAddin = False
    s = s & "|" & Workbooks.Count & " " & wb.Windows.Count & " " & (ActiveWorkbook.Name = wb.Name)
    wb.Close False
    s = s & "|" & Workbooks.Count & " " & (ActiveWorkbook.Name = keep)
    Set n = Workbooks.Add(-4167)
    n.IsAddin = True
    s = s & "|" & Workbooks.Count & " " & n.Windows.Count & " " & (ActiveWorkbook.Name = keep) & " " & n.IsAddin
    n.IsAddin = False
    s = s & "|" & Workbooks.Count & " " & n.Windows.Count & " " & (ActiveWorkbook.Name = n.Name)
    n.Close False
    Go = s
End Function
'''),
}


def run_sequence(name: str) -> str:
    """The model's answer, from a copy of the fixture in a new folder."""
    (source, target), macro = SEQUENCES[name]
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_sequence_"))
    try:
        shutil.copy(OUT / source, folder / target)
        return model_answer(macro, "Go", folder)
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def scrub_links(fixture: Path, folder: Path) -> None:
    """A link's relationships name the folder the source was in; a fixture says a neutral one instead."""
    with zipfile.ZipFile(fixture) as package:
        items = [(info, package.read(info.filename)) for info in package.infolist()]
    with zipfile.ZipFile(fixture, "w", zipfile.ZIP_DEFLATED) as package:
        for info, data in items:
            package.writestr(info, data.replace(str(folder).encode("utf-8"), b"C:\\Temp") if info.filename.endswith(".rels") else data)


def cases(reads: list[str]) -> str:
    return "\n".join(f"    Case {index}\n        v = {read}" for index, read in enumerate(reads))


def png(file: Path) -> None:
    """A small picture, 8 by 6 and a single colour."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    rows = (b"\x00" + b"\x10\x80\xf0" * 8) * 6
    file.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 8, 6, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def top_level(text: str) -> dict[str, str]:
    """A workbook's or sheet's direct children, by name (a repeated element numbered), without the parts' declarations."""
    body = re.sub(r"^.*?<(workbook|worksheet|chartsheet)\b[^>]*>", "", text, count=1, flags=re.DOTALL)
    found: dict[str, str] = {}
    position = 0
    while True:
        opening = re.compile(r"<([\w:.]+)\b[^>]*?(/?)>").search(body, position)
        if opening is None or opening.group(1) in ("workbook", "worksheet", "chartsheet"):
            return found
        end = opening.end()
        if not opening.group(2):
            depth = 1
            pattern = re.compile(rf"<(/?){re.escape(opening.group(1))}\b[^>]*?(/?)>")
            for tag in pattern.finditer(body, end):
                if tag.group(1):
                    depth -= 1
                elif not tag.group(2):
                    depth += 1
                if depth == 0:
                    end = tag.end()
                    break
        key = opening.group(1)
        number = 1
        while key in found:
            number += 1
            key = f"{opening.group(1)}#{number}"
        found[key] = body[opening.start():end]
        position = end


def display_default(element: str) -> str:
    """A sheetFormatPr without the default height and descent the display gave it: a changed sheet takes the 96 DPI ones, as
    Excel at 96 DPI does (see _dimensions), and a height the sheet set itself stays compared."""
    if "customHeight" in element:
        return element
    return re.sub(r'\s(defaultRowHeight|x14ac:dyDescent)="[^"]*"', "", element)


def relations(text: str) -> list[str]:
    """Relationships without their ids, which a save may number again."""
    return sorted(re.sub(r'\sId="[^"]*"', "", tag) for tag in re.findall(r"<Relationship\b[^>]*>", text))


def package_loss(given: Path, through: Path) -> dict[str, list[str]]:
    """What differs between two packages, by part and by element; a part only one has is ``removed`` or ``added``."""
    lost: dict[str, list[str]] = {}
    with zipfile.ZipFile(given) as one, zipfile.ZipFile(through) as two:
        names = set(one.namelist()), set(two.namelist())
        for name in sorted(names[0] | names[1]):
            if name not in names[1]:
                lost[name] = ["removed"]
            elif name not in names[0]:
                lost[name] = ["added"]
            elif name in REWRITTEN or one.read(name) == two.read(name):
                continue
            elif name.endswith(".rels"):
                before, after = relations(one.read(name).decode("utf-8")), relations(two.read(name).decode("utf-8"))
                if before != after:
                    lost[name] = [str(sorted(set(before) - set(after)))[:400], str(sorted(set(after) - set(before)))[:400]]
            elif name == "xl/workbook.xml" or re.fullmatch(r"xl/(work|chart)sheets/\w+\.xml", name):
                before_all, after_all = top_level(one.read(name).decode("utf-8")), top_level(two.read(name).decode("utf-8"))
                for key in sorted(set(before_all) | set(after_all)):
                    if key == "sheetFormatPr":
                        before_all[key], after_all[key] = display_default(before_all.get(key, "")), display_default(after_all.get(key, ""))
                    if key.split("#")[0] not in EDITED and before_all.get(key) != after_all.get(key):
                        lost[f"{name}:{key}"] = [before_all.get(key, "")[:400], after_all.get(key, "")[:400]]
            elif name.endswith((".xml", ".vml")):
                lost[name] = ["changed", one.read(name).decode("utf-8", "replace")[:300], two.read(name).decode("utf-8", "replace")[:300]]
            else:
                lost[name] = ["changed bytes", str(len(one.read(name))), str(len(two.read(name)))]
    return lost


def model_answer(source: str, entry: str, file: Path) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    try:
        return str(app.run(entry, file.as_posix()))
    except VBAUnsupportedError as failure:
        return f"reports: {str(failure)[:100]}"
    except Exception as failure:  # noqa: BLE001 - what the model does with a real file is what is being measured
        return f"crash: {type(failure).__name__}: {str(failure)[:100]}"


def python_edit(file: Path) -> str:
    try:
        app = ExcelApplication.open(file)
        app.application.active_book.sheets_[0].vba_get("Range", ["Z99"]).vba_set("Value", 7)
        app.save(file)
        return "saved"
    except VBAUnsupportedError as failure:
        return f"reports: {str(failure)[:100]}"
    except Exception as failure:  # noqa: BLE001 - as above
        return f"crash: {type(failure).__name__}: {str(failure)[:100]}"


def main() -> None:
    from fixture_workbook import strip_save_path
    from pyvbaharness import ExcelSession, HarnessConfig

    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_parts_"))
    png(folder / "p.png")
    done = OUT / "parts_loss.json"
    rows: list[dict[str, object]] = json.loads(done.read_text(encoding="utf-8"))["rows"] if done.exists() else []
    sequences: dict[str, str] = {}
    only = set(sys.argv[1:])
    rows = [row for row in rows if row["name"] not in only]
    have = {row["name"] for row in rows}
    try:
        for name, (kind, action, reads) in ROWS.items():
            if (only and name not in only) or name in have:
                continue
            original = folder / f"{name}.{kind}"
            macro = folder / f"{name}_macro.{kind}"
            python = folder / f"{name}_python.{kind}"
            control = folder / f"{name}_control.{kind}"
            source = cases(reads)
            read = READ.format(cases=source, last=len(reads) - 1)
            row: dict[str, object] = {"name": name, "format": kind, "action": action, "reads": reads}
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                made = excel.run_vba(MAKE.format(cases=source, action=action), "Make", (str(original), str(folder), FORMATS[kind]), timeout=90.0)
                if not made.ok:
                    row["failed"] = f"{made.outcome} {made.message[:120]}"
                    rows.append(row)
                    print(name, "made failed", row["failed"], flush=True)
                    continue
                first = excel.run_vba(read, "Read1", (str(original),), timeout=90.0)
                shutil.copy(original, control)
                edited = excel.run_vba(MODEL_EDIT, "Edit1", (str(control),), timeout=90.0)
                again = excel.run_vba(read, "Read1", (str(control),), timeout=90.0) if edited.ok else None
            row["control"] = str(again.value) if again is not None and again.ok else f"failed {edited.outcome} {edited.message[:80]}"
            row["excel"] = str(first.value) if first.ok else f"failed {first.outcome} {first.message[:80]}"
            row["model"] = model_answer(HELP.format(cases=source) + MODEL_READ.format(last=len(reads) - 1), "Read1", original)
            shutil.copy(original, macro)
            shutil.copy(original, python)
            row["macro_edit"] = model_answer(MODEL_EDIT, "Edit1", macro)
            row["python_edit"] = python_edit(python)
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                for label, file, edit in (("macro", macro, row["macro_edit"]), ("python", python, row["python_edit"])):
                    if str(edit).startswith(("reports", "crash")):
                        continue
                    again = excel.run_vba(read, "Read1", (str(file),), timeout=90.0)
                    row[f"after_{label}"] = str(again.value) if again.ok else f"failed {again.outcome} {again.message[:80]}"
                    if kind in PACKAGED:
                        row[f"lost_{label}"] = package_loss(original, file)
            fixture = OUT / original.name
            shutil.copy(original, fixture)
            if kind in PACKAGED:
                strip_save_path(fixture)
                scrub_links(fixture, folder)
            rows.append(row)
            print(f"{name:20} model={row['model'] == row['excel']} macro={row.get('after_macro') == row['control']} python={row.get('after_python') == row['control']} "
                  f"parts={sorted(row.get('lost_macro', {}))[:6]}", flush=True)
            done.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "help": HELP, "rows": rows}, indent=1) + "\n", encoding="utf-8")
        for name, ((source, target), macro) in SEQUENCES.items():
            run = folder / name
            run.mkdir()
            shutil.copy(OUT / source, run / target)
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                result = excel.run_vba(macro, "Go", (str(run),), timeout=90.0)
            sequences[name] = str(result.value) if result.ok else f"failed {result.outcome} {result.message[:80]}"
            print(name, sequences[name], "model", run_sequence(name), flush=True)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    done.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "help": HELP, "sequences": sequences, "rows": rows}, indent=1) + "\n",
                    encoding="utf-8")


if __name__ == "__main__":
    main()
