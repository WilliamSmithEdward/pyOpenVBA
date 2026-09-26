"""What Excel does when a macro changes a cell style, deletes one or merges another workbook's.

A style changed after cells were given it: which cells follow, part by
part, with formats given before and after the style, and when the style
leaves a part out; Normal changed, its font in ways that leave the
sheets' geometry alone (NormalFont shows the rest moving it, read on this
display); a new style's every settable property set in turn, and the
read-only ones; built-in styles changed whether cells use them or not;
whole rows and columns; protected sheets, in this workbook and another;
a style the file holds, in a workbook opened again. Style.Delete with
cells using the style, a held Style object, the name added again;
Styles.Merge, names both workbooks have among them. The order Excel's
tables write what a session makes, and what a save keeps of a file
holding entries nothing uses (planted.xlsx), a font and xfs twice
(duplicates.xlsx), and custom formats listed out of order
(renumbered.xlsx), each made here from only_good.xlsx. Workbooks are
saved along the way to show what Excel writes.

The probe is one module, which the record keeps, run with the folder of
tests/fixtures/cell_styles to open workbooks from and one to save them
in; the replay runs the same module in the model, which reports
NormalFont, Guards2, Guards3 and Elsewhere unsupported.

    python scripts/measure_style_changes.py

writes tests/fixtures/cell_styles/style_changes.json, the three files it
opens, and the workbooks Excel saved, which tests/test_excel_cell_styles.py
replays.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import copy_saved
from measure_cell_styles import HELPER, STYLE_READS, function, reads, statements

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "cell_styles"

#: What a cell answers after each change.
LOOK = ["c.Style.Name", "c.NumberFormat", "c.Font.Name", "c.Font.Size", "c.Font.Bold", "c.Font.Italic",
        "c.Font.Underline", "c.Font.Color", "c.Interior.Color", "c.Interior.Pattern", "c.Borders(9).LineStyle",
        "c.Borders(9).Weight", "c.HorizontalAlignment", "c.Locked"]
#: What the sheet's geometry answers, which Normal's font sets.
GEOMETRY = ["ws.StandardHeight", "ws.StandardWidth", "ws.Rows(1).Height", "ws.Columns(1).ColumnWidth",
            "ws.Columns(1).Width"]

#: (cell, statements): what each cell went through before Good and Comma change.
FOLLOW_CELLS = [
    ("A1", ['c.Style = "Good"']),
    ("A2", ['c.Style = "Good"', "c.Font.Italic = True"]),
    ("A3", ['c.Style = "Good"', "c.Font.Bold = True", "c.Font.Bold = False"]),
    ("A4", ['c.Style = "Good"', "c.Interior.Color = RGB(0, 0, 255)"]),
    ("A5", ['c.Style = "Good"', 'c.NumberFormat = "0.0"']),
    ("A6", ['c.Style = "Comma"']),
    ("A7", []),
    ("A8", ['c.Style = "Good"', "c.Font.Color = c.Font.Color"]),
    ("A9", ["c.Font.Bold = True"]),
    ("A10", ['c.Style = "Good"', 'c.Style = "Normal"']),
    ("A11", ["c.Font.Italic = True", 'c.Style = "Good"']),
]
#: (statements, the styles read after them): the changes, one after another.
FOLLOW_CHANGES = [
    (['wb.Styles("Good").Font.Bold = True'], ["Good"]),
    (['wb.Styles("Good").Interior.Color = RGB(255, 0, 0)'], []),
    (['wb.Styles("Good").NumberFormat = "0.00"'], ["Good"]),
    (['wb.Styles("Comma").Font.Italic = True'], ["Comma"]),
    (['wb.Styles("Good").IncludeFont = False', 'wb.Styles("Good").Font.Underline = 2'], ["Good"]),
    (['wb.Styles("Good").IncludeFont = True'], []),
    (['wb.Styles("Good").Borders(-4107).LineStyle = 1'], ["Good"]),
]

#: (cell, statements) under Normal's changes; A7 is left alone.
NORMAL_CELLS = [("A1", []), ("A2", ['c.Style = "Good"']), ("A3", ["c.Font.Italic = True"]),
                ("A4", ['c.Style = "Comma"']), ("A5", ['c.Style = "Heading 4"']), ("A6", ['c.Style = "Note"'])]
#: Normal's font changes only in ways that leave its metrics, which the sheets' geometry rests on (NormalFont).
NORMAL_CHANGES = [
    (['wb.Styles("Normal").Font.Color = RGB(0, 0, 255)'], ["Normal", "Comma", "Note"]),
    (['wb.Styles("Normal").Interior.Color = RGB(255, 255, 0)'], []),
    (['wb.Styles("Normal").NumberFormat = "0.00"'], ["Comma"]),
    (['wb.Styles("Normal").HorizontalAlignment = -4108'], ["Normal"]),
    (['wb.Styles("Normal").Font.Strikethrough = True', 'wb.Styles("Normal").Font.ThemeColor = 6'], ["Normal"]),
]

#: Groups of statements set on a new style, one group after another; the last ones are read-only.
SETS = [
    ["s.Font.Bold = True", "s.Font.Italic = True", "s.Font.Size = 14", 's.Font.Name = "Arial"', "s.Font.Underline = 2",
     "s.Font.Strikethrough = True", "s.Font.Color = RGB(1, 2, 3)"],
    ["s.Font.ThemeColor = 5", "s.Font.TintAndShade = 0.4"],
    ["s.Font.ColorIndex = 3", "s.Font.Superscript = True"],
    ["s.Interior.Color = RGB(0, 255, 0)"],
    ["s.Interior.Pattern = 17", "s.Interior.PatternColor = RGB(255, 0, 0)"],
    ["s.Interior.ThemeColor = 6", "s.Interior.TintAndShade = -0.25"],
    ["s.Interior.Pattern = -4142"],
    ["s.Borders(-4131).LineStyle = 1", "s.Borders(-4152).Weight = -4138", "s.Borders(-4160).Color = RGB(255, 0, 0)",
     "s.Borders(-4107).LineStyle = -4119"],
    ["s.Borders(5).LineStyle = 1", "s.Borders(7).LineStyle = 1"],
    ["s.Borders.LineStyle = -4142"],
    ['s.NumberFormat = "0.0%"'],
    ['s.NumberFormatLocal = "#,##0.000"'],
    ["s.HorizontalAlignment = -4108", "s.VerticalAlignment = -4160", "s.WrapText = True", "s.Orientation = 45"],
    ["s.IndentLevel = 2"],
    ["s.ShrinkToFit = True", "s.ReadingOrder = -5004", "s.AddIndent = True"],
    ["s.Locked = False", "s.FormulaHidden = True"],
    ["s.IncludeNumber = False", "s.IncludeFont = False", "s.IncludeAlignment = False", "s.IncludeBorder = False",
     "s.IncludePatterns = False", "s.IncludeProtection = False"],
    ['s.Name = "Other"', 's.Value = "Other"', 's.NameLocal = "Other"', "s.BuiltIn = True", "s.MergeCells = True",
     's.Font.FontStyle = "Bold"', "s.Font.ThemeFont = 1"],
]
#: Built-in styles changed, used by a cell or not.
BUILT_INS = ['wb.Styles("Title").Font.Size = 20', 'ws.Range("A2").Style = "Bad"', 'wb.Styles("Bad").Font.Bold = True',
             'wb.Styles("Accent1").Interior.Color = RGB(9, 9, 9)']

#: (cell, statements) before Good is deleted.
DELETE_CELLS = [("A1", ['c.Style = "Good"']), ("A2", ['c.Style = "Good"', "c.Font.Bold = True"]),
                ("A3", ['c.Style = "Good"', "c.Interior.Color = RGB(0, 0, 255)"]), ("A4", ['c.Style = "Mine"']),
                ("A5", ['c.Style = "Bad"']), ("A6", ['c.Style = "Comma"'])]

#: Formats made in an order their rows do not follow: fonts, one a built-in style's; fills; number formats, one
#: made and dropped; borders, one a built-in style's; a cell xf made and dropped.
ORDER = ["A3.Font.Italic = True", "A2.Font.Underline = 2", "A1.Font.Strikethrough = True",
         "A4.Font.Color = RGB(255, 0, 0)", "A5.Font.Bold = True", "A6.Interior.Color = RGB(255, 0, 0)",
         "A7.Interior.Color = RGB(198, 239, 206)", 'A8.NumberFormat = "0.0000"', 'A9.NumberFormat = "0.00000"',
         'A9.NumberFormat = "General"', 'A10.NumberFormat = "0.000"', "A11.Borders(8).LineStyle = 1",
         "A12.Borders(9).Weight = 4", "A12.Borders(9).ThemeColor = 5", "A13.Font.Size = 20", "A13.Font.Size = 11"]
#: What the Style object does in the corners: superscript under IncludeFont, ThemeFont, IndentLevel, xlEdgeLeft
#: and the rest after one edge or four, a property set to what it was, and the pattern colour a fill gets.
QUIRKS = [
    (['Set s = wb.Styles.Add("Sup")', 'ws.Range("A1").Style = "Sup"', "s.Font.Superscript = True"],
     ["s.Font.Superscript", 'ws.Range("A1").Font.Superscript']),
    (["s.IncludeFont = False"], ["s.Font.Superscript", "s.Font.Subscript", 'ws.Range("A1").Font.Superscript']),
    (["s.IncludeFont = True"], ["s.Font.Superscript", 'ws.Range("A1").Font.Superscript']),
    (["s.Font.Subscript = True", "s.IncludeFont = False"],
     ["s.Font.Subscript", "s.Font.Superscript", 'ws.Range("A1").Font.Subscript']),
    (["s.IncludeNumber = False"], ["s.Font.Subscript"]),
    (['Set s = wb.Styles.Add("Themed")', 'ws.Range("A2").Style = "Themed"', "s.Font.ThemeFont = 1"],
     ["s.Font.Name", 'ws.Range("A2").Font.Name', 'ws.Range("A2").Font.ThemeFont']),
    (['s.Font.Name = "Arial"'], ["s.Font.Name", 'ws.Range("A2").Font.Name']),
    (["s.Font.ThemeFont = 2"], ["s.Font.Name", "s.Font.ThemeFont", 'ws.Range("A2").Font.Name']),
    (['Set s = wb.Styles.Add("Indent")', 'ws.Range("A3").Style = "Indent"', "s.HorizontalAlignment = -4131",
      "s.IndentLevel = 2"], ["s.IndentLevel", "s.HorizontalAlignment", 'ws.Range("A3").IndentLevel']),
    (["s.IndentLevel = 0"], ["s.IndentLevel", 'ws.Range("A3").IndentLevel']),
    (["s.HorizontalAlignment = -4108", "s.IndentLevel = 3"],
     ["s.IndentLevel", "s.HorizontalAlignment", 'ws.Range("A3").IndentLevel', 'ws.Range("A3").HorizontalAlignment']),
    *(([f'Set s = wb.Styles.Add("Edge{number}")', *setting],
       [f"s.Borders({index}).{what}" for index in range(7, 13) for what in ("LineStyle", "Weight")])
      for number, setting in enumerate([["s.Borders(-4131).LineStyle = 1"], ["s.Borders(-4107).LineStyle = 1"],
                                        ["s.Borders(5).LineStyle = 1"],
                                        [f"s.Borders({edge}).LineStyle = 1" for edge in (-4131, -4152, -4160)],
                                        [f"s.Borders({edge}).LineStyle = 1" for edge in (-4131, -4152, -4160, -4107)],
                                        ["s.Borders(-4131).Weight = 4"], ["s.Borders(-4131).Color = 255"]])),
    (['Set s = wb.Styles("Heading 1")', "s.Borders(-4131).LineStyle = 1"],
     [f"s.Borders({index}).{what}" for index in range(7, 13) for what in ("LineStyle", "Weight")]),
    (['Set s = wb.Styles("Title")', "s.Font.Size = 18", 'Set s = wb.Styles("Bad")', "s.Font.Bold = False"], []),
    (['Set s = wb.Styles.Add("Filled")', "s.Interior.Color = RGB(1, 2, 3)"],
     ["s.Interior.PatternColor", "s.Interior.PatternColorIndex", "s.Interior.Pattern"]),
    (['Set s = wb.Styles("Good")', "s.Interior.Color = RGB(4, 5, 6)"],
     ["s.Interior.PatternColor", "s.Interior.PatternColorIndex"]),
    (['Set s = wb.Styles.Add("Patterned")', "s.Interior.Pattern = 17"],
     ["s.Interior.Color", "s.Interior.PatternColor", "s.Interior.PatternColorIndex"]),
    (['Set s = wb.Styles("Accent2")', "s.Interior.Pattern = 17"],
     ["s.Interior.Color", "s.Interior.PatternColor", "s.Interior.PatternColorIndex"]),
]
#: (statements, reads): whether a change reaches cells when the style includes less or nothing.
REACH = [
    (['Set s = wb.Styles.Add("P1")', 'ws.Range("A1").Style = "P1"', "s.IncludeFont = False", "s.Font.Underline = 2"],
     ['ws.Range("A1").Font.Underline']),
    (["s.Font.ThemeFont = 1"], ['ws.Range("A1").Font.Name']),
    (['Set s = wb.Styles.Add("P2")', 'ws.Range("A2").Style = "P2"',
      *(f"s.Include{part} = False" for part in ("Number", "Font", "Alignment", "Border", "Patterns", "Protection")),
      "s.Font.Underline = 2"], ['ws.Range("A2").Font.Underline']),
    (['Set s = wb.Styles("Bad")', 'ws.Range("A3").Style = "Bad"', "s.IncludeFont = False", "s.IncludePatterns = False",
      "s.Font.Bold = True"], ['ws.Range("A3").Font.Bold']),
    (['Set s = wb.Styles.Add("P4")', 'ws.Range("A4").Style = "P4"', "s.IncludeFont = False", "s.Font.Italic = True",
      "s.Interior.Color = 255"], ['ws.Range("A4").Font.Italic', 'ws.Range("A4").Interior.Color']),
]
#: (statements, reads): when a style's IndentLevel reads Null.
INDENTS = [
    (['Set s = wb.Styles.Add("I1")', "s.HorizontalAlignment = -4152", "s.IndentLevel = 1"], ["s.IndentLevel"]),
    (['Set s = wb.Styles.Add("I2")', "s.HorizontalAlignment = -4131", "s.Orientation = 45", "s.IndentLevel = 2"],
     ["s.IndentLevel", "s.HorizontalAlignment"]),
    (['Set s = wb.Styles.Add("I3")', "s.HorizontalAlignment = -4117", "s.IndentLevel = 1"],
     ["s.IndentLevel", "s.HorizontalAlignment"]),
    (['Set s = wb.Styles.Add("I4")', "s.IndentLevel = 2"], ["s.IndentLevel", "s.HorizontalAlignment"]),
    (['Set s = wb.Styles.Add("I5")', "s.HorizontalAlignment = -4131"], ["s.IndentLevel"]),
    (['Set s = wb.Styles.Add("I6")', "s.HorizontalAlignment = -4131", "s.IndentLevel = 0"], ["s.IndentLevel"]),
    (["s.Orientation = 45"], ["s.IndentLevel"]),
    (["s.Orientation = 0"], ["s.IndentLevel"]),
    (['Set s = wb.Styles.Add("I7")', "s.WrapText = True"], ["s.IndentLevel"]),
    (['Set s = wb.Styles.Add("I8")', "s.HorizontalAlignment = -4108"], ["s.IndentLevel"]),
]
#: (statements, reads): what setting xlEdgeTop and the like on a style does, and what clears the reads they give.
EDGE_SETS = [
    (['Set s = wb.Styles.Add("E1")', "s.Borders(8).LineStyle = 1"],
     [f"s.Borders({index}).LineStyle" for index in (-4131, -4152, -4160, -4107, 7, 8, 9, 10)]),
    (['Set s = wb.Styles.Add("E2")', "s.Borders(12).Weight = 4"],
     [f"s.Borders({index}).LineStyle" for index in (5, 6, 11, 12)]),
    (['Set s = wb.Styles.Add("E3")', "s.Borders(-4131).LineStyle = 1", "s.Borders(-4131).LineStyle = -4142"],
     ["s.Borders(7).LineStyle", "s.Borders(7).Weight"]),
    (['Set s = wb.Styles.Add("E4")', "s.Borders(-4131).LineStyle = 1", "s.Borders.Color = 255"],
     ["s.Borders(7).LineStyle", "s.Borders(-4131).Color", "s.Borders(-4152).LineStyle", "s.Borders(8).LineStyle"]),
    (['Set s = wb.Styles.Add("E5")', "s.Borders(-4131).LineStyle = 1", "s.Borders.Weight = 4"],
     ["s.Borders(7).LineStyle", "s.Borders(8).LineStyle", "s.Borders(-4152).LineStyle", "s.Borders(-4152).Weight"]),
    (['Set s = wb.Styles.Add("E6")', "s.Borders(-4131).Color = 255", "s.Borders(-4131).LineStyle = 1"],
     ["s.Borders(7).LineStyle"]),
]
#: (statements, reads): which font changes reach a cell once the style stops including its font: built-in
#: or not, including the fill or not, after an earlier change or not.
REACH_FONT = [
    (['ws.Range("A1").Style = "Good"', 'wb.Styles("Good").IncludeFont = False', 'wb.Styles("Good").Font.Underline = 2'],
     ['ws.Range("A1").Font.Underline']),
    (['Set s = wb.Styles.Add("Q")', *(f"s.Include{part} = False" for part in ("Number", "Alignment", "Border",
                                                                          "Protection")),
      'ws.Range("A2").Style = "Q"', "s.IncludeFont = False", "s.Font.Underline = 2"], ['ws.Range("A2").Font.Underline']),
    (['ws.Range("A3").Style = "Bad"', 'wb.Styles("Bad").Font.Bold = True', 'wb.Styles("Bad").IncludeFont = False',
      'wb.Styles("Bad").Font.Underline = 2'], ['ws.Range("A3").Font.Underline']),
    (['ws.Range("A4").Style = "Neutral"', 'wb.Styles("Neutral").IncludeFont = False',
      'wb.Styles("Neutral").Interior.Color = 255', 'wb.Styles("Neutral").Font.Underline = 2'],
     ['ws.Range("A4").Interior.Color', 'ws.Range("A4").Font.Underline']),
    (['Set s = wb.Styles.Add("R")', *(f"s.Include{part} = False" for part in ("Number", "Alignment", "Border",
                                                                          "Protection")),
      'ws.Range("A5").Style = "R"', "s.Font.Bold = True", "s.IncludeFont = False", "s.Font.Underline = 2"],
     ['ws.Range("A5").Font.Underline']),
    (['ws.Range("A6").Style = "Title"', 'wb.Styles("Title").IncludeFont = False', 'wb.Styles("Title").Font.Underline = 2'],
     ['ws.Range("A6").Font.Underline']),
    (['Set s = wb.Styles.Add("T")', 'ws.Range("A7").Style = "T"', "s.IncludeNumber = False", "s.IncludeFont = False",
      "s.Font.Underline = 2"], ['ws.Range("A7").Font.Underline']),
]

#: A style of a workbook that is not the active one, changed and deleted: which workbook's style changes.
ELSEWHERE = ['wb.Styles("Good").Font.Bold = True', 'wb.Styles("Mine").Font.Italic = True', 'wb.Styles("Bad").Delete',
             'wb.Styles("Mine").Delete', 'wb.Styles.Add "Added"']
ELSEWHERE_READS = ['wb.Styles("Good").Font.Bold', 'two.Styles("Good").Font.Bold', 'ws.Range("A1").Font.Bold',
                   'wb.Styles("Mine").Font.Italic', 'wb.Styles.Count', 'two.Styles.Count', 'wb.Styles("Bad").Name',
                   'two.Styles("Bad").Name', 'wb.Styles("Added").Name', 'two.Styles("Added").Name']

#: (cell, statements) before Good is deleted: a number format another style left, and one a macro set.
DELETE_KEPT = [("A1", ['c.Style = "Comma"', 'c.Style = "Good"']),
               ("A2", ['c.Style = "Good"', 'c.NumberFormat = "0.0"']),
               ("A3", ['c.Style = "Heading 1"', 'c.Style = "Good"'])]

#: Statements that change styles with sheets protected, each read by the error it gives.
GUARDS = ["ws2.Protect", 'wb.Styles("Good").Font.Bold = True', 'wb.Styles("Bad").Font.Bold = True', "ws2.Unprotect",
          "ws.Protect", 'wb.Styles("Good").Font.Italic = True', 'wb.Styles("Bad").Font.Italic = True',
          'wb.Styles.Add "Added"', 'wb.Styles("Added").Font.Bold = True', 'wb.Styles("Neutral").Delete',
          'wb.Styles("Good").Delete', "ws.Unprotect", "wb.Protect Structure:=True",
          'wb.Styles("Good").Font.Strikethrough = True', "wb.Unprotect"]


def _cells(cells: list[tuple[str, list[str]]]) -> list[str]:
    lines = []
    for cell, done in cells:
        lines += [f'Set c = ws.Range("{cell}")', *statements(done)]
    return lines


def module() -> str:
    """The whole probe: a function per section, fields apart by `, items by ^ and groups by ~."""
    parts = [HELPER]
    parts += ["Private Function StyleProps(s As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(STYLE_READS), "StyleProps = out", "End Function"]
    parts += ["Private Function Look(c As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(LOOK), "Look = out", "End Function"]
    parts += ["Private Function Snap(r As Object) As String", "Dim c As Object, out As String",
              "For Each c In r.Cells", 'out = out & Look(c) & "^"', "Next", "Snap = out", "End Function"]
    parts += ["Private Function Listing(wb As Object) As String", "Dim out As String, i As Long",
              "For i = 1 To wb.Styles.Count", 'out = out & wb.Styles(i).Name & "`"', "Next", "Listing = out",
              "End Function"]
    # Cells that went through different things, then Good and Comma changed one step at a time.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A11").Value = 1234.5',
            *_cells(FOLLOW_CELLS), 'out = out & Snap(ws.Range("A1:A11")) & "~"']
    for lines, names in FOLLOW_CHANGES:
        body += [*statements(lines), 'out = out & Snap(ws.Range("A1:A11")) & "~"',
                 *(f'out = out & StyleProps(wb.Styles("{name}")) & "~"' for name in names)]
    body += ['wb.SaveAs Filename:=target & "follow.xlsx"', "wb.Close False"]
    parts += function("Follow", "target As String", body)
    # Normal changed under cells given other styles, and the sheet's geometry.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A6").Value = 1234.5',
            *_cells(NORMAL_CELLS), 'out = out & Snap(ws.Range("A1:A7")) & "~"']
    for lines, names in NORMAL_CHANGES:
        body += [*statements(lines), 'out = out & Snap(ws.Range("A1:A7")) & "~"',
                 *(f'out = out & StyleProps(wb.Styles("{name}")) & "~"' for name in names)]
    body += ['wb.SaveAs Filename:=target & "normal.xlsx"', "wb.Close False"]
    parts += function("Normal", "target As String", body)
    # Normal's font size and typeface, which the sheet's geometry follows.
    parts += function("NormalFont", "", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1").Value = 1',
        *reads(GEOMETRY), *statements(['wb.Styles("Normal").Font.Size = 14']),
        *reads([*GEOMETRY, 'wb.Styles("Good").Font.Size', 'wb.Styles("Comma").Font.Size', 'ws.Range("A1").Font.Size']),
        *statements(['wb.Styles("Normal").Font.Name = "Arial"']),
        *reads([*GEOMETRY, 'wb.Styles("Good").Font.Name', 'wb.Styles("Comma").Font.Name', 'ws.Range("A1").Font.Name']),
        "wb.Close False"])
    # Every property a macro can set, on a new style a cell has; then built-in styles.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A2").Value = 1234.5',
            'Set s = wb.Styles.Add("Set")', 'ws.Range("A1").Style = "Set"', 'Set c = ws.Range("A1")']
    for group in SETS:
        body += [*statements(group), 'out = out & StyleProps(s) & "~" & Look(c) & "~"']
    body += [*statements(BUILT_INS), 'out = out & StyleProps(wb.Styles("Title")) & "~"',
             'out = out & Look(ws.Range("A2")) & "~"', 'wb.SaveAs Filename:=target & "setting.xlsx"', "wb.Close False"]
    parts += function("Setting", "target As String", body)
    # Good deleted under cells, a held Style, the name again; then more deletes.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A6").Value = 1234.5',
            'ws.Range("B1").Font.Italic = True', 'wb.Styles.Add "Mine", ws.Range("B1")', *_cells(DELETE_CELLS),
            'Set s = wb.Styles("Good")', *statements(['wb.Styles("Good").Delete']),
            *reads(["wb.Styles.Count", "s.Name", 'wb.Styles("Good").Name', 'TypeName(ws.Range("A1").Style)']),
            'out = out & "~" & Snap(ws.Range("A1:A6")) & "~"',
            *statements(['ws.Range("B2").Style = "Good"', 'wb.Styles.Add "Good"']),
            *reads(['wb.Styles("Good").BuiltIn', "wb.Styles.Count", "s.Name"]),
            *statements(['ws.Range("B3").Style = "Good"']), 'out = out & "~" & Look(ws.Range("B3")) & "~"',
            *statements(['wb.Styles("Mine").Delete', 'wb.Styles("Normal").Delete', 'wb.Styles("Bad").Delete',
                         'wb.Styles("Note").Delete', "wb.Styles(1).Delete"]),
            *reads(["wb.Styles.Count", "wb.Styles(1).Name"]),
            'out = out & "~" & Snap(ws.Range("A1:A6")) & "~" & Listing(wb) & "~"',
            'wb.SaveAs Filename:=target & "delete.xlsx"', "wb.Close False"]
    parts += function("Delete", "target As String", body)
    # Another workbook's styles merged in, some of them named as this one's.
    body = ["Dim two As Object, sheet As Object",
            "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("B1").Font.Bold = True',
            'wb.Styles.Add "Mine", ws.Range("B1")', 'wb.Styles("Good").Font.Bold = True', 'wb.Styles("Bad").Delete',
            'wb.Styles("Neutral").Font.Italic = True', 'ws.Range("A1").Style = "Good"',
            "Set two = Workbooks.Add(xlWBATWorksheet)", "Set sheet = two.Worksheets(1)",
            'sheet.Range("A1:A3").Value = 1234.5', 'sheet.Range("A1").Style = "Good"',
            'sheet.Range("A2").Style = "Neutral"', 'two.Styles.Add "Theirs"', 'two.Styles.Add "mine"',
            'sheet.Range("A3").Style = "mine"', *statements(["two.Styles.Merge wb"]),
            *reads(["two.Styles.Count", 'two.Styles("Mine").Name', 'two.Styles("Mine").Font.Bold',
                    'two.Styles("Good").Font.Bold', 'two.Styles("Bad").Name', 'two.Styles("Neutral").Font.Italic',
                    'two.Styles("Theirs").Name']),
            'out = out & "~" & Snap(sheet.Range("A1:A3")) & "~"',
            *statements(["two.Styles.Merge two", 'two.Styles.Merge "x"', "two.Styles.Merge sheet"]),
            'out = out & "~" & Listing(two) & "~"', 'two.SaveAs Filename:=target & "merge.xlsx"', "two.Close False",
            "wb.Close False"]
    parts += function("Merge", "target As String", body)
    # Whole rows and columns given styles that then change.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Columns("E").Style = "Bad"',
            'ws.Rows(5).Style = "Neutral"', *statements(['wb.Styles("Bad").Font.Bold = True',
                                                         'wb.Styles("Neutral").Font.Italic = True'])]
    for cell in ("E1", "E5", "F5", "E9", "H9"):
        body += [f'out = out & Look(ws.Range("{cell}")) & "^"']
    body += ['wb.SaveAs Filename:=target & "whole_changed.xlsx"', "wb.Close False"]
    parts += function("WholeChanged", "target As String", body)
    # A style changed while cells using it are on a protected sheet.
    parts += function("Protected", "", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1").Style = "Good"',
        "ws.Protect", *statements(['wb.Styles("Good").Font.Bold = True']), 'out = out & Look(ws.Range("A1")) & "~"',
        *statements(["ws.Unprotect", "ws.Protect AllowFormattingCells:=True", 'wb.Styles("Good").Font.Italic = True']),
        'out = out & Look(ws.Range("A1")) & "~"', "ws.Unprotect", "wb.Close False"])
    # Styles the file holds, changed in a workbook opened again.
    parts += function("Reopened", "source As String, target As String", [
        'Set wb = Workbooks.Open(source & "mixed.xlsx")', "Set ws = wb.Worksheets(1)",
        *statements(['wb.Styles("Good").Font.Bold = True', 'wb.Styles("Comma").NumberFormat = "0.000"']),
        'out = out & Snap(ws.Range("A1:A21")) & "~"', 'wb.SaveAs Filename:=target & "reopen_changed.xlsx"',
        "wb.Close False"])
    # Formats made out of row order, saved to show the order Excel's tables keep.
    parts += function("Order", "target As String", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A13").Value = 1234.5',
        *statements([f'ws.Range("{line.split(".", 1)[0]}").{line.split(".", 1)[1]}' for line in ORDER]),
        'wb.SaveAs Filename:=target & "order.xlsx"', "wb.Close False"])
    # A file holding entries nothing uses, opened, given new formats and saved.
    parts += function("Planted", "source As String, target As String", [
        'Set wb = Workbooks.Open(source & "planted.xlsx")', "Set ws = wb.Worksheets(1)",
        *statements(['ws.Range("B1").Font.Italic = True', 'ws.Range("B2").Font.Color = RGB(255, 0, 0)',
                     'ws.Range("B3").NumberFormat = "0.000"']),
        "out = out & Listing(wb)", 'wb.SaveAs Filename:=target & "planted_saved.xlsx"', "wb.Close False"])
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A3").Value = 1234.5']
    for lines, answers in QUIRKS:
        body += [*statements(lines), *reads(answers), 'out = out & "~"']
    body += ['wb.SaveAs Filename:=target & "quirks.xlsx"', "wb.Close False"]
    parts += function("Quirks", "target As String", body)
    parts += function("Guards", "", [
        "Dim ws2 As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
        "Set ws2 = wb.Worksheets.Add", 'ws.Range("A1").Style = "Good"', *statements(GUARDS),
        'out = out & "~" & Look(ws.Range("A1")) & "~"', *reads(['wb.Styles("Added").Font.Bold', "wb.Styles.Count"]),
        "wb.Close False"])
    # Cells whose style is deleted, then Normal changed under them.
    parts += function("AfterDelete", "target As String", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A3").Value = 1234.5',
        'ws.Range("A1").Style = "Good"', 'ws.Range("A2").Style = "Good"', 'ws.Range("A2").Font.Italic = True',
        'ws.Range("A3").Style = "Comma"', 'wb.Styles("Good").Delete', 'wb.Styles("Comma").Delete',
        'out = out & Snap(ws.Range("A1:A3")) & "~"',
        *statements(['wb.Styles("Normal").NumberFormat = "0.0"', 'wb.Styles("Normal").Interior.Color = RGB(255, 255, 0)',
                     'wb.Styles("Normal").Font.Color = RGB(0, 0, 255)']),
        'out = out & Snap(ws.Range("A1:A3")) & "~"', 'wb.SaveAs Filename:=target & "after_delete.xlsx"',
        "wb.Close False"])
    # A file whose own entries repeat: a font, an xf, an xf apart only in its flags; custom formats it numbers.
    parts += function("Duplicates", "source As String, target As String", [
        'Set wb = Workbooks.Open(source & "duplicates.xlsx")', "Set ws = wb.Worksheets(1)",
        *statements(['ws.Range("C1").NumberFormat = "0.0"']),
        *reads(['ws.Range("B1").NumberFormat', 'ws.Range("C1").NumberFormat', 'ws.Range("B4").Style.Name']),
        'wb.SaveAs Filename:=target & "duplicates_saved.xlsx"', "wb.Close False"])
    # Which protected sheets stop a change: another workbook's, active or not.
    parts += function("Guards2", "", [
        "Dim two As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
        'ws.Range("A1").Style = "Good"', "Set two = Workbooks.Add(xlWBATWorksheet)", "two.Worksheets(1).Protect",
        "wb.Activate", *statements(['wb.Styles("Good").Font.Bold = True']), "two.Activate",
        *statements(['wb.Styles("Good").Font.Italic = True']), "two.Close False",
        *reads(['ws.Range("A1").Font.Bold', 'ws.Range("A1").Font.Italic']), "wb.Close False"])
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A4").Value = 1234.5']
    for lines, answers in [*REACH, *INDENTS, *EDGE_SETS]:
        body += [*statements(lines), *reads(answers), 'out = out & "~"']
    body += ['wb.SaveAs Filename:=target & "reach.xlsx"', "wb.Close False"]
    parts += function("Reach", "target As String", body)
    # A number format another style left under Good, one a macro set, and a border, when Good is deleted.
    parts += function("DeleteKept", "target As String", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A3").Value = 1234.5',
        *_cells(DELETE_KEPT), 'wb.Styles("Good").Delete', 'out = out & Snap(ws.Range("A1:A3")) & "~"',
        'wb.SaveAs Filename:=target & "delete_kept.xlsx"', "wb.Close False"])
    # Another workbook's styles of its own merged in, and where a save puts them.
    parts += function("MergeOrder", "target As String", [
        "Dim two As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", 'wb.Styles.Add "Zed"', 'wb.Styles.Add "Alpha"',
        "Set two = Workbooks.Add(xlWBATWorksheet)", 'two.Styles.Add "Mid"', *statements(["two.Styles.Merge wb"]),
        'two.Worksheets(1).Range("A1").Style = "Zed"', "out = out & Listing(two)",
        'two.SaveAs Filename:=target & "merge_order.xlsx"', "two.Close False", "wb.Close False"])
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A7").Value = 1234.5']
    for lines, answers in REACH_FONT:
        body += [*statements(lines), *reads(answers), 'out = out & "~"']
    body += ['wb.SaveAs Filename:=target & "reach_font.xlsx"', "wb.Close False"]
    parts += function("ReachFont", "target As String", body)
    # The style's own workbook with a protected sheet while another, unprotected, is active; then the file whose
    # custom formats are listed out of the order of their ids.
    parts += function("Guards3", "", [
        "Dim two As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
        'ws.Range("A1").Style = "Good"', "ws.Protect", "Set two = Workbooks.Add(xlWBATWorksheet)", "two.Activate",
        *statements(['wb.Styles("Good").Font.Bold = True', 'wb.Styles("Bad").Font.Bold = True']), "two.Close False",
        *reads(['ws.Range("A1").Font.Bold', 'wb.Styles("Bad").Font.Bold']), "ws.Unprotect", "wb.Close False"])
    parts += function("Elsewhere", "", [
        "Dim two As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
        'wb.Styles.Add "Mine"', 'ws.Range("A1").Style = "Good"', "Set two = Workbooks.Add(xlWBATWorksheet)",
        "two.Activate", *statements(ELSEWHERE), *reads(ELSEWHERE_READS), "two.Close False", "wb.Close False"])
    parts += function("Renumbered", "source As String, target As String", [
        'Set wb = Workbooks.Open(source & "renumbered.xlsx")', "Set ws = wb.Worksheets(1)",
        *statements(['ws.Range("C1").NumberFormat = "0.0"']),
        *reads(['ws.Range("B1").NumberFormat', 'ws.Range("B2").NumberFormat']),
        'wb.SaveAs Filename:=target & "renumbered_saved.xlsx"', "wb.Close False"])
    # A custom format made before built-in ones a style spells out.
    parts += function("Spelled", "target As String", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
        *statements(['ws.Range("A1").NumberFormat = "0.0"', 'ws.Range("A2").Style = "Comma"',
                     'ws.Range("A3").Style = "Currency [0]"', 'ws.Range("A4").NumberFormat = "0.00%"',
                     'ws.Range("A5").NumberFormat = "$#,##0.00"']),
        'wb.SaveAs Filename:=target & "spelled.xlsx"', "wb.Close False"])
    return "\n".join(parts) + "\n"


def duplicate() -> None:
    """duplicates.xlsx: only_good.xlsx given custom formats 170 (a cell's) and 171 (nobody's), a second copy of
    Good's font, and three xfs a cell each: one on the copied font, one the same as Good's, one apart only in a
    flag."""
    with zipfile.ZipFile(OUT / "only_good.xlsx") as package:
        parts = [(info, package.read(info.filename)) for info in package.infolist()]
    good_font = '<font><sz val="11"/><color rgb="FF006100"/><name val="Aptos Narrow"/><family val="2"/>' \
                '<scheme val="minor"/></font>'
    styles_edits = [
        ('<fonts count="2"', '<numFmts count="2"><numFmt numFmtId="170" formatCode="0.0000000"/>'
                             '<numFmt numFmtId="171" formatCode="0.000000"/></numFmts><fonts count="3"'),
        ("</font></fonts>", "</font>" + good_font + "</fonts>"),
        ('<cellXfs count="2">', '<cellXfs count="6">'),
        ("</cellXfs>", '<xf numFmtId="170" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                       '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="1"/>'
                       '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/>'
                       '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1" applyFont="1"/></cellXfs>'),
    ]
    sheet_edits = [
        ('<dimension ref="A1"/>', '<dimension ref="A1:B4"/>'),
        ('<row r="1" spans="1:1" x14ac:dyDescent="0.35"><c r="A1" s="1"/></row>',
         '<row r="1" spans="1:2" x14ac:dyDescent="0.35"><c r="A1" s="1"/><c r="B1" s="2"><v>1</v></c></row>'
         + "".join(f'<row r="{row}" spans="2:2" x14ac:dyDescent="0.35"><c r="B{row}" s="{row + 1}"><v>1</v></c></row>'
                   for row in (2, 3, 4))),
    ]
    _rewrite(parts, {"xl/styles.xml": styles_edits, "xl/worksheets/sheet1.xml": sheet_edits}, "duplicates.xlsx")


def renumber() -> None:
    """renumbered.xlsx: only_good.xlsx given custom formats 172 and 171, listed in that order, a cell each."""
    with zipfile.ZipFile(OUT / "only_good.xlsx") as package:
        parts = [(info, package.read(info.filename)) for info in package.infolist()]
    styles_edits = [
        ('<fonts count="2"', '<numFmts count="2"><numFmt numFmtId="172" formatCode="0.0000000"/>'
                             '<numFmt numFmtId="171" formatCode="0.000000"/></numFmts><fonts count="2"'),
        ('<cellXfs count="2">', '<cellXfs count="4">'),
        ("</cellXfs>", '<xf numFmtId="171" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                       '<xf numFmtId="172" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                       "</cellXfs>"),
    ]
    sheet_edits = [
        ('<dimension ref="A1"/>', '<dimension ref="A1:B2"/>'),
        ('<row r="1" spans="1:1" x14ac:dyDescent="0.35"><c r="A1" s="1"/></row>',
         '<row r="1" spans="1:2" x14ac:dyDescent="0.35"><c r="A1" s="1"/><c r="B1" s="2"><v>1</v></c></row>'
         '<row r="2" spans="2:2" x14ac:dyDescent="0.35"><c r="B2" s="3"><v>1</v></c></row>'),
    ]
    _rewrite(parts, {"xl/styles.xml": styles_edits, "xl/worksheets/sheet1.xml": sheet_edits}, "renumbered.xlsx")


def _rewrite(parts: list[tuple[zipfile.ZipInfo, bytes]], edits: dict[str, list[tuple[str, str]]], name: str) -> None:
    """Write a workbook's parts to tests/fixtures/cell_styles under ``name``, with each edit made once."""
    with zipfile.ZipFile(OUT / name, "w", zipfile.ZIP_DEFLATED) as package:
        for info, data in parts:
            if info.filename in edits:
                text = data.decode("utf-8")
                for old, new in edits[info.filename]:
                    assert text.count(old) == 1, old
                    text = text.replace(old, new)
                data = text.encode("utf-8")
            package.writestr(info, data)


def plant() -> None:
    """planted.xlsx: only_good.xlsx given a number format, a font, a fill, a border and a cell xf that nothing uses,
    and a style of its own that no cell has."""
    with zipfile.ZipFile(OUT / "only_good.xlsx") as package:
        parts = [(info, package.read(info.filename)) for info in package.infolist()]
    edits = [
        ("<fonts count=\"2\"", '<numFmts count="1"><numFmt numFmtId="170" formatCode="0.0000000"/></numFmts>'
                               '<fonts count="3"'),
        ("</font></fonts>", '</font><font><b/><sz val="30"/><color rgb="FF123456"/><name val="Arial"/>'
                            '<family val="2"/></font></fonts>'),
        ('<fills count="3">', '<fills count="4">'),
        ("</fill></fills>", '</fill><fill><patternFill patternType="solid"><fgColor rgb="FF654321"/>'
                            '<bgColor indexed="64"/></patternFill></fill></fills>'),
        ('<borders count="1">', '<borders count="2">'),
        ("</border></borders>", '</border><border><left style="dotted"><color indexed="64"/></left><right/><top/>'
                                "<bottom/><diagonal/></border></borders>"),
        ('<cellStyleXfs count="2">', '<cellStyleXfs count="3">'),
        ("</cellStyleXfs>", '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'),
        ('<cellXfs count="2">', '<cellXfs count="3">'),
        ("</cellXfs>", '<xf numFmtId="170" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" '
                       'applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>'),
        ('<cellStyles count="2">', '<cellStyles count="3">'),
        ("</cellStyles>", '<cellStyle name="Spare" xfId="2"/></cellStyles>'),
    ]
    with zipfile.ZipFile(OUT / "planted.xlsx", "w", zipfile.ZIP_DEFLATED) as package:
        for info, data in parts:
            if info.filename == "xl/styles.xml":
                text = data.decode("utf-8")
                for old, new in edits:
                    assert text.count(old) == 1, old
                    text = text.replace(old, new)
                data = text.encode("utf-8")
            package.writestr(info, data)


#: Each section and whether it takes the folder to open from as well as the one to save in.
SECTIONS = {"Follow": "target", "Normal": "target", "NormalFont": "", "Setting": "target", "Delete": "target",
            "Merge": "target", "WholeChanged": "target", "Protected": "", "Reopened": "both", "Order": "target",
            "Planted": "both", "Quirks": "target", "Guards": "", "AfterDelete": "target", "Duplicates": "both",
            "Guards2": "", "Reach": "target", "DeleteKept": "target", "MergeOrder": "target", "Spelled": "target",
            "ReachFont": "target", "Guards3": "", "Renumbered": "both", "Elsewhere": ""}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    saved, opened = str(folder) + "\\", str(OUT) + "\\"
    plant()
    duplicate()
    renumber()
    code = module()
    runs: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, folders in SECTIONS.items():
            args = {"target": (saved,), "": (), "both": (opened, saved)}[folders]
            result = excel.run_vba(code, name, args=args, timeout=900.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            runs[name] = str(result.value)
    OUT.mkdir(parents=True, exist_ok=True)
    for path in folder.glob("*.xlsx"):
        copy_saved(path, OUT / path.name)
    record = {"module": code, "sections": SECTIONS, "runs": runs}
    (OUT / "style_changes.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, answer in runs.items():
        print(f"{name}: {len(answer)} characters")


if __name__ == "__main__":
    main()
