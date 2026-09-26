"""What Excel's named cell styles are and do: Workbook.Styles, the Style object, Range.Style, and the file.

A new workbook's Styles collection and every property of every style in
it; every style given to a cell of a new workbook, in the collection's
order and the other way round, with what each cell reads after; styles
given over formats and formats over styles; styles a macro adds, alone
and based on a range, with how they list; whole rows and columns; a
macro's fill over a style's; a protected sheet; the errors. Workbooks
are saved along the way, and a second set of cases saves a workbook each,
to show what Excel writes for them: in a new workbook, in one it opens
again, under the Office 2013-2022 theme, and where Normal's font is not
the theme's body font.

The probe is one module, which the record keeps, run with a folder to open
workbooks from and one to save them in; the replay runs the same module
in the model.

    python scripts/measure_cell_styles.py

writes tests/fixtures/cell_styles/: cell_styles.json, the workbooks Excel
saved, and themed_base.xlsx, a new workbook given the older theme, which
the replay opens. tests/test_excel_cell_styles.py replays them.
"""

from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import copy_saved

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "cell_styles"
DIMENSIONS = ROOT / "tests" / "fixtures" / "dimensions"
#: The older Office theme, in the folder beside the one Excel runs from.
THEME = r'Application.Path & "\..\Document Themes 16\Office 2013 - 2022 Theme.thmx"'

HELPER = '''Private Function Show(v As Variant) As String
    If IsObject(v) Then
        Show = "Object:" & TypeName(v)
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''

#: What a Style object answers.
STYLE_READS = [
    "s.Name", "s.NameLocal", "s.BuiltIn", "s.Value", "s.IncludeNumber", "s.IncludeFont", "s.IncludeAlignment",
    "s.IncludeBorder", "s.IncludePatterns", "s.IncludeProtection", "s.NumberFormat", "s.NumberFormatLocal",
    "s.HorizontalAlignment", "s.VerticalAlignment", "s.WrapText", "s.Orientation", "s.IndentLevel", "s.ShrinkToFit",
    "s.ReadingOrder", "s.AddIndent", "s.Locked", "s.FormulaHidden", "s.MergeCells", "s.Font.Name", "s.Font.Size",
    "s.Font.Bold", "s.Font.Italic", "s.Font.Underline", "s.Font.Strikethrough", "s.Font.Superscript",
    "s.Font.Subscript", "s.Font.Color", "s.Font.ColorIndex", "s.Font.ThemeColor", "s.Font.TintAndShade",
    "s.Font.ThemeFont", "s.Font.FontStyle", "s.Interior.Color", "s.Interior.ColorIndex", "s.Interior.Pattern",
    "s.Interior.PatternColor", "s.Interior.PatternColorIndex", "s.Interior.ThemeColor", "s.Interior.TintAndShade",
    "s.Interior.PatternThemeColor", "s.Interior.PatternTintAndShade", "s.Borders.Count",
    *(f"s.Borders({edge}).{what}" for edge in (-4131, -4152, -4160, -4107, 5, 6, 7, 12)
      for what in ("LineStyle", "Weight", "Color", "ColorIndex", "ThemeColor", "TintAndShade")),
]
#: What a cell answers once a style or a format lands on it.
CELL_READS = [
    "c.Style.Name", "c.NumberFormat", "c.Font.Name", "c.Font.Size", "c.Font.Bold", "c.Font.Italic",
    "c.Font.Underline", "c.Font.Color", "c.Font.ColorIndex", "c.Font.ThemeColor", "c.Font.TintAndShade",
    "c.Interior.Color", "c.Interior.ColorIndex", "c.Interior.Pattern", "c.Interior.PatternColor",
    "c.Interior.PatternColorIndex", "c.Interior.ThemeColor", "c.Interior.TintAndShade",
    "c.Interior.PatternThemeColor", "c.Interior.PatternTintAndShade",
    *(f"c.Borders({edge}).{what}" for edge in (7, 8, 9, 10) for what in ("LineStyle", "Weight", "Color", "ThemeColor")),
    "c.HorizontalAlignment", "c.IndentLevel", "c.Locked",
]

#: Formats a cell is given before or after a style.
DIRECT = ["c.Font.Bold = True", "c.Font.Italic = True", "c.Font.Color = RGB(255, 0, 0)",
          "c.Interior.Color = RGB(255, 255, 0)", "c.Borders(7).LineStyle = 1", 'c.NumberFormat = "0.00"',
          "c.HorizontalAlignment = -4108", "c.Locked = False"]
#: (cell, statements): styles over formats, formats over styles, and the ways of naming a style.
MIXED = [
    ("A1", [*DIRECT, 'c.Style = "Good"']),
    ("A2", [*DIRECT, 'c.Style = "Comma"']),
    ("A3", [*DIRECT, 'c.Style = "Normal"']),
    ("A4", [*DIRECT, 'c.Style = "Heading 1"']),
    ("A5", [*DIRECT, 'c.Style = "Note"']),
    ("A6", ['c.Style = "Good"', "c.Font.Bold = True"]),
    ("A7", ['c.Style = "Good"', 'c.NumberFormat = "0.00"']),
    ("A8", ['c.Style = "Good"', "c.Interior.Color = RGB(0, 0, 255)"]),
    ("A9", ['c.Style = "Good"', 'c.Style = "Bad"']),
    ("A10", ['c.Style = "Good"', 'c.Style = "Comma"']),
    ("A11", ['c.Style = "Comma"', 'c.Style = "Good"']),
    ("A12", ['c.Style = "good"']),
    ("A13", ['c.Style = ActiveWorkbook.Styles("Title")']),
    ("A14", ['c.Style = "Hyperlink"']),
    ("A15", ['c.Style = "Nope"']),
    ("A16", ['c.Style = ""']),
    ("A17", ["c.Style = 5"]),
    ("A18", ['c.Style = "Heading 1"', "c.Font.Bold = False", "c.Font.Bold = True"]),
    ("A19", ['c.Style = "Good"', "c.Font.Color = c.Font.Color"]),
    ("A20", ['c.Style = "Currency [0]"', 'c.Style = "Percent"']),
]
#: How a range whose cells differ reads Style.
MIXED_READS = ['ws.Range("A1:A2").Style', 'TypeName(ws.Range("A1:A2").Style)', 'ws.Range("A9:A10").Style.Name',
               'ws.Range("A21:A22").Style.Name', 'ws.Range("A1,A9").Style', 'ws.Range("A12,A19").Style.Name']

#: Names a macro adds a style under, in an order no sort gives, some of them refused.
NAMES = ["zeta", "Alpha", "_under", "1st", "Normal 2", "a", "B", "-dash", "(paren)", "#hash", "~tilde", "20%",
         "Accent10", "Heading 10", "normal x", "Z", "good2", "a:b", "Accent1 x", "Comma0", "  spaced", "note",
         "Good", "good", "zeta", "Normal", "Hyperlink"]
#: Styles a macro makes, read, and given to cells: (style, cell).
MADE = [("zeta", "A1"), ("spaced", "A2"), ("FromGood", "A3"), ("FromTwo", "A4"), ("Indented", "A5")]

#: (cell, statements): a macro's fill over a style's.
FILLS = [
    ("A1", ['c.Style = "Good"', "c.Interior.Color = RGB(0, 0, 255)"]),
    ("A2", ['c.Style = "20% - Accent1"', "c.Interior.Color = RGB(0, 0, 255)"]),
    ("A3", ['c.Style = "Good"', "c.Interior.Pattern = 17"]),
    ("A4", ['c.Style = "Good"', "c.Interior.PatternColor = RGB(255, 0, 0)"]),
    ("A5", ['c.Style = "Accent1"', "c.Interior.TintAndShade = 0.5"]),
    ("A6", ['c.Style = "Note"', "c.Interior.ColorIndex = 3"]),
    ("A7", ['c.Style = "Accent2"', "c.Interior.ThemeColor = 7"]),
    ("A8", ['c.Style = "Good"', "c.Interior.Pattern = -4142"]),
]

#: Workbooks saved each for itself: (name, open the mixed workbook instead of a new one, statements).
FILES = [
    ("only_good", False, ['ws.Range("A1").Style = "Good"']),
    ("only_comma", False, ['ws.Range("A1").Style = "Comma"']),
    ("only_accent", False, ['ws.Range("A1").Style = "20% - Accent1"']),
    ("bold_then_good", False, ['ws.Range("A1").Font.Bold = True', 'ws.Range("A2").Style = "Good"']),
    ("fill_then_note", False, ['ws.Range("A1").Interior.Color = RGB(1, 2, 3)', 'ws.Range("A2").Style = "Note"']),
    ("good_then_normal", False, ['ws.Range("A1").Style = "Good"', 'ws.Range("A1").Style = "Normal"']),
    ("good_then_cleared", False, ['ws.Range("A1").Style = "Good"', 'ws.Range("A1").ClearFormats']),
    ("good_then_bad", False, ['ws.Range("A1").Style = "Good"', 'ws.Range("A1").Style = "Bad"']),
    ("currencies", False, ['ws.Range("A1").Style = "Comma"', 'ws.Range("A2").Style = "Currency [0]"',
                           'ws.Range("A3").Style = "Currency"', 'ws.Range("A4").Style = "Comma [0]"']),
    ("added_unused", False, ['ActiveWorkbook.Styles.Add "Mine"']),
    ("added_used", False, ['ActiveWorkbook.Styles.Add "Mine"', 'ws.Range("A1").Style = "Mine"']),
    ("added_based", False, ['ws.Range("B1").Font.Italic = True', 'ws.Range("B1").NumberFormat = "0.0"',
                            'ActiveWorkbook.Styles.Add "Based", ws.Range("B1")', 'ws.Range("A1").Style = "Based"']),
    ("added_aligned", False, ['ws.Range("B1").HorizontalAlignment = -4131', 'ws.Range("B1").IndentLevel = 2',
                              'ws.Range("B1").Locked = False', 'ActiveWorkbook.Styles.Add "Aligned", ws.Range("B1")',
                              'ws.Range("A1").Style = "Aligned"']),
    ("added_names", False, ['ActiveWorkbook.Styles.Add "zeta"', 'ActiveWorkbook.Styles.Add "Alpha"',
                            'ActiveWorkbook.Styles.Add "_under"', 'ActiveWorkbook.Styles.Add "1st"',
                            'ActiveWorkbook.Styles.Add "Normal 2"', 'ws.Range("A1").Style = "zeta"',
                            'ws.Range("A2").Style = "Alpha"']),
    ("built_in_and_added", False, ['ws.Range("A1").Style = "Good"', 'ActiveWorkbook.Styles.Add "Mine"',
                                   'ws.Range("A2").Style = "Mine"', 'ws.Range("A3").Style = "Bad"']),
    ("reopen_neutral", True, ['ws.Range("C1").Style = "Neutral"']),
    ("reopen_accent", True, ['ws.Range("C1").Style = "Accent1"', 'ws.Range("C2").Style = "20% - Accent3"']),
    ("reopen_underline", True, ['ws.Range("C1").Font.Underline = 2']),
    ("reopen_added", True, ['ActiveWorkbook.Styles.Add "Mine"', 'ws.Range("C1").Style = "Mine"',
                            'ws.Range("C2").Style = "Title"']),
]
#: Styles given one to a cell of the workbook whose Normal font is Calibri under the Aptos theme.
CALIBRI = ["Good", "Title", "Heading 1", "Comma", "20% - Accent1", "Note", "Accent1", "Heading 4", "Total"]

#: Reads that go wrong, and the collection's own.
ERRORS = ['wb.Styles(0).Name', 'wb.Styles(48).Name', 'wb.Styles("Nope").Name', 'wb.Styles("Hyperlink").Name',
          'wb.Styles("Followed Hyperlink").Name', 'wb.Styles.Count', 'TypeName(wb.Styles)', 'TypeName(wb.Styles(1))',
          'TypeName(wb.Styles(1).Font)', 'TypeName(wb.Styles(1).Interior)', 'TypeName(wb.Styles(1).Borders)',
          'TypeName(wb.Styles(1).Borders(-4107))', 'wb.Styles("NORMAL").Name', 'wb.Styles(1)',
          'ws.Range("A1:B2").Style.Name', 'TypeName(ws.Range("A1:B2").Style)', 'ws.Range("A1").Style',
          'wb.Styles("Heading 1").Borders(1).LineStyle', 'wb.Styles("Heading 1").Borders(4).LineStyle',
          'wb.Styles("Heading 1").Borders(0).LineStyle', 'wb.Styles("Heading 1").Borders(13).LineStyle',
          'wb.Styles("Heading 1").Borders(-4108).LineStyle', 'wb.Styles.Creator', 'TypeName(wb.Styles.Parent)',
          'TypeName(wb.Styles(1).Parent)']


def _quoted(text: str) -> str:
    return text.replace('"', '""')


def reads(expressions: list[str]) -> list[str]:
    """Lines that read each expression into ``out``: Show's answer, or E and the error."""
    lines = []
    for expression in expressions:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & "`" Else out = out & Show(v) & "`"']
    return lines


def statements(lines: list[str]) -> list[str]:
    """Lines that run each statement, each leaving - or S and the error in ``out``."""
    found = []
    for line in lines:
        found += ["Err.Clear", line, 'If Err.Number <> 0 Then out = out & "S" & Err.Number & "`" Else out = out & "-`"']
    return found


def function(name: str, parameters: str, body: list[str]) -> list[str]:
    return [f"Public Function {name}({parameters}) As String",
            "Dim wb As Object, ws As Object, c As Object, s As Object, out As String, v As Variant, i As Long",
            "Application.DisplayAlerts = False", "On Error Resume Next", *body, "On Error GoTo 0",
            f"{name} = out", "End Function"]


def module() -> str:
    """The whole probe: a function per section, each giving its answers, fields apart by ` and items by ^."""
    parts = [HELPER]
    parts += ["Private Function StyleProps(s As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(STYLE_READS), "StyleProps = out", "End Function"]
    parts += ["Private Function CellProps(c As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(CELL_READS), "CellProps = out", "End Function"]
    parts += ["Private Function Listing(wb As Object) As String", "Dim out As String, i As Long",
              "For i = 1 To wb.Styles.Count", 'out = out & wb.Styles(i).Name & "`"', "Next", "Listing = out",
              "End Function"]
    # Every style of a new workbook, then every one given to a cell, one way round and the other.
    parts += function("EveryStyle", "", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "out = Listing(wb) & \"~\"",
        "For Each s In wb.Styles", 'out = out & StyleProps(s) & "^"', "Next", "wb.Close False"])
    for name, order, file in (("EveryCell", "1 To wb.Styles.Count", "every_style.xlsx"),
                              ("EveryCellBack", "wb.Styles.Count To 1 Step -1", "every_style_reverse.xlsx")):
        parts += function(name, "target As String", [
            "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", "Dim row As Long",
            f"For i = {order}", "row = row + 1", "Set c = ws.Cells(row, 1)", "c.Value = 1234.5",
            "c.Style = wb.Styles(i).Name", 'out = out & CellProps(c) & "^"', "Next",
            f'wb.SaveAs Filename:=target & "{file}"', "wb.Close False"])
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for cell, lines in MIXED:
        body += [f'Set c = ws.Range("{cell}")', "c.Value = 1234.5", *statements(lines), 'out = out & CellProps(c) & "^"']
    body += ['ws.Range("A21").Style = "Good"', 'out = out & "~"', *reads(MIXED_READS),
             'wb.SaveAs Filename:=target & "mixed.xlsx"', "wb.Close False"]
    parts += function("Mixed", "target As String", body)
    # Styles a macro makes.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            *statements([f'wb.Styles.Add "{_quoted(name)}"' for name in NAMES]), 'out = out & "~" & Listing(wb) & "~"',
            'ws.Range("B1").Style = "Good"', 'ws.Range("B1").Font.Bold = True', 'ws.Range("B2").Font.Italic = True',
            'ws.Range("B2").NumberFormat = "0.0"', 'ws.Range("B3").HorizontalAlignment = -4131',
            'ws.Range("B3").IndentLevel = 2', 'ws.Range("B3").WrapText = True', 'ws.Range("B3").Locked = False',
            *statements(['wb.Styles.Add "FromGood", ws.Range("B1")', 'wb.Styles.Add "FromTwo", ws.Range("B2:B1")',
                         'wb.Styles.Add "Indented", ws.Range("B3")', 'wb.Styles.Add "FromText", "B1"',
                         'wb.Styles.Add "FromAnother", Workbooks.Add(xlWBATWorksheet).Worksheets(1).Range("A1")']),
            'out = out & "~" & TypeName(wb.Styles.Add("Returned")) & "~"']
    for name, cell in MADE:
        body += [f'out = out & StyleProps(wb.Styles("{name}")) & "^"']
    body.append('out = out & "~"')
    for name, cell in MADE:
        body += [f'Set c = ws.Range("{cell}")', "c.Value = 1234.5", f'c.Style = "{name}"',
                 'out = out & CellProps(c) & "^"']
    body += ['wb.SaveAs Filename:=target & "custom.xlsx"', "wb.Close False"]
    parts += function("Custom", "target As String", body)
    # Whole rows and columns, a macro's fill over a style's, and a protected sheet.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            *statements(['ws.Columns("E").Style = "Bad"', 'ws.Rows(5).Style = "Neutral"',
                         'ws.Range("G1:G3").Style = "Title"'])]
    for cell in ("E1", "E5", "F5", "G2", "H9"):
        body += [f'Set c = ws.Range("{cell}")', 'out = out & CellProps(c) & "^"']
    body += ['out = out & "~"', *reads(['ws.Columns("E").Style.Name', 'ws.Rows(5).Style.Name',
                                        'ws.Range("E4:E6").Style.Name', 'ws.Cells.Style.Name']),
             'wb.SaveAs Filename:=target & "whole.xlsx"', "wb.Close False"]
    parts += function("Whole", "target As String", body)
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for cell, lines in FILLS:
        body += [f'Set c = ws.Range("{cell}")', *statements(lines), 'out = out & CellProps(c) & "^"']
    body += ['wb.SaveAs Filename:=target & "fills.xlsx"', "wb.Close False"]
    parts += function("Fills", "target As String", body)
    parts += function("Protected", "", [
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
        *statements(["ws.Protect", 'ws.Range("A1").Style = "Good"', "ws.Unprotect",
                     "ws.Protect AllowFormattingCells:=True", 'ws.Range("A2").Style = "Good"', "ws.Unprotect"]),
        *reads(['ws.Range("A1").Style.Name', 'ws.Range("A2").Style.Name']), "wb.Close False"])
    parts += function("Errors", "", ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
                                     *reads(ERRORS), "wb.Close False"])
    # A workbook saved for each case, some of them the mixed one opened again.
    body = []
    for name, reopen, lines in FILES:
        body += ['Set wb = Workbooks.Open(source & "mixed.xlsx")' if reopen else
                 "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *statements(lines),
                 'out = out & wb.Styles.Count & "^"', f'wb.SaveAs Filename:=target & "{name}.xlsx"', "wb.Close False"]
    parts += function("Files", "source As String, target As String", body)
    # The older theme: every style, then the reads of three; and Normal in Calibri under the Aptos theme.
    parts += function("Themed", "source As String, target As String", [
        'Set wb = Workbooks.Open(source & "themed_base.xlsx")', "Set ws = wb.Worksheets(1)",
        "out = Listing(wb) & \"~\"",
        *(f'out = out & StyleProps(wb.Styles("{name}")) & "^"' for name in ("Normal", "Title", "Good", "Comma")),
        "For i = 1 To wb.Styles.Count", "ws.Cells(i, 1).Style = wb.Styles(i).Name", "Next",
        'wb.SaveAs Filename:=target & "themed.xlsx"', "wb.Close False"])
    parts += function("Calibri", "source As String, target As String", [
        'Set wb = Workbooks.Open(source & "calibri.xlsx")', "Set ws = wb.Worksheets(1)",
        *(f'out = out & StyleProps(wb.Styles("{name}")) & "^"' for name in ("Normal", "Good", "Comma")),
        *(f'ws.Range("J{row}").Style = "{name}"' for row, name in enumerate(CALIBRI, start=1)),
        'wb.SaveAs Filename:=target & "calibri.xlsx"', "wb.Close False"])
    return "\n".join(parts) + "\n"


#: The theme change the replay does not make: a new workbook given the older theme and saved.
BASE = f'''Public Function ThemedBase(target As String) As String
Dim wb As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
wb.ApplyTheme {THEME}
wb.SaveAs Filename:=target & "themed_base.xlsx"
wb.Close False
ThemedBase = "ok"
End Function
'''


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    saved, opened = str(folder) + "\\", str(folder) + "\\"
    shutil.copy(DIMENSIONS / "calibri.xlsx", folder / "calibri_source.xlsx")
    code = module()
    runs: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(BASE, "ThemedBase", args=(saved,), timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        for name, args in (("EveryStyle", ()), ("EveryCell", (saved,)), ("EveryCellBack", (saved,)),
                           ("Mixed", (saved,)), ("Custom", (saved,)), ("Whole", (saved,)), ("Fills", (saved,)),
                           ("Protected", ()), ("Errors", ()), ("Files", (opened, saved)),
                           ("Themed", (opened, saved))):
            result = excel.run_vba(code, name, args=args, timeout=900.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            runs[name] = str(result.value)
        # The Calibri workbook opens from its own copy, which the save then replaces.
        (folder / "calibri.xlsx").write_bytes((folder / "calibri_source.xlsx").read_bytes())
        result = excel.run_vba(code, "Calibri", args=(opened, saved), timeout=300.0)
        assert result.ok, f"Calibri: {result.outcome}: {result.message} {result.error}"
        runs["Calibri"] = str(result.value)
    OUT.mkdir(parents=True, exist_ok=True)
    for path in folder.glob("*.xlsx"):
        if path.name != "calibri_source.xlsx":
            copy_saved(path, OUT / path.name)
    record = {"module": code, "style_reads": STYLE_READS, "cell_reads": CELL_READS, "mixed": MIXED,
              "mixed_reads": MIXED_READS, "names": NAMES, "made": MADE, "fills": FILLS, "files": FILES,
              "calibri": CALIBRI, "errors": ERRORS, "runs": runs}
    (OUT / "cell_styles.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, answer in runs.items():
        print(f"{name}: {len(answer)} characters")


if __name__ == "__main__":
    main()
