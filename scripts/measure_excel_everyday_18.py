"""An eighteenth sweep: Range as everyday macros use it, each group run bound late and typed, read line by line from live Excel.

A line is a statement and an expression read after it, under On Error, and the lines of a group share one sheet in order, as a macro's do.
No window is touched and nothing prints or sends. A line that does not compile is found by halving and left out of that run.

    python scripts/measure_excel_everyday_18.py

writes tests/fixtures/excel_everyday_18.json, keeps each group as it finishes and asks only the rest.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_18.json'
SEPARATOR = '¦'
PRE = '''ws.Range("A1:D1").Value = Array("Name", "Team", "Qty", "When")
ws.Range("A2:D2").Value = Array("Ann", "x", 5, #1/2/2020#)
ws.Range("A3:D3").Value = Array("Bob", "y", 7.5, #2/3/2020#)
ws.Range("A4:D4").Value = Array("Cy", "x", -2, #3/4/2020#)
ws.Range("A5:D5").Value = Array("Di", "", 0, "")
ws.Range("F1").Formula = "=SUM(C2:C5)"
'''
GROUPS: dict[str, list[tuple[str, str]]] = {
    'navigate': [('', line) for line in [
        'ws.Range("A1").Address', 'ws.Range("A1:B2").Address', 'ws.Range("A1,C3").Address', 'ws.Range("A1").Address(False, False)', 'ws.Range("A1").Address(True, False)', 'ws.Range("A1").Address(ReferenceStyle:=xlR1C1)', 'ws.Range("A1").Address(External:=True) <> ""', 'ws.Range("B2").Address(RowAbsolute:=False, ColumnAbsolute:=False, ReferenceStyle:=xlR1C1, RelativeTo:=ws.Range("A1"))',
        'ws.Range("B2").Offset(1, 1).Address', 'ws.Range("B2").Offset(-1).Address', 'ws.Range("B2").Offset(-2).Address', 'ws.Range("B2").Offset(, 1).Address', 'ws.Range("A1:B2").Offset(1, 1).Address', 'ws.Range("B2").Offset(1.5).Address', 'ws.Range("B2").Offset("1").Address', 'ws.Range("B2").Resize(2, 2).Address', 'ws.Range("B2").Resize(0).Address', 'ws.Range("B2").Resize(, 3).Address', 'ws.Range("A1:B2").Resize(1).Address', 'ws.Range("A1,C3").Resize(2).Address',
        'ws.Cells(2, 3).Address', 'ws.Cells(2, "C").Address', 'ws.Cells(1).Address', 'ws.Cells(0, 1).Address', 'ws.Cells.Count', 'ws.Cells(1048577, 1).Address', 'ws.Cells(1, 17385).Address', 'ws.Range("B2:C3").Cells(2, 2).Address', 'ws.Range("B2:C3").Cells(3).Address', 'ws.Range("B2:C3").Cells(9, 9).Address', 'ws.Range("B2:C3")(1, 1).Address', 'ws.Range("B2:C3").Item(2).Address', 'ws.Range("A1:A3").Rows(2).Address', 'ws.Range("A1:C1").Columns(2).Address', 'ws.Rows(2).Address', 'ws.Rows("2:3").Address', 'ws.Columns(2).Address', 'ws.Columns("B:C").Address', 'ws.Columns("B").Address', 'ws.Range("A1:B2").Rows.Count', 'ws.Range("A1:B2").Columns.Count', 'ws.Range("A1:B2").Count', 'ws.Range("A1,C3:D4").Count', 'ws.Range("A1,C3:D4").Areas.Count', 'ws.Range("A1,C3:D4").Areas(2).Address', 'ws.Range("A1,C3:D4").Areas(3).Address', 'ws.Range("A1,C3:D4").Rows.Count', 'ws.Range("A1").Row', 'ws.Range("B2:C3").Row', 'ws.Range("B2:C3").Column', 'ws.Range("A1").EntireRow.Address', 'ws.Range("A1").EntireColumn.Address', 'ws.Range("A1:B2").EntireRow.Address', 'ws.Range("A1").End(xlDown).Address', 'ws.Range("A1").End(xlToRight).Address', 'ws.Range("A5").End(xlUp).Address', 'ws.Range("D5").End(xlToLeft).Address', 'ws.Range("Z9").End(xlUp).Address', 'ws.Range("Z9").End(xlToLeft).Address', 'ws.Range("A1").End(5).Address', 'ws.Range("A1").CurrentRegion.Address', 'ws.Range("F1").CurrentRegion.Address', 'ws.Range("H9").CurrentRegion.Address', 'ws.UsedRange.Address', 'ws.Range("A1:B2").Parent.Name <> ""', 'ws.Range("A1").Worksheet.Name <> ""', 'Application.Union(ws.Range("A1"), ws.Range("A2")).Address', 'Application.Union(ws.Range("A1"), ws.Range("C3")).Address', 'Application.Intersect(ws.Range("A1:B2"), ws.Range("B2:C3")).Address', 'TypeName(Application.Intersect(ws.Range("A1"), ws.Range("C3")))', 'ws.Range("1:1").Address', 'ws.Range("A:A").Address', 'ws.Range("A1:A1").Address', 'ws.Range("B2:A1").Address', 'ws.Range("nope").Address', 'ws.Range("").Address', 'ws.Range(ws.Cells(1, 1), ws.Cells(2, 2)).Address', 'ws.Range("A1", "B2").Address', 'ws.Range("A1", ws.Range("B2")).Address', 'ws.Range("A1:B2 B2:C3").Address', 'ws.Range("A1").Value', 'ws.Range("C2").Value', 'ws.Range("D2").Value', 'TypeName(ws.Range("D2").Value)', 'ws.Range("D2").Value2', 'TypeName(ws.Range("D2").Value2)', 'ws.Range("D2").Text', 'ws.Range("C3").Text', 'TypeName(ws.Range("A1:B2").Value)', 'UBound(ws.Range("A1:B2").Value, 2)', 'ws.Range("A1:B2").Value(1, 2)', 'ws.Range("F1").Formula', 'ws.Range("F1").FormulaR1C1', 'ws.Range("F1").HasFormula', 'ws.Range("A1").HasFormula', 'ws.Range("A1:F1").HasFormula', 'ws.Range("C2").Formula', 'ws.Range("D2").Formula', 'ws.Range("Z9").Value', 'TypeName(ws.Range("Z9").Value)', 'ws.Range("Z9").Text', 'ws.Range("Z9").Formula', 'ws.Range("A1:A2").Text', 'TypeName(ws.Range("A1:A2").Text)', 'ws.Range("A1:B1").Formula(1, 2)',
    ]],
    'writes': [
        ('ws.Range("H1").Value = 5', 'ws.Range("H1").Value'), ('ws.Range("H1").Value = "5"', 'TypeName(ws.Range("H1").Value)'), ('ws.Range("H1").Value = "abc"', 'ws.Range("H1").Value'), ('ws.Range("H1").Value = "=1+1"', 'ws.Range("H1").Value'), ('ws.Range("H1").Value = "1/2/2020"', 'TypeName(ws.Range("H1").Value)'), ('ws.Range("H1").Value = "50%"', 'ws.Range("H1").Value'), ('ws.Range("H1").Value = "$5"', 'ws.Range("H1").Value'), ('ws.Range("H1").Value = True', 'TypeName(ws.Range("H1").Value)'), ('ws.Range("H1").Value = Empty', 'TypeName(ws.Range("H1").Value)'), ('ws.Range("H1").Value = ""', 'TypeName(ws.Range("H1").Value)'), ("ws.Range(\"H1\").Value = \"'5\"", 'ws.Range("H1").Value'), ("ws.Range(\"H1\").Value = \"'5\"", 'ws.Range("H1").PrefixCharacter'), ('ws.Range("H1").Value = #1/2/2020#', 'ws.Range("H1").Value2'), ('ws.Range("H1").Value = #1/2/2020#', 'ws.Range("H1").NumberFormat'), ('ws.Range("H1").Value = CCur(1.5)', 'ws.Range("H1").NumberFormat'), ('ws.Range("H1").NumberFormat = "General"', 'ws.Range("H1").Text'), ('ws.Range("H1").Value2 = #1/2/2020#', 'ws.Range("H1").Value2'), ('ws.Range("H1:H2").Value = 7', 'ws.Range("H2").Value'), ('ws.Range("H1").ClearContents', 'ws.Range("H1").Value'), ('ws.Range("H2").Clear', 'ws.Range("H2").Value'), ('ws.Range("H1").Formula = "=1+1"', 'ws.Range("H1").Value'), ('ws.Range("H1").Formula = "abc"', 'ws.Range("H1").Value'), ('ws.Range("H1").Formula = "=nosuch()"', 'ws.Range("H1").Text'), ('ws.Range("H1").Formula = "=1+"', 'ws.Range("H1").Formula'), ('ws.Range("H1").FormulaR1C1 = "=R1C3"', 'ws.Range("H1").Formula'), ('ws.Range("H1").FormulaR1C1 = "=RC[-1]"', 'ws.Range("H1").Formula'), ('ws.Range("H1:H2").Formula = "=A1"', 'ws.Range("H2").Formula'), ('ws.Range("H1:H2").FormulaR1C1 = "=RC[-1]"', 'ws.Range("H2").Formula'), ('ws.Range("H1:H2").Clear', 'ws.Range("H1").Formula'), ('ws.Range("H1").Formula = 5', 'TypeName(ws.Range("H1").Value)'), ('ws.Range("H1:I1").Value = Array(1, 2)', 'ws.Range("I1").Value'), ('ws.Range("H2:H3").Value = Application.Transpose(Array(3, 4))', 'ws.Range("H3").Value'), ('ws.Range("H1:I3").ClearContents', 'ws.Range("I1").Value'), ('ws.Range("H1").Value = 1: ws.Range("H1").Value = ws.Range("H1").Value + 1', 'ws.Range("H1").Value'), ('ws.Range("H2:H3").Value = ws.Range("A2:A3").Value', 'ws.Range("H3").Value'), ('ws.Range("H1:I2").Value = ws.Range("A2:B3").Value', 'ws.Range("I2").Value'), ('ws.Range("H1:I2").ClearContents', ''), ('ws.Range("H1").NumberFormat = "0.00"', 'ws.Range("H1").NumberFormat'), ('ws.Range("H1").Value = 1.5', 'ws.Range("H1").Text'), ('ws.Range("H1").NumberFormat = "0%"', 'ws.Range("H1").Text'), ('ws.Range("H1").NumberFormat = "@"', 'ws.Range("H1").Text'), ('ws.Range("H1").Value = 5', 'TypeName(ws.Range("H1").Value)'), ('ws.Range("H1").NumberFormat = "bad["', 'ws.Range("H1").NumberFormat'), ('ws.Range("H1").Clear', 'ws.Range("H1").NumberFormat'), ('ws.Range("H1:H2").NumberFormat = "0.0"', 'ws.Range("H2").NumberFormat'), ('ws.Range("H3").NumberFormat = "0"', 'ws.Range("H2:H3").NumberFormat'), ('ws.Range("H1:H3").Clear', ''), ('ws.Range("H1").Font.Bold = True', 'ws.Range("H1").Font.Bold'), ('ws.Range("H2").Font.Bold = False', 'ws.Range("H1:H2").Font.Bold'), ('ws.Range("H1").Interior.Color = vbRed', 'ws.Range("H1").Interior.Color'), ('ws.Range("H1").Font.Color = RGB(1, 2, 3)', 'ws.Range("H1").Font.Color'), ('ws.Range("H1").HorizontalAlignment = xlCenter', 'ws.Range("H1").HorizontalAlignment'), ('ws.Range("H1").WrapText = True', 'ws.Range("H1").WrapText'), ('ws.Range("H1").ClearFormats', 'ws.Range("H1").Font.Bold'), ('ws.Range("H1").ColumnWidth = 20', 'ws.Range("H1").ColumnWidth'), ('ws.Range("H1").RowHeight = 30', 'ws.Range("H1").RowHeight'), ('ws.Columns("H").Hidden = True', 'ws.Range("H1").EntireColumn.Hidden'), ('ws.Columns("H").Hidden = False', 'ws.Range("H1").Width > 0'), ('ws.Rows(1).Hidden = True', 'ws.Range("A1").Height'), ('ws.Rows(1).Hidden = False', 'ws.Range("A1").Top'), ('ws.Range("H1").Name = "Spot"', 'ws.Range("Spot").Address'), ('', 'ws.Range("H1").Name.Name'), ('ws.Range("H1").Value = "x"', 'ActiveWorkbook.Names("Spot").RefersToRange.Value'), ('ActiveWorkbook.Names("Spot").Delete', 'ws.Range("Spot").Address'), ('ws.Range("H1").Clear', ''), ('ws.Range("H1").AddComment "n"', 'ws.Range("H1").Comment.Text'), ('ws.Range("H1").AddComment "m"', 'ws.Range("H1").Comment.Text'), ('ws.Range("H1").ClearComments', 'TypeName(ws.Range("H1").Comment)'), ('ws.Range("H1").Locked = False', 'ws.Range("H1").Locked'), ('ws.Range("H1").Locked = True', ''),
    ],
    'structure': [
        ('ws.Range("H1:H2").Merge', 'ws.Range("H1").MergeCells'), ('', 'ws.Range("H2").MergeCells'), ('', 'ws.Range("H2").MergeArea.Address'), ('ws.Range("H2").Value = 5', 'ws.Range("H1").Value'), ('', 'ws.Range("H2").Value'), ('ws.Range("H1").UnMerge', 'ws.Range("H2").MergeCells'), ('ws.Range("H1:I2").Merge True', 'ws.Range("H2").MergeArea.Address'), ('ws.Range("H1:I2").UnMerge', ''), ('ws.Range("H1:I2").Clear', ''), ('ws.Range("A2").Insert', 'ws.Range("A3").Value'), ('ws.Range("A2").Delete', 'ws.Range("A2").Value'), ('ws.Range("A2").Insert xlShiftToRight', 'ws.Range("B2").Value'), ('ws.Range("A2").Delete xlShiftToLeft', 'ws.Range("A2").Value'), ('ws.Rows(2).Insert', 'ws.Range("A3").Value'), ('', 'ws.Range("F1").Formula'), ('ws.Rows(2).Delete', 'ws.Range("A2").Value'), ('', 'ws.Range("F1").Formula'), ('ws.Columns("B").Insert', 'ws.Range("C1").Value'), ('', 'ws.Range("G1").Formula'), ('ws.Columns("B").Delete', 'ws.Range("B1").Value'), ('ws.Range("A2:B2").Insert Shift:=xlDown', 'ws.Range("A3").Value'), ('ws.Range("A2:B2").Delete Shift:=xlUp', 'ws.Range("A2").Value'), ('ws.Range("A2").EntireRow.Insert', 'ws.Range("A3").Value'), ('ws.Range("A2").EntireRow.Delete', 'ws.Range("A2").Value'), ('ws.Range("C2:C3").Copy ws.Range("H1")', 'ws.Range("H2").Value'), ('ws.Range("C2:C3").Cut ws.Range("I1")', 'ws.Range("C2").Value'), ('', 'ws.Range("I2").Value'), ('', 'ws.Range("F1").Formula'), ('ws.Range("I1:I2").Cut ws.Range("C2")', 'ws.Range("F1").Formula'), ('ws.Range("H1:H2").ClearContents', ''), ('ws.Range("F1").Copy ws.Range("H1")', 'ws.Range("H1").Formula'), ('ws.Range("H1").Clear', ''), ('ws.Range("A2:A3").Copy', 'Application.CutCopyMode'), ('ws.Range("H1").PasteSpecial xlPasteValues', 'ws.Range("H2").Value'), ('', 'Application.CutCopyMode'), ('Application.CutCopyMode = False', 'Application.CutCopyMode'), ('ws.Range("H1:H2").Clear', ''), ('ws.Range("C2:C3").FillDown', 'ws.Range("C3").Value'), ('ws.Range("C3").Value = 7.5', ''), ('ws.Range("H1").Value = 1: ws.Range("H1").AutoFill ws.Range("H1:H3")', 'ws.Range("H3").Value'), ('ws.Range("H1").Value = 1: ws.Range("H2").Value = 2: ws.Range("H1:H2").AutoFill ws.Range("H1:H4")', 'ws.Range("H4").Value'), ('ws.Range("H1:H4").Clear', ''), ('ws.Range("C2:C5").Replace "5", "6"', 'ws.Range("C2").Value'), ('ws.Range("C2").Value = 5', ''), ('ws.Range("A1:D5").Columns(2).Hidden = True', 'ws.Range("B1").EntireColumn.Hidden'), ('ws.Range("A1:D5").Columns(2).Hidden = False', ''), ('ws.Range("A1:A5").TextToColumns ws.Range("H1")', 'ws.Range("H2").Value'), ('ws.Range("H1:H5").Clear', ''), ('ws.Range("A1:D5").RemoveDuplicates 2, xlYes', 'ws.Range("A3").Value'), ('', 'ws.UsedRange.Address'),
    ],
    'reads': [('', line) for line in [
        'Application.WorksheetFunction.Sum(ws.Range("C2:C5"))', 'Application.Sum(ws.Range("C2:C5"))', 'WorksheetFunction.CountA(ws.Range("A1:A9"))', 'WorksheetFunction.CountIf(ws.Range("B2:B5"), "x")', 'WorksheetFunction.VLookup("Bob", ws.Range("A2:C5"), 3, False)', 'WorksheetFunction.VLookup("Zed", ws.Range("A2:C5"), 3, False)', 'Application.VLookup("Zed", ws.Range("A2:C5"), 3, False)', 'TypeName(Application.VLookup("Zed", ws.Range("A2:C5"), 3, False))', 'WorksheetFunction.Match("Bob", ws.Range("A2:A5"), 0)', 'WorksheetFunction.Max(ws.Range("C2:C5"))', 'ws.Range("C2:C5").Find(7.5).Address', 'ws.Range("A2:A5").Find("Bob").Row', 'TypeName(ws.Range("A2:A5").Find("Zed"))', 'ws.Range("A2:A5").Find("b", LookAt:=xlPart, MatchCase:=False).Address', 'ws.Range("A2:A5").Find("b", LookAt:=xlPart, MatchCase:=True) Is Nothing', 'ws.Range("A:A").Find("*", SearchDirection:=xlPrevious, LookIn:=xlValues, LookAt:=xlPart).Address', 'ws.Cells(ws.Rows.Count, 1).End(xlUp).Row', 'ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column', 'ws.Range("A1:D5").SpecialCells(xlCellTypeConstants).Count', 'ws.Range("A1:D5").SpecialCells(xlCellTypeBlanks).Address', 'ws.Range("A1:F5").SpecialCells(xlCellTypeFormulas).Address', 'ws.Range("A1:D5").SpecialCells(xlCellTypeFormulas).Address', 'ws.Range("A1:D5").SpecialCells(xlCellTypeConstants, xlNumbers).Address', 'ws.Range("A1:D5").SpecialCells(xlCellTypeConstants, xlTextValues).Count', 'ws.Range("A1").Font.Name <> ""', 'ws.Range("A1").Font.Size > 0', 'ws.Range("A1").Font.Bold', 'ws.Range("A1").Interior.ColorIndex', 'ws.Range("A1").Interior.Pattern', 'ws.Range("A1").Borders(xlEdgeTop).LineStyle', 'ws.Range("A1").HorizontalAlignment', 'ws.Range("A1").VerticalAlignment', 'ws.Range("A1").NumberFormat', 'ws.Range("D2").NumberFormat', 'ws.Range("A1:D1").NumberFormat', 'ws.Range("A1:D2").NumberFormat', 'TypeName(ws.Range("A1:D2").NumberFormat)', 'ws.Range("A1").WrapText', 'ws.Range("A1").Locked', 'ws.Range("A1").FormulaHidden', 'ws.Range("A1").Hidden', 'ws.Rows(1).Hidden', 'ws.Range("A1").IndentLevel', 'ws.Range("A1").Orientation', 'ws.Range("A1").ShrinkToFit', 'ws.Range("A1").Style.Name', 'ws.Range("A1").MergeCells', 'ws.Range("A1").HasArray', 'ws.Range("A1").ListObject Is Nothing', 'ws.Range("A1").Hyperlinks.Count', 'TypeName(ws.Range("A1").Comment)', 'ws.Range("A1").Validation.Type', 'ws.Range("A1").FormatConditions.Count', 'ws.Range("A1").PageBreak', 'ws.Range("A1").OutlineLevel', 'ws.Range("A1").Characters(1, 2).Text', 'ws.Range("A1").Characters.Count', 'ws.Range("C2").Characters.Text', 'ws.Range("A1").Precedents.Address', 'ws.Range("F1").Precedents.Address', 'ws.Range("C2").Dependents.Address', 'ws.Range("A1").Dependents.Address', 'ws.Range("F1").DirectPrecedents.Address', 'ws.Range("A1").Parent Is ws', 'ws.Range("A1") Is ws.Range("A1")', 'TypeName(ws.Range("A1"))', 'TypeName(ws.Cells)', 'TypeName(ws.Rows)', 'TypeName(ws.Range("A1").Value2)', 'ws.Range("A1") = "Name"', 'ws.Range("C2") + 1', 'ws.Range("A1") & "!"', 'CStr(ws.Range("C2"))', 'Len(ws.Range("A1"))', 'IsEmpty(ws.Range("Z9"))', 'IsEmpty(ws.Range("A1"))', 'IsNumeric(ws.Range("C2"))', 'IsNumeric(ws.Range("A1"))', 'IsDate(ws.Range("D2"))', 'IsError(ws.Range("A1"))', 'IsNull(ws.Range("A1:A2").Font.Size)', 'IsArray(ws.Range("A1:A2").Value)', 'IsArray(ws.Range("A1").Value)', 'WorksheetFunction.Sum(ws.Range("A1:A2"))',
    ]],
    'text': [
        ('ws.Range("A1:A2").Value = "same"', 'ws.Range("A1:A2").Text'), ('ws.Range("B1").Value = 5', 'ws.Range("A1:A2").Text'),
        ('ws.Range("B2").Value = "5"', 'ws.Range("A1:A2").Text'), ('ws.Range("C1").Value = 5', 'ws.Range("A1:A2").Text'),
        ('ws.Range("C2").Value = 5', 'ws.Range("A1:A2").Text'), ('ws.Range("C2").NumberFormat = "0.00"', 'ws.Range("A1:A2").Text'),
        ('ws.Range("D1:E1").Value = Array("Name", "Team")', 'ws.Range("A1:A2").Text'), ('ws.Range("F1").Value = "same"', 'ws.Range("A1:A2").Text'),
        ('ws.Range("H1:H2").Merge', 'ws.Range("A1:A2").Text'), ('ws.Range("H1").Value = "m"', 'ws.Range("A1:A2").Text'),
        ('ws.Range("I1").Value = 1', 'ws.Range("A1:A2").Text'), ('ws.Range("I2").Value = 1', 'ws.Range("A1:A2").Text'),
        ('ws.Rows(9).Hidden = True', 'ws.Range("A1:A2").Text'), ('ws.Range("J8:J9").Value = "same"', 'ws.Range("A1:A2").Text'),
        ('ws.Range("J10").Value = "other"', 'ws.Range("A1:A2").Text'),
        ('', 'ws.Range("A1:A2").Text'), ('', 'ws.Range("B1:B2").Text'), ('', 'ws.Range("C1:C2").Text'), ('', 'ws.Range("Z1:Z9").Text'),
        ('', 'ws.Range("A1,F1").Text'), ('', 'ws.Range("A1,D1").Text'), ('', 'ws.Range("H1:H2").Text'), ('', 'ws.Range("I1:I2").Text'),
        ('', 'ws.Range("A1:A3").Text'), ('', 'ws.Range("J8:J9").Text'), ('', 'ws.Rows(30).Text'), ('', 'ws.Columns(30).Text'), ('', 'ws.Cells.Text'),
        ('', 'TypeName(ws.Range("A1:A3").Text)'), ('', 'IsNull(ws.Range("A1:A3").Text)'),
    ],
    'walks': [('', line) for line in [
        'ws.Range("A1,C3").Resize().Address', 'ws.Range("A1,C3").Resize(1).Address', 'ws.Range("A1,C3").Resize(1, 1).Address',
        'ws.Range("A1,C3").Resize(, 1).Address', 'ws.Range("A1,A1").Resize(1).Address', 'ws.Range("A1").Resize().Address',
        'ws.Range("A1,C3").Offset(1).Address', 'ws.Range("A1,C3").Cells(1, 1).Address', 'ws.Range("A1,C3").Item(2).Address',
        'ws.Cells(0, 1).Address', 'ws.Cells(1, 0).Address', 'ws.Cells(-1, 1).Address', 'ws.Cells(0).Address', 'ws.Cells(-1).Address',
        'ws.Cells(17179869185#).Address', 'ws.Cells(1, "XFE").Address', 'ws.Cells(1, "").Address', 'ws.Range("B2").Item(0, 1).Address',
        'ws.Range("B2").Item(-1, 1).Address', 'ws.Range("B2").Item(0).Address', 'ws.Range("B2").Item(-1).Address',
        'ws.Range("A1").Item(0, 1).Address', 'ws.Range("A1").Item(1, 0).Address', 'ws.Range("A1").Item(0).Address',
        'ws.Range("A1").Item(-1).Address', 'ws.Range("A1").Cells(0, 1).Address', 'ws.Range("A1").Cells(0).Address', 'ws.Range("B2")(0, 1).Address',
        'ws.Range("A1")(0, 1).Address', 'ws.Range("XFD1").Item(1, 2).Address', 'ws.Range("A1048576").Item(2, 1).Address',
        'ws.Range("XFD1048576").Item(2).Address', 'ws.Cells(1, "2").Address', 'ws.Cells(1, "a").Address', 'ws.Cells(1, "$A").Address',
        'ws.Cells(1, "A1").Address', 'ws.Cells(1, " A").Address', 'ws.Cells(1, "XFD").Address', 'ws.Cells(1, "2.5").Address',
        'ws.Cells(1, "-1").Address', 'ws.Cells(1, "0").Address', 'ws.Cells("2", 1).Address', 'ws.Cells("x", 1).Address', 'ws.Cells("", 1).Address',
        'ws.Cells("2").Address', 'ws.Cells("x").Address', 'ws.Cells("A1").Address', 'ws.Cells(1, 2.5).Address', 'ws.Cells(2.5, 1).Address',
        'ws.Cells(1.5).Address', 'ws.Cells(2.5).Address', 'ws.Cells(0.4, 1).Address', 'ws.Cells(0.6, 1).Address', 'ws.Cells(True, 1).Address',
        'ws.Cells(1, True).Address', 'ws.Cells(False, 1).Address', 'ws.Cells(Empty, 1).Address', 'ws.Cells(1, Empty).Address',
        'ws.Cells(Null, 1).Address', 'ws.Cells(1, Null).Address', 'ws.Cells(#1/2/1900#, 1).Address', 'ws.Cells(1, Array(1)).Address',
        'ws.Cells(ws.Range("A1"), 1).Address', 'ws.Cells(2147483648#, 1).Address', 'ws.Cells(1, 2147483648#).Address',
        'ws.Cells(17179869184#).Address', 'ws.Range("A2:B3").Item(0).Address', 'ws.Range("B3:C4").Item(-1).Address',
        'ws.Range("B3:D4").Item(-1).Address', 'ws.Range("A1:B2").Item(5).Address', 'ws.Range("A1:B2").Item(3, 3).Address',
        'ws.Range("B2:C3").Item(0, 0).Address', 'ws.Range("B2").Item(1, "B").Address', 'ws.Range("B2").Item(1, "A").Address',
        'ws.Rows(2).Item(0).Address', 'ws.Rows(1).Item(0).Address', 'ws.Rows(2).Item(-1).Address', 'ws.Columns(2).Item(0).Address',
        'ws.Columns(1).Item(0).Address', 'ws.Rows(1048576).Item(2).Address', 'ws.Columns(16384).Item(2).Address', 'ws.Rows("2:3").Item(0).Address',
        'ws.Rows(2).Cells(0).Address', 'ws.Rows(2).Cells(0, 1).Address', 'ws.Rows(0).Address', 'ws.Rows(-1).Address', 'ws.Rows(1048577).Address',
        'ws.Columns(0).Address', 'ws.Columns(16385).Address', 'ws.Columns("XFE").Address', 'ws.Rows("0").Address', 'ws.Rows("").Address',
        'ws.Columns("").Address', 'ws.Range("A1").Offset(-1).Address', 'ws.Range("A1").Offset(, -1).Address',
        'ws.Range("A1").Offset(1048576).Address', 'ws.Range("A1,C3").Rows(1).Address', 'ws.Range("A1,C3").Rows.Count',
        'ws.Range("A1,C3").Columns.Count', 'ws.Range("A1,C3").Cells.Count', 'ws.Range("A1,C3").Count', 'ws.Range("A1,C3").Row',
        'ws.Range("A1,C3").Column', 'ws.Range("A1,C3").EntireRow.Address', 'ws.Range("A1,C3").EntireColumn.Address',
    ]],
    'arguments': [('', line) for line in [
        'ws.Range("A1:B2").Value(1, 2)', 'ws.Range("A1:B2").Value(1)', 'ws.Range("D1:E1").Value2(1, 2)', 'ws.Range("D1:E1").Value2(1)',
        'ws.Range("D1:E1").Formula(1, 2)', 'ws.Range("D1:E1").Formula(1)', 'ws.Range("D1:E1").FormulaR1C1(1, 2)', 'ws.Range("D1:E1").Formula2(1, 2)',
        'ws.Range("D1").Formula(1)', 'ws.Range("D1").Text(1)', 'ws.Range("D1").Row(1)', 'ws.Range("D1").Count(1)', 'ws.Range("D1").HasFormula(1)',
        'ws.Range("D1").NumberFormat(1)', 'ws.Range("D1").ColumnWidth(1)', 'ws.Range("D1").Font(1).Bold', 'ws.Range("D1").Font.Bold(1)',
        'ws.Range("D1").Interior(1).Color', 'ws.Range("D1").Worksheet(1).Name', 'ws.Range("D1").Parent(1).Name', 'ws.Range("D1").Address(1)',
        'ws.Range("D1").EntireRow(1).Address', 'ws.Range("D1").CurrentRegion(1).Address', 'ws.Name(1)', 'ws.UsedRange(1).Address', 'ws.Index(1)',
        'ws.Parent.Name(1)', 'ws.Parent.Sheets.Count(1)', 'ws.Range("A1:B2").Value(10)', 'ws.Range("A1:B2").Value(xlRangeValueDefault)', 'ws.Range("A1:B2").Value(0)', 'ws.Range("A1:B2").Value(-1)', 'ws.Range("A1:B2").Value(9)',
        'ws.Range("A1:B2").Value(13)', 'ws.Range("A1:B2").Value(10.4)', 'ws.Range("A1:B2").Value("10")', 'ws.Range("A1:B2").Value("x")', 'ws.Range("A1:B2").Value(True)', 'ws.Range("A1:B2").Value(Null)', 'ws.Range("A1:B2").Value(Empty)', 'ws.Range("A1:B2").Value(Array(10))',
        'ws.Range("A1:B2").Value(r)', 'TypeName(ws.Range("A1").Value(10))', 'ws.Range("A1").Value(1)', 'ws.Range("A1").Value(RangeValueDataType:=10)',
        'ws.Range("A1:B2").Value2(RangeValueDataType:=10)', 'ws.Range("A1").Value(1, 1)',
    ]],
    'value_sets': [
        ('ws.Range("D1").Value(10) = 7', 'ws.Range("D1").Value'), ('ws.Range("D2").Value(1) = 7', 'ws.Range("D2").Value'),
        ('ws.Range("D3").Value("x") = 7', 'ws.Range("D3").Value'), ('ws.Range("D4").Value(Null) = 7', 'ws.Range("D4").Value'),
        ('ws.Range("D5").Value() = 7', 'ws.Range("D5").Value'),
    ],
}
#: Typed, these do not compile, each checked once on its own: finding them by halving takes longer than a run may.
REFUSED_TYPED = frozenset({
    'ws.Range("A1:B2").Value(1, 2)', 'ws.Range("D1").Row(1)', 'ws.Range("D1").Count(1)', 'ws.Name(1)', 'ws.Index(1)', 'ws.Parent.Name(1)',
    'ws.Parent.Sheets.Count(1)', 'ws.Range("A1:B2").Value2(RangeValueDataType:=10)', 'ws.Range("A1").Value(1, 1)',
})
HELPERS = f'''Private Function Shown(ByVal x As Variant) As String
    On Error Resume Next
    Shown = CStr(x)
    If Err.Number Then Shown = "?" & Err.Number: Err.Clear
End Function
'''


def line_of(statement: str, expression: str) -> str:
    run = f'{statement}\ns = s & Err.Number & "!"\nErr.Clear\n' if statement else 's = s & "0!"\n'
    read = f'v = Empty\nv = {expression}\ne = Err.Number\nErr.Clear\ns = s & e & "=" & TypeName(v) & ":" & Shown(v)\n' if expression else ''
    return run + read + f's = s & "{SEPARATOR}"\n'


def macro(name: str, *, typed: bool, leaving_out: frozenset[int] = frozenset()) -> str:
    body = ''.join(f's = s & "skipped{SEPARATOR}"\n' if at in leaving_out else line_of(*pair) for at, pair in enumerate(GROUPS[name]))
    declared = 'Worksheet' if typed else 'Object'
    return (HELPERS + f'Function Probe() As String\nDim ws As {declared}, s As String, v As Variant, e As Long\nSet ws = Worksheets.Add\n' + PRE
            + 'On Error Resume Next\n' + body + 'On Error GoTo 0\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True\nProbe = s\nEnd Function\n')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    def run(source: str) -> object:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', timeout=180.0, line_numbers=False)
        if result.dialogs or not result.ok:
            kept = [dialog.message[:70] for dialog in result.dialogs if dialog.classification != 'compile-error']
            if kept or not result.dialogs:
                print('  stopped by', kept, str(result.message)[:70], flush=True)
            return None
        return result.value

    def answered(name: str, typed: bool) -> tuple[object, list[int]]:
        count = len(GROUPS[name])
        bad: set[int] = set()
        known = {at for at, pair in enumerate(GROUPS[name]) if typed and pair[1] in REFUSED_TYPED
                 and run(macro(name, typed=typed, leaving_out=frozenset(range(count)) - {at})) is None}
        answer = run(macro(name, typed=typed, leaving_out=frozenset(known)))
        if answer is not None:
            return answer, sorted(known)

        def search(lines: list[int]) -> None:
            if run(macro(name, typed=typed, leaving_out=frozenset(range(count)) - frozenset(lines))) is not None:
                return
            if len(lines) == 1:
                bad.add(lines[0])
                return
            middle = len(lines) // 2
            search(lines[:middle])
            search(lines[middle:])

        search(list(range(count)))
        return run(macro(name, typed=typed, leaving_out=frozenset(bad))), sorted(bad)

    probes: list[dict[str, object]] = json.loads(OUT.read_text(encoding='utf-8'))['probes'] if OUT.exists() else []
    done = {(probe['name'], probe['binding']) for probe in probes}
    for name in GROUPS:
        for binding in ('late', 'typed'):
            if (name, binding) in done:
                continue
            answer, bad = answered(name, binding == 'typed')
            probes.append({'name': name, 'binding': binding, 'asked': [f'{one} / {two}' for one, two in GROUPS[name]], 'macro': macro(name, typed=binding == 'typed', leaving_out=frozenset(bad)), 'answer': answer, 'not_compiling': bad})
            OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=1, ensure_ascii=True) + '\n', encoding='utf-8')
            print(name, binding, len(GROUPS[name]), 'does not compile:', [GROUPS[name][at] for at in bad], flush=True)


if __name__ == '__main__':
    main()
