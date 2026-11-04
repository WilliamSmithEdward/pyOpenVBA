"""A sixteenth sweep of everyday Excel macros, on what macros use most: loops over ranges, Find, AutoFilter, Sort, Copy and PasteSpecial, dates, number
formats, strings, and opening and saving workbooks. Each line of a group is read with its type and value or its error, so a difference names one thing.

    python scripts/measure_excel_everyday_16.py

writes tests/fixtures/excel_everyday_16.json, which tests/test_excel_everyday_16.py replays line by line.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_16.json'
PRE = ('Dim ws As Worksheet, s As String, v As Variant, r As Range, c As Range, i As Long, n As Long, o As Object, t As String, a As Variant, f As String\n'
       'Set ws = Worksheets.Add\nws.Name = "Sweep"\nws.Activate\n'
       'ws.Range("A1:D1").Value = Array("Name", "Team", "Qty", "When")\n'
       'ws.Range("A2:C2").Value = Array("ann", "North", 10): ws.Range("D2").Value = #1/15/2020#\n'
       'ws.Range("A3:C3").Value = Array("Bob", "South", 3): ws.Range("D3").Value = #3/1/2020#\n'
       'ws.Range("A4:C4").Value = Array("cy", "North", 7): ws.Range("D4").Value = #2/10/2020#\n'
       'ws.Range("A5:C5").Value = Array("Di", "East", 20): ws.Range("D5").Value = #12/31/2019#\n'
       'ws.Range("A6:C6").Value = Array("Ed", "South", 1): ws.Range("D6").Value = #7/4/2020#\n'
       'ws.Range("A7:C7").Value = Array("ann", "West", 7): ws.Range("D7").Value = #1/1/2021#\n'
       'ws.Columns("A:Z").ColumnWidth = 30\n')
#: Each group's lines share one sheet, in order.
GROUPS: dict[str, list[str]] = {
    'loops': [
        'ws.UsedRange.Address', 'ws.UsedRange.Rows.Count', 'ws.UsedRange.Columns.Count', 'ws.Cells(ws.Rows.Count, 1).End(xlUp).Row', 'ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column', 'ws.Range("A1").End(xlDown).Address', 'ws.Range("A1").End(xlToRight).Address',
        'ws.Range("A1").CurrentRegion.Address', 'ws.Range("B2").CurrentRegion.Rows.Count', 'ws.Range("A2").Offset(1, 1).Address', 'ws.Range("A2").Offset(-5, 0).Address', 'ws.Range("A2").Resize(2, 3).Address', 'ws.Range("A2:B3").Resize(, 1).Address', 'ws.Range("A1:D7").Rows(2).Address', 'ws.Range("A1:D7").Columns(2).Address', 'ws.Range("A1:D7").Cells(2, 2).Address', 'ws.Range("A1:D7").Cells(9).Address', 'ws.Range("B2:C3")(1, 1).Address', 'ws.Range("B2:C3")(3).Address', 'ws.Range("B2:C3").Item(2, 2).Address',
        'ws.Range("A1:B2").Count', 'ws.Range("A:A").Count', 'ws.Rows(1).Count', 'ws.Range("A1:B2").CountLarge', 'ws.Range("A1").Row', 'ws.Range("B3:C4").Column', 'ws.Range("A1").Address(False, False)', 'ws.Range("A1").Address(True, False)', 'ws.Range("A1").Address(, , xlR1C1)', 'ws.Range("B2").Address(, , xlR1C1, , ws.Range("A1"))', 'ws.Range("A1").Address(External:=True) <> ""', 'ws.Range("A1,C3").Address', 'ws.Range("A1,C3").Areas.Count', 'ws.Range("A1,C3").Count', 'ws.Range("A1,C3").Areas(2).Address',
        'Union(ws.Range("A1"), ws.Range("A2")).Address', 'Union(ws.Range("A1"), ws.Range("C3")).Address', 'Intersect(ws.Range("A1:B2"), ws.Range("B2:C3")).Address', 'Intersect(ws.Range("A1"), ws.Range("B2")) Is Nothing', 'Application.Union(ws.Range("A1:A2"), ws.Range("A2:A3")).Address', 'ws.Range("A1:D7").SpecialCells(xlCellTypeConstants).Address', 'ws.Range("A1:D7").SpecialCells(xlCellTypeConstants, xlNumbers).Count', 'ws.Range("A1:D7").SpecialCells(xlCellTypeBlanks).Count', 'ws.Range("F1:F2").SpecialCells(xlCellTypeBlanks).Count', 'ws.Cells.SpecialCells(xlCellTypeLastCell).Address',
        'Total(ws.Range("C2:C7"))', 'Names(ws.Range("A2:A7"))', 'RowsOf(ws.Range("A2:B3"))', 'ColumnsOf(ws.Range("A2:B3"))', 'TypeName(ws.Range("A1").Value)', 'TypeName(ws.Range("C2").Value)', 'TypeName(ws.Range("D2").Value)', 'TypeName(ws.Range("D2").Value2)', 'TypeName(ws.Range("F1").Value)', 'TypeName(ws.Range("A1:B2").Value)', 'UBound(ws.Range("A1:B2").Value, 1)', 'LBound(ws.Range("A1:B2").Value, 2)', 'ws.Range("A1:B2").Value(2, 1)', 'TypeName(ws.Range("A1:A1").Value)', 'TypeName(ws.Range("A1,C3").Value)', 'ws.Range("A1").Value = ws.Range("A1").Value', 'ws.Range("C2").Formula', 'ws.Range("C2").HasFormula', 'ws.Range("A1:C1").HasFormula', 'ws.Range("D2").Formula', 'ws.Range("D2").FormulaR1C1', 'ws.Range("A1").FormulaLocal', 'ws.Range("A1").Text', 'ws.Range("C2").Text', 'ws.Range("D2").Text', 'ws.Range("D2").NumberFormat', 'ws.Range("A1:D1").NumberFormat', 'ws.Range("C2:D2").NumberFormat', 'IsNull(ws.Range("C2:D2").NumberFormat)',
    ],
    'find': [
        'ws.Cells.Find("ann", LookIn:=xlValues, LookAt:=xlWhole).Address', 'ws.Cells.Find("ANN", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find("ANN", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True) Is Nothing', 'ws.Cells.Find("nor", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address', 'ws.Cells.Find("nor", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) Is Nothing', 'ws.Cells.Find("z*", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) Is Nothing', 'ws.Cells.Find("a?n", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find("~*", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing', 'ws.Cells.Find(7, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find("7", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find(7, After:=ws.Range("C4"), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find(7, After:=ws.Range("C7"), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find(7, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False).Address', 'ws.Cells.Find("ann", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False).Address', 'ws.Cells.Find("North", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False).Address', 'ws.Range("B2:B7").Find("South", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Range("B3").Find("South", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Range("A1:A2").Find("South", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) Is Nothing', 'ws.Cells.Find(#3/1/2020#, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) Is Nothing', 'ws.Cells.Find("3/1/2020", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'ws.Cells.Find("", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False).Address', 'FindAll("North")', 'FindAll("ann")', 'FindAll("zzz")', 'FindAll(7)',
        'ws.Range("A1:C7").Replace("North", "N", xlWhole, MatchCase:=False)', 'ws.Range("B2").Value', 'ws.Range("A1:C7").Replace("zzz", "N", xlWhole, MatchCase:=False)', 'ws.Range("A1:C7").Replace("o", "0", xlPart, MatchCase:=False)', 'ws.Range("B3").Value', 'ws.Range("A3").Value', 'ws.Range("A1:C7").Replace("ANN", "x", xlWhole, MatchCase:=True)', 'ws.Range("A2").Value',
        'WorksheetFunction.CountIf(ws.Range("A1:A7"), "x")', 'ws.Range("A2").Find("x", LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False).Address',
    ],
    'autofilter': [
        'ws.AutoFilterMode', 'ws.FilterMode', 'ws.AutoFilter Is Nothing', 'Do1(ws.Range("A1:D7").AutoFilter(2, "North"))', 'ws.AutoFilterMode', 'ws.FilterMode', 'ws.AutoFilter.Range.Address', 'ws.AutoFilter.Filters.Count', 'ws.AutoFilter.Filters(2).On', 'ws.AutoFilter.Filters(1).On', 'ws.AutoFilter.Filters(2).Criteria1', 'ws.AutoFilter.Filters(2).Operator', 'ws.Rows(3).Hidden', 'ws.Rows(2).Hidden', 'ws.Rows(8).Hidden', 'ws.Range("A1:D7").SpecialCells(xlCellTypeVisible).Address', 'ws.Range("A1:D7").SpecialCells(xlCellTypeVisible).Count', 'ws.Range("A1:D7").SpecialCells(xlCellTypeVisible).Areas.Count', 'ws.Range("A1:A7").SpecialCells(xlCellTypeVisible).Cells.Count', 'WorksheetFunction.Subtotal(103, ws.Range("A2:A7"))', 'WorksheetFunction.Subtotal(9, ws.Range("C2:C7"))', 'WorksheetFunction.Sum(ws.Range("C2:C7"))', 'ws.Cells(ws.Rows.Count, 1).End(xlUp).Row', 'Visible(ws.Range("A2:A7"))',
        'Do1(ws.Range("A1:D7").AutoFilter(3, ">5"))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(2))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.ShowAllData())', 'Visible(ws.Range("A2:A7"))', 'ws.FilterMode',
        'Do1(ws.Range("A1:D7").AutoFilter(3, ">5", xlAnd, "<10"))', 'Visible(ws.Range("A2:A7"))', 'ws.AutoFilter.Filters(3).Criteria2', 'ws.AutoFilter.Filters(3).Operator', 'Do1(ws.Range("A1:D7").AutoFilter(3))', 'Do1(ws.Range("A1:D7").AutoFilter(2, "North", xlOr, "East"))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(2))', 'Do1(ws.Range("A1:D7").AutoFilter(2, Array("North", "West"), xlFilterValues))', 'Visible(ws.Range("A2:A7"))', 'TypeName(ws.AutoFilter.Filters(2).Criteria1)', 'ws.AutoFilter.Filters(2).Operator', 'Do1(ws.Range("A1:D7").AutoFilter(2))', 'Do1(ws.Range("A1:D7").AutoFilter(3, "2", xlTop10Items))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(3, "50", xlTop10Percent))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(3, "1", xlBottom10Items))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(3))', 'Do1(ws.Range("A1:D7").AutoFilter(1, "a*"))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(1, "<>ann"))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(1, "="))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(1, "<>"))', 'Visible(ws.Range("A2:A7"))', 'Do1(ws.Range("A1:D7").AutoFilter(1))', 'Do1(ws.Range("A1:D7").AutoFilter(9, "x"))', 'Do1(ws.Range("A1:D7").AutoFilter(0, "x"))', 'Do1(ws.Range("A1:D7").AutoFilter(1, "zzz"))', 'Visible(ws.Range("A2:A7"))', 'ws.Range("A1:D7").SpecialCells(xlCellTypeVisible).Address', 'Do1(ws.Range("A1:D7").AutoFilter())', 'ws.AutoFilterMode', 'ws.FilterMode',
    ],
    'sort': [
        'Do1(ws.Range("A1:D7").Sort(ws.Range("C1"), xlAscending, Header:=xlYes))', 'Col(3)', 'Col(1)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("C1"), xlDescending, Header:=xlYes))', 'Col(3)', 'Col(1)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("A1"), xlAscending, Header:=xlYes))', 'Col(1)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("A1"), xlAscending, Header:=xlYes, MatchCase:=True))', 'Col(1)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("D1"), xlAscending, Header:=xlYes))', 'Col(1)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("B1"), xlAscending, ws.Range("C1"), , xlDescending, Header:=xlYes))', 'Col(2)', 'Col(3)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("C1")))', 'Col(3)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("C1"), Header:=xlNo))', 'Col(3)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("C1"), Header:=xlGuess))', 'Col(3)', 'Do1(ws.Range("A2:D7").Sort(ws.Range("A2")))', 'Col(1)', 'Do1(ws.Range("A1:D7").Sort(ws.Range("A1"), Orientation:=xlLeftToRight, Header:=xlNo))', 'Row1()', 'Do1(ws.Range("A1:D7").Sort(ws.Range("A1"), Orientation:=xlTopToBottom, Header:=xlYes))', 'Row1()', 'Do1(ws.Range("A1:D7").Sort("Qty", Header:=xlYes))', 'Do1(ws.Range("A1:D7").Sort(ws.Range("Z1"), Header:=xlYes))', 'Do1(ws.Range("A1:D7").Sort(Key1:=ws.Range("A1:A2"), Header:=xlYes))',
        'ws.Sort.SortFields.Count', 'Do1(ws.Sort.SortFields.Clear())', 'ws.Sort.SortFields.Add(ws.Range("B2:B7"), xlSortOnValues, xlAscending).Order', 'ws.Sort.SortFields.Add(ws.Range("C2:C7"), xlSortOnValues, xlDescending).Key.Address', 'ws.Sort.SortFields.Count', 'Do1(ws.Sort.SetRange(ws.Range("A1:D7")))', 'ws.Sort.Rng.Address', 'ws.Sort.Header', 'ws.Sort.Header', 'ws.Sort.MatchCase', 'ws.Sort.Orientation', 'ws.Sort.SortMethod', 'Do1(ws.Sort.Apply())', 'Col(2)', 'Col(3)', 'ws.Sort.SortFields.Count',
        'Do1(FillMixed())', 'Do1(ws.Range("H1:H9").Sort(ws.Range("H1"), xlAscending, Header:=xlNo))', 'ColH()', 'Do1(ws.Range("H1:H9").Sort(ws.Range("H1"), xlDescending, Header:=xlNo))', 'ColH()',
    ],
    'paste': [],
    'dates': [
        'DateSerial(2020, 2, 30)', 'DateSerial(2020, 13, 1)', 'DateSerial(99, 1, 1)', 'DateSerial(29, 1, 1)', 'DateSerial(30, 1, 1)', 'DateSerial(0, 1, 1)', 'DateSerial(2020, 0, 0)', 'TimeSerial(25, 0, 0)', 'TimeSerial(1, 61, 0)', 'TimeSerial(0, 0, -1)', 'DateAdd("d", 1, #1/31/2020#)', 'DateAdd("m", 1, #1/31/2020#)', 'DateAdd("yyyy", 1, #2/29/2020#)', 'DateAdd("q", 1, #11/30/2020#)', 'DateAdd("w", 1, #1/31/2020#)', 'DateAdd("ww", 1, #1/31/2020#)', 'DateAdd("h", 25, #1/31/2020#)', 'DateAdd("n", 1, #1/31/2020#)', 'DateAdd("s", 1, #1/31/2020#)', 'DateAdd("y", 1, #1/31/2020#)', 'DateAdd("x", 1, #1/31/2020#)', 'DateAdd("m", 1.5, #1/31/2020#)', 'DateAdd("d", 1, "1/31/2020")', 'DateAdd("d", 1, "x")', 'DateAdd("yyyy", 9000, #1/1/2020#)',
        'DateDiff("d", #1/1/2020#, #1/31/2020#)', 'DateDiff("m", #1/31/2020#, #2/1/2020#)', 'DateDiff("yyyy", #12/31/2019#, #1/1/2020#)', 'DateDiff("q", #1/1/2020#, #12/31/2020#)', 'DateDiff("w", #1/1/2020#, #1/31/2020#)', 'DateDiff("ww", #1/1/2020#, #1/31/2020#)', 'DateDiff("ww", #1/1/2020#, #1/31/2020#, vbMonday)', 'DateDiff("h", #1/1/2020#, #1/2/2020#)', 'DateDiff("n", #1/1/2020#, #1/2/2020#)', 'DateDiff("s", #1/1/2020#, #1/2/2020#)', 'DateDiff("d", #1/31/2020#, #1/1/2020#)', 'DateDiff("d", "1/1/2020", "1/3/2020")', 'DatePart("yyyy", #3/4/2020#)', 'DatePart("q", #3/4/2020#)', 'DatePart("m", #3/4/2020#)', 'DatePart("y", #3/4/2020#)', 'DatePart("d", #3/4/2020#)', 'DatePart("w", #3/4/2020#)', 'DatePart("w", #3/4/2020#, vbMonday)', 'DatePart("ww", #3/4/2020#)', 'DatePart("ww", #1/1/2021#, vbMonday, vbFirstFourDays)', 'DatePart("h", #3/4/2020 5:06:07 AM#)', 'Year(#3/4/2020#)', 'Month(#3/4/2020#)', 'Day(#3/4/2020#)', 'Hour(#5:06:07 PM#)', 'Minute(#5:06:07 PM#)', 'Second(#5:06:07 PM#)', 'Weekday(#3/4/2020#)', 'Weekday(#3/4/2020#, vbMonday)', 'WeekdayName(1)', 'WeekdayName(1, True)', 'WeekdayName(1, False, vbMonday)', 'WeekdayName(8)', 'MonthName(1)', 'MonthName(1, True)', 'MonthName(13)',
        'CDate("1/2/2020")', 'CDate("2020-01-02")', 'CDate("Jan 2, 2020")', 'CDate("2 January 2020")', 'CDate("13/1/2020")', 'CDate("1/13/2020")', 'CDate("x")', 'CDate(43831)', 'CDate(43831.5)', 'CDate("12:30")', 'CDate("12:30 PM")', 'CDate("")', 'CDate(Empty)', 'CDbl(#1/1/2020#)', 'CLng(#1/1/2020#)', 'IsDate("1/2/2020")', 'IsDate("13/13/2020")', 'IsDate("x")', 'IsDate(43831)', 'IsDate(#1/1/2020#)', 'IsDate("Jan")', 'IsDate("12:30")', 'IsDate(Empty)', 'DateValue("1/2/2020")', 'DateValue("1/2/2020 3:00 PM")', 'TimeValue("3:00 PM")', 'TimeValue("15:00:30")', 'DateValue("x")', '#1/2/2020# + 1', '#1/2/2020# - #1/1/2020#', 'TypeName(#1/2/2020# - #1/1/2020#)', '#1/2/2020# > #1/1/2020#', '#1/2/2020# = "1/2/2020"', 'CStr(#1/2/2020#)', 'CStr(#3:04:05 PM#)', 'CStr(#1/2/2020 3:04:05 PM#)', 'CStr(#12:00:00 AM#)',
        'Format(#1/2/2020#, "yyyy-mm-dd")', 'Format(#1/2/2020#, "mm/dd/yy")', 'Format(#1/2/2020#, "d-mmm-yy")', 'Format(#1/2/2020#, "dddd, mmmm d, yyyy")', 'Format(#1/2/2020#, "ddd")', 'Format(#1/2/2020#, "w")', 'Format(#1/2/2020#, "ww")', 'Format(#1/2/2020#, "q")', 'Format(#1/2/2020#, "y")', 'Format(#1/2/2020 3:04:05 PM#, "hh:mm:ss")', 'Format(#1/2/2020 3:04:05 PM#, "h:m:s AM/PM")', 'Format(#1/2/2020 3:04:05 PM#, "hh:nn AMPM")', 'Format(#1/2/2020 3:04:05 PM#, "ttttt")', 'Format(#1/2/2020#, "Short Date")', 'Format(#1/2/2020#, "Long Date")', 'Format(#1/2/2020#, "Medium Date")', 'Format(#3:04:05 PM#, "Short Time")', 'Format(#3:04:05 PM#, "Long Time")', 'Format(#3:04:05 PM#, "Medium Time")', 'Format(#1/2/2020#, "General Date")', 'Format(#1/2/2020 3:04:05 PM#)', 'Format(#1/2/2020#)', 'Format("1/2/2020", "yyyy")', 'Format("x", "yyyy")', 'Format(43831, "yyyy-mm-dd")', 'Format(0, "yyyy-mm-dd")', 'FormatDateTime(#1/2/2020 3:04:05 PM#)', 'FormatDateTime(#1/2/2020#, vbLongDate)', 'FormatDateTime(#1/2/2020#, vbShortDate)', 'FormatDateTime(#3:04:05 PM#, vbLongTime)', 'FormatDateTime(#3:04:05 PM#, vbShortTime)',
    ],
    'number_formats': [],
    'strings': [
        'Left("abc", 2)', 'Left("abc", 9)', 'Left("abc", 0)', 'Left("abc", -1)', 'Left(Null, 1)', 'Left(123, 2)', 'Right("abc", 2)', 'Mid("abc", 2)', 'Mid("abc", 2, 1)', 'Mid("abc", 9)', 'Mid("abc", 0)', 'Mid("abc", 2, -1)', 'Len("abc")', 'Len(123)', 'Len(Null)', 'Len(Empty)', 'Len(1.5)', 'LenB("abc")', 'InStr("abc", "b")', 'InStr(2, "abcb", "b")', 'InStr(3, "abcb", "b")', 'InStr(9, "abc", "b")', 'InStr(0, "abc", "b")', 'InStr("abc", "")', 'InStr("", "a")', 'InStr(1, "aBc", "b", vbTextCompare)', 'InStr(1, "aBc", "b", vbBinaryCompare)', 'InStr(Null, "a")', 'InStr("a", Null)', 'InStrRev("abcb", "b")', 'InStrRev("abcb", "b", 3)', 'InStrRev("abcb", "z")', 'InStrRev("abc", "")', 'InStrRev("aBc", "b", -1, vbTextCompare)', 'Replace("aXbX", "X", "-")', 'Replace("aXbX", "X", "-", 3)', 'Replace("aXbX", "X", "-", 1, 1)', 'Replace("aXbX", "x", "-", , , vbTextCompare)', 'Replace("abc", "", "-")', 'Replace("", "a", "-")', 'Replace("abc", "b", "")', 'Replace(Null, "a", "b")', 'Replace("abc", "a", "b", 9)', 'Replace("abc", "a", "b", 0)', 'Trim("  a  ")', 'LTrim("  a  ") & "|"', 'RTrim("  a  ") & "|"', 'Trim(Chr(9) & "a")', 'UCase("aé")', 'LCase("AB")', 'StrConv("ab cd", vbProperCase)', 'StrConv("AB", vbLowerCase)', 'StrConv("ab", vbUpperCase)', 'StrReverse("abc")', 'Space(2) & "|"', 'Space(-1)', 'String(2, "ab")', 'String(2, 65)', 'String(0, "a")', 'String(-1, "a")', 'Asc("A")', 'Asc("")', 'AscW("A")', 'Chr(65)', 'Chr(256)', 'Chr(-1)', 'ChrW(65)', 'Hex(255)', 'Hex(-1)', 'Oct(8)', 'Val("12abc")', 'Val("1.5e2")', 'Val("  1 2")', 'Val("&HFF")', 'Val("abc")', 'Val("")', 'Str(5)', 'Str(-5)', 'Str(1.5)', 'CStr(1.5)', 'CStr(True)', 'CStr(Null)', 'CStr(Empty) & "|"', 'CInt("5")', 'CInt("x")', 'CInt(2.5)', 'CInt(3.5)', 'CInt(40000)', 'CLng("1,000")', 'CLng(2.5)', 'CDbl("1.5")', 'CDbl("1,000.5")', 'CDbl("$5")', 'CDbl("5%")', 'CDbl("1e3")', 'CDbl("")', 'CCur("1.23456")', 'CBool("True")', 'CBool("false")', 'CBool("x")', 'CBool(0)', 'CBool("1")', 'CByte(255)', 'CByte(256)', 'CSng(1.5)', 'CDec("1.5")', 'CVar(1)', 'Fix(-1.5)', 'Int(-1.5)', 'Round(2.5)', 'Round(3.5)', 'Round(2.345, 2)', 'Round(-2.5)', 'Abs(-1)', 'Sgn(-2)', 'Sqr(4)', 'Sqr(-1)', 'Rnd(0) >= 0', 'Log(1)', 'Log(0)', 'Exp(0)',
        'StrComp("a", "b")', 'StrComp("a", "A")', 'StrComp("a", "A", vbTextCompare)', 'StrComp("a", Null)', '"a" = "A"', '"a" < "B"', '"abc" Like "a*"', '"abc" Like "A*"', '"a1" Like "a#"', '"b" Like "[a-c]"', '"b" Like "[!a-c]"', 'Join(Array("a", "b"))', 'Join(Array("a", "b"), "-")', 'Join(Array(1, 2), "")', 'Join(Array())', 'Join(Split("a b"), ",")', 'UBound(Split("a,b,c", ","))', 'UBound(Split("", ","))', 'UBound(Split("a,b,c", ",", 2))', 'Split("a,b,c", ",", 2)(1)', 'Split("a,b", ",", -1)(0)', 'UBound(Split("a b"))', 'UBound(Split("aXb", "x", , vbTextCompare))', 'UBound(Split("abc", ""))', 'Split(Null)', 'Join(Filter(Array("ab", "cd"), "a"), ",")', 'Join(Filter(Array("ab", "cd"), "a", False), ",")', 'UBound(Filter(Array("ab"), "z"))', 'IsNumeric("1")', 'IsNumeric("1,000")', 'IsNumeric("$1")', 'IsNumeric("1e3")', 'IsNumeric("x")', 'IsNumeric("")', 'IsNumeric(Empty)', 'IsNumeric(Null)', 'IsNumeric(True)', 'IsNumeric("&H10")', 'IsEmpty(Empty)', 'IsEmpty("")', 'IsNull(Null)', 'IsArray(Array())', 'IsMissing(1)', 'IsObject(ws)', 'IsError(CVErr(5))', 'VarType(1)', 'VarType(1&)', 'VarType("a")', 'VarType(1.5)', 'VarType(Null)', 'VarType(Empty)', 'VarType(True)', 'VarType(#1/1/2020#)', 'VarType(CCur(1))', 'VarType(Array(1))', 'VarType(ws)', 'TypeName(1)', 'TypeName(Null)', 'TypeName(Empty)', 'TypeName(ws)', 'TypeName(Array(1))', 'TypeName(Nothing)', 'RGB(1, 2, 3)', 'RGB(300, 0, 0)', 'IIf(True, 1, 2)', 'Choose(2, "a", "b")', 'Switch(True, "a")', 'Nz',
    ],
    'files': [],
}
#: Between one line's statements and the next's, so that a line can be left out.
MARK = '\n\x00\n'
READ = 'v = Empty\nv = {one}\ns = s & Err.Number & "=" & TypeName(v) & ":" & Show(v) & "¦"\nErr.Clear'
#: Number format and value pairs: both Range.Text and VBA's Format.
FORMATS = ['General', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.0%', '$#,##0.00', '0.00E+00', '# ?/?', '# ??/??', '@', '[Red]0', '0;(0)', '#,##0;(#,##0);"-"', '0.0,', '#,##0,,', '_(* #,##0_)', '000', '#', '#.##', '?.??', '"x"0', '0"x"', '\\a0', 'mm/dd/yyyy', 'd-mmm-yy', 'yyyy', 'h:mm AM/PM', 'hh:mm:ss', '[h]:mm', 'mm:ss', 'dddd', 'mmmm', 'General;General', ';;;', '0.00_)', '[>=100]0;[<0]"neg";0.0', '[$-409]mmm', '$0', 'x']
VALUES = ['0', '1234.567', '-1234.567', '0.5', '1E+15', '"abc"', 'True', '#1/2/2020 3:04:05 PM#']
LINES = {
    'paste': [
        ('ws.Range("F2").Formula = "=C2*2"', 'ws.Range("F2").Value'), ('ws.Range("F2").Font.Bold = True', ''), ('ws.Range("F2").NumberFormat = "0.0"', ''),
        ('ws.Range("F2").Copy ws.Range("G2")', 'ws.Range("G2").Formula & ws.Range("G2").Text & ws.Range("G2").Font.Bold'), ('ws.Range("F2").Copy ws.Range("G3:G4")', 'ws.Range("G4").Formula'), ('ws.Range("F2").Copy', 'Application.CutCopyMode'),
        ('ws.Range("H2").PasteSpecial xlPasteValues', 'ws.Range("H2").Formula & ws.Range("H2").Text & ws.Range("H2").Font.Bold'), ('ws.Range("H3").PasteSpecial xlPasteFormulas', 'ws.Range("H3").Formula & ws.Range("H3").Text & ws.Range("H3").Font.Bold'), ('ws.Range("H4").PasteSpecial xlPasteFormats', 'ws.Range("H4").Formula & ws.Range("H4").Font.Bold & ws.Range("H4").NumberFormat'), ('ws.Range("H5").PasteSpecial xlPasteAll', 'ws.Range("H5").Formula & ws.Range("H5").Font.Bold'), ('ws.Range("H6").PasteSpecial xlPasteValuesAndNumberFormats', 'ws.Range("H6").Formula & ws.Range("H6").NumberFormat & ws.Range("H6").Font.Bold'), ('ws.Range("H7").PasteSpecial xlPasteFormulasAndNumberFormats', 'ws.Range("H7").Formula & ws.Range("H7").NumberFormat & ws.Range("H7").Font.Bold'), ('ws.Range("H8").PasteSpecial xlPasteAllExceptBorders', 'ws.Range("H8").Formula'), ('', 'Application.CutCopyMode'), ('', 'Selection.Address'),
        ('ws.Range("J1:J2").Value = 2', ''), ('ws.Range("K1:K2").Value = 3', ''), ('ws.Range("J1:J2").Copy', ''), ('ws.Range("K1:K2").PasteSpecial xlPasteValues, xlPasteSpecialOperationAdd', 'ws.Range("K1").Value'), ('ws.Range("K1:K2").PasteSpecial xlPasteValues, xlPasteSpecialOperationMultiply', 'ws.Range("K1").Value'), ('ws.Range("K1:K2").PasteSpecial xlPasteValues, xlPasteSpecialOperationSubtract', 'ws.Range("K1").Value'), ('ws.Range("K1:K2").PasteSpecial xlPasteValues, xlPasteSpecialOperationDivide', 'ws.Range("K1").Value'),
        ('ws.Range("L1").Value = 1', ''), ('ws.Range("L3").Value = 3', ''), ('ws.Range("M1:M3").Value = 9', ''), ('ws.Range("L1:L3").Copy', ''), ('ws.Range("M1").PasteSpecial xlPasteValues, SkipBlanks:=True', 'ws.Range("M2").Value'), ('ws.Range("M1:M3").Value = 9', ''), ('ws.Range("L1:L3").Copy', ''), ('ws.Range("M1").PasteSpecial xlPasteValues, SkipBlanks:=False', 'IsEmpty(ws.Range("M2").Value)'),
        ('ws.Range("A1:C1").Copy', ''), ('ws.Range("O1").PasteSpecial Transpose:=True', 'ws.Range("O3").Value & ws.Range("P1").Value'), ('ws.Range("A1:A2").Copy', ''), ('ws.Range("Q1:R4").PasteSpecial xlPasteValues', 'ws.Range("R4").Value'), ('ws.Range("Q6:Q8").PasteSpecial xlPasteValues', 'ws.Range("Q8").Value & "|" & ws.Range("Q7").Value'), ('Application.CutCopyMode = False', 'Application.CutCopyMode'), ('ws.Range("A1").PasteSpecial xlPasteValues', 'ws.Range("A1").Value'), ('ws.Paste ws.Range("Z1")', 'ws.Range("Z1").Value'),
        ('ws.Columns(1).ColumnWidth = 12', ''), ('ws.Range("A1").Copy', ''), ('ws.Range("S1").PasteSpecial xlPasteColumnWidths', 'ws.Columns(19).ColumnWidth'), ('ws.Range("A1:A2").Copy ws.Range("T1:T5")', 'ws.Range("T5").Value & "|" & ws.Range("T4").Value'), ('ws.Range("A1:B1").Copy ws.Range("U1:U3")', 'ws.Range("U3").Value & ws.Range("V1").Value'), ('ws.Rows(2).Copy ws.Rows(20)', 'ws.Range("A20").Value & ws.Range("D20").Text'), ('ws.Columns(1).Copy ws.Columns(24)', 'ws.Range("X2").Value & ws.Columns(24).ColumnWidth'), ('ws.Range("A2").Copy ws.Rows(30)', ''), ('ws.Range("A2:B3").Cut ws.Range("A30")', 'ws.Range("A30").Value & IsEmpty(ws.Range("A2").Value) & Application.CutCopyMode'), ('ws.Range("A30:B31").Cut', 'Application.CutCopyMode'), ('ws.Paste ws.Range("A2")', 'ws.Range("A2").Value & Application.CutCopyMode'), ('ws.Range("A2:B3").Value = ws.Range("A2:B3").Value', 'ws.Range("A3").Value'), ('ws.Range("A2:A3").Value = ws.Range("B2:B3").Value', 'ws.Range("A3").Value'), ('ws.Range("A2:A3").Value = ws.Range("C2:C7").Value', 'ws.Range("A3").Value'), ('ws.Range("A2:A5").Value = ws.Range("C2:C3").Value', 'CStr(ws.Range("A4").Text)'), ('ws.Range("AA1:AA3").Value = 5', 'ws.Range("AA3").Value'), ('ws.Range("AB1:AB3").Formula = "=AA1+1"', 'ws.Range("AB3").Formula & ws.Range("AB3").Value'), ('ws.Range("AC1:AC2").Value = Array(1, 2)', 'ws.Range("AC2").Value'), ('ws.Range("AC1:AD1").Value = Array(1, 2)', 'ws.Range("AD1").Value'), ('ws.Range("AE1:AE2").FillDown', ''),
    ],
    'files': [
        ('f = Environ("TEMP") & "\\sweep16a.xlsx"', ''), ('On Error Resume Next: Kill f: Err.Clear', ''), ('', 'Dir(f)'), ('Set o = Workbooks.Add(xlWBATWorksheet)', 'o.Saved'), ('n = Workbooks.Count', 'o.Path'), ('', 'Len(o.Name) > 0'), ('', 'o.FileFormat'), ('o.Worksheets(1).Range("A1").Value = 5', 'o.Saved'), ('o.Worksheets(1).Range("B1").Formula = "=A1*2"', ''), ('o.SaveAs f, xlOpenXMLWorkbook', 'o.Saved'), ('', 'o.Name'), ('', 'o.FullName = f'), ('', 'o.Path = Environ("TEMP")'), ('', 'o.FileFormat'), ('', 'Dir(f)'), ('', 'FileLen(f) > 0'), ('', 'o.ReadOnly'), ('o.Worksheets(1).Range("A1").Value = 6', 'o.Saved'), ('o.Save', 'o.Saved'), ('o.Close', 'Workbooks.Count = n - 1'), ('', 'o.Name'), ('Set o = Workbooks.Open(f)', 'o.Name'), ('', 'o.Worksheets(1).Range("A1").Value'), ('', 'o.Worksheets(1).Range("B1").Formula'), ('', 'o.Worksheets(1).Range("B1").Value'), ('', 'o.Saved'), ('', 'ActiveWorkbook.Name'), ('', 'ActiveSheet.Name'), ('', 'Workbooks("sweep16a.xlsx").Name'), ('', 'Workbooks("SWEEP16A.XLSX").Name'), ('', 'Workbooks("sweep16a").Name'), ('Set c = Nothing', 'TypeName(Workbooks.Open(f))'), ('', 'Workbooks.Count = n'), ('o.Close False', 'Workbooks.Count = n - 1'),
        ('Set o = Workbooks.Open(f, ReadOnly:=True)', 'o.ReadOnly'), ('o.Close False', ''), ('Set o = Workbooks.Open(Filename:=f, UpdateLinks:=0)', 'o.Name'), ('o.Worksheets(1).Range("A1").Value = 7', 'o.Saved'), ('o.Saved = True', 'o.Saved'), ('o.Close', ''), ('Set o = Workbooks.Open(f)', 'o.Worksheets(1).Range("A1").Value'), ('o.Worksheets(1).Range("A1").Value = 8', ''), ('o.Close SaveChanges:=True', ''), ('Set o = Workbooks.Open(f)', 'o.Worksheets(1).Range("A1").Value'), ('o.Worksheets(1).Range("A1").Value = 9', ''), ('o.Close False', ''), ('Set o = Workbooks.Open(f)', 'o.Worksheets(1).Range("A1").Value'), ('o.SaveCopyAs Environ("TEMP") & "\\sweep16b.xlsx"', 'Dir(Environ("TEMP") & "\\sweep16b.xlsx")'), ('', 'o.Name'), ('o.Close False', ''), ('Kill Environ("TEMP") & "\\sweep16b.xlsx"', 'Dir(Environ("TEMP") & "\\sweep16b.xlsx")'),
        ('Set o = Nothing', ''), ('Workbooks.Open "nofile_sweep16.xlsx"', ''), ('', 'Workbooks("nofile_sweep16.xlsx").Name'), ('', 'Workbooks(99).Name'), ('', 'Dir("nofile_sweep16.xlsx") = ""'), ('Kill "nofile_sweep16.xlsx"', ''),
        ('t = Environ("TEMP") & "\\sweep16c.csv"', ''), ('Set o = Workbooks.Add(xlWBATWorksheet)', ''), ('o.Worksheets(1).Range("A1:C1").Value = Array("a b", 1.5, "x,y")', ''), ('o.Worksheets(1).Range("A2").Value = #1/2/2020#', ''), ('o.Worksheets(1).Range("B2").Formula = "=1+1"', ''), ('o.SaveAs t, xlCSV', 'o.Name'), ('', 'o.FileFormat'), ('', 'o.Worksheets(1).Name'), ('o.Close False', ''), ('Open t For Input As #1', ''), ('', 'FileText()'), ('Close #1', ''), ('Set o = Workbooks.Open(t)', 'o.Worksheets(1).Range("C1").Value'), ('', 'o.Worksheets(1).Range("A1").Value'), ('', 'o.Worksheets(1).Range("B1").Value'), ('', 'TypeName(o.Worksheets(1).Range("A2").Value)'), ('', 'o.Worksheets(1).Range("B2").Value'), ('', 'o.FileFormat'), ('o.Close False', ''), ('Kill t', ''), ('Kill f', 'Dir(f)'),
    ],
}
GROUPS.pop('paste')
GROUPS.pop('files')
GROUPS.pop('number_formats')
HELPERS = """
Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        Show = "object"
    ElseIf IsArray(v) Then
        Show = "array"
    ElseIf IsError(v) Then
        Show = "error"
    ElseIf IsNull(v) Then
        Show = "null"
    Else
        Show = CStr(v)
    End If
End Function

Private Function Total(ByVal r As Range) As Double
    Dim c As Range
    For Each c In r
        Total = Total + c.Value
    Next
End Function

Private Function Names(ByVal r As Range) As String
    Dim c As Range
    For Each c In r.Cells
        Names = Names & c.Value & ","
    Next
End Function

Private Function RowsOf(ByVal r As Range) As String
    Dim c As Range
    For Each c In r.Rows
        RowsOf = RowsOf & c.Address & ","
    Next
End Function

Private Function ColumnsOf(ByVal r As Range) As String
    Dim c As Range
    For Each c In r.Columns
        ColumnsOf = ColumnsOf & c.Address & ","
    Next
End Function

Private Function FindAll(ByVal what As Variant) As String
    Dim r As Range, first As String, guard As Long
    Set r = ActiveSheet.Cells.Find(what, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If r Is Nothing Then Exit Function
    first = r.Address
    Do
        FindAll = FindAll & r.Address & ","
        Set r = ActiveSheet.Cells.FindNext(r)
        guard = guard + 1
    Loop While Not r Is Nothing And r.Address <> first And guard < 20
End Function

Private Function Visible(ByVal r As Range) As String
    Dim c As Range
    For Each c In r
        If Not c.EntireRow.Hidden Then Visible = Visible & c.Row & ","
    Next
End Function

Private Function Do1(ByVal v As Variant) As String
    Do1 = "ran:" & TypeName(v)
End Function

Private Function Col(ByVal which As Long) As String
    Dim i As Long
    For i = 1 To 7
        Col = Col & ActiveSheet.Cells(i, which).Text & ","
    Next
End Function

Private Function Row1() As String
    Dim i As Long
    For i = 1 To 4
        Row1 = Row1 & ActiveSheet.Cells(1, i).Text & ","
    Next
End Function

Private Function ColH() As String
    Dim i As Long
    For i = 1 To 9
        ColH = ColH & TypeName(ActiveSheet.Cells(i, 8).Value) & ActiveSheet.Cells(i, 8).Text & ","
    Next
End Function

Private Function FillMixed() As Boolean
    With ActiveSheet
        .Range("H1").Value = "b": .Range("H2").Value = 10: .Range("H3").Value = "A": .Range("H4").Value = True: .Range("H5").Value = -1
        .Range("H6").Value = "a": .Range("H7").Formula = "=1/0": .Range("H9").Value = "2": .Range("H9").NumberFormat = "@"
    End With
    FillMixed = True
End Function

Private Function FileText() As String
    Dim line As String
    Do While Not EOF(1)
        Line Input #1, line
        FileText = FileText & line & "/"
    Loop
End Function
"""


def _read_line(one: str) -> str:
    return READ.replace('{one}', one)


def _body(name: str) -> tuple[list[str], str]:
    """What a group asks, and the VBA that asks it."""
    if name.startswith('number_formats_'):
        asked, lines = [], []
        part = int(name.rsplit('_', 1)[1])
        for code in FORMATS[part * FORMATS_IN_A_PART:(part + 1) * FORMATS_IN_A_PART]:
            for value in VALUES:
                quoted = code.replace('"', '""')
                for form in (f'ws.Range("A1").Text', f'Format({value}, "{quoted}")'):
                    asked.append(f'{value} in {code} by {form[:6]}')
                    lines.append(f'ws.Range("A1").NumberFormat = "{quoted}"\nws.Range("A1").Value = {value}\n' + _read_line(form))
        return asked, MARK.join(lines)
    if name in LINES:
        asked, lines = [], []
        for statement, expression in LINES[name]:
            asked.append(f'{statement} / {expression}')
            run = f'{statement}\ns = s & Err.Number & "!"\nErr.Clear\n' if statement else 's = s & "0!"\n'
            lines.append(run + _read_line(expression or '""'))
        return asked, MARK.join(lines)
    return GROUPS[name], MARK.join(_read_line(one) for one in GROUPS[name])


#: One procedure of all of them is too large for VBA.
FORMATS_IN_A_PART = 8
NAMES = [*GROUPS, 'paste', *(f'number_formats_{part}' for part in range((len(FORMATS) + FORMATS_IN_A_PART - 1) // FORMATS_IN_A_PART)), 'files']


def _declared(source: str, typed: bool) -> str:
    return source if typed else source.replace('Dim ws As Worksheet,', 'Dim ws As Object,', 1)


def macro(name: str, *, typed: bool = False, leaving_out: frozenset[int] = frozenset()) -> tuple[list[str], str]:
    asked, body = _body(name)
    if leaving_out:
        # A line left out still adds its place to the answer, so that the lines stay in step.
        parts = body.split(MARK)
        body = '\n'.join('s = s & "skipped¦"' if index in leaving_out else part for index, part in enumerate(parts))
    else:
        body = body.replace(MARK, '\n')
    cleanup = '\nOn Error Resume Next\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\nws.AutoFilterMode = False\nws.Delete\nApplication.DisplayAlerts = True\nOn Error GoTo 0'
    source = HELPERS + '\nFunction Probe() As String\n' + PRE + 'On Error Resume Next\n' + body + '\nOn Error GoTo 0' + cleanup + '\nProbe = s\nEnd Function\n'
    return asked, _declared(source, typed)


def _run(source: str) -> object:
    """Its answer, or None if it does not compile."""
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=240.0, line_numbers=False)
    if any(dialog.classification == 'compile-error' for dialog in result.dialogs):
        return None
    assert result.ok and not result.dialogs, (result.error, result.message, result.dialogs)
    return result.value


def _typed(name: str, count: int, *, typed: bool = True) -> tuple[object, list[int]]:
    """The answer, with the lines that will not compile that way left out, found by halving."""
    bad: set[int] = set()
    answer = _run(macro(name, typed=typed)[1])
    if answer is not None:
        return answer, []

    def compiles(keep: list[int]) -> bool:
        leave = frozenset(range(count)) - frozenset(keep)
        return _run(macro(name, typed=typed, leaving_out=leave)[1]) is not None

    def search(lines: list[int]) -> None:
        if compiles(lines):
            return
        if len(lines) == 1:
            bad.add(lines[0])
            return
        middle = len(lines) // 2
        search(lines[:middle])
        search(lines[middle:])

    search(list(range(count)))
    answer = _run(macro(name, typed=typed, leaving_out=frozenset(bad))[1])
    return answer, sorted(bad)


def main() -> None:
    held = json.loads(OUT.read_text(encoding='utf-8'))['probes'] if OUT.exists() else []
    probes: list[dict[str, object]] = list(held)
    done = {probe['name'] for probe in probes}
    for name in NAMES:
        if name in done:
            continue
        asked, source = macro(name)
        late, not_compiling_late = _typed(name, len(asked), typed=False)
        assert late is not None, name
        source = macro(name, leaving_out=frozenset(not_compiling_late))[1]
        typed, not_compiling = _typed(name, len(asked))
        probes.append({'name': name, 'asked': asked, 'late_macro': source, 'typed_macro': macro(name, typed=True, leaving_out=frozenset(not_compiling))[1],
                       'late': late, 'typed': typed, 'not_compiling_typed': not_compiling, 'not_compiling_late': not_compiling_late})
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=1, ensure_ascii=True) + '\n', encoding='utf-8')
        print(name, len(asked), 'typed will not compile:', [asked[index][:50] for index in not_compiling], flush=True)


if __name__ == '__main__':
    main()
