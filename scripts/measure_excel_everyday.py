"""Everyday Excel macros: ranges, formats, sheets, names, Find, Sort, filters, worksheet functions, VBA's own
functions, links, notes, shapes, charts, tables, validation and protection.

A sweep of what macros commonly do, each probe on a sheet of its own that it
deletes afterwards, recorded as its answer or the run-time error it stopped
with (see vba_probe_fixture.py). It found ranges across two sheets,
Application.ReferenceStyle, the separators and International; the rest
already agreed.

    python scripts/measure_excel_everyday.py

writes tests/fixtures/excel_everyday.json, which tests/test_excel_everyday.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday.json'
PRE = "Dim ws As Worksheet, r As Range, s As String, i As Long, v As Variant\nSet ws = Worksheets.Add\n"

CASES: dict[str, str] = {
    "last_row_end": 'For i = 1 To 10: If i <> 5 Then ws.Cells(i, 1).Value = i\nNext\n'
                    's = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row & "|" & ws.Range("A1").End(xlDown).Row & "|" & '
                    'ws.Range("A6").End(xlUp).Row & "|" & ws.Range("A1").End(xlToRight).Column',
    "used_range": 'ws.Range("B3").Value = 1: ws.Range("D7").Value = 2\ns = ws.UsedRange.Address\n'
                  'ws.Range("D7").ClearContents\ns = s & "|" & ws.UsedRange.Address',
    "current_region": 'ws.Range("B2:D3").Value = 1: ws.Range("B5:D5").Value = 2: ws.Range("E4").Value = 3\n'
                      's = ws.Range("C2").CurrentRegion.Address & "|" & ws.Range("B5").CurrentRegion.Address',
    "offset_resize": 's = ws.Range("B2").Offset(1, 2).Resize(2, 3).Address & "|" & ws.Range("B2:C3").Resize(, 1).Address\n'
                     'On Error Resume Next\nSet r = ws.Range("A1").Offset(-1, 0)\ns = s & "|" & Err.Number',
    "union_intersect": 's = Union(ws.Range("A1:B2"), ws.Range("B2:C3")).Address & "|" & '
                       'Intersect(ws.Range("A1:C3"), ws.Range("B2:D4")).Address & "|" & '
                       '(Intersect(ws.Range("A1"), ws.Range("C3")) Is Nothing) & "|" & Union(ws.Range("A1:A2"), ws.Range("A3:A4")).Areas.Count',
    "address_variants": 'Set r = ws.Range("B2:C3")\ns = r.Address(False, True) & "|" & r.Address(RowAbsolute:=False) & "|" & '
                        'r.Address(ReferenceStyle:=xlR1C1) & "|" & r.Address(False, False, xlR1C1, , ws.Range("A1")) & "|" & '
                        'ws.Range("A:A").Address & "|" & ws.Range("2:3").Address & "|" & ws.Cells.Address',
    "cells_loop_sum": 'For i = 1 To 10: ws.Cells(i, 1).Value = i * i: Next\n'
                      's = WorksheetFunction.Sum(ws.Range("A1:A10")) & "|" & WorksheetFunction.Average(ws.Range("A1:A10")) & "|" & '
                      'WorksheetFunction.Max(ws.Range("A1:A10"))',
    "value_array_roundtrip": 'ws.Range("A1:C2").Value = 5\nv = ws.Range("A1:C2").Value\n'
                             's = TypeName(v) & "|" & LBound(v, 1) & UBound(v, 1) & LBound(v, 2) & UBound(v, 2) & "|" & TypeName(v(1, 1))\n'
                             'v(2, 3) = "x"\nws.Range("E1:G2").Value = v\ns = s & "|" & ws.Range("G2").Value & "|" & ws.Range("E1").Value',
    "value_1d_array_to_row": 'ws.Range("A1:C1").Value = Array(1, "b", True)\ns = ws.Range("A1").Value & ws.Range("B1").Value & ws.Range("C1").Value\n'
                             'ws.Range("A3:A5").Value = Array(7, 8, 9)\ns = s & "|" & ws.Range("A3").Value & ws.Range("A4").Value & ws.Range("A5").Value\n'
                             'ws.Range("A7:A9").Value = WorksheetFunction.Transpose(Array(7, 8, 9))\ns = s & "|" & ws.Range("A7").Value & ws.Range("A8").Value & ws.Range("A9").Value',
    "value2_dates": 'ws.Range("A1").Value = DateSerial(2024, 1, 15)\ns = TypeName(ws.Range("A1").Value) & "|" & ws.Range("A1").Value2 & "|" & '
                    'ws.Range("A1").Text & "|" & ws.Range("A1").NumberFormat',
    "formula_writes": 'ws.Range("A1:A3").Value = 2\nws.Range("B1").Formula = "=SUM(A1:A3)"\nws.Range("B2").FormulaR1C1 = "=R[-1]C*2"\n'
                      's = ws.Range("B1").Value & "|" & ws.Range("B1").FormulaR1C1 & "|" & ws.Range("B2").Formula & "|" & ws.Range("B2").Value & "|" & ws.Range("B1").HasFormula',
    "fill_down": 'ws.Range("A1").Value = 1\nws.Range("A2").Formula = "=A1+1"\nws.Range("A2:A5").FillDown\n'
                 's = ws.Range("A5").Value & "|" & ws.Range("A5").Formula',
    "copy_paste_values": 'ws.Range("A1:A3").Value = 3\nws.Range("B1").Formula = "=A1*2"\nws.Range("B1").Copy ws.Range("C1")\n'
                         'ws.Range("B1").Copy\nws.Range("D1").PasteSpecial xlPasteValues\nApplication.CutCopyMode = False\n'
                         's = ws.Range("C1").Formula & "|" & ws.Range("D1").Formula & "|" & ws.Range("D1").Value',
    "insert_delete_rows": 'For i = 1 To 5: ws.Cells(i, 1).Value = i: Next\nws.Rows(2).Insert\nws.Rows(4).Delete\n'
                          'For i = 1 To 5: s = s & ws.Cells(i, 1).Value & ",": Next',
    "delete_rows_backwards": 'For i = 1 To 8: ws.Cells(i, 1).Value = i: Next\nFor i = 8 To 1 Step -1\n'
                             'If ws.Cells(i, 1).Value Mod 2 = 0 Then ws.Rows(i).Delete\nNext\nFor i = 1 To 5: s = s & ws.Cells(i, 1).Value & ",": Next',
    "clear_variants": 'ws.Range("A1:A3").Value = 1: ws.Range("A1:A3").Font.Bold = True\nws.Range("A1").ClearContents\nws.Range("A2").Clear\n'
                      'ws.Range("A3").ClearFormats\ns = ws.Range("A1").Font.Bold & "|" & ws.Range("A2").Font.Bold & "|" & IsEmpty(ws.Range("A2").Value) & "|" & '
                      'ws.Range("A3").Value & "|" & ws.Range("A3").Font.Bold',
    "merge_cells": 'ws.Range("A1").Value = "x": ws.Range("B2").Value = "y"\nApplication.DisplayAlerts = False\nws.Range("A1:B2").Merge\n'
                   'Application.DisplayAlerts = True\ns = ws.Range("A1").MergeCells & "|" & ws.Range("B2").MergeArea.Address & "|" & ws.Range("A1").Value & "|" & '
                   'IsEmpty(ws.Range("B2").Value)\nws.Range("A1").UnMerge\ns = s & "|" & ws.Range("B2").MergeCells',
    "find_findnext": 'ws.Range("A2").Value = "x": ws.Range("A5").Value = "x": ws.Range("C3").Value = "xx"\n'
                     'Dim first As String\nSet r = ws.Cells.Find(What:="x", LookAt:=xlWhole)\nIf Not r Is Nothing Then\nfirst = r.Address\n'
                     'Do\ns = s & r.Address & ","\nSet r = ws.Cells.FindNext(r)\nLoop While r.Address <> first\nEnd If',
    "replace_partial": 'ws.Range("A1").Value = "banana": ws.Range("A2").Value = "apple"\n'
                       'ws.Range("A1:A2").Replace What:="a", Replacement:="o", LookAt:=xlPart, MatchCase:=False\n'
                       's = ws.Range("A1").Value & "|" & ws.Range("A2").Value',
    "sort_two_keys": 'ws.Range("A1:B1").Value = Array("k", "v")\nws.Range("A2:B2").Value = Array("b", 2)\nws.Range("A3:B3").Value = Array("a", 3)\n'
                     'ws.Range("A4:B4").Value = Array("b", 1)\nws.Range("A1:B4").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, '
                     'Key2:=ws.Range("B1"), Order2:=xlDescending, Header:=xlYes\nFor i = 2 To 4: s = s & ws.Cells(i, 1).Value & ws.Cells(i, 2).Value & ",": Next',
    "autofilter_visible": 'ws.Range("A1").Value = "n"\nFor i = 2 To 6: ws.Cells(i, 1).Value = i: Next\nws.Range("A1:A6").AutoFilter 1, ">3"\n'
                          's = ws.Range("A1:A6").SpecialCells(xlCellTypeVisible).Count & "|" & ws.Rows(3).Hidden & "|" & ws.AutoFilterMode\n'
                          'ws.AutoFilterMode = False\ns = s & "|" & ws.Rows(3).Hidden',
    "special_cells": 'ws.Range("A1").Value = 1: ws.Range("A2").Value = "t": ws.Range("A3").Formula = "=1+1"\n'
                     's = ws.Range("A1:A3").SpecialCells(xlCellTypeConstants, xlNumbers).Address & "|" & '
                     'ws.Range("A1:A3").SpecialCells(xlCellTypeFormulas).Address & "|" & ws.Cells.SpecialCells(xlCellTypeLastCell).Address',
    "names_add": 'ws.Range("A1:A3").Value = 4\nThisWorkbook.Names.Add Name:="Probe_Rng", RefersTo:="=" & ws.Name & "!$A$1:$A$3"\n'
                 's = ThisWorkbook.Names("Probe_Rng").RefersToRange.Address & "|" & Application.Evaluate("SUM(Probe_Rng)") & "|" & '
                 'ws.Range("Probe_Rng").Count\nThisWorkbook.Names("Probe_Rng").Delete',
    "sheet_add_rename_delete": 'Dim n As Long: n = Worksheets.Count\nDim w2 As Worksheet\nSet w2 = Worksheets.Add(After:=Worksheets(Worksheets.Count))\n'
                               'w2.Name = "ProbeSheet"\ns = (Worksheets.Count - n) & "|" & Worksheets("probesheet").Index - Worksheets.Count & "|" & w2.Previous.Name = ws.Name\n'
                               'Application.DisplayAlerts = False\nw2.Delete\nApplication.DisplayAlerts = True\ns = s & "|" & (Worksheets.Count - n)',
    "interior_font": 'Set r = ws.Range("A1")\nr.Interior.Color = RGB(255, 0, 0)\nr.Font.Bold = True\nr.Font.Size = 14\nr.Font.Color = vbBlue\n'
                     's = r.Interior.Color & "|" & r.Interior.ColorIndex & "|" & r.Font.Bold & "|" & r.Font.Size & "|" & r.Font.Color & "|" & r.Font.ColorIndex',
    "number_format_text": 'ws.Range("A1").Value = 1234.5\nws.Range("A1").NumberFormat = "#,##0.00"\ns = ws.Range("A1").Text\n'
                          'ws.Range("A1").NumberFormat = "0%"\ns = s & "|" & ws.Range("A1").Text\nws.Range("A2").Value = DateSerial(2024, 3, 5)\n'
                          'ws.Range("A2").NumberFormat = "dddd, mmmm d, yyyy"\ns = s & "|" & ws.Range("A2").Text',
    "borders": 'Set r = ws.Range("B2:C3")\nr.Borders(xlEdgeBottom).LineStyle = xlContinuous\nr.Borders(xlEdgeBottom).Weight = xlThick\n'
               's = r.Borders(xlEdgeBottom).LineStyle & "|" & r.Borders(xlEdgeBottom).Weight & "|" & IsNull(r.Borders.LineStyle) & "|" & '
               'ws.Range("B3").Borders(xlEdgeBottom).Weight',
    "column_width_row_height": 'ws.Columns("A").ColumnWidth = 20\nws.Rows(2).RowHeight = 30\ns = ws.Columns("A").ColumnWidth & "|" & ws.Rows(2).RowHeight & "|" & '
                               'ws.Range("A1").Width & "|" & ws.Range("A2").Height',
    "alignment_wrap": 'Set r = ws.Range("A1")\nr.HorizontalAlignment = xlCenter\nr.VerticalAlignment = xlTop\nr.WrapText = True\nr.IndentLevel = 0\n'
                      's = r.HorizontalAlignment & "|" & r.VerticalAlignment & "|" & r.WrapText & "|" & r.Orientation',
    "conditional_format": 'ws.Range("A1:A3").Value = Array(1, 6, 9)\nDim fc As Object\nSet fc = ws.Range("A1:A3").FormatConditions.Add(xlCellValue, xlGreater, "=5")\n'
                          'fc.Interior.Color = vbYellow\ns = ws.Range("A1:A3").FormatConditions.Count & "|" & fc.Formula1 & "|" & fc.Type & "|" & '
                          'ws.Range("A2").DisplayFormat.Interior.Color & "|" & ws.Range("A1").DisplayFormat.Interior.Color',
    "wf_lookup": 'ws.Range("A1:B3").Value = 0\nws.Range("A1").Value = "a": ws.Range("A2").Value = "b": ws.Range("A3").Value = "c"\n'
                 'ws.Range("B1").Value = 10: ws.Range("B2").Value = 20: ws.Range("B3").Value = 30\n'
                 's = WorksheetFunction.VLookup("b", ws.Range("A1:B3"), 2, False) & "|" & WorksheetFunction.Match("c", ws.Range("A1:A3"), 0) & "|" & '
                 'WorksheetFunction.Index(ws.Range("B1:B3"), 2) & "|" & WorksheetFunction.CountIf(ws.Range("B1:B3"), ">15") & "|" & '
                 'WorksheetFunction.SumIfs(ws.Range("B1:B3"), ws.Range("A1:A3"), "<>a")\nOn Error Resume Next\n'
                 'v = WorksheetFunction.VLookup("z", ws.Range("A1:B3"), 2, False)\ns = s & "|" & Err.Number\nErr.Clear\n'
                 'v = Application.VLookup("z", ws.Range("A1:B3"), 2, False)\ns = s & "|" & IsError(v) & "|" & CLng(v)',
    "application_index_array": 'Dim a As Variant\na = ws.Range("A1:C2").Value\nws.Range("A1:C1").Value = Array(1, 2, 3)\nws.Range("A2:C2").Value = Array(4, 5, 6)\n'
                               'a = ws.Range("A1:C2").Value\nv = Application.Index(a, 0, 2)\ns = TypeName(v) & "|" & LBound(v, 1) & UBound(v, 1) & "|" & v(2, 1)\n'
                               'v = Application.Index(a, 2, 0)\ns = s & "|" & LBound(v) & UBound(v) & "|" & v(3)\nv = Application.Transpose(a)\ns = s & "|" & UBound(v, 1) & UBound(v, 2)',
    "evaluate_expr": 'ws.Range("A1:A3").Value = 2\ns = Application.Evaluate("=SUM(1,2)*3") & "|" & ws.Evaluate("SUM(A1:A3)") & "|" & TypeName(ws.Evaluate("A1:A3")) & "|" & '
                     'ws.Evaluate("ROWS(A1:A3)")',
    "split_join": 'v = Split("a,b,,c", ",")\ns = UBound(v) & "|" & Join(v, ";") & "|" & UBound(Split("", ",")) & "|" & Join(Split("x y  z"), "+")',
    "format_function": 's = Format(1234.567, "#,##0.00") & "|" & Format(#3/5/2024#, "yyyy-mm-dd ddd") & "|" & Format(0.5, "Percent") & "|" & '
                       'Format("abc", ">") & "|" & Format(1234.5, "Currency") & "|" & Format(0, "0.00;-0.00;\\z\\e\\r\\o") & "|" & Format(-3, "0;(0)")',
    "string_funcs": 's = InStrRev("a.b.c", ".") & "|" & StrReverse("abc") & "|" & Replace("aaaa", "a", "b", 2, 2) & "|" & Trim("  x  ") & "|" & '
                    'Space(2) & "|" & String(3, "z") & "|" & Left("abc", 5) & "|" & Mid("abcdef", 2, 3) & "|" & UCase("äb") & "|" & StrComp("a", "B", vbTextCompare)',
    "date_funcs": 's = DateAdd("m", 1, #1/31/2024#) & "|" & DateDiff("d", #1/1/2024#, #3/1/2024#) & "|" & Weekday(#3/5/2024#, vbMonday) & "|" & '
                  'DateSerial(2024, 2, 30) & "|" & Format(#3/5/2024#, "dddd") & "|" & Year(#12/31/1999#) & "|" & DatePart("q", #8/1/2024#) & "|" & '
                  'DateValue("March 5, 2024") & "|" & TimeSerial(25, 0, 0)',
    "arrays": 'Dim b() As Long\nReDim b(1 To 2, 1 To 2)\nb(1, 2) = 5\nReDim Preserve b(1 To 2, 1 To 4)\ns = b(1, 2) & "|" & UBound(b, 2)\n'
              'Erase b\ns = s & "|" & IsArray(b)\nv = Filter(Array("ab", "cd", "abc"), "ab")\ns = s & "|" & UBound(v) & "|" & Join(v, ",")\n'
              'v = Array()\ns = s & "|" & UBound(v)',
    "collection": 'Dim c As New Collection\nc.Add 1, "one": c.Add 2, "two": c.Add 3, Before:=1\ns = c.Count & "|" & c(1) & "|" & c("two")\nc.Remove "one"\n'
                  'For Each v In c: s = s & "|" & v: Next\nOn Error Resume Next\nv = c("zzz")\ns = s & "|" & Err.Number',
    "dictionary": 'Dim d As Object\nSet d = CreateObject("Scripting.Dictionary")\nd.CompareMode = vbTextCompare\nd.Add "A", 1\nd("b") = 2\n'
                  's = d.Count & "|" & d.Exists("a") & "|" & Join(d.Keys, ",") & "|" & d.Item("B")\nd.Remove "a"\ns = s & "|" & d.Count\n'
                  'On Error Resume Next\nd.Add "b", 3\ns = s & "|" & Err.Number',
    "errors": 'On Error Resume Next\nErr.Raise vbObjectError + 513, "Src", "Custom"\ns = Err.Number - vbObjectError & "|" & Err.Source & "|" & Err.Description\n'
              'Err.Clear\nv = 1 / 0\ns = s & "|" & Err.Number & "|" & Err.Description\nErr.Clear\nv = CInt("x")\ns = s & "|" & Err.Number',
    "conversions": 's = CInt(2.5) & "|" & CInt(3.5) & "|" & CLng("1e3") & "|" & Val("12abc") & "|" & CStr(1 / 3) & "|" & CDbl("1,234.5") & "|" & '
                   'Round(2.5) & "|" & Round(3.45, 1) & "|" & Int(-2.5) & "|" & Fix(-2.5) & "|" & CBool("True") & "|" & CDate("2024-03-05")',
    "like_operator": 's = ("abc" Like "a*") & "|" & ("abc" Like "a?c") & "|" & ("a1c" Like "a#c") & "|" & ("abc" Like "[a-c]bc") & "|" & '
                     '("abc" Like "[!x]*") & "|" & ("ABC" Like "abc")',
    "math_funcs": 's = Sgn(-3) & "|" & Abs(-2.5) & "|" & Sqr(16) & "|" & Round(Exp(1), 6) & "|" & Round(Log(10), 6) & "|" & (7 Mod 3) & "|" & (7 \\ 2) & "|" & '
                  '(-7 Mod 3) & "|" & (2 ^ 10) & "|" & Hex(255) & "|" & Oct(8) & "|" & Asc("A") & "|" & Chr(66) & "|" & AscW("€")',
    "hyperlinks": 'ws.Hyperlinks.Add Anchor:=ws.Range("A1"), Address:="https://example.com/", TextToDisplay:="ex"\n'
                  's = ws.Hyperlinks.Count & "|" & ws.Hyperlinks(1).Address & "|" & ws.Range("A1").Value & "|" & ws.Range("A1").Hyperlinks.Count',
    "comments": 'ws.Range("A1").AddComment "hi"\ns = ws.Comments.Count & "|" & ws.Range("A1").Comment.Text & "|" & (ws.Range("A2").Comment Is Nothing)\n'
                'ws.Range("A1").Comment.Delete\ns = s & "|" & ws.Comments.Count',
    "shapes": 'Dim sh As Shape\nSet sh = ws.Shapes.AddShape(msoShapeRectangle, 10, 20, 50, 40)\nsh.Name = "Box"\n'
              's = ws.Shapes.Count & "|" & sh.Type & "|" & sh.Left & "|" & sh.Top & "|" & sh.Width & "|" & ws.Shapes("Box").Height\nsh.Delete\ns = s & "|" & ws.Shapes.Count',
    "charts": 'ws.Range("A1:B3").Value = 1\nDim co As ChartObject\nSet co = ws.ChartObjects.Add(10, 10, 200, 100)\nco.Chart.SetSourceData ws.Range("A1:B3")\n'
              'co.Chart.ChartType = xlColumnClustered\ns = ws.ChartObjects.Count & "|" & co.Chart.SeriesCollection.Count & "|" & co.Chart.ChartType',
    "listobjects": 'ws.Range("A1:B1").Value = Array("h1", "h2")\nws.Range("A2:B3").Value = 1\nDim lo As ListObject\n'
                   'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)\ns = lo.ListRows.Count & "|" & lo.DataBodyRange.Address & "|" & '
                   'lo.ListColumns(2).Name & "|" & lo.HeaderRowRange.Address\nlo.ListRows.Add\ns = s & "|" & lo.Range.Address',
    "validation": 'With ws.Range("A1").Validation\n.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="a,b,c"\n'
                  's = .Type & "|" & .Formula1 & "|" & .InCellDropdown\nEnd With',
    "protect": 'ws.Protect "pw"\ns = ws.ProtectContents & "|"\nOn Error Resume Next\nws.Range("A1").Value = 1\ns = s & Err.Number\nErr.Clear\n'
               'ws.Unprotect "pw"\nws.Range("A1").Value = 1\ns = s & "|" & Err.Number & "|" & ws.ProtectContents',
    "outline_group": 'ws.Rows("2:4").Group\ns = ws.Rows(3).OutlineLevel & "|" & ws.Rows(1).OutlineLevel\nws.Outline.ShowLevels RowLevels:=1\n'
                     's = s & "|" & ws.Rows(3).Hidden\nws.Rows("2:4").Ungroup\ns = s & "|" & ws.Rows(3).OutlineLevel',
    "page_setup": 'ws.PageSetup.Orientation = xlLandscape\nws.PageSetup.Zoom = 80\ns = ws.PageSetup.Orientation & "|" & ws.PageSetup.Zoom',
    "counts_whole": 's = ws.Range("A:A").Count & "|" & ws.Cells.CountLarge & "|" & ws.Range("A1:C3").Cells.Count & "|" & ws.Range("A1:C3").Rows.Count & "|" & '
                    'Union(ws.Range("A1:A2"), ws.Range("C1:C5")).Rows.Count & "|" & ws.Range("B2:D4").Columns(2).Address',
    "cells_column_letter": 's = ws.Cells(1, "C").Address & "|" & ws.Range("A1").Item(2, 3).Address & "|" & ws.Range("C3").Cells(0, 0).Address & "|" & '
                           'ws.Range("A1:B2").Cells(5).Address',
    "entire_row_hidden": 'ws.Rows(2).EntireRow.Hidden = True\nws.Columns("B").EntireColumn.Hidden = True\ns = ws.Range("A2").EntireRow.Hidden & "|" & '
                         'ws.Range("B1").EntireColumn.Hidden & "|" & ws.Range("A2").Height & "|" & ws.Range("B1").Width',
    "today_formula": 'ws.Range("A1").Formula = "=TODAY()"\ns = TypeName(ws.Range("A1").Value) & "|" & ws.Range("A1").NumberFormat & "|" & '
                     '(Int(ws.Range("A1").Value2) = CLng(Date))',
    "wf_transpose_dims": 'v = WorksheetFunction.Transpose(Array(1, 2, 3))\ns = LBound(v, 1) & UBound(v, 1) & LBound(v, 2) & UBound(v, 2) & "|" & '
                         'TypeName(Application.Max(ws.Range("A1:A3")))',
    "range_other_sheet_error": 'Dim w2 As Worksheet\nSet w2 = Worksheets.Add\nOn Error Resume Next\nSet r = ws.Range(w2.Cells(1, 1), w2.Cells(2, 2))\n'
                               's = Err.Number\nErr.Clear\nSet r = Union(ws.Range("A1"), w2.Range("A1"))\ns = s & "|" & Err.Number\nApplication.DisplayAlerts = False\n'
                               'w2.Delete\nApplication.DisplayAlerts = True',
    "with_block_nested": 'With ws.Range("A1")\n.Value = 3\nWith .Font\n.Italic = True\nEnd With\ns = .Value & "|" & .Font.Italic & "|" & .Offset(1).Address\nEnd With',
    "select_case": 'For i = 1 To 12\nSelect Case i\nCase 1, 3 To 5: s = s & "a"\nCase Is > 10: s = s & "b"\nCase Else: s = s & "c"\nEnd Select\nNext',
    "do_loops": 'i = 0\nDo Until i >= 3: i = i + 1: Loop\ns = i\nDo: i = i + 2: Loop While i < 10\ns = s & "|" & i\nWhile i > 5: i = i - 4: Wend\ns = s & "|" & i',
    "range_value_types": 'ws.Range("A1").Value = "123"\nws.Range("A2").Value = "'"'"'123"\nws.Range("A3").Value = "1/2"\nws.Range("A4").Value = "TRUE"\n'
                         'ws.Range("A5").Value = "=1+1"\nFor i = 1 To 5: s = s & TypeName(ws.Cells(i, 1).Value) & ":" & ws.Cells(i, 1).Text & ",": Next',
    "count_functions": 'ws.Range("A1").Value = 1: ws.Range("A2").Value = "x": ws.Range("A4").Value = True\n'
                       's = WorksheetFunction.Count(ws.Range("A1:A4")) & "|" & WorksheetFunction.CountA(ws.Range("A1:A4")) & "|" & '
                       'WorksheetFunction.CountBlank(ws.Range("A1:A4")) & "|" & Application.CountIf(ws.Range("A1:A4"), "x")',
    "range_text_numbers": 'ws.Range("A1").Value = 0.1 + 0.2\nws.Range("A2").Value = 1E+20\nws.Range("A3").Value = 123456789012\nws.Columns("A").ColumnWidth = 30\n'
                          's = ws.Range("A1").Text & "|" & ws.Range("A2").Text & "|" & ws.Range("A3").Text & "|" & ws.Range("A1").Value',
    "workbook_props": 's = (ActiveWorkbook Is ThisWorkbook) & "|" & ws.Parent.Sheets.Count = Worksheets.Count & "|" & TypeName(ws.Parent) & "|" & TypeName(Application.ActiveSheet)',
    "application_props": 'Application.ScreenUpdating = False\ns = Application.ScreenUpdating & "|" & Application.Calculation & "|" & Application.ReferenceStyle & "|" & '
                         'Application.DecimalSeparator & "|" & Application.International(xlListSeparator)\nApplication.ScreenUpdating = True',
    "error_values_in_cells": 'ws.Range("A1").Formula = "=1/0"\nws.Range("A2").Formula = "=NA()"\ns = IsError(ws.Range("A1").Value) & "|" & CStr(ws.Range("A1").Value) & "|" & '
                             'ws.Range("A2").Text & "|" & (ws.Range("A1").Value = CVErr(xlErrDiv0)) & "|" & WorksheetFunction.IsNA(ws.Range("A2"))',
    "string_number_coercion": 's = ("10" > "9") & "|" & (10 > "9") & "|" & ("a" & 1 + 2) & "|" & ("1" + 2) & "|" & ("1" & 2) & "|" & TypeName("1" + 2)',
    "array_formula_entry": 'ws.Range("A1:A3").Value = 2\nws.Range("B1").FormulaArray = "=SUM(A1:A3*2)"\ns = ws.Range("B1").Value & "|" & ws.Range("B1").HasArray & "|" & '
                           'ws.Range("B1").FormulaArray',
    "spill_formula2": 'ws.Range("A1").Formula2 = "=SEQUENCE(3)"\ns = ws.Range("A3").Value & "|" & ws.Range("A1").HasSpill & "|" & ws.Range("A1").SpillingToRange.Address',
    "resize_value_assign": 'Dim arr(1 To 2, 1 To 3) As Variant\narr(1, 1) = 1: arr(2, 3) = "z"\nws.Range("B2").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr\n'
                           's = ws.Range("D3").Value & "|" & IsEmpty(ws.Range("C2").Value) & "|" & ws.Range("B2").CurrentRegion.Address',
    "find_after_wrap": 'ws.Range("A1").Value = "k": ws.Range("A3").Value = "k"\nSet r = ws.Range("A1:A5").Find("k", After:=ws.Range("A3"))\ns = r.Address\n'
                       'Set r = ws.Range("A1:A5").Find("k", SearchDirection:=xlPrevious)\ns = s & "|" & r.Address',
    "insert_shift": 'ws.Range("A1:C1").Value = Array(1, 2, 3)\nws.Range("B1").Insert Shift:=xlToRight\ns = ws.Range("C1").Value & ws.Range("D1").Value & "|"\n'
                    'ws.Range("A1").Delete Shift:=xlToLeft\ns = s & ws.Range("A1").Value & ws.Range("B1").Value',
    "rows_columns_props": 's = ws.Range("B2:D5").Row & "|" & ws.Range("B2:D5").Column & "|" & ws.Range("B2:D5").Rows.Count & "|" & ws.Range("B2:D5").Columns.Count & "|" & '
                          'ws.Range("B2:D5").Rows(2).Address & "|" & ws.Range("B2:D5").EntireColumn.Address',
    "goal_text_to_columns": 'ws.Range("A1").Value = "a,b,c"\nws.Range("A1").TextToColumns DataType:=xlDelimited, Comma:=True\ns = ws.Range("C1").Value',
    "worksheet_function_text": 's = WorksheetFunction.Text(0.25, "0%") & "|" & WorksheetFunction.Text(#3/5/2024#, "mmm d") & "|" & WorksheetFunction.Proper("hello world") & "|" & '
                               'WorksheetFunction.Trim("  a   b ") & "|" & WorksheetFunction.Substitute("aaa", "a", "b", 2) & "|" & WorksheetFunction.Rept("ab", 2)',
}



def macro(name: str) -> str:
    cleanup = "\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "Function Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
