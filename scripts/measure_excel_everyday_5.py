"""A fifth sweep of everyday Excel macros: finding things, reports laid out and made ready to print, windows and panes, links
and notes, text and dates worked through in loops, and procedures calling each other.

Like scripts/measure_excel_everyday.py, each probe works on a sheet of its own that it deletes afterwards, and is recorded
with its answer or the run-time error it stopped with (see vba_probe_fixture.py).

    python scripts/measure_excel_everyday_5.py

writes tests/fixtures/excel_everyday_5.json, which tests/test_excel_everyday_5.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_5.json'
PRE = ("Dim ws As Worksheet, ws2 As Worksheet, r As Range, c As Range, s As String, i As Long, j As Long, n As Long\n"
       "Dim v As Variant, o As Object, a As Variant, t As String, d As Double, k As Variant\nSet ws = Worksheets.Add\n"
       "ws.Name = \"Sweep\"\n")
#: Staff: a header row and six rows of a name, a team, a start date, a salary and a note.
LIST = ('ws.Range("A1:E1").Value = Array("Name", "Team", "Start", "Salary", "Note")\n'
        'ws.Range("A2:E2").Value = Array("Ann Lee", "North", DateSerial(2020, 1, 15), 52000, "new")\n'
        'ws.Range("A3:E3").Value = Array("Bob Ray", "South", DateSerial(2019, 6, 1), 61000.5, "")\n'
        'ws.Range("A4:E4").Value = Array("cy Young", "North", DateSerial(2021, 3, 31), 47000, "part time")\n'
        'ws.Range("A5:E5").Value = Array("Di Fox", "East", DateSerial(2018, 12, 25), 75000, "lead")\n'
        'ws.Range("A6:E6").Value = Array("Ed Poe", "South", DateSerial(2022, 2, 28), 39000, "new")\n'
        'ws.Range("A7:E7").Value = Array("Flo Ng", "North", DateSerial(2020, 2, 29), 58000.25, "")\n')
CASES: dict[str, str] = {
    # Finding things.
    'find_loop': LIST + 'Set c = ws.Columns(2).Find("North", LookIn:=xlValues, LookAt:=xlWhole)\nIf Not c Is Nothing Then\nt = c.Address\nDo\ns = s & c.Row & ","\nSet c = ws.Columns(2).FindNext(c)\nLoop While c.Address <> t\nEnd If',
    'find_previous': LIST + 'Set c = ws.Columns(5).Find("new", After:=ws.Range("E1"), LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)\ns = c.Address\nSet c = ws.Columns(5).FindPrevious(c)\ns = s & c.Address',
    'find_partial_case': LIST + 'Set c = ws.UsedRange.Find("YOUNG", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)\ns = c.Address\nSet c = ws.UsedRange.Find("YOUNG", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=True)\ns = s & (c Is Nothing)',
    'find_by_rows_columns': LIST + 'Set c = ws.UsedRange.Find("North", LookIn:=xlValues, SearchOrder:=xlByColumns)\ns = c.Address\nSet c = ws.UsedRange.Find("new", LookIn:=xlValues, SearchOrder:=xlByRows)\ns = s & c.Address',
    'find_formula_value': 'ws.Range("A1").Formula = "=1+1"\nws.Range("A2").Value = 2\nSet c = ws.Cells.Find("2", LookIn:=xlValues, LookAt:=xlWhole)\ns = c.Address\nSet c = ws.Cells.Find("1+1", LookIn:=xlFormulas, LookAt:=xlPart)\ns = s & c.Address',
    'find_replace_count': LIST + 'Dim b As Boolean\nb = ws.Columns(2).Replace("North", "N", xlWhole, xlByRows, True)\ns = b & ws.Range("B2").Value & ws.Range("B3").Value & Application.CountIf(ws.Columns(2), "N")',
    'match_index': LIST + 'n = Application.Match("Di Fox", ws.Columns(1), 0)\ns = n & Application.Index(ws.Range("D1:D7"), n)\nv = Application.Match("Nobody", ws.Columns(1), 0)\ns = s & IsError(v) & CStr(v)',
    'countifs_sumifs': LIST + 's = Application.CountIfs(ws.Columns(2), "North", ws.Columns(4), ">50000") & "|" & Application.SumIfs(ws.Columns(4), ws.Columns(2), "South") & "|" & Application.AverageIf(ws.Columns(2), "North", ws.Columns(4))',
    'large_small_rank': LIST + 's = Application.Large(ws.Range("D2:D7"), 2) & "|" & Application.Small(ws.Range("D2:D7"), 1) & "|" & Application.Rank(47000, ws.Range("D2:D7")) & "|" & Application.Median(ws.Range("D2:D7"))',
    # Text worked through a loop.
    'proper_names': LIST + 'For Each c In ws.Range("A2:A7")\nc.Value = StrConv(c.Value, vbProperCase)\ns = s & Left(c.Value, 1)\nNext',
    'split_first_last': LIST + 'For i = 2 To 7\na = Split(ws.Cells(i, 1).Value, " ")\nws.Cells(i, 6).Value = a(1) & ", " & a(0)\ns = s & UBound(a) & Len(ws.Cells(i, 6).Value)\nNext\ns = s & ws.Range("F4").Value',
    'instr_mid': 't = "alpha,beta,,gamma"\nn = 1\nDo\ni = InStr(n, t, ",")\nIf i = 0 Then Exit Do\ns = s & Mid(t, n, i - n) & "/"\nn = i + 1\nLoop\ns = s & Mid(t, n) & InStrRev(t, ",") & Replace(t, ",", ";", 1, 2)',
    'like_patterns': 'For Each k In Array("A1", "a1", "AB", "12", "X-9", "")\ns = s & (k Like "[A-Z]#") & (k Like "*#") & (k Like "?-?") & "|"\nNext',
    'format_function': 's = Format(1234.567, "#,##0.00") & "|" & Format(0.256, "0.0%") & "|" & Format(DateSerial(2024, 3, 5), "dddd d mmm yyyy") & "|" & Format(-5, "0;(0)") & "|" & Format("abc", ">") & "|" & Format(5, "000") & "|" & Format(1.5, "Currency")',
    'trim_clean': 'ws.Range("A1").Value = "  two  spaces  "\ns = "[" & Trim(ws.Range("A1").Value) & "][" & Application.Trim(ws.Range("A1").Value) & "][" & Application.Clean("a" & Chr(7) & "b") & "]" & Application.Substitute("aXbXc", "X", "-", 2) & Application.Rept("ab", 2)',
    'number_text_conversion': 's = Val("12abc") & "|" & CDbl("1.5") & "|" & CStr(1.50) & "|" & Str(7) & "|" & CInt(2.5) & CInt(3.5) & "|" & Round(2.567, 1) & "|" & Int(-1.5) & Fix(-1.5) & "|" & Hex(255) & Oct(8) & "|" & IsNumeric("1e3") & IsNumeric("x")',
    'text_numbers_in_cells': 'ws.Range("A1").NumberFormat = "@"\nws.Range("A1").Value = 123\nws.Range("A2").Value = "\'456"\nws.Range("A3").Value = "789"\ns = TypeName(ws.Range("A1").Value) & TypeName(ws.Range("A2").Value) & TypeName(ws.Range("A3").Value) & ws.Range("A2").PrefixCharacter & Application.Sum(ws.Range("A1:A3")) & Application.Count(ws.Range("A1:A3")) & Application.CountA(ws.Range("A1:A3"))',
    # Dates.
    'date_steps': 'd = DateSerial(2024, 1, 31)\ns = Format(DateAdd("m", 1, d), "yyyy-mm-dd") & "|" & Format(DateAdd("q", 1, d), "yyyy-mm-dd") & "|" & DateDiff("d", d, DateSerial(2024, 3, 1)) & "|" & DateDiff("m", d, DateSerial(2024, 3, 1)) & "|" & Weekday(d) & "|" & Weekday(d, vbMonday) & "|" & DatePart("q", d) & "|" & MonthName(2) & "|" & WeekdayName(1)',
    'date_in_cells': LIST + 's = TypeName(ws.Range("C2").Value) & TypeName(ws.Range("C2").Value2) & ws.Range("C2").Value2 & "|" & ws.Range("C2").Text & "|" & ws.Range("C2").NumberFormat\nws.Range("C2").NumberFormat = "yyyy-mm"\ns = s & ws.Range("C2").Text & Year(ws.Range("C7").Value) & Month(ws.Range("C7").Value) & Day(ws.Range("C7").Value)',
    'workday_eomonth': 's = CDbl(Application.EoMonth(DateSerial(2024, 2, 10), 0)) & "|" & CDbl(Application.WorkDay(DateSerial(2024, 3, 1), 5)) & "|" & Application.NetworkDays(DateSerial(2024, 3, 1), DateSerial(2024, 3, 31)) & "|" & Application.Text(45000, "yyyy-mm-dd") & "|" & Application.Weeknum(DateSerial(2024, 3, 1))',
    'date_filter_loop': LIST + 'For i = 2 To 7\nIf ws.Cells(i, 3).Value >= DateSerial(2020, 1, 1) And ws.Cells(i, 3).Value < DateSerial(2021, 1, 1) Then s = s & i\nNext\nws.Range("A1:E7").AutoFilter 3, ">=" & CLng(DateSerial(2020, 1, 1))\ns = s & "|" & ws.Range("A2:A7").SpecialCells(xlCellTypeVisible).Count\nws.AutoFilterMode = False',
    'time_values': 'ws.Columns(1).ColumnWidth = 20\nws.Range("A1").Value = TimeSerial(13, 30, 0)\nws.Range("A2").Formula = "=A1+TIME(12,0,0)"\ns = ws.Range("A1").Text & "|" & ws.Range("A1").Value2 & "|" & ws.Range("A2").Value & "|" & Hour(ws.Range("A1").Value) & Minute(ws.Range("A1").Value) & "|" & Format(ws.Range("A2").Value, "[h]:mm")',
    # A report laid out.
    'report_header': LIST + 'ws.Rows(1).Insert\nws.Range("A1").Value = "Staff"\nWith ws.Range("A1:E1")\n.HorizontalAlignment = xlCenterAcrossSelection\n.Font.Size = 16\n.Interior.ColorIndex = 15\nEnd With\ns = ws.Range("A2").Value & ws.Range("C1").HorizontalAlignment & ws.Range("A1").Interior.Color & ws.UsedRange.Address',
    'borders_loop': LIST + 'For Each k In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)\nWith ws.Range("A1:E7").Borders(k)\n.LineStyle = xlContinuous\n.Weight = xlThin\n.ColorIndex = 3\nEnd With\nNext\ns = ws.Range("C3").Borders(xlEdgeTop).LineStyle & ws.Range("C3").Borders(xlEdgeTop).Color & ws.Range("A1:E7").Borders.Count\nws.Range("A1:E7").Borders.LineStyle = xlNone\ns = s & ws.Range("C3").Borders(xlEdgeTop).LineStyle',
    'subtotal_totals': LIST + 'ws.Range("D8").Formula = "=SUBTOTAL(9,D2:D7)"\nws.Range("A8").Value = "Total"\nws.Rows(8).Font.Bold = True\ns = ws.Range("D8").Value & ws.Range("E8").Font.Bold\nws.Rows(3).Hidden = True\nws.Calculate\ns = s & "|" & ws.Range("D8").Value & "|" & Application.Subtotal(109, ws.Range("D2:D7"))',
    'percent_of_total': LIST + 'For i = 2 To 7\nws.Cells(i, 6).Formula = "=D" & i & "/SUM($D$2:$D$7)"\nNext\nws.Range("F2:F7").NumberFormat = "0.0%"\ns = ws.Range("F5").Text & ws.Range("F5").Formula & ws.Range("F5").FormulaR1C1 & ws.Range("F5").NumberFormat',
    'r1c1_block': LIST + 'ws.Range("F2:F7").FormulaR1C1 = "=RC[-2]*1.1"\nws.Range("G2:G7").Formula = "=F2-D2"\ns = ws.Range("F7").Formula & ws.Range("G7").FormulaR1C1 & Round(ws.Range("G7").Value, 2) & ws.Range("F2").HasFormula & ws.Range("F2:G7").HasFormula',
    'autofit_widths': LIST + 'ws.Columns("A:E").ColumnWidth = 5\ns = ws.Columns(1).ColumnWidth & ws.Range("A:E").ColumnWidth\nws.Columns(2).ColumnWidth = 20\ns = s & "|" & IsNull(ws.Range("A:B").ColumnWidth)\nws.Rows(1).RowHeight = 30\ns = s & ws.Range("A1").RowHeight & ws.Rows(1).Height',
    'copy_formats_only': LIST + 'ws.Range("A1:E1").Font.Bold = True\nws.Range("A1:E1").Interior.Color = RGB(200, 200, 0)\nws.Range("A1:E1").Copy\nws.Range("A10").PasteSpecial xlPasteFormats\nApplication.CutCopyMode = False\ns = ws.Range("C10").Font.Bold & ws.Range("C10").Interior.Color & "[" & ws.Range("C10").Value & "]" & Application.CutCopyMode',
    'copy_destination_widths': LIST + 'Set ws2 = Worksheets.Add\nws.Columns(1).ColumnWidth = 22\nws.Range("A1:E7").Copy ws2.Range("B2")\nws2.Range("B2").PasteSpecial xlPasteColumnWidths\ns = ws2.Range("B3").Value & ws2.Columns(2).ColumnWidth & ws2.UsedRange.Address\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'transpose_paste': LIST + 'ws.Range("A1:B3").Copy\nws.Range("H1").PasteSpecial Transpose:=True\nApplication.CutCopyMode = False\ns = ws.Range("I1").Value & ws.Range("J2").Value\na = Application.Transpose(ws.Range("A2:A4").Value)\ns = s & LBound(a) & UBound(a) & a(2)',
    'paste_add_values': 'ws.Range("A1:A3").Value = 2\nws.Range("B1").Value = 10\nws.Range("B1").Copy\nws.Range("A1:A3").PasteSpecial xlPasteValues, xlPasteSpecialOperationAdd\nApplication.CutCopyMode = False\ns = ws.Range("A2").Value\nws.Range("A1:A3").Value = ws.Range("A1:A3").Value\ns = s & ws.Range("A3").HasFormula',
    # Ready to print.
    'page_setup': LIST + 'With ws.PageSetup\n.Orientation = xlLandscape\n.Zoom = False\n.FitToPagesWide = 1\n.FitToPagesTall = False\n.PrintTitleRows = "$1:$1"\n.PrintArea = "$A$1:$E$7"\n.CenterFooter = "Page &P of &N"\n.LeftMargin = Application.InchesToPoints(0.5)\n.PaperSize = xlPaperA4\ns = .Orientation & .Zoom & .FitToPagesWide & .FitToPagesTall & .PrintTitleRows & .PrintArea & .CenterFooter & .LeftMargin & .PaperSize\nEnd With',
    'page_breaks': LIST + 'ws.HPageBreaks.Add ws.Range("A4")\nws.VPageBreaks.Add ws.Range("C1")\ns = ws.HPageBreaks.Count & ws.VPageBreaks.Count & ws.HPageBreaks(1).Location.Address\nws.ResetAllPageBreaks\ns = s & ws.HPageBreaks.Count',
    # Windows.
    'freeze_panes': LIST + 'ws.Activate\nws.Range("B2").Select\nActiveWindow.FreezePanes = True\ns = ActiveWindow.FreezePanes & ActiveWindow.SplitRow & ActiveWindow.SplitColumn\nActiveWindow.FreezePanes = False\ns = s & ActiveWindow.FreezePanes',
    'window_settings': 'ws.Activate\nActiveWindow.DisplayGridlines = False\nActiveWindow.DisplayHeadings = False\nActiveWindow.Zoom = 85\ns = ActiveWindow.DisplayGridlines & ActiveWindow.DisplayHeadings & ActiveWindow.Zoom\nActiveWindow.DisplayGridlines = True\nActiveWindow.DisplayHeadings = True\nActiveWindow.Zoom = 100\nActiveWindow.ScrollRow = 5\nActiveWindow.ScrollColumn = 3\ns = s & ActiveWindow.ScrollRow & ActiveWindow.ScrollColumn',
    'select_active': LIST + 'ws.Activate\nws.Range("B2:C3").Select\ns = Selection.Address & ActiveCell.Address & Selection.Count\nws.Range("C3").Activate\ns = s & ActiveCell.Address & Selection.Address\nSelection.Offset(1, 0).Select\ns = s & Selection.Address\nApplication.Goto ws.Range("D5")\ns = s & ActiveCell.Address',
    # Links, notes and names.
    'hyperlinks_cell': 'ws.Hyperlinks.Add ws.Range("A1"), "https://example.com", , "Tip", "Site"\nws.Hyperlinks.Add ws.Range("A2"), "", "Sweep!C3", , "Jump"\ns = ws.Hyperlinks.Count & ws.Range("A1").Value & ws.Range("A1").Hyperlinks(1).Address & ws.Range("A1").Hyperlinks(1).ScreenTip & ws.Range("A2").Hyperlinks(1).SubAddress\nws.Range("A1").Hyperlinks.Delete\ns = s & ws.Hyperlinks.Count & ws.Range("A1").Value',
    'hyperlink_formula': 'ws.Range("A1").Formula = "=HYPERLINK(""https://example.com"",""go"")"\ns = ws.Range("A1").Value & ws.Hyperlinks.Count & ws.Range("A1").HasFormula',
    'notes_loop': LIST + 'For i = 2 To 4\nws.Cells(i, 1).AddComment "Row " & i\nNext\nFor Each o In ws.Comments\ns = s & o.Parent.Address & o.Text & o.Visible & "|"\nNext\nws.Range("A3").ClearComments\ns = s & ws.Comments.Count & (ws.Range("A3").Comment Is Nothing)\nws.Range("A2").Comment.Text "Changed"\ns = s & ws.Range("A2").Comment.Text',
    'names_scopes': LIST + 'ThisWorkbook.Names.Add "Pay", "=Sweep!$D$2:$D$7"\nws.Names.Add "Local", "=Sweep!$A$2:$A$7"\ns = Application.Sum(Range("Pay")) & ws.Range("Local").Count & ThisWorkbook.Names("Pay").RefersTo & ws.Names("Local").Name\nws.Range("G1").Formula = "=SUM(Pay)"\ns = s & ws.Range("G1").Value\nThisWorkbook.Names("Pay").Delete\nws.Calculate\ns = s & CStr(ws.Range("G1").Value)\nws.Names("Local").Delete',
    'names_constant': 'ThisWorkbook.Names.Add "Rate", "=0.2"\nws.Range("A1").Formula = "=100*Rate"\ns = ws.Range("A1").Value & ThisWorkbook.Names("Rate").RefersTo & Evaluate("Rate")\nThisWorkbook.Names("Rate").Delete',
    'offset_name_dynamic': LIST + 'ThisWorkbook.Names.Add "Dyn", "=OFFSET(Sweep!$A$1,0,0,COUNTA(Sweep!$A:$A),1)"\ns = Range("Dyn").Address & Application.CountA(Range("Dyn"))\nws.Range("A8").Value = "New"\ns = s & Range("Dyn").Address\nThisWorkbook.Names("Dyn").Delete',
    # Procedures and objects.
    'arrays_2d': 'ReDim a(1 To 3, 1 To 2)\nFor i = 1 To 3\na(i, 1) = i\na(i, 2) = i * i\nNext\nws.Range("A1").Resize(3, 2).Value = a\ns = ws.Range("B3").Value & LBound(a, 2) & UBound(a, 2)\nv = ws.Range("A1:B3").Value\ns = s & UBound(v, 1) & v(2, 2)\nReDim Preserve a(1 To 3, 1 To 3)\ns = s & UBound(a, 2) & a(3, 2) & IsEmpty(a(3, 3))',
    'array_function_cells': 'ws.Range("A1:C1").Value = Array(1, 2, 3)\nws.Range("A2:A4").Value = Application.Transpose(Array(4, 5, 6))\ns = ws.Range("C1").Value & ws.Range("A4").Value\nws.Range("A6:B7").FormulaArray = "=A1:B2*2"\ns = s & ws.Range("B7").Value & ws.Range("A6").HasArray & ws.Range("A6").CurrentArray.Address\nws.Range("A6:B7").ClearContents\ns = s & ws.Range("A6").HasArray',
    'dictionary_nested': LIST + 'Dim dd As Object\nSet dd = CreateObject("Scripting.Dictionary")\nFor i = 2 To 7\nIf Not dd.Exists(ws.Cells(i, 2).Value) Then dd.Add ws.Cells(i, 2).Value, New Collection\ndd(ws.Cells(i, 2).Value).Add ws.Cells(i, 1).Value\nNext\nFor Each k In dd.Keys\ns = s & k & dd(k).Count & dd(k)(1) & "|"\nNext\ns = s & dd.Count',
    'collection_keys': 'Dim cc As New Collection\ncc.Add "a", "x"\ncc.Add "b", "y"\ncc.Add "c", Before:=1\ns = cc(1) & cc("y") & cc.Count\ncc.Remove "x"\ns = s & cc.Count\nOn Error Resume Next\ncc.Add "z", "y"\ns = s & Err.Number\nErr.Clear\nv = cc("nope")\ns = s & Err.Number',
    'range_as_variant': LIST + 'Set r = ws.Range("D2:D7")\nFor Each c In r.Cells\nIf c.Value > 50000 Then n = n + 1\nNext\ns = n & r.Rows.Count & r.Columns.Count & r.Cells(2).Address & r.Item(3).Address & r(4).Address & r.Cells(1, 1).Address & r.Offset(1).Address & r.Resize(2).Address',
    'entire_row_column': LIST + 's = ws.Range("B3").EntireRow.Address & ws.Range("B3").EntireColumn.Address & ws.Range("B3:C4").Rows.Count & ws.Rows("2:4").Count & ws.Columns("B:C").Count & ws.Range("B3").End(xlToRight).Address & ws.Range("B3").End(xlDown).Address & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row',
    'error_goto_resume': 'On Error GoTo Oops\nn = 1 / 0\ns = s & "after"\nGoTo Done\nOops:\ns = s & Err.Number & Err.Description & Err.Source\nResume Next\nDone:',
    'err_raise_custom': 'On Error Resume Next\nErr.Raise 513, "Mine", "Bad thing"\ns = Err.Number & Err.Source & Err.Description\nErr.Clear\ns = s & Err.Number\nv = Worksheets("Nope").Name\ns = s & Err.Number & Err.Description',
    'select_case_like': 'For Each k In Array(5, 15, "x", 0.5, Empty)\nSelect Case k\nCase Is < 1: s = s & "a"\nCase 1 To 10: s = s & "b"\nCase 11, 15: s = s & "c"\nCase "x": s = s & "d"\nCase Else: s = s & "e"\nEnd Select\nNext',
    'with_nested_loops': LIST + 'With ws\nFor i = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row\nFor j = 4 To 4\nWith .Cells(i, j)\nIf .Value > 60000 Then .Font.Color = vbRed: s = s & .Row\nEnd With\nNext j\nNext i\nEnd With\ns = s & ws.Range("D5").Font.Color & ws.Range("D2").Font.ColorIndex',
    'do_until_exit': LIST + 'i = 2\nDo Until ws.Cells(i, 1).Value = ""\nIf ws.Cells(i, 2).Value = "East" Then Exit Do\ni = i + 1\nLoop\ns = i\nWhile i > 3\ni = i - 1\nWend\ns = s & i\nFor i = 10 To 1 Step -3\ns = s & i\nNext\ns = s & i',
    'sheets_index_order': 'Set ws2 = Worksheets.Add(After:=ws)\nws2.Name = "Sweep2"\ns = ws.Index & ws2.Index - ws.Index & ws2.Previous.Name & ws.Next.Name\nws2.Move Before:=ws\ns = s & ws2.Next.Name & Sheets("Sweep2").Visible\nws2.Visible = xlSheetHidden\ns = s & ws2.Visible\nws2.Visible = xlSheetVisible\nws2.Tab.Color = RGB(1, 2, 3)\ns = s & ws2.Tab.Color\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'workbook_in_memory': 'Dim w As Workbook\nSet w = Workbooks.Add\nw.Worksheets(1).Range("A1").Value = 5\ns = (w.Name <> ThisWorkbook.Name) & w.Worksheets(1).Range("A1").Value & w.Saved & (ActiveWorkbook.Name = w.Name) & w.Sheets.Count & Workbooks.Count - Workbooks.Count\nw.Close False\ns = s & (ActiveWorkbook Is Nothing)',
    'evaluate_shortcuts': LIST + 's = Evaluate("SUM(D2:D7)") & "|" & ws.Evaluate("COUNTA(A:A)") & "|" & [1+2] & "|" & Evaluate("A2") & "|" & TypeName(Evaluate("A2:A3")) & "|" & TypeName(ws.Evaluate("{1,2,3}")) & UBound(ws.Evaluate("{1,2,3}"))',
    'calculation_modes': 'Application.Calculation = xlCalculationManual\nws.Range("A1").Value = 1\nws.Range("A2").Formula = "=A1+1"\nws.Range("A1").Value = 5\ns = ws.Range("A2").Value\nws.Calculate\ns = s & ws.Range("A2").Value\nws.Range("A1").Value = 7\nApplication.Calculate\ns = s & ws.Range("A2").Value\nApplication.Calculation = xlCalculationAutomatic\ns = s & Application.Calculation',
    'cell_properties': LIST + 'With ws.Range("D3")\ns = .Text & "|" & .Value & "|" & .Value2 & "|" & .Formula & "|" & .HasFormula & "|" & .Address(False, False) & "|" & .Address(True, False) & "|" & .Address(ReferenceStyle:=xlR1C1) & "|" & .Row & .Column & "|" & .Parent.Name & "|" & .Worksheet.Name & .NumberFormatLocal\nEnd With',
    'clear_vs_delete': LIST + 'ws.Range("A3:E3").Font.Bold = True\nws.Range("A3:E3").ClearContents\ns = ws.Range("B3").Font.Bold & "[" & ws.Range("B3").Value & "]" & ws.Range("B4").Value\nws.Range("A3:E3").Delete xlShiftUp\ns = s & ws.Range("B3").Value & ws.Range("B3").Font.Bold\nws.Range("B2:B3").Insert xlShiftToRight\ns = s & ws.Range("C2").Value & ws.UsedRange.Address',
    'conditional_loop_colors': LIST + 'For i = 2 To 7\nSelect Case ws.Cells(i, 2).Value\nCase "North": ws.Cells(i, 2).Interior.ColorIndex = 4\nCase "South": ws.Cells(i, 2).Interior.Color = vbYellow\nCase Else: ws.Cells(i, 2).Interior.Pattern = xlGray25\nEnd Select\nNext\ns = ws.Range("B2").Interior.Color & "|" & ws.Range("B3").Interior.ColorIndex & "|" & ws.Range("B5").Interior.Pattern & ws.Range("B5").Interior.ColorIndex',
    'validate_list_and_check': LIST + 'With ws.Range("G2:G7").Validation\n.Delete\n.Add xlValidateList, xlValidAlertStop, xlBetween, "North,South,East"\n.IgnoreBlank = True\n.InCellDropdown = True\n.InputTitle = "Team"\n.ErrorMessage = "Pick one"\ns = .Type & .Formula1 & .AlertStyle & .InputTitle & .ErrorMessage\nEnd With\nws.Range("G2").Value = "West"\ns = s & ws.Range("G2").Validation.Value\nws.Range("G2").Value = "East"\ns = s & ws.Range("G2").Validation.Value\ns = s & ws.Range("A1:A7").SpecialCells(xlCellTypeConstants).Count',
}


def macro(name: str) -> str:
    cleanup = "\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "Function Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
