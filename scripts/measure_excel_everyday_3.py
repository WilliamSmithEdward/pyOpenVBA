"""A third sweep of everyday Excel macros: newer worksheet functions, tables, shapes, charts, pivots, workbooks,
the Sort object, AutoFilter, hyperlinks, notes, protection, conditional formats, outlines, page setup, number
formats and VBA's Format, Split and arrays.

Like scripts/measure_excel_everyday.py, each probe works on a sheet of its
own that it deletes afterwards, and is recorded as its answer or the
run-time error it stopped with (see vba_probe_fixture.py). Macros that make
workbooks of their own close them.

    python scripts/measure_excel_everyday_3.py

writes tests/fixtures/excel_everyday_3.json, which tests/test_excel_everyday_3.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_3.json'
#: Each probe's sheet is named, so what names it in a reference reads the same in Excel and the model.
PRE = ("Dim ws As Worksheet, r As Range, c As Range, s As String, i As Long, j As Long, n As Long, v As Variant\n"
       "Dim o As Object, w As Workbook, t As String, a As Variant\nSet ws = Worksheets.Add\nws.Name = \"Sweep\"\n")
#: A small list most probes start from: a header row and five rows of a name, a region and an amount.
LIST = ('ws.Range("A1:C1").Value = Array("Name", "Region", "Amount")\n'
        'ws.Range("A2:C2").Value = Array("ann", "East", 10)\nws.Range("A3:C3").Value = Array("bob", "West", 25)\n'
        'ws.Range("A4:C4").Value = Array("cy", "East", 5)\nws.Range("A5:C5").Value = Array("dee", "North", 40)\n'
        'ws.Range("A6:C6").Value = Array("ed", "West", 15)\n')

CASES: dict[str, str] = {
    # --- worksheet functions, newer and older
    'wf_sumifs_countifs': LIST + 's = WorksheetFunction.SumIfs(ws.Range("C2:C6"), ws.Range("B2:B6"), "West", '
                                 'ws.Range("C2:C6"), ">12") & "|" & WorksheetFunction.CountIfs(ws.Range("B2:B6"), "*t", '
                                 'ws.Range("C2:C6"), "<=15") & "|" & WorksheetFunction.AverageIfs(ws.Range("C2:C6"), '
                                 'ws.Range("B2:B6"), "<>East")',
    'wf_xlookup': LIST + 's = WorksheetFunction.XLookup("cy", ws.Range("A2:A6"), ws.Range("C2:C6")) & "|" & '
                         'WorksheetFunction.XLookup("zz", ws.Range("A2:A6"), ws.Range("C2:C6"), "none") & "|" & '
                         'WorksheetFunction.XLookup(20, ws.Range("C2:C6"), ws.Range("A2:A6"), , -1) & "|" & '
                         'WorksheetFunction.XLookup("West", ws.Range("B2:B6"), ws.Range("A2:A6"), , 0, -1)',
    'wf_textjoin_concat': LIST + 's = WorksheetFunction.TextJoin(",", True, ws.Range("A2:A4"), "", "z") & "|" & '
                                 'WorksheetFunction.Concat(ws.Range("B2:B3")) & "|" & '
                                 'WorksheetFunction.TextJoin("-", False, "a", "", "b")',
    'wf_unique_sort_filter': LIST + 'v = WorksheetFunction.Unique(ws.Range("B2:B6"))\n'
                                    's = UBound(v, 1) & "|" & v(3, 1) & "|"\n'
                                    'v = WorksheetFunction.Sort(ws.Range("C2:C6"), , -1)\ns = s & v(1, 1) & "," & v(5, 1) & "|"\n'
                                    'v = ws.Evaluate("FILTER(A2:A6,C2:C6>12)")\ns = s & UBound(v, 1) & "," & v(1, 1)',
    'wf_sequence': 'v = WorksheetFunction.Sequence(3, 2, 10, 5)\ns = UBound(v, 1) & "x" & UBound(v, 2) & "|" & v(1, 2) & '
                   '"|" & v(3, 2) & "|" & TypeName(v(1, 1))',
    'wf_iferror_ifs': 's = WorksheetFunction.IfError(ws.Evaluate("1/0"), "bad") & "|" & ws.Evaluate("IFS(2>3,""a"",2<3,""b"")") '
                      '& "|" & WorksheetFunction.IfNa(ws.Evaluate("NA()"), 0)',
    'wf_dates_workdays': 's = WorksheetFunction.EoMonth(DateSerial(2024, 1, 31), 1) & "|" & '
                         'WorksheetFunction.NetworkDays(DateSerial(2024, 3, 1), DateSerial(2024, 3, 31)) & "|" & '
                         'Format(WorksheetFunction.WorkDay(DateSerial(2024, 12, 23), 3), "yyyy-mm-dd") & "|" & '
                         'WorksheetFunction.WeekNum(DateSerial(2024, 1, 7), 2) & "|" & '
                         'WorksheetFunction.IsoWeekNum(DateSerial(2021, 1, 3)) & "|" & '
                         'WorksheetFunction.Days(DateSerial(2024, 3, 1), DateSerial(2024, 2, 1))',
    'wf_ranking_stats': LIST + 's = WorksheetFunction.Large(ws.Range("C2:C6"), 2) & "|" & WorksheetFunction.Small(ws.Range("C2:C6"), 1) '
                               '& "|" & WorksheetFunction.Rank_Eq(25, ws.Range("C2:C6")) & "|" & '
                               'WorksheetFunction.Percentile_Inc(ws.Range("C2:C6"), 0.5) & "|" & '
                               'WorksheetFunction.Quartile_Inc(ws.Range("C2:C6"), 1) & "|" & '
                               'Round(WorksheetFunction.StDev_S(ws.Range("C2:C6")), 6) & "|" & WorksheetFunction.Median(ws.Range("C2:C6"))',
    'wf_text_functions': 's = WorksheetFunction.Proper("hello wORLD") & "|" & WorksheetFunction.Substitute("a-b-c", "-", "+", 2) & '
                         '"|" & WorksheetFunction.Rept("ab", 3) & "|" & WorksheetFunction.Trim("  a   b  ") & "|" & '
                         'WorksheetFunction.Find("b", "abcb", 3) & "|" & WorksheetFunction.Search("B", "abcb") & "|" & '
                         'Len(WorksheetFunction.Clean("a" & Chr(7) & "b"))',
    'wf_financial': 's = Round(WorksheetFunction.Pmt(0.05 / 12, 360, 200000), 4) & "|" & Round(WorksheetFunction.Fv(0.04, 10, -1000), 4) '
                    '& "|" & Round(WorksheetFunction.Npv(0.1, 100, 200, 300), 4) & "|" & '
                    'Round(WorksheetFunction.Irr(Array(-1000, 300, 400, 500)), 6) & "|" & '
                    'Round(WorksheetFunction.Rate(36, -300, 10000), 8)',
    'wf_rounding_more': 's = WorksheetFunction.Ceiling_Math(-4.5) & "|" & WorksheetFunction.Floor_Math(4.5, 2) & "|" & '
                        'WorksheetFunction.Even(3.1) & "|" & WorksheetFunction.Odd(-2) & "|" & '
                        'WorksheetFunction.Ceiling(7, 3) & "|" & ws.Evaluate("TRUNC(-3.99)")',
    'wf_hlookup_lookup': 'ws.Range("A1:C2").Value = Array(1, 2, 3)\nws.Range("A2:C2").Value = Array("x", "y", "z")\n'
                         's = WorksheetFunction.HLookup(2, ws.Range("A1:C2"), 2, False) & "|" & '
                         'WorksheetFunction.Lookup(2.5, ws.Range("A1:C1"), ws.Range("A2:C2"))',
    'wf_counting_blanks': 'ws.Range("A1").Value = 1: ws.Range("A3").Formula = "=""""": ws.Range("A4").Value = "x"\n'
                          's = WorksheetFunction.CountBlank(ws.Range("A1:A5")) & "|" & WorksheetFunction.CountA(ws.Range("A1:A5")) '
                          '& "|" & WorksheetFunction.Count(ws.Range("A1:A5"))',
    'evaluate_arrays': LIST + 's = ws.Evaluate("SUM(IF(C2:C6>12,C2:C6))") & "|"\nv = ws.Evaluate("ROW(1:3)")\n'
                              's = s & UBound(v, 1) & "," & v(3, 1) & "|" & ws.Evaluate("SUMPRODUCT(--(B2:B6=""East""))")',
    # --- tables
    'listrows_add': LIST + 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C6"), , xlYes)\n'
                           'o.ListRows.Add.Range.Value = Array("fay", "South", 7)\n'
                           's = o.DataBodyRange.Address & "|" & ws.Range("A7").Value & "|" & o.ListRows.Count\n'
                           'o.ListRows(1).Delete\ns = s & "|" & o.DataBodyRange.Address & "|" & ws.Range("A2").Value',
    'listobject_sort': LIST + 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C6"), , xlYes)\n'
                              'o.Sort.SortFields.Clear\no.Sort.SortFields.Add Key:=o.ListColumns("Amount").Range, '
                              'SortOn:=xlSortOnValues, Order:=xlDescending\no.Sort.Header = xlYes\no.Sort.Apply\n'
                              's = ws.Range("A2").Value & ws.Range("A3").Value & ws.Range("A6").Value',
    'listcolumn_formula': LIST + 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C6"), , xlYes)\n'
                                 'Set c = o.ListColumns.Add.Range\no.HeaderRowRange.Cells(1, 4).Value = "Double"\n'
                                 'o.ListColumns(4).DataBodyRange.Formula = "=[@Amount]*2"\n'
                                 's = ws.Range("D3").Value & "|" & ws.Range("D3").Formula & "|" & o.ListColumns.Count',
    'listobject_filter': LIST + 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C6"), , xlYes)\n'
                                'o.Range.AutoFilter Field:=3, Criteria1:=">12"\n'
                                's = o.Range.SpecialCells(xlCellTypeVisible).Rows.Count & "|" & ws.Rows(2).Hidden & "|" & '
                                'o.AutoFilter.Filters(3).Criteria1',
    'listobject_resize_total': LIST + 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B6"), , xlYes)\n'
                                      'o.Resize ws.Range("A1:C6")\no.ShowTotals = True\n'
                                      's = o.Range.Address & "|" & o.ListColumns.Count & "|" & o.TotalsRowRange.Address',
    # --- shapes, charts and pivots
    'shape_rectangle_text': 'Set o = ws.Shapes.AddShape(msoShapeRectangle, 10, 20, 100, 50)\no.Name = "Box1"\n'
                            'o.TextFrame.Characters.Text = "Hello"\ns = ws.Shapes.Count & "|" & ws.Shapes("Box1").TextFrame.Characters.Text '
                            '& "|" & o.Left & "," & o.Top & "," & o.Width & "," & o.Height & "|" & o.Type',
    'shape_textbox_delete': 'Set o = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 80, 30)\n'
                            'o.TextFrame2.TextRange.Text = "note"\ns = o.TextFrame2.TextRange.Text & "|" & ws.Shapes.Count & "|"\n'
                            'o.Delete\ns = s & ws.Shapes.Count',
    'shape_top_left_cell': 'Set o = ws.Shapes.AddShape(msoShapeOval, 100, 40, 50, 50)\n'
                           's = o.TopLeftCell.Address & "|" & o.BottomRightCell.Address & "|" & o.AutoShapeType',
    'chart_type_title': 'ws.Range("A1:B4").Value = 3\nSet o = ws.ChartObjects.Add(10, 10, 300, 200)\n'
                        'o.Chart.SetSourceData ws.Range("A1:B4")\no.Chart.ChartType = xlLine\no.Chart.HasTitle = True\n'
                        'o.Chart.ChartTitle.Text = "Trend"\ns = o.Chart.ChartType & "|" & o.Chart.ChartTitle.Text & "|" & (o.Name Like "Chart*")',
    'chart_new_series': 'ws.Range("A1:A3").Value = WorksheetFunction.Transpose(Array(1, 2, 3))\n'
                        'Set o = ws.ChartObjects.Add(0, 0, 200, 150)\nWith o.Chart.SeriesCollection.NewSeries\n'
                        '.Values = ws.Range("A1:A3")\n.Name = "First"\nEnd With\n'
                        's = o.Chart.SeriesCollection.Count & "|" & o.Chart.SeriesCollection(1).Name',
    'pivot_create_sum': LIST + 'Set o = ThisWorkbook.PivotCaches.Create(xlDatabase, ws.Range("A1:C6")).CreatePivotTable(ws.Range("F1"), '
                               '"SweepPivot")\no.PivotFields("Region").Orientation = xlRowField\n'
                               'o.AddDataField o.PivotFields("Amount"), "Total", xlSum\n'
                               's = o.GetPivotData("Total", "Region", "West") & "|" & o.TableRange1.Address & "|" & '
                               'o.RowFields.Count',
    # --- workbooks and sheets
    'workbooks_add_close': 'n = Workbooks.Count\nSet w = Workbooks.Add\ns = (Workbooks.Count = n + 1) & "|" & w.Worksheets.Count & "|" & '
                           '(w.Name Like "Book*")\nw.Close SaveChanges:=False\ns = s & "|" & (Workbooks.Count = n)',
    'sheet_copy_new_book': 'ws.Range("A1").Value = "moved"\nn = Workbooks.Count\nws.Copy\nSet w = ActiveWorkbook\n'
                           's = (Workbooks.Count = n + 1) & "|" & w.Worksheets.Count & "|" & w.Worksheets(1).Name & "|" & '
                           'w.Worksheets(1).Range("A1").Value\nw.Close SaveChanges:=False\nws.Activate',
    'copy_between_books': 'Set w = Workbooks.Add\nw.Worksheets(1).Range("A1:B2").Value = 4\n'
                          'w.Worksheets(1).Range("A1:B2").Copy ws.Range("C3")\ns = ws.Range("D4").Value & "|" & '
                          'WorksheetFunction.Sum(ws.Range("C3:D4"))\nw.Close SaveChanges:=False',
    # Both workbooks have a sheet named Sweep, which a function has to keep apart.
    'functions_of_another_book': 'ws.Range("A1").Value = 100\nSet w = Workbooks.Add\nw.Worksheets(1).Name = "Sweep"\n'
                                 'w.Worksheets(1).Range("A1:B2").Value = 4\nw.Worksheets(1).Range("C1").Value = "x"\n'
                                 'ws.Parent.Activate\nSet r = w.Worksheets(1).Range("A1:B2")\n'
                                 's = WorksheetFunction.Sum(r) & "|" & Application.Sum(r) & "|" & '
                                 'WorksheetFunction.CountA(w.Worksheets(1).Range("A1:C2")) & "|" & '
                                 'WorksheetFunction.Max(r, 9) & "|" & WorksheetFunction.VLookup(4, r, 2, False) & "|" & '
                                 'WorksheetFunction.Sum(w.Worksheets(1).Range("A1"), ws.Range("A1")) & "|" & '
                                 'WorksheetFunction.Sum(w.Worksheets(1).Columns(1)) & "|" & '
                                 'WorksheetFunction.CountIf(w.Worksheets(1).Range("A1:C2"), 4)\n'
                                 'w.Worksheets(1).Rows(2).Hidden = True\ns = s & "|" & WorksheetFunction.Subtotal(109, r)\n'
                                 'v = WorksheetFunction.Index(w.Worksheets(1).Range("A1:C2"), 1, 0)\n'
                                 's = s & "|" & UBound(v, 2) & v(1, 3)\nw.Close SaveChanges:=False',
    'manual_calculation': 'ws.Range("A1").Value = 1: ws.Range("B1").Formula = "=A1*10"\n'
                          'Application.Calculation = xlCalculationManual\nws.Range("A1").Value = 2\ns = ws.Range("B1").Value & "|"\n'
                          'ws.Calculate\ns = s & ws.Range("B1").Value & "|" & Application.Calculation\n'
                          'Application.Calculation = xlCalculationAutomatic',
    'statusbar_and_screen': 'Application.StatusBar = "Working"\ns = Application.StatusBar & "|"\nApplication.StatusBar = False\n'
                            's = s & Application.StatusBar & "|" & TypeName(Application.ScreenUpdating)',
    'statusbar_types': 'Application.StatusBar = False\ns = TypeName(Application.StatusBar) & "|" & Application.StatusBar & "|"\n'
                       'Application.StatusBar = 12\ns = s & TypeName(Application.StatusBar) & "|" & Application.StatusBar & "|"\n'
                       'Application.StatusBar = True\ns = s & TypeName(Application.StatusBar) & "|" & Application.StatusBar\n'
                       'Application.StatusBar = False',
    # --- sorting and filtering
    'sort_object_two_keys': LIST + 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B2:B6"), Order:=xlAscending\n'
                                   'ws.Sort.SortFields.Add Key:=ws.Range("C2:C6"), Order:=xlDescending\n'
                                   'With ws.Sort\n.SetRange ws.Range("A1:C6")\n.Header = xlYes\n.Apply\nEnd With\n'
                                   'For i = 2 To 6: s = s & ws.Cells(i, 1).Value & ",": Next',
    'sort_by_cell_color': LIST + 'ws.Range("A4").Interior.Color = RGB(255, 0, 0): ws.Range("A6").Interior.Color = RGB(255, 0, 0)\n'
                                 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add(Key:=ws.Range("A2:A6"), SortOn:=xlSortOnCellColor, '
                                 'Order:=xlAscending).SortOnValue.Color = RGB(255, 0, 0)\n'
                                 'With ws.Sort\n.SetRange ws.Range("A1:C6")\n.Header = xlYes\n.Apply\nEnd With\n'
                                 'For i = 2 To 6: s = s & ws.Cells(i, 1).Value & ",": Next',
    'autofilter_top_items': LIST + 'ws.Range("A1:C6").AutoFilter Field:=3, Criteria1:="2", Operator:=xlTop10Items\n'
                                   's = ws.Range("A1:C6").SpecialCells(xlCellTypeVisible).Rows.Count & "|" & ws.Rows(5).Hidden & '
                                   'ws.Rows(3).Hidden & ws.Rows(2).Hidden\nws.AutoFilterMode = False',
    'autofilter_value_list': LIST + 'ws.Range("A1:C6").AutoFilter Field:=2, Criteria1:=Array("East", "North"), '
                                    'Operator:=xlFilterValues\ns = ws.Rows(3).Hidden & ws.Rows(5).Hidden & "|" & '
                                    'ws.AutoFilter.Filters(2).Count & "|" & ws.AutoFilter.Range.Address\nws.AutoFilterMode = False',
    'autofilter_dates': 'ws.Range("A1").Value = "When"\nFor i = 2 To 6: ws.Cells(i, 1).Value = DateSerial(2024, i, 1): Next\n'
                        'ws.Range("A1:A6").AutoFilter Field:=1, Criteria1:=">=" & CLng(DateSerial(2024, 4, 1))\n'
                        's = ws.Range("A1:A6").SpecialCells(xlCellTypeVisible).Count & "|" & ws.Rows(3).Hidden\n'
                        'ws.AutoFilterMode = False',
    'replace_whole_case': 'ws.Range("A1:A3").Value = WorksheetFunction.Transpose(Array("a", "A", "ab"))\n'
                          'ws.Range("A1:A3").Replace What:="a", Replacement:="z", LookAt:=xlWhole, MatchCase:=True\n'
                          's = ws.Range("A1").Value & ws.Range("A2").Value & ws.Range("A3").Value',
    # --- links, notes and protection
    'hyperlinks_add': 'ws.Hyperlinks.Add Anchor:=ws.Range("A1"), Address:="https://example.com/x", TextToDisplay:="site"\n'
                      'ws.Hyperlinks.Add Anchor:=ws.Range("A2"), Address:="", SubAddress:="Sweep!C5", TextToDisplay:="jump"\n'
                      's = ws.Hyperlinks.Count & "|" & ws.Range("A1").Hyperlinks(1).Address & "|" & ws.Range("A1").Value & "|" & '
                      'ws.Hyperlinks(2).SubAddress & "|" & ws.Range("A2").Text',
    'notes_add_clear': 'ws.Range("B2").AddComment "first note"\ns = ws.Range("B2").Comment.Text & "|" & ws.Comments.Count & "|"\n'
                       'ws.Range("B2").Comment.Text Text:="changed"\ns = s & ws.Range("B2").Comment.Text & "|"\n'
                       'ws.Range("B2").ClearComments\ns = s & ws.Comments.Count',
    'protect_with_password': 'ws.Range("A1").Locked = False\nws.Protect Password:="pw", AllowFormattingCells:=True\n'
                             's = ws.ProtectContents & "|" & ws.Protection.AllowFormattingCells & "|"\nws.Range("A1").Value = 5\n'
                             'On Error Resume Next\nws.Range("B1").Value = 5\ns = s & Err.Number & "|"\nErr.Clear\n'
                             'ws.Unprotect "wrong"\ns = s & Err.Number & "|"\nErr.Clear\nOn Error GoTo 0\nws.Unprotect "pw"\n'
                             's = s & ws.ProtectContents & "|" & ws.Range("A1").Value',
    # --- names, merges and validation
    'name_refers_r1c1': 'ThisWorkbook.Names.Add Name:="SweepCell", RefersTo:="=Sweep!$B$3"\n'
                        's = ThisWorkbook.Names("SweepCell").RefersToR1C1 & "|" & ThisWorkbook.Names("SweepCell").Visible\n'
                        'ThisWorkbook.Names("SweepCell").Delete',
    'merge_across_rows': 'ws.Range("A1:C2").Merge Across:=True\ns = ws.Range("B2").MergeArea.Address & "|" & '
                         'ws.Range("A1").MergeArea.Address & "|" & ws.Range("A1:C2").MergeCells',
    'center_across_selection': 'ws.Range("A1").Value = "title": ws.Range("A1:C1").HorizontalAlignment = xlCenterAcrossSelection\n'
                               's = ws.Range("A1").HorizontalAlignment & "|" & ws.Range("B1").HorizontalAlignment & "|" & '
                               'ws.Range("A1").Text',
    'validation_whole_number': 'ws.Range("A1").Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, '
                               'Operator:=xlBetween, Formula1:="1", Formula2:="10"\nws.Range("A1").Validation.ErrorMessage = "1 to 10"\n'
                               'ws.Range("A1").Value = 11\ns = ws.Range("A1").Validation.Value & "|" & ws.Range("A1").Validation.Formula2 '
                               '& "|" & ws.Range("A1").Validation.ErrorMessage & "|" & ws.Range("A1").Validation.Operator',
    # --- conditional formats
    'cf_color_scale': 'ws.Range("A1:A5").Value = 1\nws.Range("A1:A5").FormatConditions.AddColorScale ColorScaleType:=3\n'
                      's = ws.Range("A1:A5").FormatConditions.Count & "|" & ws.Range("A1:A5").FormatConditions(1).Type',
    'cf_databar_iconset': 'ws.Range("A1:A5").FormatConditions.AddDatabar\nws.Range("B1:B5").FormatConditions.AddIconSetCondition\n'
                          's = ws.Range("A1:A5").FormatConditions(1).Type & "|" & ws.Range("B1:B5").FormatConditions(1).Type',
    # --- outlines and page setup
    'group_show_levels': 'ws.Rows("2:4").Group\nws.Outline.ShowLevels RowLevels:=1\ns = ws.Rows(3).Hidden & "|" & ws.Rows(3).OutlineLevel '
                         '& "|"\nws.Outline.ShowLevels RowLevels:=2\ns = s & ws.Rows(3).Hidden\nws.Rows("2:4").Ungroup',
    'page_setup_titles': 'With ws.PageSetup\n.Orientation = xlLandscape\n.Zoom = False\n.FitToPagesWide = 1\n.FitToPagesTall = False\n'
                         '.PrintTitleRows = "$1:$2"\nEnd With\ns = ws.PageSetup.Orientation & "|" & ws.PageSetup.FitToPagesWide & "|" & '
                         'ws.PageSetup.PrintTitleRows & "|" & ws.Names("Print_Titles").RefersTo',
    # --- number formats and values
    'number_format_local_dates': 'ws.Range("A1").Value = DateSerial(2024, 3, 9)\nws.Range("A1").NumberFormatLocal = "dd/mm/yyyy"\n'
                                 's = ws.Range("A1").NumberFormat & "|" & ws.Range("A1").Text',
    'conditional_number_format': 'ws.Columns(1).ColumnWidth = 15\nws.Range("A1").Value = 15300: ws.Range("A2").Value = 120\n'
                                 'ws.Range("A1:A2").NumberFormat = "[>=1000]#,##0,""K"";0"\ns = ws.Range("A1").Text & "|" & ws.Range("A2").Text',
    'fraction_scientific': 'ws.Columns(1).ColumnWidth = 15\nws.Range("A1").Value = 2.75: ws.Range("A1").NumberFormat = "# ?/?"\n'
                           'ws.Range("A2").Value = 123456: ws.Range("A2").NumberFormat = "0.00E+00"\n'
                           'ws.Range("A3").Value = 0.000123: ws.Range("A3").NumberFormat = "##0.0E+0"\n'
                           's = ws.Range("A1").Text & "|" & ws.Range("A2").Text & "|" & ws.Range("A3").Text',
    'mixed_array_types': 'ws.Range("A1:D1").Value = Array(CCur(1.5), DateSerial(2024, 1, 2), True, "t")\n'
                         's = TypeName(ws.Range("A1").Value) & TypeName(ws.Range("B1").Value) & TypeName(ws.Range("C1").Value) & '
                         'TypeName(ws.Range("D1").Value) & "|" & TypeName(ws.Range("A1").Value2) & "|" & ws.Range("A1").NumberFormat',
    'spill_range_read': 'ws.Range("A1").Formula2 = "=SEQUENCE(3)"\ns = ws.Range("A1").SpillingToRange.Address & "|" & '
                        'ws.Range("A3").Value & "|" & ws.Range("A2").HasSpill & "|" & ws.Range("A1").HasSpill',
    'cell_error_values': 'ws.Range("A1").Formula = "=NA()": ws.Range("A2").Formula = "=1/0"\nv = ws.Range("A1").Value\n'
                         's = IsError(v) & "|" & CStr(v) & "|" & CStr(ws.Range("A2").Value) & "|" & ws.Range("A2").Text & "|" & '
                         'TypeName(ws.Range("A2").Value)',
    # --- VBA's Format, Split and arrays
    'format_number_specs': 's = Format(1234.5, "#,##0.00") & "|" & Format(0.5, "Percent") & "|" & Format(-1, "0;(0)") & "|" & '
                           'Format("abc", ">") & "|" & Format(DateSerial(2024, 7, 4) + TimeSerial(9, 8, 7), "yyyy-mm-dd hh:nn:ss") & '
                           '"|" & Format(1234.5, "Standard") & "|" & Format(0, "Yes/No") & "|" & Format(12, "000.0")',
    'split_limit_filter': 'a = Split("a,b,,c", ",", 2)\ns = UBound(a) & "|" & a(1) & "|"\na = Split("xAyaz", "a", , vbTextCompare)\n'
                          's = s & UBound(a) & "|" & Join(Filter(Array("apple", "berry", "grape"), "ap"), ",") & "|" & '
                          'Join(Filter(Array("apple", "berry", "grape"), "ap", False), ",")',
    'multi_dim_array_range': 'ReDim a(1 To 2, 1 To 3)\nFor i = 1 To 2: For j = 1 To 3: a(i, j) = i * 10 + j: Next: Next\n'
                             'ws.Range("A1").Resize(2, 3).Value = a\ns = UBound(a, 2) & "|" & ws.Range("C2").Value & "|"\n'
                             'v = ws.Range("A1:C2").Value\ns = s & LBound(v, 1) & UBound(v, 2) & "|" & v(2, 1)',
    'transpose_two_dims': 'ReDim a(1 To 2, 1 To 3)\na(1, 3) = "x"\nv = WorksheetFunction.Transpose(a)\n'
                          's = UBound(v, 1) & "x" & UBound(v, 2) & "|" & v(3, 1) & "|" & IsEmpty(v(1, 1))',
    'err_vbobjecterror': 'On Error Resume Next\nErr.Raise vbObjectError + 513, "Sweep", "mine"\ns = Err.Number & "|" & '
                         '(Err.Number - vbObjectError) & "|" & Err.Description',
    'worksheetfunction_error_text': 'On Error Resume Next\nv = WorksheetFunction.VLookup(1, ws.Range("A1:B2"), 2, False)\n'
                                    's = Err.Number & "|" & Err.Description',
    'hidden_rows_visible_count': LIST + 'ws.Range("A3,A5").EntireRow.Hidden = True\n'
                                        's = ws.Range("A1:A6").SpecialCells(xlCellTypeVisible).Count & "|" & '
                                        'WorksheetFunction.Subtotal(109, ws.Range("C2:C6")) & "|" & WorksheetFunction.Sum(ws.Range("C2:C6"))',
    'hidden_several_areas': 'ws.Range("A3,A5").EntireRow.Hidden = True\n'
                            's = ws.Rows(3).Hidden & ws.Rows(4).Hidden & ws.Rows(5).Hidden & "|" & '
                            'TypeName(ws.Range("A3,A5").EntireRow.Hidden) & ":" & ws.Range("A3,A5").EntireRow.Hidden & "|" & '
                            'TypeName(ws.Range("A3:A5").EntireRow.Hidden) & "|"\nws.Rows(5).Hidden = False\n'
                            's = s & ws.Range("A3,A5").EntireRow.Hidden & ws.Range("A5,A3").EntireRow.Hidden & "|"\n'
                            'ws.Range("B1,D1").EntireColumn.Hidden = True\n'
                            's = s & ws.Columns(2).Hidden & ws.Columns(3).Hidden & ws.Columns(4).Hidden & "|"\n'
                            'On Error Resume Next\nws.Range("A3,A5").Hidden = False\ns = s & Err.Number & "|"\nErr.Clear\n'
                            'ws.Range("A3:B3").Hidden = False\ns = s & Err.Number & "|"\nErr.Clear\n'
                            'ws.Rows("3:3,5:5").Hidden = False\ns = s & Err.Number & "|" & ws.Rows(3).Hidden & ws.Rows(5).Hidden\n'
                            # Rows of several areas an address or Union made, rather than EntireRow.
                            'Err.Clear\nws.Range("7:7,9:9").Hidden = True\n'
                            's = s & "|" & Err.Number & ":" & ws.Rows(7).Hidden & ws.Rows(9).Hidden\n'
                            'Err.Clear\nv = ws.Range("7:7,9:9").Hidden\ns = s & "|" & Err.Number\n'
                            'Err.Clear\nUnion(ws.Rows(11), ws.Rows(13)).Hidden = True\n'
                            's = s & "|" & Err.Number & ":" & ws.Rows(11).Hidden & ws.Rows(13).Hidden\n'
                            'Err.Clear\nv = Union(ws.Rows(11), ws.Rows(13)).Hidden\ns = s & "|" & Err.Number & ":" & v',
    'used_range_after_delete': LIST + 'ws.Rows("4:6").Delete\ns = ws.UsedRange.Address & "|" & ws.Range("A1").CurrentRegion.Address',
    'find_by_format_values': LIST + 'Set r = ws.Range("C:C").Find(What:=25)\ns = r.Address & "|"\n'
                                    'Set r = ws.Range("A:C").Find(What:="e", LookAt:=xlPart, SearchOrder:=xlByColumns)\n'
                                    's = s & r.Address & "|"\nSet r = ws.Range("A:C").Find(What:="E", MatchCase:=True)\n'
                                    's = s & r.Address',
}


def macro(name: str) -> str:
    cleanup = "\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "Function Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
