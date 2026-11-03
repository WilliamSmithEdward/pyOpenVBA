"""A sixth sweep of everyday Excel macros: formulas put in by macro and what they work out, reports summed and sorted, rules
that colour, shapes and charts tidied, and procedures that call each other and pass things by reference.

Like scripts/measure_excel_everyday.py, each probe works on a sheet of its own that it deletes afterwards, and is recorded
with its answer or the run-time error it stopped with (see vba_probe_fixture.py).

    python scripts/measure_excel_everyday_6.py

writes tests/fixtures/excel_everyday_6.json, which tests/test_excel_everyday_6.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_6.json'
PRE = ("Dim ws As Worksheet, ws2 As Worksheet, r As Range, c As Range, s As String, i As Long, j As Long, n As Long\n"
       "Dim v As Variant, o As Object, a As Variant, t As String, d As Double, k As Variant\nSet ws = Worksheets.Add\n"
       "ws.Name = \"Sweep\"\n")
#: Orders: a header row and eight rows of a customer, a region, a product, a quantity and a price.
LIST = ('ws.Range("A1:E1").Value = Array("Cust", "Region", "Item", "Qty", "Price")\n'
        'ws.Range("A2:E2").Value = Array("Ann", "North", "Pen", 10, 1.5)\n'
        'ws.Range("A3:E3").Value = Array("Bob", "South", "Ink", 3, 12.25)\n'
        'ws.Range("A4:E4").Value = Array("Cy", "North", "Pad", 7, 4)\n'
        'ws.Range("A5:E5").Value = Array("Di", "East", "Pen", 20, 1.5)\n'
        'ws.Range("A6:E6").Value = Array("Ed", "South", "Pad", 1, 4)\n'
        'ws.Range("A7:E7").Value = Array("Flo", "North", "Ink", 2, 12.25)\n'
        'ws.Range("A8:E8").Value = Array("Gus", "East", "Pad", 5, 4)\n'
        'ws.Range("A9:E9").Value = Array("Hal", "West", "Pen", 8, 1.5)\n')
#: Procedures the probes that need them call, outside the probe.
HELPERS = '''
Private Type Pair
    Left As Long
    Label As String
End Type

Private Enum Size
    Small = 1
    Large = 10
    Huge
End Enum

Private Sub Twice(ByRef x As Long)
    x = x * 2
End Sub

Private Sub Kept(ByVal x As Long)
    x = x * 2
End Sub

Private Function Total(ParamArray items() As Variant) As Double
    Dim one As Variant
    For Each one In items
        Total = Total + one
    Next
End Function

Private Function Named(Optional ByVal who As String = "nobody", Optional ByVal times As Long = 1) As String
    Named = who & times
End Function

Private Function Doubled(ByVal r As Range) As Variant
    Dim out() As Variant, i As Long
    ReDim out(1 To r.Rows.Count, 1 To 1)
    For i = 1 To r.Rows.Count
        out(i, 1) = r.Cells(i, 1).Value * 2
    Next
    Doubled = out
End Function

Private Sub Painted(ByVal r As Range, Optional ByVal colour As Long = 65535)
    r.Interior.Color = colour
End Sub

Private Function Fails() As Long
    Fails = 1 / 0
End Function

Private Function Caught() As String
    On Error Resume Next
    Caught = Fails()
    Caught = "after " & Err.Number
End Function

Private Function Paired() As String
    Dim p As Pair, many(1 To 2) As Pair
    p.Left = 3
    p.Label = "x"
    many(2) = p
    many(2).Left = 4
    Paired = p.Left & many(2).Left & many(2).Label & Len(many(1).Label)
End Function
'''
CASES: dict[str, str] = {
    # Formulas put in by macro.
    'vlookup_exact_approx': LIST + 'ws.Range("G1").Formula = "=VLOOKUP(""Di"",A:E,4,FALSE)"\nws.Range("G2").Formula = "=VLOOKUP(""Nobody"",A:E,4,FALSE)"\nws.Range("G3").Formula = "=IFERROR(G2,""none"")"\nws.Range("I1:J3").Value = Application.Transpose(Application.Transpose(ws.Range("D2:E4").Value))\ns = ws.Range("G1").Value & CStr(ws.Range("G2").Value) & ws.Range("G3").Value & ws.Range("G2").Text & IsError(ws.Range("G2").Value)',
    'index_match_formula': LIST + 'ws.Range("G1").Formula = "=INDEX(E2:E9,MATCH(""Gus"",A2:A9,0))"\nws.Range("G2").Formula = "=MATCH(5,D2:D9,0)"\nws.Range("G3").Formula = "=COUNTIF(B:B,""North"")"\nws.Range("G4").Formula = "=SUMIF(B:B,""East"",D:D)"\nws.Range("G5").Formula = "=SUMPRODUCT(D2:D9,E2:E9)"\ns = ws.Range("G1").Value & "|" & ws.Range("G2").Value & "|" & ws.Range("G3").Value & "|" & ws.Range("G4").Value & "|" & ws.Range("G5").Value',
    'text_formulas': LIST + 'ws.Range("G2").Formula = "=UPPER(LEFT(A2,1))&LOWER(MID(A2,2,10))&LEN(A2)"\nws.Range("G3").Formula = "=TEXT(D3*E3,""#,##0.00"")&"" ""&RIGHT(B3,2)"\nws.Range("G4").Formula = "=SUBSTITUTE(TRIM(""  a  b ""),"" "",""-"")&FIND(""b"",""abc"")&SEARCH(""C"",""abc"")"\nws.Range("G5").Formula = "=CONCATENATE(A5,""/"",B5)&REPT(""*"",2)&EXACT(""a"",""A"")"\ns = ws.Range("G2").Value & "|" & ws.Range("G3").Value & "|" & ws.Range("G4").Value & "|" & ws.Range("G5").Value',
    'date_formulas': 'ws.Range("A1").Value = DateSerial(2024, 2, 29)\nws.Range("A2").Formula = "=EDATE(A1,12)"\nws.Range("A3").Formula = "=EOMONTH(A1,1)"\nws.Range("A4").Formula = "=YEAR(A1)&MONTH(A1)&DAY(A1)&WEEKDAY(A1)"\nws.Range("A5").Formula = "=DATEDIF(A1,DATE(2025,3,1),""m"")"\nws.Range("A6").Formula = "=TEXT(A1,""dd/mm/yy"")"\ns = ws.Range("A1").Text & "|" & ws.Range("A2").Value2 & "|" & ws.Range("A3").Value2 & "|" & ws.Range("A4").Value & "|" & ws.Range("A5").Value & "|" & ws.Range("A6").Value & "|" & ws.Range("A2").NumberFormat',
    'logic_formulas': 'ws.Range("A1").Value = 5\nws.Range("A2").Formula = "=IF(A1>3,""big"",""small"")"\nws.Range("A3").Formula = "=AND(A1>1,A1<3)"\nws.Range("A4").Formula = "=OR(A1>1,A1<3)"\nws.Range("A5").Formula = "=NOT(A1=5)"\nws.Range("A6").Formula = "=ISNUMBER(A2)&ISTEXT(A2)&ISBLANK(B1)&ISERROR(1/0)"\nws.Range("A7").Formula = "=CHOOSE(2,""a"",""b"")"\ns = ws.Range("A2").Value & ws.Range("A3").Value & ws.Range("A4").Value & ws.Range("A5").Value & ws.Range("A6").Value & ws.Range("A7").Value & TypeName(ws.Range("A3").Value)',
    'round_math_formulas': 'ws.Range("A1").Formula = "=ROUND(2.345,2)&""|""&ROUNDUP(2.341,2)&""|""&ROUNDDOWN(2.349,2)&""|""&INT(-1.5)&""|""&MOD(7,3)&""|""&ABS(-2)&""|""&POWER(2,10)&""|""&SQRT(16)&""|""&CEILING(7,5)&""|""&FLOOR(7,5)"\ns = ws.Range("A1").Value',
    'relative_fill_formulas': LIST + 'ws.Range("F2").Formula = "=D2*E2"\nws.Range("F2").AutoFill ws.Range("F2:F9")\nws.Range("G2:G9").Formula = "=F2/SUM($F$2:$F$9)"\ns = ws.Range("F9").Formula & ws.Range("G9").Formula & Round(Application.Sum(ws.Range("G2:G9")), 6)\nws.Range("F2:F9").Copy ws.Range("H2")\ns = s & ws.Range("H3").Formula\nws.Range("F3").Copy\nws.Range("I5").PasteSpecial xlPasteFormulas\nApplication.CutCopyMode = False\ns = s & ws.Range("I5").Formula',
    'formula_error_kinds': 'ws.Range("A1").Formula = "=1/0"\nws.Range("A2").Formula = "=NoSuchName"\nws.Range("A3").Formula = "=""a""+1"\nws.Range("A4").Formula = "=SQRT(-1)"\nws.Range("A5").Formula = "=NA()"\nws.Range("A6").Formula = "=INDEX(A1:A2,5)"\nFor i = 1 To 6\nv = ws.Cells(i, 1).Value\ns = s & CStr(v) & "/" & CLng(v) & "|"\nNext',
    'error_compare': 'ws.Range("A1").Formula = "=1/0"\nv = ws.Range("A1").Value\ns = (v = CVErr(xlErrDiv0)) & IsError(v) & TypeName(v) & (ws.Range("A1").Text = "#DIV/0!") & CVErr(2007)\nOn Error Resume Next\nn = 0\nn = ws.Range("A1").Value\ns = s & Err.Number & n\nErr.Clear\nIf ws.Range("A1").Value > 1 Then s = s & "big"\ns = s & Err.Number',
    'application_vs_worksheetfunction': LIST + 'On Error Resume Next\nv = WorksheetFunction.Match("Nobody", ws.Columns(1), 0)\ns = Err.Number & "|" & Err.Description & "|"\nErr.Clear\nv = Application.Match("Nobody", ws.Columns(1), 0)\ns = s & Err.Number & IsError(v) & "|"\nv = WorksheetFunction.Sum(ws.Range("D2:D9"))\ns = s & v & WorksheetFunction.Max(ws.Range("D2:D9")) & WorksheetFunction.CountA(ws.Columns(1))',
    # Reports summed and sorted.
    'sort_two_keys': LIST + 'With ws.Sort\n.SortFields.Clear\n.SortFields.Add Key:=ws.Range("B2:B9"), Order:=xlAscending\n.SortFields.Add Key:=ws.Range("D2:D9"), Order:=xlDescending\n.SetRange ws.Range("A1:E9")\n.Header = xlYes\n.Apply\nEnd With\nFor i = 2 To 9\ns = s & ws.Cells(i, 1).Value & ","\nNext',
    'sort_range_method': LIST + 'ws.Range("A1:E9").Sort Key1:=ws.Range("E1"), Order1:=xlDescending, Key2:=ws.Range("A1"), Order2:=xlAscending, Header:=xlYes\nFor i = 2 To 9\ns = s & ws.Cells(i, 1).Value\nNext',
    'autofilter_top_and_multi': LIST + 'ws.Range("A1:E9").AutoFilter 4, "5", xlTop10Items\ns = ws.Range("A2:A9").SpecialCells(xlCellTypeVisible).Count & ws.AutoFilter.Filters(4).On & ws.AutoFilter.Range.Address\nws.ShowAllData\nws.Range("A1:E9").AutoFilter 2, Array("North", "East"), xlFilterValues\ns = s & "|" & ws.Range("A2:A9").SpecialCells(xlCellTypeVisible).Count & ws.FilterMode\nws.Range("A1:E9").AutoFilter\ns = s & ws.AutoFilterMode',
    'autofilter_or_and_sum': LIST + 'ws.Range("A1:E9").AutoFilter 4, ">=5", xlAnd, "<=10"\ns = Application.Subtotal(109, ws.Range("D2:D9")) & "|" & Application.Subtotal(103, ws.Range("A2:A9"))\nFor i = 2 To 9\nIf ws.Rows(i).Hidden Then s = s & i\nNext\nws.AutoFilterMode = False\ns = s & "|" & ws.Rows(3).Hidden',
    'unique_report': LIST + 'ws.Range("B1:B9").Copy ws.Range("G1")\nws.Range("G1:G9").RemoveDuplicates 1, xlYes\nn = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row\nFor i = 2 To n\nws.Cells(i, 8).Formula = "=SUMIF($B$2:$B$9,G" & i & ",$D$2:$D$9)"\ns = s & ws.Cells(i, 7).Value & ws.Cells(i, 8).Value & ","\nNext\ns = s & n',
    'advanced_filter_copy': LIST + 'ws.Range("G1").Value = "Region"\nws.Range("G2").Value = "North"\nws.Range("A1:E9").AdvancedFilter xlFilterCopy, ws.Range("G1:G2"), ws.Range("I1")\ns = ws.Range("I1").Value & ws.Range("I2").Value & ws.Cells(ws.Rows.Count, 9).End(xlUp).Row\nws.Columns("B").AdvancedFilter xlFilterCopy, , ws.Range("O1"), True\ns = s & ws.Cells(ws.Rows.Count, 15).End(xlUp).Row',
    'subtotal_method': LIST + 'ws.Range("A1:E9").Sort Key1:=ws.Range("B1"), Header:=xlYes\nws.Range("A1:E9").Subtotal 2, xlSum, Array(4)\ns = ws.UsedRange.Address & ws.Cells(ws.Rows.Count, 2).End(xlUp).Value\nws.Outline.ShowLevels 2\ns = s & ws.Rows(2).Hidden\nws.Range("A1").CurrentRegion.RemoveSubtotal\ns = s & ws.UsedRange.Address',
    'consolidate_loop_array': LIST + 'v = ws.Range("A2:E9").Value\nReDim a(1 To UBound(v, 1), 1 To 2)\nFor i = 1 To UBound(v, 1)\na(i, 1) = UCase(v(i, 3))\na(i, 2) = v(i, 4) * v(i, 5)\nNext\nws.Range("G1").Resize(UBound(a, 1), 2).Value = a\ns = ws.Range("G3").Value & ws.Range("H3").Value & Application.Sum(ws.Range("H1:H8"))',
    # Rules that colour.
    'cf_cell_and_formula': LIST + 'With ws.Range("D2:D9").FormatConditions\n.Delete\n.Add xlCellValue, xlGreater, "=5"\n.Item(1).Interior.Color = 255\n.Add xlExpression, , "=$E2>10"\n.Item(2).Font.Bold = True\ns = .Count & .Item(1).Formula1 & .Item(1).Type & .Item(1).Operator & .Item(2).Formula1 & .Item(2).Type & .Item(1).Priority\nEnd With\ns = s & ws.Range("D2").DisplayFormat.Interior.Color & ws.Range("D3").DisplayFormat.Font.Bold & ws.Range("D3").Interior.Color\nws.Cells.FormatConditions.Delete\ns = s & ws.Range("D2:D9").FormatConditions.Count',
    'cf_top_dup_bars': LIST + 'ws.Range("D2:D9").FormatConditions.AddTop10\nws.Range("D2:D9").FormatConditions(1).TopBottom = xlTop10Top\nws.Range("D2:D9").FormatConditions(1).Rank = 2\nws.Range("A2:A9").FormatConditions.AddUniqueValues\nws.Range("E2:E9").FormatConditions.AddDatabar\nws.Range("E2:E9").FormatConditions.AddColorScale 3\nws.Range("E2:E9").FormatConditions.AddIconSetCondition\ns = ws.Range("D2:D9").FormatConditions.Count & ws.Range("E2:E9").FormatConditions.Count & TypeName(ws.Range("E2").FormatConditions(1)) & TypeName(ws.Range("E2").FormatConditions(2)) & TypeName(ws.Range("E2").FormatConditions(3)) & ws.Range("D2:D9").FormatConditions(1).Rank',
    # Styles and looks.
    'named_styles': 'Dim st As Object\nSet st = ThisWorkbook.Styles.Add("SweepStyle")\nst.Font.Bold = True\nst.Interior.Color = 65535\nst.NumberFormat = "0.0"\nws.Range("A1").Style = "SweepStyle"\nws.Range("A1").Value = 2\ns = ws.Range("A1").Text & ws.Range("A1").Font.Bold & ws.Range("A1").Interior.Color & ws.Range("A1").Style.Name\nws.Range("A1").Style = "Normal"\ns = s & ws.Range("A1").Font.Bold & ws.Range("A1").Text\nst.Delete',
    'builtin_styles_formats': 'ws.Range("A1").Style = "Currency"\nws.Range("A2").Style = "Percent"\nws.Range("A3").Style = "Comma"\nws.Range("A1:A3").Value = 0.5\nws.Columns(1).ColumnWidth = 15\ns = ws.Range("A1").Text & "|" & ws.Range("A2").Text & "|" & ws.Range("A3").Text & "|" & ws.Range("A1").NumberFormat & "|" & ws.Range("A2").NumberFormat',
    'number_format_texts': 'ws.Columns(1).ColumnWidth = 25\nFor Each k In Array("0", "0.00", "#,##0", "0%", "0.0E+00", "# ?/?", "yyyy-mm-dd", "hh:mm", "@", "$#,##0.00;[Red]($#,##0.00)", "General", "000-00", """x""0", "0.0,", ";;;")\nws.Range("A1").NumberFormat = k\nws.Range("A1").Value = -1234.5\ns = s & ws.Range("A1").Text & "|"\nNext',
    'color_helpers': 'ws.Range("A1").Interior.ColorIndex = 6\nws.Range("A2").Font.ColorIndex = 3\nws.Range("A3").Interior.Color = RGB(10, 20, 30)\nws.Range("A4").Interior.ThemeColor = xlThemeColorAccent2\nws.Range("A4").Interior.TintAndShade = 0.4\ns = ws.Range("A1").Interior.Color & "|" & ws.Range("A2").Font.Color & "|" & ws.Range("A3").Interior.Color & "|" & ws.Range("A3").Interior.ColorIndex & "|" & ws.Range("A4").Interior.ThemeColor & "|" & Round(ws.Range("A4").Interior.TintAndShade, 3) & "|" & ws.Range("A4").Interior.Color\nws.Range("A1:A4").Interior.Pattern = xlNone\ns = s & "|" & ws.Range("A1").Interior.ColorIndex',
    'merge_center_read': 'ws.Range("A1").Value = "T"\nws.Range("A1:C1").Merge\nws.Range("A1:C1").HorizontalAlignment = xlCenter\ns = ws.Range("B1").MergeCells & ws.Range("B1").MergeArea.Address & ws.Range("A1").MergeArea.Count & "[" & ws.Range("B1").Value & "]" & ws.Range("A1").HorizontalAlignment\nws.Range("A1:C1").UnMerge\ns = s & ws.Range("B1").MergeCells & ws.UsedRange.Address',
    # Shapes and charts tidied.
    'shapes_loop': 'Set o = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 50)\no.Name = "Box"\no.TextFrame2.TextRange.Text = "Hi"\no.Fill.ForeColor.RGB = RGB(1, 2, 3)\no.Line.Visible = msoFalse\nSet o = ws.Shapes.AddLine(0, 0, 50, 50)\nSet o = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 1, 1, 40, 20)\nFor Each o In ws.Shapes\ns = s & o.Name & o.Type & "|"\nNext\ns = s & ws.Shapes("Box").TextFrame2.TextRange.Text & ws.Shapes("Box").Fill.ForeColor.RGB & ws.Shapes("Box").Line.Visible & ws.Shapes.Count\nws.Shapes("Box").Delete\ns = s & ws.Shapes.Count',
    'shape_position_cells': 'Set o = ws.Shapes.AddShape(1, ws.Range("C3").Left, ws.Range("C3").Top, 50, 20)\ns = o.TopLeftCell.Address & o.BottomRightCell.Address & o.Placement\no.Left = o.Left + 100\ns = s & o.TopLeftCell.Address\no.Placement = xlFreeFloating\ns = s & o.Placement & o.ZOrderPosition & o.Rotation & o.Visible\no.Rotation = 45\no.Visible = msoFalse\ns = s & o.Rotation & o.Visible',
    'chart_made_recorded': LIST + 'ws.Range("A1:A9,D1:D9").Select\nws.Range("A1:B9").Select\nws.Shapes.AddChart2(201, xlColumnClustered).Select\nActiveChart.SetSourceData ws.Range("C1:D9")\nActiveChart.ChartTitle.Text = "Qty"\nActiveChart.HasLegend = False\nActiveChart.Axes(xlValue).HasMajorGridlines = False\nActiveChart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(9, 9, 9)\nActiveChart.ChartGroups(1).GapWidth = 80\ns = ActiveChart.ChartTitle.Text & ActiveChart.SeriesCollection.Count & ActiveChart.SeriesCollection(1).Formula & ActiveChart.Parent.Name & ActiveChart.ChartGroups(1).GapWidth\nws.Range("A1").Select\ns = s & TypeName(ActiveChart)',
    'chart_objects_loop': LIST + 'Set o = ws.ChartObjects.Add(10, 10, 200, 120)\no.Chart.ChartType = xlLine\no.Chart.SetSourceData ws.Range("D1:E9")\no.Name = "One"\nSet o = ws.ChartObjects.Add(10, 150, 200, 120)\no.Chart.ChartType = xlPie\no.Chart.SetSourceData ws.Range("A1:A9,D1:D9")\nFor Each o In ws.ChartObjects\ns = s & o.Name & o.Chart.ChartType & o.Chart.SeriesCollection.Count & o.Index & "|"\nNext\ns = s & ws.ChartObjects("One").TopLeftCell.Address\nws.ChartObjects.Delete\ns = s & ws.ChartObjects.Count',
    'pivot_small': LIST + 'Set ws2 = Worksheets.Add\nSet o = ThisWorkbook.PivotCaches.Create(xlDatabase, ws.Range("A1:E9")).CreatePivotTable(ws2.Range("A3"), "Pt")\no.PivotFields("Region").Orientation = xlRowField\no.AddDataField o.PivotFields("Qty"), "Total", xlSum\ns = o.TableRange1.Address & ws2.Range("A5").Value & ws2.Range("B5").Value & o.DataBodyRange.Address & o.RowFields.Count & o.GetPivotData("Total", "Region", "North")\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    # Procedures.
    'byref_byval': 'n = 3\nTwice n\ns = n\nKept n\ns = s & n\nTwice (n)\ns = s & n\nCall Twice(n)\ns = s & n\ni = 5\nTwice i\ns = s & i',
    'paramarray_optional': 's = Total(1, 2, 3.5) & "|" & Total() & "|" & Named() & "|" & Named("a") & "|" & Named(times:=3) & "|" & Named("b", 2)',
    'range_to_function': 'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\nv = Doubled(ws.Range("A1:A3"))\ns = UBound(v, 1) & v(3, 1)\nws.Range("B1:B3").Value = Doubled(ws.Range("A1:A3"))\ns = s & ws.Range("B2").Value\nPainted ws.Range("A1")\nPainted ws.Range("A2"), 255\ns = s & ws.Range("A1").Interior.Color & ws.Range("A2").Interior.Color',
    'error_through_calls': 's = Caught()\nOn Error Resume Next\nn = Fails()\ns = s & "|" & Err.Number & Err.Description',
    'types_enums': 's = Paired() & Small & Large & Huge & "|"\nDim b As Byte, x As Integer, l As Long, f As Single, y As Currency, z As Date, q As Boolean\nb = 200\nx = 32000\nf = 1.5\ny = 1.23456\nz = #1/2/2024#\nq = 5\ns = s & TypeName(b) & TypeName(x * 2&) & TypeName(f + 1) & TypeName(y * 2) & TypeName(z + 1) & q & y & CDbl(z) & VarType(s) & VarType(Empty) & VarType(Null)',
    'overflow_and_division': 'Dim x As Integer\nOn Error Resume Next\nx = 32767\nx = x + 1\ns = Err.Number & "|" & x & "|"\nErr.Clear\nn = 300 * 300\ns = s & Err.Number & "|"\nErr.Clear\nd = 7 \\ 2\ns = s & d & 7 Mod 3 & 2 ^ 3 & -2 ^ 2 & 7 / 2 & CLng(2.5) & CLng(3.5) & "|"\nd = 1 / 0\ns = s & Err.Number',
    'strings_compare_and_loop': 's = ("a" = "A") & StrComp("a", "A", vbTextCompare) & StrComp("a", "b") & LCase("AB") & Space(2) & String(2, "x") & Asc("A") & Chr(66) & StrReverse("ab") & Len("") & Left("abc", 9) & Right("abc", 0) & Mid("abc", 2) & Join(Array(1, "a", 2.5), "-") & UBound(Split("", ",")) & UBound(Split("a,b", ","))',
    'array_helpers': 'a = Array(3, 1, 2)\ns = LBound(a) & UBound(a) & IsArray(a) & IsArray(1) & TypeName(a)\nReDim b(2) As String\nb(0) = "x"\nReDim Preserve b(3)\ns = s & b(0) & UBound(b)\nErase b\nDim f(1 To 2, 0 To 1) As Long\nf(2, 1) = 7\ns = s & UBound(f, 1) & LBound(f, 2) & f(2, 1)\nOn Error Resume Next\nf(3, 3) = 1\ns = s & Err.Number & Application.Max(a) & Application.Sum(a) & Filter(Array("ab", "cd"), "a")(0)',
    'with_object_lifetimes': LIST + 'Set r = ws.Range("A2")\nws.Rows(1).Insert\ns = r.Address\nws.Rows(1).Delete\ns = s & r.Address\nSet c = ws.Range("Z1")\nc.EntireColumn.Delete\ns = s & "|"\nOn Error Resume Next\ns = s & c.Address\ns = s & Err.Number\nErr.Clear\nSet ws2 = Worksheets.Add\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True\ns = s & ws2.Name\ns = s & Err.Number',
    'is_nothing_empty_null': 'Dim e As Variant\nSet o = Nothing\ns = (o Is Nothing) & IsEmpty(e) & IsNull(Null) & IsEmpty(ws.Range("A1").Value) & (ws.Range("A1").Value = "") & (ws.Range("A1").Value = 0) & TypeName(ws.Range("A1").Value) & Len(ws.Range("A1").Value) & IsObject(ws) & IsDate("1/2/2024") & IsDate("x") & (Null = 1) & TypeName(Null & "a")',
    'used_range_shapes': LIST + 's = ws.UsedRange.Rows.Count & ws.UsedRange.Columns.Count & ws.Cells.SpecialCells(xlCellTypeLastCell).Address\nws.Range("H20").Value = 1\ns = s & ws.UsedRange.Address\nws.Range("H20").ClearContents\ns = s & ws.UsedRange.Address & ws.Range("A1").CurrentRegion.Address\nws.Range("A1:E1").EntireRow.Hidden = True\ns = s & ws.Range("A1").EntireRow.Hidden & ws.Range("A2").Top & ws.Range("B1").Height',
    'special_cells_kinds': LIST + 'ws.Range("F2").Formula = "=D2*E2"\nws.Range("F3").Formula = "=1/0"\ns = ws.UsedRange.SpecialCells(xlCellTypeFormulas).Address & ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors).Address & ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers).Count & ws.UsedRange.SpecialCells(xlCellTypeConstants, xlTextValues).Count & ws.UsedRange.SpecialCells(xlCellTypeBlanks).Count\nOn Error Resume Next\nSet r = ws.Range("Z1:Z2").SpecialCells(xlCellTypeConstants)\ns = s & Err.Number & Err.Description',
    'protect_edit_flow': 'ws.Range("A1").Locked = False\nws.Protect Password:="p", UserInterfaceOnly:=False, AllowFormattingCells:=True\ns = ws.ProtectContents & ws.Protection.AllowFormattingCells & ws.Range("A1").Locked & ws.Range("A2").Locked\nws.Range("A1").Value = 1\nOn Error Resume Next\nws.Range("A2").Value = 1\ns = s & Err.Number\nErr.Clear\nws.Range("A2").Font.Bold = True\ns = s & Err.Number\nErr.Clear\nws.Unprotect "x"\ns = s & Err.Number\nErr.Clear\nws.Unprotect "p"\ns = s & ws.ProtectContents',
    'workbook_sheets_add_names': 'n = Worksheets.Count\nSet ws2 = Worksheets.Add(After:=Worksheets(Worksheets.Count))\ns = Worksheets.Count - n & ws2.Index - n\nOn Error Resume Next\nws2.Name = "Sweep"\ns = s & Err.Number\nErr.Clear\nws2.Name = "a/b"\ns = s & Err.Number\nErr.Clear\nws2.Name = String(32, "x")\ns = s & Err.Number\nErr.Clear\nOn Error GoTo 0\nws2.Name = "Other"\ns = s & Worksheets("other").Name & Sheets(ws2.Index).Name & ws2.CodeName\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
}


def macro(name: str) -> str:
    cleanup = "\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return HELPERS + "\nFunction Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
