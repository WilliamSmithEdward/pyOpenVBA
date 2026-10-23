"""A second sweep of everyday Excel macros: copying and pasting, cuts, inserts, sheet copies, names, filters,
lookups, VBA's conversions and strings, regular expressions, dictionaries, colours, validation and AutoFit.

Like scripts/measure_excel_everyday.py, each probe works on a sheet of its
own that it deletes afterwards, and is recorded as its answer or the
run-time error it stopped with (see vba_probe_fixture.py). Macros that make
sheets of their own delete them too.

    python scripts/measure_excel_everyday_2.py

writes tests/fixtures/excel_everyday_2.json, which tests/test_excel_everyday_2.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_2.json'
#: Each probe's sheet is named, so what names it in a reference reads the same in Excel and the model.
PRE = ("Dim ws As Worksheet, r As Range, c As Range, s As String, i As Long, j As Long, n As Long, v As Variant\n"
       "Dim o As Object, w As Worksheet, t As String\nSet ws = Worksheets.Add\nws.Name = \"Sweep\"\n")

CASES: dict[str, str] = {
    # --- copying, pasting and moving cells
    'copy_destination_formats': 'ws.Range("A1").Value = 1234.5: ws.Range("A1").NumberFormat = "#,##0.00": '
                                'ws.Range("A1").Font.Bold = True\nws.Range("A1").Copy Destination:=ws.Range("C3")\n'
                                's = ws.Range("C3").Text & "|" & ws.Range("C3").Font.Bold & "|" & ws.Range("C3").NumberFormat',
    'paste_transpose': 'ws.Range("A1:C1").Value = Array(1, 2, 3)\nws.Range("A1:C1").Copy\n'
                       'ws.Range("E1").PasteSpecial Paste:=xlPasteAll, Transpose:=True\nApplication.CutCopyMode = False\n'
                       's = ws.Range("E1").Value & ws.Range("E2").Value & ws.Range("E3").Value & "|" & ws.Range("F1").Value',
    'paste_formats_only': 'ws.Range("A1").Value = 5: ws.Range("A1").Interior.Color = RGB(255, 0, 0): '
                          'ws.Range("B1").Value = 7\nws.Range("A1").Copy\nws.Range("B1").PasteSpecial xlPasteFormats\n'
                          'Application.CutCopyMode = False\ns = ws.Range("B1").Value & "|" & ws.Range("B1").Interior.Color',
    'cut_moves_references': 'ws.Range("A1").Value = 3: ws.Range("B1").Formula = "=A1*2"\n'
                            'ws.Range("A1").Cut ws.Range("D4")\ns = ws.Range("B1").Formula & "|" & ws.Range("B1").Value & '
                            '"|" & IsEmpty(ws.Range("A1").Value)',
    'insert_copy_origin': 'ws.Range("A1").Font.Bold = True: ws.Range("A2").Font.Italic = True\n'
                          'ws.Rows(2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove\n'
                          's = ws.Range("A2").Font.Bold & ws.Range("A2").Font.Italic & "|" & ws.Range("A3").Font.Italic\n'
                          'ws.Rows(2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow\n'
                          's = s & "|" & ws.Range("A2").Font.Bold & ws.Range("A2").Font.Italic',
    'insert_from_below_rows': 'ws.Range("A1").Font.Bold = True: ws.Range("A2").Font.Italic = True: '
                              'ws.Range("A3").Font.Underline = xlUnderlineStyleSingle: ws.Rows(2).RowHeight = 30\n'
                              'ws.Rows("2:3").Insert CopyOrigin:=xlFormatFromRightOrBelow\n'
                              's = ws.Range("A2").Font.Bold & ws.Range("A2").Font.Italic & ws.Range("A3").Font.Italic & '
                              '"|" & ws.Rows(2).RowHeight & "," & ws.Rows(3).RowHeight & "," & ws.Rows(4).RowHeight',
    'insert_from_below_cells': 'ws.Range("B1").Interior.Color = RGB(255, 0, 0): ws.Range("B2").Interior.Color = RGB(0, 255, 0)\n'
                               'ws.Range("B2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow\n'
                               's = ws.Range("B2").Interior.Color & "|" & ws.Range("B3").Interior.Color\n'
                               'ws.Range("D1").Font.Bold = True: ws.Range("E1").Font.Italic = True\n'
                               'ws.Range("E1").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow\n'
                               's = s & "|" & ws.Range("E1").Font.Bold & ws.Range("E1").Font.Italic',
    'insert_from_below_columns': 'ws.Columns(1).ColumnWidth = 5: ws.Columns(2).ColumnWidth = 20: ws.Range("B1").Font.Bold = True\n'
                                 'ws.Columns(2).Insert CopyOrigin:=xlFormatFromRightOrBelow\n'
                                 's = ws.Columns(2).ColumnWidth & "," & ws.Columns(3).ColumnWidth & "|" & '
                                 'ws.Range("B1").Font.Bold\nws.Columns(2).Insert\ns = s & "|" & ws.Columns(2).ColumnWidth & '
                                 '"," & ws.Range("B1").Font.Bold',
    'insert_from_below_top': 'ws.Range("A1").Font.Bold = True\nws.Rows(1).Insert CopyOrigin:=xlFormatFromRightOrBelow\n'
                             's = ws.Range("A1").Font.Bold & "|"\nws.Rows(1).Insert\ns = s & ws.Range("A1").Font.Bold',
    'advanced_filter_blank_target': 'ws.Range("A1:B1").Value = Array("h", "k")\n'
                                    'ws.Range("A2:B5").Value = WorksheetFunction.Transpose(WorksheetFunction.Transpose('
                                    'ws.Evaluate("{1,2;1,2;3,4;1,5}")))\n'
                                    'ws.Range("A1:B5").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("E1"), '
                                    'Unique:=True\ns = ws.Range("E1").CurrentRegion.Address & "|" & ws.Range("F1").Value & '
                                    'ws.Range("F4").Value & "|" & ws.Names("Extract").RefersTo',
    'advanced_filter_blank_row': 'ws.Range("A1:B1").Value = Array("h", "k")\nws.Range("A2:B3").Value = 7\n'
                                 'ws.Range("A1:B3").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("E1:G1")\n'
                                 's = ws.Range("E1").CurrentRegion.Address & "|" & ws.Range("G1").Value & "|" & '
                                 'ws.Names("Extract").RefersTo',
    'delete_column_ref_error':'ws.Range("B1").Value = 4: ws.Range("C1").Formula = "=B1+1"\nws.Columns(2).Delete\n'
                               's = ws.Range("B1").Formula & "|" & ws.Range("B1").Text',
    'merge_area_props': 'ws.Range("B2").Value = "x": ws.Range("B2:C3").Merge\n'
                        's = ws.Range("C3").MergeArea.Address & "|" & ws.Range("C3").MergeCells & "|" & '
                        'ws.Range("B2:C3").MergeCells & "|" & ws.Range("C3").Value & "|"\nws.Range("B2").UnMerge\n'
                        's = s & ws.Range("B2").Value & "," & IsEmpty(ws.Range("C3").Value)',
    'relative_formula_block': 'ws.Range("A1:A5").Value = 3\nws.Range("B1:B5").Formula = "=A1*2"\n'
                              's = ws.Range("B5").Formula & "|" & ws.Range("B5").Value\n'
                              'ws.Range("C1:C3").FormulaR1C1 = "=RC[-2]+R1C1"\ns = s & "|" & ws.Range("C3").Formula',
    'clear_formats_keeps_values': 'ws.Range("A1").Value = 0.5: ws.Range("A1").NumberFormat = "0%": '
                                  'ws.Range("A1").Font.Bold = True\nws.Range("A1").ClearFormats\n'
                                  's = ws.Range("A1").Text & "|" & ws.Range("A1").Font.Bold & "|" & ws.Range("A1").NumberFormat',
    # --- sheets and names
    'sheet_copy_after': 'ws.Name = "SweepCopy": ws.Range("A1").Value = 9\nws.Copy After:=ws\n'
                        'Set w = Worksheets(ws.Index + 1)\ns = w.Name & "|" & w.Range("A1").Value\n'
                        'Application.DisplayAlerts = False: w.Delete: Application.DisplayAlerts = True',
    'sheet_move_before': 'Set w = Worksheets.Add(After:=ws)\nw.Move Before:=ws\ns = (w.Index = ws.Index - 1) & "|"\n'
                         'Application.DisplayAlerts = False: w.Delete: Application.DisplayAlerts = True\n'
                         's = s & (Worksheets.Count > 0)',
    'sheet_scoped_name': 'ws.Names.Add Name:="LocalX", RefersTo:="=$A$1"\nws.Range("A1").Value = 42\n'
                         's = ws.Names("LocalX").Name & "|" & ws.Range("LocalX").Value & "|" & ws.Names("LocalX").RefersTo\n'
                         'ws.Names("LocalX").Delete',
    'range_name_property': 'ws.Range("B2:C3").Name = "SweepBlock"\ns = ThisWorkbook.Names("SweepBlock").RefersToRange.Address & '
                           '"|" & Mid(ws.Range("B2:C3").Name, 1, 1)\nThisWorkbook.Names("SweepBlock").Delete',
    'named_constant_evaluate': 'ThisWorkbook.Names.Add Name:="SweepTax", RefersTo:="=0.2"\n'
                               's = Evaluate("SweepTax*100") & "|" & ThisWorkbook.Names("SweepTax").Value\n'
                               'ThisWorkbook.Names("SweepTax").Delete',
    'very_hidden_sheet': 'ws.Visible = xlSheetVeryHidden\ns = ws.Visible & "|"\nws.Visible = xlSheetVisible\ns = s & ws.Visible',
    # --- finding, filtering and sorting
    'find_whole_values': 'ws.Range("A1").Formula = "=2+3": ws.Range("A2").Value = "5x": ws.Range("A3").Value = 5\n'
                         'Set r = ws.Range("A1:A3").Find(What:="5", LookIn:=xlValues, LookAt:=xlWhole, After:=ws.Range("A3"))\n'
                         's = r.Address\nSet r = ws.Range("A1:A3").FindNext(r)\ns = s & "|" & r.Address\n'
                         'Set r = ws.Range("A1:A3").Find(What:="5", LookIn:=xlFormulas, LookAt:=xlWhole, After:=ws.Range("A3"))\n'
                         's = s & "|" & r.Address',
    'last_cell_find_star': 'ws.Range("B2").Value = 1: ws.Range("E3").Value = 2: ws.Range("C7").Value = 3\n'
                           's = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row & "|" & '
                           'ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column',
    'special_cells_constants': 'ws.Range("A1").Value = 1: ws.Range("A2").Value = "t": ws.Range("A3").Formula = "=1"\n'
                               's = ws.Range("A1:A3").SpecialCells(xlCellTypeConstants, xlNumbers).Address & "|" & '
                               'ws.Range("A1:A3").SpecialCells(xlCellTypeConstants).Address & "|" & '
                               'ws.Range("A1:A3").SpecialCells(xlCellTypeFormulas).Address',
    'autofilter_or_count': 'ws.Range("A1").Value = "k": For i = 2 To 9: ws.Cells(i, 1).Value = i Mod 4: Next\n'
                           'ws.Range("A1:A9").AutoFilter Field:=1, Criteria1:="1", Operator:=xlOr, Criteria2:="3"\n'
                           's = ws.Range("A1:A9").SpecialCells(xlCellTypeVisible).Count & "|" & ws.Rows(3).Hidden\n'
                           'ws.AutoFilterMode = False',
    'remove_duplicates_two': 'ws.Range("A1:B1").Value = Array("a", "b")\nws.Range("A2:B2").Value = Array(1, 2)\n'
                             'ws.Range("A3:B3").Value = Array(1, 2)\nws.Range("A4:B4").Value = Array(1, 3)\n'
                             'ws.Range("A1:B4").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes\n'
                             's = ws.Range("A1").CurrentRegion.Rows.Count & "|" & ws.Range("B3").Value',
    'sort_descending_header': 'ws.Range("A1").Value = "n": ws.Range("A2").Value = 2: ws.Range("A3").Value = 9: '
                              'ws.Range("A4").Value = 5\nws.Range("A1:A4").Sort Key1:=ws.Range("A1"), '
                              'Order1:=xlDescending, Header:=xlYes\n'
                              's = ws.Range("A1").Value & ws.Range("A2").Value & ws.Range("A3").Value & ws.Range("A4").Value',
    'advanced_filter_unique': 'ws.Range("A1").Value = "h": ws.Range("A2:A5").Value = WorksheetFunction.Transpose(Array(1, 2, 1, 3))\n'
                              'ws.Range("A1:A5").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("C1"), Unique:=True\n'
                              's = ws.Range("C1").CurrentRegion.Rows.Count & "|" & ws.Range("C4").Value',
    # --- worksheet functions and lookups
    'countif_wildcards': 'ws.Range("A1:A4").Value = WorksheetFunction.Transpose(Array("apple", "apricot", "berry", "Avocado"))\n'
                         's = WorksheetFunction.CountIf(ws.Range("A1:A4"), "a*") & "|" & '
                         'WorksheetFunction.CountIf(ws.Range("A1:A4"), "?????") & "|" & '
                         'WorksheetFunction.CountIf(ws.Range("A1:A4"), "<>berry")',
    'sumproduct_evaluate': 'For i = 1 To 5: ws.Cells(i, 1).Value = i: ws.Cells(i, 2).Value = i * 10: Next\n'
                           's = ws.Evaluate("SUMPRODUCT((A1:A5>2)*B1:B5)") & "|" & ws.Evaluate("SUM(A1:A5*B1:B5)")',
    'vlookup_not_found': 'ws.Range("A1:B2").Value = 1: ws.Range("A2").Value = 2\n'
                         'v = Application.VLookup(7, ws.Range("A1:B2"), 2, False)\n'
                         's = IsError(v) & "|" & CStr(v)\nOn Error Resume Next\n'
                         'v = WorksheetFunction.VLookup(7, ws.Range("A1:B2"), 2, False)\ns = s & "|" & Err.Number',
    'match_approximate': 'ws.Range("A1:A4").Value = WorksheetFunction.Transpose(Array(10, 20, 30, 40))\n'
                         's = Application.Match(25, ws.Range("A1:A4"), 1) & "|" & Application.Match(40, ws.Range("A1:A4"), 0) & "|" & '
                         'IsError(Application.Match(5, ws.Range("A1:A4"), 1))',
    'index_match_two_way': 'ws.Range("B1:C1").Value = Array("x", "y"): ws.Range("A2:A3").Value = WorksheetFunction.Transpose(Array("p", "q"))\n'
                           'ws.Range("B2:C3").Value = 7: ws.Range("C3").Value = 9\n'
                           's = WorksheetFunction.Index(ws.Range("B2:C3"), WorksheetFunction.Match("q", ws.Range("A2:A3"), 0), '
                           'WorksheetFunction.Match("y", ws.Range("B1:C1"), 0))',
    'round_family': 's = Round(2.5) & "|" & Round(3.5) & "|" & WorksheetFunction.Round(2.5, 0) & "|" & '
                    'WorksheetFunction.RoundUp(2.11, 1) & "|" & WorksheetFunction.RoundDown(-2.19, 1) & "|" & '
                    'Round(1.005, 2) & "|" & WorksheetFunction.MRound(17, 5) & "|" & Int(-2.5) & Fix(-2.5)',
    'worksheet_text_formats': 's = WorksheetFunction.Text(0.256, "0.0%") & "|" & WorksheetFunction.Text(45000, "dddd") & "|" & '
                              'WorksheetFunction.Text(-1234.5, "#,##0.00;(#,##0.00)") & "|" & Format(0.256, "0.0%") & "|" & '
                              'Format(1234.5, "Currency")',
    'application_errors': 'On Error Resume Next\nv = WorksheetFunction.Match("zz", ws.Range("A1:A3"), 0)\n'
                          's = Err.Number & "|"\nErr.Clear\nv = Application.Match("zz", ws.Range("A1:A3"), 0)\n'
                          's = s & Err.Number & "|" & TypeName(v) & "|"\nErr.Clear\nv = WorksheetFunction.Sum(ws.Range("A1:A3"))\n'
                          's = s & v & "|" & Err.Number',
    # --- VBA conversions and strings
    'cstr_doubles': 's = CStr(1 / 3) & "|" & CStr(1E+16) & "|" & CStr(0.1 + 0.2) & "|" & Str(42) & "|" & CStr(-0.000001) & '
                    '"|" & CStr(123456789012345#) & "|" & CStr(CSng(1.1))',
    'val_function': 's = Val("12abc") & "|" & Val("&H10") & "|" & Val(" 1 2 3") & "|" & Val("1e3x") & "|" & Val("-.5") & '
                    '"|" & Val("abc")',
    'isnumeric_isdate': 's = IsNumeric("1e5") & IsNumeric("$1,000") & IsNumeric("1,2,3") & IsNumeric("") & IsNumeric(" 12 ") & '
                        '"|" & IsDate("2024-02-30") & IsDate("2024-02-29") & IsDate("12:61") & IsDate("Jan 5")',
    'cdate_strings': 's = Format(CDate("1/2/2024"), "yyyy-mm-dd") & "|" & Format(CDate("2024-03-04 5:06 PM"), "yyyy-mm-dd hh:nn") & '
                     '"|" & CDbl("1,234.5") & "|" & CLng("&HFF") & "|" & CInt(2.5) & CInt(3.5)',
    'string_functions_more': 's = StrConv("hello world", vbProperCase) & "|" & InStrRev("a.b.c", ".") & "|" & '
                             'Replace("aaaa", "a", "b", 2, 2) & "|" & StrReverse("abc") & "|" & String(3, "x") & Space(2) & "|" & '
                             'LCase(UCase("MiXeD")) & "|" & Len(Trim("  a  ")) & "|" & StrComp("a", "A", vbTextCompare)',
    'format_named_dates': 's = Format(DateSerial(2024, 1, 5), "Long Date") & "|" & Format(DateSerial(2024, 1, 5), "dddd, mmm d") & '
                          '"|" & Format(TimeSerial(14, 5, 9), "Long Time") & "|" & Format(DateSerial(2024, 1, 5), "Short Date") & '
                          '"|" & Format(0.75, "hh:mm AM/PM")',
    'date_arithmetic': 's = DateDiff("m", DateSerial(2024, 1, 31), DateSerial(2024, 3, 1)) & "|" & '
                       'DateAdd("q", 1, DateSerial(2024, 11, 30)) & "|" & DatePart("ww", DateSerial(2024, 1, 7)) & "|" & '
                       'WeekdayName(Weekday(DateSerial(2024, 1, 7))) & "|" & MonthName(2, True) & "|" & '
                       'DateDiff("d", DateSerial(2024, 2, 28), DateSerial(2024, 3, 1))',
    'choose_switch_iif': 's = Choose(2, "a", "b", "c") & "|" & Switch(False, 1, True, 2) & "|" & IIf(1 > 2, "y", "n") & "|" & '
                         'IsNull(Choose(5, "a"))',
    'variant_types': 'Dim e As Variant\ns = VarType(e) & "|" & VarType(Null) & "|" & VarType(1&) & "|" & VarType(1.5) & "|" & '
                     'VarType("x") & "|" & VarType(Array(1)) & "|" & TypeName(CDec(1)) & "|" & TypeName(CCur(1)) & "|" & '
                     'TypeName(Nothing) & "|" & IsMissing(e)',
    'redim_preserve': 'Dim a() As Long\nReDim a(1 To 3)\na(2) = 5\nReDim Preserve a(1 To 5)\ns = UBound(a) & "|" & a(2) & "|"\n'
                      'Erase a\nOn Error Resume Next\nn = UBound(a)\ns = s & Err.Number',
    'error_handling_flow': 'On Error GoTo Handler\nErr.Raise 1001, "Sweep", "custom"\ns = "not reached"\nGoTo Done\n'
                           'Handler:\ns = Err.Number & "|" & Err.Source & "|" & Err.Description\nResume Done\nDone:\n'
                           'On Error Resume Next\nn = 1 / 0\ns = s & "|" & Err.Number & "|" & Err.Description',
    'collection_keys': 'Set o = New Collection\no.Add "one", "k1"\no.Add "two", "k2"\no.Add "zero", "k0", Before:=1\n'
                       's = o(1) & "|" & o("k2") & "|" & o.Count & "|"\nOn Error Resume Next\no.Add "dup", "k1"\n'
                       's = s & Err.Number\nErr.Clear\nv = o("nokey")\ns = s & "|" & Err.Number',
    'dictionary_compare_text': 'Set o = CreateObject("Scripting.Dictionary")\no.CompareMode = vbTextCompare\no.Add "Key", 1\n'
                               's = o.Exists("KEY") & "|" & o.Count & "|"\no("key") = 5\ns = s & o("KEY") & "|" & Join(o.Keys, ",") & '
                               '"|" & o.Item("missing") & "|" & o.Count',
    'regexp_basic': 'Set o = CreateObject("VBScript.RegExp")\no.Pattern = "(\\d+)-(\\d+)": o.Global = True\n'
                    's = o.Test("a 12-34 b") & "|" & o.Replace("12-34 and 5-6", "$2:$1") & "|"\n'
                    'Set v = o.Execute("1-2, 33-44")\ns = s & v.Count & "|" & v(1).SubMatches(0) & "|" & v(0).FirstIndex',
    'fso_path_parts': 'Set o = CreateObject("Scripting.FileSystemObject")\n'
                      's = o.GetBaseName("C:\\a\\b\\report.final.xlsx") & "|" & o.GetExtensionName("C:\\a\\b\\report.final.xlsx") & '
                      '"|" & o.BuildPath("C:\\a", "b.txt") & "|" & o.GetFileName("C:\\a\\b\\c.txt") & "|" & '
                      'o.GetParentFolderName("C:\\a\\b\\c.txt")',
    'like_patterns_more': 's = ("abc" Like "a?c") & ("a1c" Like "a#c") & ("abc" Like "a[!b]c") & ("aXc" Like "a[A-Z]c") & '
                          '("ab" Like "a*b*") & ("" Like "*") & ("a*" Like "a[*]")',
    'number_to_string_concat': 'ws.Range("A1").Value = 0.1: ws.Range("A2").Value = 1E+20: ws.Range("A3").Value = 123.45\n'
                               's = ws.Range("A1").Value & "|" & ws.Range("A2").Value & "|" & ws.Range("A3").Value * 2 & "|" & '
                               '(ws.Range("A1").Value + 0.2)',
    # --- cell values and formats
    'leading_zeros_text': 'ws.Range("A1").NumberFormat = "@": ws.Range("A1").Value = "00123"\n'
                          'ws.Range("A2").Value = "00123"\ns = TypeName(ws.Range("A1").Value) & "|" & ws.Range("A1").Text & "|" & '
                          'TypeName(ws.Range("A2").Value) & "|" & ws.Range("A2").Value',
    'number_formats_builtin': 'ws.Range("A1").Value = -1234.5: ws.Range("A1").NumberFormat = "#,##0;[Red](#,##0)"\n'
                              'ws.Range("A2").Value = 0.12345: ws.Range("A2").NumberFormat = "0.00%"\n'
                              'ws.Range("A3").Value = DateSerial(2024, 2, 9) + TimeSerial(7, 5, 0): '
                              'ws.Range("A3").NumberFormat = "mm/dd/yyyy hh:mm"\nws.Columns(1).ColumnWidth = 20\n'
                              's = ws.Range("A1").Text & "|" & ws.Range("A2").Text & "|" & ws.Range("A3").Text',
    'iso_date_string_typed': 'ws.Range("A1").Value = "2024-01-15": ws.Range("A2").Value = "15/01/2024": '
                             'ws.Range("A3").Value = "Jan 15, 2024"\ns = TypeName(ws.Range("A1").Value) & "|" & '
                             'TypeName(ws.Range("A2").Value) & "|" & TypeName(ws.Range("A3").Value) & "|" & ws.Range("A1").NumberFormat',
    'big_numbers_in_cells': 'ws.Range("A1").Value = 12345678901234567890#: ws.Range("A2").Value = 0.1 + 0.2\n'
                            'ws.Columns(1).ColumnWidth = 30\ns = ws.Range("A1").Text & "|" & ws.Range("A1").Value2 & "|" & '
                            'ws.Range("A2").Value & "|" & ws.Range("A2").Text',
    'formula_with_quotes': 'ws.Range("A1").Formula = "=""a""&""b""&LEN(""xyz"")"\n'
                           's = ws.Range("A1").Value & "|" & ws.Range("A1").Formula & "|" & Len(ws.Range("A1").Formula)',
    'long_text_cell': 'ws.Range("A1").Value = String(32767, "a")\ns = Len(ws.Range("A1").Value) & "|"\nOn Error Resume Next\n'
                      'ws.Range("A2").Value = String(32768, "b")\ns = s & Err.Number & "|" & Len(ws.Range("A2").Value)',
    'color_index_round_trip': 'ws.Range("A1").Font.ColorIndex = 3: ws.Range("A2").Interior.Color = RGB(0, 0, 255)\n'
                              's = ws.Range("A1").Font.Color & "|" & ws.Range("A2").Interior.ColorIndex & "|" & '
                              'ws.Range("A3").Interior.ColorIndex & "|" & ws.Range("A3").Font.ColorIndex',
    'border_around_weight': 'ws.Range("B2:C3").BorderAround Weight:=xlThick, ColorIndex:=5\n'
                            's = ws.Range("B2").Borders(xlEdgeTop).Weight & "|" & ws.Range("C3").Borders(xlEdgeRight).LineStyle & "|" & '
                            'ws.Range("B2").Borders(xlEdgeRight).LineStyle & "|" & ws.Range("B2").Borders(xlEdgeLeft).ColorIndex',
    'hidden_column_width': 'ws.Columns(2).ColumnWidth = 12: ws.Columns(2).Hidden = True\n'
                           's = ws.Columns(2).ColumnWidth & "|" & ws.Columns(2).Width & "|"\nws.Columns(2).Hidden = False\n'
                           's = s & ws.Columns(2).ColumnWidth',
    'autofit_text_and_numbers': 'ws.Range("A1").Value = "Quarterly revenue": ws.Range("A2").Value = 1234567.89: '
                                'ws.Range("B1").Value = 12345: ws.Range("B1").Font.Size = 16\nws.Columns("A:B").AutoFit\n'
                                's = ws.Columns(1).ColumnWidth & "|" & ws.Columns(2).ColumnWidth',
    'format_widens_column': 'ws.Range("A1").Value = 1234567.5: ws.Range("A1").NumberFormat = "$#,##0.00"\n'
                            'ws.Range("B1").Value = 0.5: ws.Range("B1").Style = "Percent"\n'
                            's = ws.Columns(1).ColumnWidth & "|" & ws.Range("A1").Text & "|" & ws.Columns(2).ColumnWidth & "|" & '
                            'ws.Range("B1").Text',
    'comma_style_text': 'ws.Range("A1").Value = 1234.5: ws.Range("A1").Style = "Comma"\n'
                        'ws.Range("A2").Value = -2: ws.Range("A2").Style = "Currency"\n'
                        's = "[" & ws.Range("A1").Text & "]|[" & ws.Range("A2").Text & "]"',
    'narrow_column_text': 'ws.Columns(1).ColumnWidth = 4: ws.Range("A1").Value = 123456: ws.Range("A2").Value = 3.14159: '
                          'ws.Range("A3").Value = DateSerial(2024, 1, 1)\n'
                          's = ws.Range("A1").Text & "|" & ws.Range("A2").Text & "|" & ws.Range("A3").Text',
    # --- validation, conditional formats, tables and charts
    'validation_list': 'ws.Range("A1").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="a,b,c"\n'
                       's = ws.Range("A1").Validation.Formula1 & "|" & ws.Range("A1").Validation.Type & "|" & '
                       'ws.Range("A1").Validation.InCellDropdown',
    'format_conditions_count': 'ws.Range("A1:A5").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="3"\n'
                               'ws.Range("A1:A5").FormatConditions.Add Type:=xlExpression, Formula1:="=A1=1"\n'
                               's = ws.Range("A1:A5").FormatConditions.Count & "|" & ws.Range("A1:A5").FormatConditions(1).Formula1 & "|"\n'
                               'ws.Range("A1:A5").FormatConditions(1).Delete\ns = s & ws.Range("A1:A5").FormatConditions.Count',
    'listobject_totals': 'ws.Range("A1:B1").Value = Array("Item", "Qty")\nws.Range("A2:B3").Value = 2\n'
                         'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)\no.ShowTotals = True\n'
                         'o.ListColumns(2).TotalsCalculation = xlTotalsCalculationSum\n'
                         's = o.Range.Address & "|" & ws.Range("B4").Value & "|" & o.ListRows.Count & "|" & (o.Name Like "Table*")',
    'chart_series_count': 'ws.Range("A1:B3").Value = 1\nSet o = ws.ChartObjects.Add(10, 10, 200, 120)\n'
                          'o.Chart.SetSourceData ws.Range("A1:B3")\ns = o.Chart.SeriesCollection.Count & "|" & ws.ChartObjects.Count',
    'print_area_name': 'ws.PageSetup.PrintArea = "$A$1:$C$5"\ns = ws.PageSetup.PrintArea & "|" & ws.Names("Print_Area").RefersTo',
    # --- counting and addressing
    'count_large_and_areas': 'Set r = Union(ws.Range("A1:B2"), ws.Range("D4"), ws.Range("F1:F3"))\nFor Each c In r.Areas\n'
                             's = s & c.Address(False, False) & ";"\nNext\n'
                             's = s & "|" & r.Count & "|" & ws.Columns(1).Cells.CountLarge & "|" & ws.Range("A:B").Columns.Count',
    'cells_property_forms': 's = ws.Cells(3, "C").Address & "|" & ws.Range("B2").Cells(2, 2).Address & "|" & '
                            'ws.Range("B2:D4")(5).Address & "|" & ws.Range("B2:D4").Cells(10).Address & "|" & ws.Range("C3").Item(1, 0).Address',
    'entire_row_column_props': 's = ws.Range("B3").EntireRow.Address & "|" & ws.Range("B3:C4").EntireColumn.Address & "|" & '
                               'ws.Range("B3").Row & ws.Range("B3").Column & "|" & ws.Range("B3:D7").Rows.Count & ws.Range("B3:D7").Columns.Count',
    'value_assign_shapes': 'ws.Range("A1:B2").Value = Array(1, 2)\ns = ws.Range("A2").Value & ws.Range("B2").Value & "|"\n'
                           'ws.Range("D1:D3").Value = Array(5, 6, 7)\ns = s & ws.Range("D1").Value & ws.Range("D2").Value & ws.Range("D3").Value',
    'formula_array_and_ctrl': 'ws.Range("A1:A3").Value = WorksheetFunction.Transpose(Array(1, 2, 3))\n'
                              'ws.Range("B1").FormulaArray = "=SUM(A1:A3*2)"\ns = ws.Range("B1").Value & "|" & ws.Range("B1").HasArray & "|" & '
                              'ws.Range("B1").FormulaArray',
    'used_range_after_format': 'ws.Range("C5").Interior.Color = RGB(1, 2, 3)\ns = ws.UsedRange.Address & "|"\n'
                               'ws.Range("C5").ClearFormats\ns = s & ws.UsedRange.Address',
}


def macro(name: str) -> str:
    cleanup = "\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "Function Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
