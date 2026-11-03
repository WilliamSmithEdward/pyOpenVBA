"""An eighth sweep of everyday Excel macros: more of what ends Copy and Cut, the members next to those the seventh found.

    python scripts/measure_excel_everyday_8.py

writes tests/fixtures/excel_everyday_8.json, which tests/test_excel_everyday_7.py replays with the seventh.
"""
from __future__ import annotations

from pathlib import Path

from measure_excel_everyday_7 import AFTER, LIST, PRE
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_8.json'
EDITS = {
    'font_color': 'ws.Range("Z1").Font.Color = 255',
    'font_color_index': 'ws.Range("Z1").Font.ColorIndex = 3',
    'font_name': 'ws.Range("Z1").Font.Name = "Arial"',
    'font_size': 'ws.Range("Z1").Font.Size = 14',
    'font_italic': 'ws.Range("Z1").Font.Italic = True',
    'font_underline': 'ws.Range("Z1").Font.Underline = xlUnderlineStyleSingle',
    'interior_pattern': 'ws.Range("Z1").Interior.Pattern = xlSolid',
    'interior_color_index': 'ws.Range("Z1").Interior.ColorIndex = 3',
    'interior_clear': 'ws.Range("Z1").Interior.Pattern = xlNone',
    'horizontal': 'ws.Range("Z1").HorizontalAlignment = xlCenter',
    'vertical': 'ws.Range("Z1").VerticalAlignment = xlTop',
    'wrap': 'ws.Range("Z1").WrapText = True',
    'indent': 'ws.Range("Z1").IndentLevel = 1',
    'add_indent': 'ws.Range("Z1").AddIndent = True',
    'orientation': 'ws.Range("Z1").Orientation = 45',
    'shrink': 'ws.Range("Z1").ShrinkToFit = True',
    'formula_hidden': 'ws.Range("Z1").FormulaHidden = True',
    'border_one': 'ws.Range("Z1").Borders(xlEdgeTop).LineStyle = xlContinuous',
    'border_color': 'ws.Range("Z1").Borders.Color = 255',
    'border_around': 'ws.Range("Z1").BorderAround xlContinuous',
    'characters': 'ws.Range("A1").Characters(1, 1).Font.Bold = True',
    'number_format_local': 'ws.Range("Z1").NumberFormatLocal = "0.00"',
    'entire_font': 'ws.Rows(20).Font.Bold = True',
    'calculate_full': 'Application.CalculateFull',
    'calculate_rebuild': 'Application.CalculateFullRebuild',
    'enable_calculation': 'ws.EnableCalculation = False\nws.EnableCalculation = True',
    'dirty': 'ws.Range("A1").Dirty',
    'calc_same_mode': 'Application.Calculation = xlCalculationAutomatic',
    'calc_iteration': 'Application.Iteration = False',
    'clear_comments': 'ws.Range("Z1").ClearComments',
    'clear_notes': 'ws.Range("Z1").ClearNotes',
    'clear_hyperlinks': 'ws.Range("Z1").ClearHyperlinks',
    'clear_outline': 'ws.Range("Z1").ClearOutline',
    'note_delete': 'ws.Range("Z1").AddComment "x"\nws.Range("Z1").Comment.Delete',
    'hyperlinks_delete': 'ws.Hyperlinks.Add ws.Range("Z1"), "https://example.com"\nws.Hyperlinks.Delete',
    'validation_delete': 'ws.Range("Z1").Validation.Delete',
    'format_condition_delete': 'ws.Range("Z1").FormatConditions.Delete',
    'ungroup': 'ws.Rows(20).Group\nws.Rows(20).Ungroup',
    'show_levels': 'ws.Outline.ShowLevels 1',
    'sort_object': 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add ws.Range("D2:D5")\nws.Sort.SetRange ws.Range("A1:E5")\nws.Sort.Header = xlYes\nws.Sort.Apply',
    'sort_fields_only': 'ws.Sort.SortFields.Clear',
    'show_all_data': 'ws.Range("A1:E5").AutoFilter\nOn Error Resume Next\nws.ShowAllData\nOn Error GoTo 0',
    'autofilter_off': 'ws.AutoFilterMode = False',
    'advanced_filter': 'ws.Range("A1:E5").AdvancedFilter xlFilterCopy, , ws.Range("R1")',
    'fill_right': 'ws.Range("Y1").Value = 1\nws.Range("Y1:Z1").FillRight',
    'fill_up': 'ws.Range("Z3").Value = 1\nws.Range("Z1:Z3").FillUp',
    'fill_left': 'ws.Range("Z1").Value = 1\nws.Range("Y1:Z1").FillLeft',
    'data_series': 'ws.Range("Z1").Value = 1\nws.Range("Z1:Z3").DataSeries xlColumns, xlLinear, , 1',
    'justify': 'ws.Range("Z1").Justify',
    'consolidate': 'ws.Range("R1").Consolidate "Sweep!R1C4:R5C4", xlSum',
    'sheet_add_after': 'Set ws2 = Worksheets.Add(After:=ws)\nws.Activate\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'sheet_move': 'Set ws2 = Worksheets.Add\nws2.Move After:=ws\nSet ws2 = ActiveSheet\nws.Activate\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'sheet_visible': 'Set ws2 = Worksheets.Add\nws2.Visible = xlSheetHidden\nws.Activate\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'activate_other_sheet': 'Worksheets(Worksheets.Count).Activate\nws.Activate',
    'protect_only': 'ws.Protect',
    'workbook_protect': 'ThisWorkbook.Protect\nThisWorkbook.Unprotect',
    'table_resize': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.Resize ws.Range("A1:E6")',
    'table_row_add': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.ListRows.Add',
    'table_row_delete': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.ListRows(4).Delete',
    'table_column_add': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.ListColumns.Add',
    'table_style': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.TableStyle = "TableStyleLight1"',
    'table_totals': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.ShowTotals = True',
    'table_unlist': 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E5"), , xlYes)\no.Unlist',
    'shape_delete': 'Set o = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50)\no.Delete',
    'shape_move': 'Set o = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50)\no.Left = 40',
    'chart_set_data': 'Set o = ws.ChartObjects.Add(10, 10, 200, 120)\no.Chart.SetSourceData ws.Range("D1:E5")',
    'pivot': 'Set ws2 = Worksheets.Add(After:=ws)\nSet o = ThisWorkbook.PivotCaches.Create(xlDatabase, ws.Range("A1:E5")).CreatePivotTable(ws2.Range("A3"))\nws.Activate\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'names_delete': 'ThisWorkbook.Names.Add "SweepName", ws.Range("A1")\nThisWorkbook.Names("SweepName").RefersTo = "=1"\nThisWorkbook.Names("SweepName").Delete',
    'page_break_reset': 'ws.ResetAllPageBreaks',
    'page_setup_margin': 'ws.PageSetup.LeftMargin = 40',
    'page_setup_header': 'ws.PageSetup.CenterHeader = "x"',
    'print_area': 'ws.PageSetup.PrintArea = "$A$1:$B$2"',
    'row_autofit': 'ws.Rows(1).AutoFit',
    'hide_column': 'ws.Columns(26).Hidden = True',
    'style_add': 'ThisWorkbook.Styles.Add "SweepStyle"\nThisWorkbook.Styles("SweepStyle").Delete',
    'formula_array': 'ws.Range("Z1").FormulaArray = "=1+1"',
    'formula_r1c1': 'ws.Range("Z1").FormulaR1C1 = "=1+1"',
    'value2': 'ws.Range("Z1").Value2 = 1',
    'row_delete_in_table_free': 'ws.Range("Z1").Value = ""',
}
CASES = {
    **{f'copy_then_{name}': f'{LIST}ws.Range("B3").Copy\n{edit}{AFTER}' for name, edit in EDITS.items()},
    **{f'cut_then_{name}': f'{LIST}ws.Range("B3").Cut\n{edit}{AFTER}' for name, edit in EDITS.items()},
}


def macro(name: str) -> str:
    cleanup = ("\nOn Error Resume Next\nws.Unprotect\nOn Error GoTo 0\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\n"
               "ws.Delete\nApplication.DisplayAlerts = True")
    return "\nFunction Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
