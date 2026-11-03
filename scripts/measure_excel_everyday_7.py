"""A seventh sweep of everyday Excel macros: what ends Copy and Cut. Each ordinary edit runs with a cell on the clipboard, and
the probe reads CutCopyMode and pastes what is left, so an edit that ends the mode, or changes what comes out, shows.

Like scripts/measure_excel_everyday.py, each probe works on a sheet of its own that it deletes afterwards, and is recorded
with its answer or the run-time error it stopped with (see vba_probe_fixture.py).

    python scripts/measure_excel_everyday_7.py

writes tests/fixtures/excel_everyday_7.json, which tests/test_excel_everyday_7.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_7.json'
PRE = ("Dim ws As Worksheet, ws2 As Worksheet, r As Range, c As Range, s As String, i As Long, j As Long, n As Long\n"
       "Dim v As Variant, o As Object, a As Variant, t As String, d As Double, k As Variant\nSet ws = Worksheets.Add\n"
       "ws.Name = \"Sweep\"\n")
LIST = ('ws.Range("A1:E1").Value = Array("Cust", "Region", "Item", "Qty", "Price")\n'
        'ws.Range("A2:E2").Value = Array("Ann", "North", "Pen", 10, 1.5)\n'
        'ws.Range("A3:E3").Value = Array("Bob", "South", "Ink", 3, 12.25)\n'
        'ws.Range("A4:E4").Value = Array("Cy", "North", "Pad", 7, 4)\n'
        'ws.Range("A5:E5").Value = Array("Di", "East", "Pen", 20, 1.5)\n')
#: Reads the mode, then pastes at H1 whatever is left, and says where the source ended up.
AFTER = ('\ns = Application.CutCopyMode\nOn Error Resume Next\nws.Paste ws.Range("H1")\ns = s & "|" & Err.Number & "|" & ws.Range("H1").Value & '
         '"|" & ws.Range("B3").Value\nErr.Clear\nOn Error GoTo 0')
#: An edit away from the cells on the clipboard, and one that reaches them.
EDITS = {
    'value': 'ws.Range("Z1").Value = 1',
    'formula': 'ws.Range("Z1").Formula = "=1+1"',
    'value_source': 'ws.Range("B3").Value = "new"',
    'clear_contents': 'ws.Range("Z1").ClearContents',
    'clear_contents_source': 'ws.Range("B3").ClearContents',
    'clear': 'ws.Range("Z1").Clear',
    'clear_formats': 'ws.Range("Z1").ClearFormats',
    'bold': 'ws.Range("Z1").Font.Bold = True',
    'interior': 'ws.Range("Z1").Interior.Color = 255',
    'number_format': 'ws.Range("Z1").NumberFormat = "0.00"',
    'column_width': 'ws.Columns(26).ColumnWidth = 20',
    'row_height': 'ws.Rows(20).RowHeight = 30',
    'hide_row': 'ws.Rows(20).Hidden = True',
    'autofit': 'ws.Columns(1).AutoFit',
    'merge': 'ws.Range("Y1:Z1").Merge',
    'unmerge': 'ws.Range("Y1:Z1").UnMerge',
    'select': 'ws.Range("Z1").Select',
    'activate_cell': 'ws.Activate\nws.Range("A1").Activate',
    'sheet_add': 'Set ws2 = Worksheets.Add\nws.Activate\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'sheet_rename': 'ws.Name = "Sweep2"\nws.Name = "Sweep"',
    'sheet_copy': 'ws.Copy After:=ws\nSet ws2 = ActiveSheet\nws.Activate\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'sheet_hide_tab': 'ws.Tab.Color = 255',
    'calculate': 'Application.Calculate',
    'calculate_sheet': 'ws.Calculate',
    'calculate_range': 'ws.Range("A1:B2").Calculate',
    'calc_mode': 'Application.Calculation = xlCalculationManual\nApplication.Calculation = xlCalculationAutomatic',
    'screen': 'Application.ScreenUpdating = False\nApplication.ScreenUpdating = True',
    'events': 'Application.EnableEvents = False\nApplication.EnableEvents = True',
    'name_add': 'ThisWorkbook.Names.Add "SweepName", ws.Range("A1")\nThisWorkbook.Names("SweepName").Delete',
    'hyperlink': 'ws.Hyperlinks.Add ws.Range("Z1"), "https://example.com"',
    'note': 'ws.Range("Z1").AddComment "x"',
    'validation': 'ws.Range("Z1").Validation.Add xlValidateWholeNumber, , xlBetween, "1", "5"',
    'format_condition': 'ws.Range("Z1").FormatConditions.Add xlCellValue, xlGreater, "=1"',
    'sort': 'ws.Range("A1:E5").Sort ws.Range("D1"), xlDescending, Header:=xlYes',
    'autofilter': 'ws.Range("A1:E5").AutoFilter 2, "North"\nws.AutoFilterMode = False',
    'remove_duplicates': 'ws.Range("A1:E5").RemoveDuplicates 2, xlYes',
    'table': 'ws.ListObjects.Add xlSrcRange, ws.Range("A1:E5"), , xlYes',
    'group': 'ws.Rows(20).Group',
    'shape': 'ws.Shapes.AddShape msoShapeRectangle, 10, 10, 50, 50',
    'chart': 'ws.ChartObjects.Add 10, 10, 200, 120',
    'copy_elsewhere': 'ws.Range("A1").Copy ws.Range("Z1")',
    'cut_elsewhere': 'ws.Range("A1").Cut ws.Range("Z1")',
    'paste_special_values': 'ws.Range("Z1").PasteSpecial xlPasteValues',
    'paste_twice': 'ws.Paste ws.Range("Z1")',
    'protect': 'ws.Protect\nws.Unprotect',
    'workbook_add': 'Set o = Workbooks.Add\no.Close False\nws.Activate',
    'find': 'Set r = ws.Cells.Find("Pad", LookIn:=xlValues, LookAt:=xlWhole)',
    'replace': 'ws.Cells.Replace "Pad", "Pod", xlWhole',
    'autofill': 'ws.Range("Z1").Value = 1\nws.Range("Z1").AutoFill ws.Range("Z1:Z3")',
    'fill_down': 'ws.Range("Z1").Value = 1\nws.Range("Z1:Z3").FillDown',
    'text_to_columns': 'ws.Range("Y1").Value = "a b"\nws.Range("Y1").TextToColumns ws.Range("Y1"), Space:=True',
    'resize_nothing': 'Set r = ws.Range("B3").Resize(2, 2)',
    'cells_value_read': 'v = ws.Range("A1:E5").Value',
    'array_write': 'ws.Range("Y1:Z1").Value = Array(1, 2)',
    'range_style': 'ws.Range("Z1").Style = "Good"',
    'borders': 'ws.Range("Z1").Borders.LineStyle = xlContinuous',
    'freeze_nothing': 'ws.Range("Z1").Locked = False',
    'subtotal': 'ws.Range("A1:E5").Subtotal 2, xlSum, Array(4)',
    'page_setup': 'ws.PageSetup.Orientation = xlLandscape',
    'page_break': 'ws.HPageBreaks.Add ws.Range("A20")',
}
CASES = {
    **{f'copy_then_{name}': f'{LIST}ws.Range("B3").Copy\n{edit}{AFTER}' for name, edit in EDITS.items()},
    **{f'cut_then_{name}': f'{LIST}ws.Range("B3").Cut\n{edit}{AFTER}' for name, edit in EDITS.items()},
    # The mode itself.
    'mode_set_values': LIST + 'ws.Range("B3").Copy\nApplication.CutCopyMode = xlCut\ns = Application.CutCopyMode\nApplication.CutCopyMode = xlCopy\ns = s & Application.CutCopyMode\nApplication.CutCopyMode = True\ns = s & Application.CutCopyMode\nApplication.CutCopyMode = False\ns = s & Application.CutCopyMode',
    'mode_set_nothing_on_it': 'Application.CutCopyMode = False\ns = Application.CutCopyMode\nApplication.CutCopyMode = xlCopy\ns = s & Application.CutCopyMode\nApplication.CutCopyMode = True\ns = s & Application.CutCopyMode',
    'paste_nothing_on_it': 'Application.CutCopyMode = False\nOn Error Resume Next\nws.Paste ws.Range("A1")\ns = Err.Number & "|" & ws.Range("A1").Value',
    'paste_special_nothing_on_it': 'Application.CutCopyMode = False\nOn Error Resume Next\nws.Range("A1").PasteSpecial xlPasteValues\ns = Err.Number & "|" & ws.Range("A1").Value',
    'paste_keeps_copy': LIST + 'ws.Range("B3").Copy\nws.Paste ws.Range("H1")\ns = Application.CutCopyMode\nws.Paste ws.Range("H2")\ns = s & ws.Range("H1").Value & ws.Range("H2").Value & Application.CutCopyMode',
    'paste_ends_cut': LIST + 'ws.Range("B3").Cut\nws.Paste ws.Range("H1")\ns = Application.CutCopyMode\nOn Error Resume Next\nws.Paste ws.Range("H2")\ns = s & Err.Number & ws.Range("H1").Value & ws.Range("H2").Value',
    'paste_special_on_a_cut': LIST + 'ws.Range("B3").Cut\nOn Error Resume Next\nws.Range("H1").PasteSpecial xlPasteValues\ns = Err.Number & "|" & Application.CutCopyMode & ws.Range("H1").Value & ws.Range("B3").Value',
    'paste_selection_after': LIST + 'ws.Activate\nws.Range("B3").Copy\nws.Range("H1").Select\nws.Paste\ns = Selection.Address & ActiveCell.Address\nws.Range("A1:A4").Copy\nws.Range("J1").Select\nws.Paste\ns = s & Selection.Address & ActiveCell.Address',
    'paste_selection_grows': LIST + 'ws.Activate\nws.Range("B3:C3").Copy\nws.Range("H1:K2").Select\nws.Paste\ns = Selection.Address & ws.Range("K2").Value',
    'paste_special_selection_after': LIST + 'ws.Activate\nws.Range("A1:A4").Copy\nws.Range("J1").Select\nSelection.PasteSpecial xlPasteValues\ns = Selection.Address & ActiveCell.Address',
    'copy_destination_keeps_clipboard': LIST + 'ws.Range("B3").Copy\nws.Range("A1").Copy ws.Range("H1")\ns = Application.CutCopyMode\nOn Error Resume Next\nws.Paste ws.Range("H2")\ns = s & Err.Number & ws.Range("H1").Value & ws.Range("H2").Value',
    'selection_copy': LIST + 'ws.Activate\nws.Range("B3:C3").Select\nSelection.Copy\nws.Range("H1").Select\nActiveSheet.Paste\ns = ws.Range("H1").Value & ws.Range("I1").Value & Selection.Address',
    'selection_cut': LIST + 'ws.Activate\nws.Range("B3:C3").Select\nSelection.Cut\nws.Range("H1").Select\nActiveSheet.Paste\ns = ws.Range("H1").Value & ws.Range("I1").Value & Selection.Address & ws.Range("B3").Value',
    'copy_then_source_changes': LIST + 'ws.Range("B3").Copy\nws.Range("B3").Value = "new"\nws.Paste ws.Range("H1")\ns = Application.CutCopyMode & ws.Range("H1").Value',
    'copy_formula_then_paste': LIST + 'ws.Range("F2").Formula = "=D2*E2"\nws.Range("F2").Copy\nws.Range("D2").Value = 2\nws.Paste ws.Range("F3")\ns = ws.Range("F3").Formula & ws.Range("F3").Value',
    'copy_other_book_closed': 'Set o = Workbooks.Add\no.Worksheets(1).Range("A1").Value = 5\no.Worksheets(1).Range("A1").Copy\no.Close False\nws.Activate\ns = Application.CutCopyMode\nOn Error Resume Next\nws.Paste ws.Range("A1")\ns = s & "|" & Err.Number & "|" & ws.Range("A1").Value',
    'copy_sheet_deleted': LIST + 'Set ws2 = Worksheets.Add\nws2.Range("A1").Value = 5\nws2.Range("A1").Copy\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True\nws.Activate\ns = Application.CutCopyMode\nOn Error Resume Next\nws.Paste ws.Range("H1")\ns = s & "|" & Err.Number & "|" & ws.Range("H1").Value',
    'cut_sheet_deleted': LIST + 'Set ws2 = Worksheets.Add\nws2.Range("A1").Value = 5\nws2.Range("A1").Cut\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True\nws.Activate\ns = Application.CutCopyMode\nOn Error Resume Next\nws.Paste ws.Range("H1")\ns = s & "|" & Err.Number & "|" & ws.Range("H1").Value',
}


def macro(name: str) -> str:
    cleanup = ("\nOn Error Resume Next\nws.Unprotect\nOn Error GoTo 0\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\n"
               "ws.Delete\nApplication.DisplayAlerts = True")
    return "\nFunction Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
