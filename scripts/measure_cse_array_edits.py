"""Measure native legacy CSE array copies, cuts and refused structural edits."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/cse_array_edits.json'
HELPER = '''Function ArrayAddress(cell As Range) As String
On Error GoTo Bad
ArrayAddress = cell.CurrentArray.Address(False,False)
Exit Function
Bad:
ArrayAddress = "!" & Err.Number
End Function
Function ReadSaved() As String
Dim cell As Range
ReadSaved = "0"
For Each cell In ActiveSheet.Range("A1:H8")
ReadSaved = ReadSaved & "|" & cell.Address(False,False) & ":" & CStr(cell.Value2) & ":" & cell.Formula & ":" & cell.HasArray & ":" & ArrayAddress(cell)
Next
End Function
'''
OPERATIONS = {
    'copy_whole': 'ws.Range("B2:C4").Copy Destination:=ws.Range("E2")',
    'copy_anchor': 'ws.Range("B2").Copy Destination:=ws.Range("E2")',
    'copy_child': 'ws.Range("B3").Copy Destination:=ws.Range("E2")',
    'copy_right_child': 'ws.Range("C3").Copy Destination:=ws.Range("E2")',
    'copy_partial': 'ws.Range("B2:B3").Copy Destination:=ws.Range("E2")',
    'copy_children': 'ws.Range("B3:C4").Copy Destination:=ws.Range("E2")',
    'copy_enclosing': 'ws.Range("A1:D5").Copy Destination:=ws.Range("E1")',
    'copy_tiled': 'ws.Range("B3:C4").Copy Destination:=ws.Range("E2:H5")',
    'copy_overlap': 'ws.Range("B2:C4").Copy Destination:=ws.Range("C2")',
    'copy_self': 'ws.Range("B2:C4").Copy Destination:=ws.Range("B2")',
    'replace_whole': 'ws.Range("E2:F4").Value = 9\nws.Range("E2:F4").Copy Destination:=ws.Range("B2")',
    'replace_child': 'ws.Range("E2").Value = 9\nws.Range("E2").Copy Destination:=ws.Range("B3")',
    'cut_whole': 'ws.Range("B2:C4").Cut Destination:=ws.Range("E2")',
    'cut_child': 'ws.Range("B3").Cut Destination:=ws.Range("E2")',
    'cut_overlap': 'ws.Range("B2:C4").Cut Destination:=ws.Range("C2")',
    'cut_over_array': 'ws.Range("E2:F4").Value = 9\nws.Range("E2:F4").Cut Destination:=ws.Range("B2")',
    'cut_over_child': 'ws.Range("E2").Value = 9\nws.Range("E2").Cut Destination:=ws.Range("B3")',
    'insert_above': 'ws.Rows(1).Insert',
    'insert_inside': 'ws.Rows(3).Insert',
    'delete_inside': 'ws.Rows(3).Delete',
    'delete_whole': 'ws.Rows("2:4").Delete',
    'shift_inside': 'ws.Range("B3").Insert Shift:=xlShiftDown',
    'shift_anchor': 'ws.Range("B2").Insert Shift:=xlShiftDown',
    'sort_whole': 'ws.Range("B2:C4").Sort Key1:=ws.Range("B2"), Order1:=xlDescending, Header:=xlNo',
    'sort_children': 'ws.Range("B3:C4").Sort Key1:=ws.Range("B3"), Order1:=xlDescending, Header:=xlNo',
    'sort_object': 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B2"), Order:=xlDescending\nws.Sort.SetRange ws.Range("B2:C4")\nws.Sort.Header = xlNo\nws.Sort.Apply',
    'clipboard_copy_child': 'ws.Range("B3").Copy\nws.Paste Destination:=ws.Range("E2")',
    'clipboard_cut_whole': 'ws.Range("B2:C4").Cut\nws.Paste Destination:=ws.Range("E2")',
    'shift_whole_down': 'ws.Range("B1:C1").Insert Shift:=xlShiftDown',
    'shift_whole_right': 'ws.Range("A2:A4").Insert Shift:=xlShiftToRight',
    'delete_whole_cells': 'ws.Range("B2:C4").Delete Shift:=xlShiftUp',
    'delete_before': 'ws.Range("B1:C1").Delete Shift:=xlShiftUp',
    'delete_leading_part': 'ws.Range("B2:C2").Delete Shift:=xlShiftUp',
    'two_arrays_copy': 'ws.Range("E2:F4").FormulaArray = "=2*ROW(A1:A3)"\nApplication.Calculate\nws.Range("B2:F4").Copy Destination:=ws.Range("B5")',
    'replace_array_with_array': 'ws.Range("E2:F4").FormulaArray = "=2*ROW(A1:A3)"\nApplication.Calculate\nws.Range("E2:F4").Copy Destination:=ws.Range("B2")',
    'cross_sheet_copy': 'Dim other As Worksheet\nSet other = Worksheets.Add\nws.Range("B3:C4").Copy Destination:=other.Range("E2")',
    'cross_sheet_cut': 'Dim other As Worksheet\nSet other = Worksheets.Add\nws.Range("B2:C4").Cut Destination:=other.Range("E2")',
    'single_array_copy': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").FormulaArray = "=SUM(A1:A3)"\nApplication.Calculate\nws.Range("B2").Copy Destination:=ws.Range("E2")',
    'sort_single_array_object': 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B2"), Order:=xlDescending\nws.Sort.SetRange ws.Range("B2")\nws.Sort.Header = xlNo\nws.Sort.Apply',
    'sort_header_only': 'ws.Range("B2:C4").ClearContents\nws.Range("B2:C2").FormulaArray = "=COLUMN(A1:B1)"\nws.Range("B3").Value = 2\nws.Range("B4").Value = 1\nws.Range("B2:C4").Sort Key1:=ws.Range("B2"), Header:=xlYes',
    'paste_all_child': 'ws.Range("B3").Copy\nws.Range("E2").PasteSpecial xlPasteAll',
    'paste_formulas_child': 'ws.Range("B3").Copy\nws.Range("E2").PasteSpecial xlPasteFormulas',
    'paste_values_child': 'ws.Range("B3").Copy\nws.Range("E2").PasteSpecial xlPasteValues',
    'paste_all_whole': 'ws.Range("B2:C4").Copy\nws.Range("E2").PasteSpecial xlPasteAll',
    'paste_transposed': 'ws.Range("B2:C4").Copy\nws.Range("E2").PasteSpecial Transpose:=True',
    'paste_values_over_child': 'ws.Range("E2").Value = 9\nws.Range("E2").Copy\nws.Range("B3").PasteSpecial xlPasteValues',
    'paste_add_child': 'ws.Range("B3").Copy\nws.Range("E2").PasteSpecial Operation:=xlPasteSpecialOperationAdd',
    'paste_add_whole': 'ws.Range("E2:F4").Value = 5\nws.Range("B2:C4").Copy\nws.Range("E2").PasteSpecial Operation:=xlPasteSpecialOperationAdd',
    'paste_skip_child': 'ws.Range("B3").Copy\nws.Range("E2").PasteSpecial SkipBlanks:=True',
    'paste_transposed_child': 'ws.Range("B3:C4").Copy\nws.Range("E2").PasteSpecial Transpose:=True',
    'paste_formulas_over_whole': 'ws.Range("E2:F4").Value = 9\nws.Range("E2:F4").Copy\nws.Range("B2").PasteSpecial xlPasteFormulas',
    'paste_values_whole': 'ws.Range("B2:C4").Copy\nws.Range("E2").PasteSpecial xlPasteValues',
    'paste_add_over_array': 'ws.Range("E2:F4").Value = 9\nws.Range("E2:F4").Copy\nws.Range("B2").PasteSpecial Operation:=xlPasteSpecialOperationAdd',
    'paste_formats_child': 'ws.Range("E2").Font.Bold = True\nws.Range("E2").Copy\nws.Range("B3").PasteSpecial xlPasteFormats',
    'copy_invalid_column_ref': 'ws.Range("E2:F4").FormulaArray = "=COLUMN(A1:B1)"\nApplication.Calculate\nws.Range("E2:F4").Copy Destination:=ws.Range("B2")',
    'array_error_argument': 'ws.Range("E2:F4").FormulaArray = "=ROW(1/0)"',
    'replace_dynamic_whole': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nApplication.Calculate\nws.Range("B2:C4").FormulaArray = "=ROW(A1:A3)"',
    'replace_dynamic_anchor': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nApplication.Calculate\nws.Range("B2").FormulaArray = "=SUM(1,2)"',
    'replace_dynamic_child': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nApplication.Calculate\nws.Range("B3").FormulaArray = "=SUM(1,2)"',
    'replace_dynamic_styled': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").NumberFormat = "0.00"\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nApplication.Calculate\nws.Range("B2:C4").FormulaArray = "=ROW(A1:A3)"',
    'replace_dynamic_partial': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nApplication.Calculate\nws.Range("B3:C4").FormulaArray = "=SUM(1,2)"',
    'replace_dynamic_enclosing': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nApplication.Calculate\nws.Range("A1:D5").FormulaArray = "=ROW(A1:A5)"',
    'replace_dynamic_anchor_band': 'ws.Range("B2:C4").ClearContents\nws.Range("B2").Formula2 = "=SEQUENCE(3,2)"\nws.Range("G1").Formula = "=C4*10"\nApplication.Calculate\nws.Range("B2:B3").FormulaArray = "=SUM(1,2)"',
}


def macro(operation: str, *, manual: bool, calculate: bool) -> str:
    return HELPER + '''Function Probe() As String
Dim ws As Worksheet, n As Long
Application.Calculation = xlCalculationAutomatic
Set ws = Worksheets.Add
ws.Range("B2:C4").FormulaArray = "=ROW(A1:A3)*10+COLUMN(A1:B1)"
Application.Calculate
''' + ('Application.Calculation = xlCalculationManual\n' if manual else '') + '''On Error Resume Next
''' + operation + '''
n = Err.Number
On Error GoTo 0
''' + ('Application.Calculate\n' if calculate else '') + '''Probe = CStr(n) & Mid(ReadSaved(),2)
End Function
'''


def main() -> None:
    selected = [name for name in sys.argv[1:] if not name.startswith('--')]
    previous = json.loads(OUT.read_text(encoding='utf-8')) if ('--files-only' in sys.argv or selected) and OUT.exists() else None
    probes = previous['probes'] if previous is not None else []
    files = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        if previous is not None:
            assert previous['excel_version_build'] == version.value
        chosen = {} if '--files-only' in sys.argv else {name: OPERATIONS[name] for name in selected} if selected else OPERATIONS
        for name, operation in chosen.items():
            for mode, manual, calculate in [('auto', False, False), ('manual', True, False), ('calculated', True, True)]:
                excel.new_document()
                source = macro(operation, manual=manual, calculate=calculate)
                result = excel.run_vba(source, 'Probe', timeout=30.0)
                assert result.ok and not result.dialogs, (name, mode, result)
                probes = [probe for probe in probes if probe['name'] != name + '_' + mode]
                probes.append({'name': name + '_' + mode, 'macro': source, 'result': result.value})
                print(name, mode, result.value, flush=True)
        folder = OUT.parent / 'cse_array_files'
        folder.mkdir(exist_ok=True)
        for name in ['copy_child_auto', 'copy_children_manual', 'copy_tiled_calculated', 'cut_whole_auto',
                     'cut_overlap_manual', 'cross_sheet_cut_auto', 'shift_whole_down_manual',
                     'paste_transposed_manual', 'paste_add_whole_auto', 'replace_array_with_array_auto',
                     'replace_dynamic_anchor_manual', 'replace_dynamic_partial_auto', 'replace_dynamic_styled_auto',
                     'replace_dynamic_anchor_band_manual']:
            excel.new_document()
            probe = next(probe for probe in probes if probe['name'] == name)
            result = excel.run_vba(probe['macro'], 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, result
            path = folder / (name + '.xlsx')
            source = ('Sub SaveProbe()\nApplication.DisplayAlerts = False\nApplication.CalculateBeforeSave = False\n'
                      'ActiveWorkbook.SaveAs "' + str(path).replace('"', '""') + '", xlOpenXMLWorkbook\nEnd Sub')
            saved = excel.run_vba(source, 'SaveProbe', timeout=30.0)
            assert saved.ok and not saved.dialogs, saved
            read = excel.run_vba(HELPER, 'ReadSaved', timeout=30.0)
            assert read.ok and not read.dialogs, read
            files.append({'name': name, 'file': path.relative_to(OUT.parent).as_posix(), 'result': read.value,
                          'calculation': -4105 if name.endswith('_auto') else -4135})
    OUT.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(),
                              'read_saved': HELPER, 'probes': probes, 'files': files}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
