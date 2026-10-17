"""Measure cached dynamic-array output through edits in manual calculation mode."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
import re
from zipfile import ZipFile

from pyvbaharness import ExcelSession, HarnessConfig

from measure_spill_edits import OPERATIONS, macro

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/manual_spill_edits.json'
READ_SAVED = '''Function ReadSaved() As String
Dim cell As Range
ReadSaved = "0"
For Each cell In ActiveSheet.Range("A1:F6")
ReadSaved = ReadSaved & "|" & cell.Address(False,False) & ":" & CStr(cell.Value2) & ":" & cell.Formula2 & ":" & cell.HasSpill
Next
End Function'''


def cases() -> dict[str, str]:
    probes: dict[str, str] = {}
    operations = {name: operation for name, operation in OPERATIONS.items() if not name.endswith('_immediate')}
    operations.update({
        'write_child': 'ws.Range("B3").Value = 9',
        'clear_child': 'ws.Range("B3").ClearContents',
        'copy_constant_to_child': 'ws.Range("E1").Value = 9\nws.Range("E1").Copy Destination:=ws.Range("B3")',
        'copy_constant_to_anchor': 'ws.Range("E1").Value = 9\nws.Range("E1").Copy Destination:=ws.Range("B2")',
        'write_empty_child': 'ws.Range("B3").Value = ""',
        'shift_anchor_down': 'ws.Range("B2").Insert Shift:=xlShiftDown',
        'delete_into_obstruction': 'ws.Range("B6").Value = 9\nws.Rows("3:4").Delete',
        'two_spills_insert': 'ws.Range("B5").Formula2 = "=SEQUENCE(2,2,10)"\nApplication.Calculate\nws.Rows(3).Insert',
        'two_spills_delete': 'ws.Range("B5").Formula2 = "=SEQUENCE(2,2,10)"\nApplication.Calculate\nws.Rows(3).Delete',
        'copy_tiled': 'ws.Range("B2:C4").Copy Destination:=ws.Range("E2:F7")',
        'copy_text_error': 'ws.Range("B2").Formula2 = "={TRUE,\"\"word\"\";1,#N/A}"\nApplication.Calculate\nws.Range("B2:C3").Copy Destination:=ws.Range("E2")',
        'copy_scalar': 'ws.Range("B2").Formula2 = "=SUM(1,2)"\nApplication.Calculate\nws.Range("B2").Copy Destination:=ws.Range("E2")',
    })
    for name, operation in operations.items():
        for calculate in [False, True]:
            source = macro('Application.Calculation = xlCalculationManual\n' + operation)
            source = source.replace('Set ws = Worksheets.Add', 'Application.Calculation = xlCalculationAutomatic\nSet ws = Worksheets.Add')
            if not calculate:
                source = source.replace('On Error GoTo 0\nApplication.Calculate', 'On Error GoTo 0')
            probes[name + ('_calculated' if calculate else '')] = source
    for name, operation in {'input_only': '', 'input_insert_inside': OPERATIONS['insert_inside'],
                            'input_insert_above': OPERATIONS['insert_above'], 'input_copy': OPERATIONS['copy_whole'],
                            'input_cut': OPERATIONS['cut_whole']}.items():
        for calculate in [False, True]:
            source = macro('Application.Calculation = xlCalculationManual\nws.Range("A1").Value = 4\n' + operation)
            source = source.replace('Set ws = Worksheets.Add', 'Application.Calculation = xlCalculationAutomatic\nSet ws = Worksheets.Add\nws.Range("A1").Value = 3')
            source = source.replace('=SEQUENCE(3,2)', '=SEQUENCE(A1,2)')
            if not calculate:
                source = source.replace('On Error GoTo 0\nApplication.Calculate', 'On Error GoTo 0')
            probes[name + ('_calculated' if calculate else '')] = source
    probes['save_flag_preserved'] = '''Function Probe() As String
Application.CalculateBeforeSave = False
Application.Calculation = xlCalculationManual
Probe = CStr(Application.CalculateBeforeSave)
Application.Calculation = xlCalculationAutomatic
Probe = Probe & ":" & Application.CalculateBeforeSave
Application.CalculateBeforeSave = True
Application.Calculation = xlCalculationManual
Probe = Probe & ":" & Application.CalculateBeforeSave
End Function'''
    probes['save_flag_null'] = '''Function Probe() As String
On Error Resume Next
Application.CalculateBeforeSave = Null
Probe = CStr(Err.Number)
End Function'''
    probes['save_flag_coercion'] = '''Function Probe() As String
Application.CalculateBeforeSave = 2
Probe = CStr(Application.CalculateBeforeSave)
Application.CalculateBeforeSave = "False"
Probe = Probe & ":" & Application.CalculateBeforeSave
End Function'''
    for name in ['copy_overlap', 'copy_overlap_calculated', 'write_child', 'write_child_calculated',
                 'delete_inside', 'delete_inside_calculated', 'input_insert_inside_calculated']:
        source = probes[name].replace('Application.Calculate\nOn Error Resume Next',
                                     'ws.Range("G1").Formula = "=B3*10"\nApplication.Calculate\nOn Error Resume Next')
        probes['dependent_' + name] = source.replace('Probe = CStr(n)',
                                                     'Probe = CStr(n) & ":" & CStr(ws.Range("G1").Value2)')
    for name, operation in {'application': 'Application.Calculate', 'worksheet': 'ws.Calculate',
                            'dependent_range': 'ws.Range("G1").Calculate', 'anchor_range': 'ws.Range("B2").Calculate'}.items():
        source = macro('Application.Calculation = xlCalculationManual\nws.Range("A1").Value = 10\n' + operation)
        source = source.replace('Set ws = Worksheets.Add', 'Application.Calculation = xlCalculationAutomatic\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1')
        source = source.replace('=SEQUENCE(3,2)', '=SEQUENCE(3,2,A1)')
        source = source.replace('Application.Calculate\nOn Error Resume Next',
                                'ws.Range("G1").Formula = "=B3*10"\nApplication.Calculate\nOn Error Resume Next')
        source = source.replace('On Error GoTo 0\nApplication.Calculate', 'On Error GoTo 0')
        probes['scope_' + name] = source.replace('Probe = CStr(n)',
                                                'Probe = CStr(n) & ":" & CStr(ws.Range("G1").Value2)')
    return probes


def main() -> None:
    probes = []
    files = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        for name, source in cases().items():
            excel.new_document()
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            probes.append({'name': name, 'macro': source, 'result': result.value})
            print(name, result.value, flush=True)
        folder = OUT.parent / 'manual_spill_cache'
        folder.mkdir(exist_ok=True)
        for name in ['input_insert_inside', 'input_copy', 'copy_overlap', 'write_child', 'clear_child',
                     'copy_partial', 'cut_whole', 'copy_anchor']:
            for calculate in ([False, True] if name == 'input_insert_inside' else [False]):
                excel.new_document()
                result = excel.run_vba(cases()[name], 'Probe', timeout=30.0)
                assert result.ok and not result.dialogs, result
                path = folder / (name + ('_recalculate' if calculate else '') + '.xlsx')
                source = ('Sub SaveProbe()\nApplication.DisplayAlerts = False\nApplication.CalculateBeforeSave = '
                          + str(calculate) + '\nActiveWorkbook.SaveAs "' + str(path).replace('"', '""')
                          + '", xlOpenXMLWorkbook\nEnd Sub')
                saved = excel.run_vba(source, 'SaveProbe', timeout=30.0)
                assert saved.ok and not saved.dialogs, saved
                read = excel.run_vba(READ_SAVED, 'ReadSaved', timeout=30.0)
                assert read.ok and not read.dialogs, read
                with ZipFile(path) as package:
                    calc_pr = re.search(r'<calcPr\b[^>]*/>', package.read('xl/workbook.xml').decode('utf-8'))
                    assert calc_pr is not None
                files.append({'name': path.stem, 'case': name, 'calculate_before_save': calculate,
                              'file': path.relative_to(OUT.parent).as_posix(), 'result': read.value,
                              'calc_pr': calc_pr.group(0)})
    OUT.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(),
                              'probes': probes, 'read_saved': READ_SAVED, 'files': files}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
