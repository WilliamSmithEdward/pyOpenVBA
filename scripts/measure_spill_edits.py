"""Measure native Excel structural edits through dynamic-array spills."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/spill_edits.json'
OPERATIONS = {
    'insert_above': 'ws.Rows(1).Insert',
    'insert_at_anchor': 'ws.Rows(2).Insert',
    'insert_inside': 'ws.Rows(3).Insert',
    'insert_below': 'ws.Rows(5).Insert',
    'delete_above': 'ws.Rows(1).Delete',
    'delete_anchor': 'ws.Rows(2).Delete',
    'delete_inside': 'ws.Rows(3).Delete',
    'delete_whole': 'ws.Rows("2:4").Delete',
    'insert_column_anchor': 'ws.Columns(2).Insert',
    'insert_column_inside': 'ws.Columns(3).Insert',
    'delete_column_anchor': 'ws.Columns(2).Delete',
    'delete_column_inside': 'ws.Columns(3).Delete',
    'insert_cells_down': 'ws.Range("B3").Insert Shift:=xlShiftDown',
    'delete_cells_up': 'ws.Range("B3").Delete Shift:=xlShiftUp',
    'insert_cells_right': 'ws.Range("C2").Insert Shift:=xlShiftToRight',
    'delete_cells_left': 'ws.Range("C2").Delete Shift:=xlShiftToLeft',
    'sort_whole': 'ws.Range("B2:C4").Sort Key1:=ws.Range("B2"), Order1:=xlDescending, Header:=xlNo',
    'sort_children': 'ws.Range("B3:C4").Sort Key1:=ws.Range("B3"), Order1:=xlDescending, Header:=xlNo',
    'sort_anchor_column': 'ws.Range("B2:B4").Sort Key1:=ws.Range("B2"), Order1:=xlDescending, Header:=xlNo',
    'cut_whole': 'ws.Range("B2:C4").Cut Destination:=ws.Range("E2")',
    'cut_child': 'ws.Range("B3").Cut Destination:=ws.Range("E2")',
    'copy_whole': 'ws.Range("B2:C4").Copy Destination:=ws.Range("E2")',
    'copy_child': 'ws.Range("B3").Copy Destination:=ws.Range("E2")',
    'copy_anchor': 'ws.Range("B2").Copy Destination:=ws.Range("E2")',
    'copy_partial': 'ws.Range("B2:B3").Copy Destination:=ws.Range("E2")',
    'copy_children': 'ws.Range("B3:C4").Copy Destination:=ws.Range("E2")',
    'copy_overlap': 'ws.Range("B2:C4").Copy Destination:=ws.Range("C2")',
    'cut_anchor': 'ws.Range("B2").Cut Destination:=ws.Range("E2")',
    'sort_object': 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B2"), Order:=xlDescending\nws.Sort.SetRange ws.Range("B2:C4")\nws.Sort.Header = xlNo\nws.Sort.Apply',
    'sort_unrelated': 'ws.Range("E2").Value = 5\nws.Range("E3").Value = 1\nws.Range("E2:F4").Sort Key1:=ws.Range("E2"), Header:=xlNo',
    'sort_blocked': 'ws.Range("C3").Value = 9\nws.Range("B2:C4").Sort Key1:=ws.Range("B2"), Order1:=xlDescending, Header:=xlNo',
    'copy_whole_immediate': 'ws.Range("B2:C4").Copy Destination:=ws.Range("E2")',
    'copy_overlap_immediate': 'ws.Range("B2:C4").Copy Destination:=ws.Range("C2")',
    'copy_anchor_immediate': 'ws.Range("B2").Copy Destination:=ws.Range("E2")',
    'cut_whole_immediate': 'ws.Range("B2:C4").Cut Destination:=ws.Range("E2")',
}

def macro(operation: str, *, immediate: bool = False) -> str:
    source = '''Function Probe() As String
Dim ws As Worksheet, cell As Range, n As Long
Set ws = Worksheets.Add
ws.Range("B2").Formula2 = "=SEQUENCE(3,2)"
Application.Calculate
On Error Resume Next
''' + operation + '''
n = Err.Number
On Error GoTo 0
Application.Calculate
Probe = CStr(n)
For Each cell In ws.Range("A1:F6")
Probe = Probe & "|" & cell.Address(False,False) & ":" & CStr(cell.Value2) & ":" & cell.Formula2 & ":" & cell.HasSpill
Next
End Function
'''
    if immediate:
        source = source.replace('On Error GoTo 0\nApplication.Calculate', 'On Error GoTo 0')
        source = source.replace('Probe = CStr(n)', 'Probe = CStr(n) & ":" & ws.Range("E3").Value2 & ":" & ws.Range("E3").HasSpill & ":" & ws.Range("D3").Value2')
    return source

def main() -> None:
    selected = sys.argv[1:]
    previous = json.loads(OUT.read_text(encoding='utf-8')) if selected and OUT.exists() else None
    probes = previous['probes'] if previous is not None else []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name in selected or OPERATIONS:
            source = macro(OPERATIONS[name], immediate=name.endswith('_immediate'))
            excel.new_document()
            version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
            assert version.ok and not version.dialogs, version
            version_value = version.value
            if previous is not None:
                assert previous['excel_version_build'] == version_value
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            probes = [probe for probe in probes if probe['name'] != name]
            probes.append({'name': name, 'macro': source, 'result': result.value})
            print(name, result.value, flush=True)
            OUT.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')

if __name__ == '__main__':
    main()
