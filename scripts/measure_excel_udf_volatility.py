"""Measure native VBA UDF volatility registration and calculation scope."""
from __future__ import annotations
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_udf_volatility.json'
FUNCTIONS = '''Private Hits As Long
Private Enabled As Variant
Function Tick() As Long
Application.Volatile
Hits = Hits + 1
Tick = Hits
End Function
Function ConfiguredTick(ByVal mode As Boolean) As Long
Application.Volatile mode
Hits = Hits + 1
ConfiguredTick = Hits
End Function
Function SwitchedTick() As Long
Application.Volatile Enabled
Hits = Hits + 1
SwitchedTick = Hits
End Function
Function NestedTick() As Long
RegisterVolatile
Hits = Hits + 1
NestedTick = Hits
End Function
Sub RegisterVolatile()
Application.Volatile
End Sub
Function FailingTick() As Variant
Application.Volatile
Hits = Hits + 1
FailingTick = CVErr(xlErrValue)
End Function
Function PlainTick() As Long
Hits = Hits + 1
PlainTick = Hits
End Function
Function InputTick(value As Variant) As Long
Hits = Hits + 1
InputTick = value * Hits
End Function
'''
SETUP = 'Application.Calculation = xlCalculationManual\nDim ws As Worksheet\nSet ws = Worksheets.Add\n'
CASES = {
    'calculate_and_cached_reads': 'ws.Range("A1").Formula = "=Tick()"\nws.Range("A1").Calculate\nProbe = ws.Range("A1").Value & ":" & ws.Range("A1").Value & ":" & Hits\nApplication.Calculate\nProbe = Probe & "|" & ws.Range("A1").Value & ":" & Hits\nApplication.Calculate\nProbe = Probe & "|" & ws.Range("A1").Value & ":" & Hits',
    'plain_function': 'ws.Range("A1").Formula = "=PlainTick()"\nws.Range("A1").Calculate\nProbe = CStr(Hits)\nApplication.Calculate\nProbe = Probe & "|" & Hits\nws.Calculate\nProbe = Probe & "|" & Hits',
    'false_opt_out': 'ws.Range("A1").Formula = "=ConfiguredTick(FALSE)"\nws.Range("A1").Calculate\nProbe = CStr(Hits)\nApplication.Calculate\nProbe = Probe & "|" & Hits',
    'disable_after_registration': 'Enabled = True\nws.Range("A1").Formula = "=SwitchedTick()"\nws.Range("A1").Calculate\nProbe = CStr(Hits)\nEnabled = False\nApplication.Calculate\nProbe = Probe & "|" & Hits\nApplication.Calculate\nProbe = Probe & "|" & Hits',
    'nested_registration': 'ws.Range("A1").Formula = "=NestedTick()"\nws.Range("A1").Calculate\nProbe = CStr(Hits)\nApplication.Calculate\nProbe = Probe & "|" & Hits',
    'error_result_stays_volatile': 'ws.Range("A1").Formula = "=FailingTick()"\nws.Range("A1").Calculate\nProbe = Hits & ":" & ws.Range("A1").Text\nApplication.Calculate\nProbe = Probe & "|" & Hits & ":" & ws.Range("A1").Text',
    'two_cells_true_then_false': 'ws.Range("A1").Formula = "=ConfiguredTick(TRUE)"\nws.Range("A2").Formula = "=ConfiguredTick(FALSE)"\nws.Range("A1").Calculate\nws.Range("A2").Calculate\nProbe = CStr(Hits)\nApplication.Calculate\nProbe = Probe & "|" & Hits & ":" & ws.Range("A1").Value & ":" & ws.Range("A2").Value',
    'two_cells_false_then_true': 'ws.Range("A1").Formula = "=ConfiguredTick(FALSE)"\nws.Range("A2").Formula = "=ConfiguredTick(TRUE)"\nws.Range("A1").Calculate\nws.Range("A2").Calculate\nProbe = CStr(Hits)\nApplication.Calculate\nProbe = Probe & "|" & Hits & ":" & ws.Range("A1").Value & ":" & ws.Range("A2").Value',
    'outside_udf_does_not_register': 'Application.Volatile\nws.Range("A1").Formula = "=PlainTick()"\nws.Range("A1").Calculate\nProbe = CStr(Hits)\nApplication.Calculate\nProbe = Probe & "|" & Hits',
    'automatic_unrelated_edit': 'ws.Range("A1").Formula = "=Tick()"\nws.Range("A1").Calculate\nApplication.Calculation = xlCalculationAutomatic\nProbe = CStr(Hits)\nws.Range("D1").Value = 5\nProbe = Probe & "|" & Hits',
    'worksheet_scope': 'Dim other As Worksheet\nSet other = Worksheets.Add\nws.Range("A1").Formula = "=Tick()"\nws.Range("A1").Calculate\nProbe = CStr(Hits)\nother.Calculate\nProbe = Probe & "|" & Hits\nws.Calculate\nProbe = Probe & "|" & Hits',
    'dependent_recalculation': 'ws.Range("A1").Formula = "=Tick()"\nws.Range("B1").Formula = "=A1*10"\nws.Calculate\nProbe = ws.Range("A1").Value & ":" & ws.Range("B1").Value & ":" & Hits\nApplication.Calculate\nProbe = Probe & "|" & ws.Range("A1").Value & ":" & ws.Range("B1").Value & ":" & Hits',
    'automatic_formula_entry': 'Application.Calculation = xlCalculationAutomatic\nws.Range("A1").Formula = "=Tick()"\nProbe = Hits & ":" & ws.Range("A1").Value\nws.Range("D1").Value = 5\nProbe = Probe & "|" & Hits & ":" & ws.Range("A1").Value',
    'manual_entry_without_range': 'ws.Range("A1").Formula = "=PlainTick()"\nProbe = Hits & ":" & ws.Range("A1").Value\nApplication.Calculate\nProbe = Probe & "|" & Hits\nApplication.Calculate\nProbe = Probe & "|" & Hits',
    'global_false_call_order_warm': 'ws.Range("A1").Formula = "=ConfiguredTick(TRUE)"\nws.Range("A2").Formula = "=ConfiguredTick(FALSE)"\nws.Range("A1").Calculate\nws.Range("A2").Calculate\nApplication.Calculate\nHits = 0\nApplication.Calculate\nProbe = CStr(Hits)',
    'global_true_call_order_warm': 'ws.Range("A1").Formula = "=ConfiguredTick(FALSE)"\nws.Range("A2").Formula = "=ConfiguredTick(TRUE)"\nws.Range("A1").Calculate\nws.Range("A2").Calculate\nApplication.Calculate\nHits = 0\nApplication.Calculate\nProbe = CStr(Hits)',
    'volatile_dependency_warm': 'ws.Range("A1").Formula = "=Tick()"\nws.Range("B1").Formula = "=A1*10"\nws.Calculate\nApplication.Calculate\nHits = 0\nApplication.Calculate\nProbe = ws.Range("A1").Value & ":" & ws.Range("B1").Value & ":" & Hits',
    'nested_registration_warm': 'ws.Range("A1").Formula = "=NestedTick()"\nApplication.Calculate\nHits = 0\nApplication.Calculate\nProbe = CStr(Hits)',
    'unrelated_sheet_edit': 'Dim other As Worksheet\nSet other = Worksheets.Add\nws.Range("A1").Formula = "=Tick()"\nApplication.Calculation = xlCalculationAutomatic\nProbe = CStr(Hits)\nother.Range("D1").Value = 5\nProbe = Probe & "|" & Hits',
    'events_disabled_edit': 'ws.Range("A1").Formula = "=Tick()"\nApplication.Calculation = xlCalculationAutomatic\nApplication.EnableEvents = False\nProbe = CStr(Hits)\nws.Range("D1").Value = 5\nProbe = Probe & "|" & Hits',
    'mixed_functions_last_false': 'ws.Range("A1").Formula = "=Tick()+ConfiguredTick(FALSE)"\nApplication.Calculate\nHits = 0\nApplication.Calculate\nProbe = Hits & ":" & ws.Range("A1").Value',
    'mixed_functions_last_true': 'ws.Range("A1").Formula = "=ConfiguredTick(FALSE)+Tick()"\nApplication.Calculate\nHits = 0\nApplication.Calculate\nProbe = Hits & ":" & ws.Range("A1").Value',
    'manual_input_edit_stays_cached': 'ws.Range("A1").Value = 1\nws.Range("B1").Formula = "=InputTick(A1)"\nApplication.Calculate\nHits = 0\nws.Range("A1").Value = 5\nProbe = Hits & ":" & ws.Range("B1").Value\nApplication.Calculate\nProbe = Probe & "|" & Hits & ":" & ws.Range("B1").Value',
    'workbook_scoped_registration': 'Dim other As Workbook, part As Object\nws.Range("A1").Formula = "=Tick()"\nSet other = Workbooks.Add\nSet part = other.VBProject.VBComponents.Add(1)\npart.Name = "OtherCode"\npart.CodeModule.AddFromString "Function Tick() As Long" & vbCrLf & "Static LocalHits As Long" & vbCrLf & "Application.Volatile" & vbCrLf & "LocalHits = LocalHits + 1" & vbCrLf & "Tick = 100 + LocalHits" & vbCrLf & "End Function"\nother.Worksheets(1).Range("A1").Formula = "=Tick()"\nApplication.Calculate\nProbe = ws.Range("A1").Value & ":" & other.Worksheets(1).Range("A1").Value\nApplication.Calculate\nProbe = Probe & "|" & ws.Range("A1").Value & ":" & other.Worksheets(1).Range("A1").Value\nApplication.Calculation = xlCalculationAutomatic\nProbe = Probe & "|" & ws.Range("A1").Value & ":" & other.Worksheets(1).Range("A1").Value\nother.Worksheets(1).Range("D1").Value = 5\nProbe = Probe & "|" & ws.Range("A1").Value & ":" & other.Worksheets(1).Range("A1").Value',
}

def macro(name: str) -> str:
    return FUNCTIONS + '\nFunction Probe() As String\n' + SETUP + CASES[name] + '\nEnd Function'

def main() -> None:
    selected = sys.argv[1:]
    previous = json.loads(OUT.read_text(encoding='utf-8')) if selected and OUT.exists() else None
    probes = previous['probes'] if previous is not None else []
    version_value = None
    for name in selected or CASES:
        source = macro(name)
        with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
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
    OUT.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')

if __name__ == '__main__':
    main()
