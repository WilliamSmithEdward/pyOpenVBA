"""Measure allocation, copies, preservation and borrows of record arrays."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_arrays.json'
DECLARATIONS = '''Public Type RecordData
Value As Long
Values(0 To 1) As Long
End Type
Private Stored() As RecordData
Private Failure As Long
Public Sub Bump(ByRef number As Long)
number = number + 10
End Sub
Public Sub Borrow(ByRef number As Long)
On Error Resume Next
ReDim Preserve Stored(0 To 2)
Failure = Err.Number
On Error GoTo 0
number = number + 10
End Sub'''
CASES = {
    'fixed': 'Dim items(0 To 1) As RecordData\nitems(1).Value = 5\nProbe = CStr(items(1).Value)',
    'dynamic': 'Dim items() As RecordData\nReDim items(0 To 1)\nitems(1).Value = 5\nProbe = CStr(items(1).Value)',
    'byref_field': 'Dim items(0 To 1) As RecordData\nitems(1).Value = 5\nBump items(1).Value\nProbe = CStr(items(1).Value)',
    'array_assignment': 'Dim first(0 To 1) As RecordData, second() As RecordData\nfirst(1).Value = 5\nsecond = first\nsecond(1).Value = 15\nProbe = CStr(first(1).Value) & ":" & CStr(second(1).Value)',
    'preserve': 'ReDim Stored(0 To 1)\nStored(1).Value = 5\nReDim Preserve Stored(0 To 2)\nProbe = CStr(Stored(1).Value) & ":" & CStr(Stored(2).Value)',
    'erase_fixed': 'Dim items(0 To 1) As RecordData\nitems(1).Value = 5\nErase items\nProbe = CStr(items(1).Value)',
    'borrowed_field_resize': 'ReDim Stored(0 To 1)\nStored(1).Value = 5\nBorrow Stored(1).Value\nProbe = CStr(Failure) & ":" & CStr(Stored(1).Value)',
    'borrowed_nested_array_resize': 'ReDim Stored(0 To 1)\nStored(1).Values(1) = 5\nBorrow Stored(1).Values(1)\nProbe = CStr(Failure) & ":" & CStr(Stored(1).Values(1))',
}
def literal(line: str) -> str:
    return '"' + line.replace('"', '""') + '"'
def macro(name: str) -> str:
    code = DECLARATIONS + '\nFunction Check() As String\n' + CASES[name].replace('Probe =', 'Check =') + '\nEnd Function'
    build = '\n'.join('sourceText = sourceText & ' + literal(line) + ' & vbCrLf' for line in code.split('\n'))
    return ('Function Probe() As String\nDim target As Workbook, part As Object\nDim sourceText As String\n'
            'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n'
            'Set part = target.VBProject.VBComponents.Add(1)\npart.Name = "Records"\n' + build + '\npart.CodeModule.AddFromString sourceText\n'
            'Probe = CStr(Application.Run("\'" & target.Name & "\'!Records.Check"))\ntarget.Close False\nEnd Function')
def main() -> None:
    probes = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        for name in CASES:
            source = macro(name)
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            print(name, result.value, flush=True)
            probes.append({'name': name, 'macro': source, 'result': result.value})
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
    OUT.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')
if __name__ == '__main__':
    main()
