"""Measure VBA array value copies, argument copies and element locks."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/array_copy.json'
DECLARATIONS = '''Private Stored() As Long
Private Replacement() As Long
Private Failure As Long
Public Sub ChangeValue(ByVal items As Variant)
items(1) = 15
End Sub
Public Sub ChangeReference(ByRef items As Variant)
items(1) = 15
End Sub
Public Function MakeValues() As Variant
MakeValues = Stored
End Function
Public Sub Borrow(ByRef number As Long)
On Error Resume Next
Stored = Replacement
Failure = Err.Number
On Error GoTo 0
number = number + 10
End Sub'''
CASES = {
    'variant_assignment': 'first = Array(0, 5)\nsecond = first\nsecond(1) = 15\nProbe = CStr(first(1)) & ":" & CStr(second(1))',
    'typed_dynamic_assignment': 'Dim left() As Long, right() As Long\nReDim left(0 To 1)\nleft(1) = 5\nright = left\nright(1) = 15\nProbe = CStr(left(1)) & ":" & CStr(right(1))',
    'fixed_to_dynamic_assignment': 'Dim left(0 To 1) As Long, right() As Long\nleft(1) = 5\nright = left\nright(1) = 15\nProbe = CStr(left(1)) & ":" & CStr(right(1))',
    'byval_variant': 'first = Array(0, 5)\nChangeValue first\nProbe = CStr(first(1))',
    'byref_variant': 'first = Array(0, 5)\nChangeReference first\nProbe = CStr(first(1))',
    'parenthesized_variant': 'first = Array(0, 5)\nChangeReference (first)\nProbe = CStr(first(1))',
    'function_result': 'ReDim Stored(0 To 1)\nStored(1) = 5\nfirst = MakeValues()\nfirst(1) = 15\nProbe = CStr(Stored(1)) & ":" & CStr(first(1))',
    'object_reference': 'Dim held As New Collection\nfirst = Array(held)\nsecond = first\nsecond(0).Add 1\nProbe = CStr(first(0).Count)',
    'element_locked_assignment': 'ReDim Stored(0 To 1)\nReDim Replacement(0 To 1)\nStored(1) = 5\nReplacement(1) = 1\nBorrow Stored(1)\nProbe = CStr(Failure) & ":" & CStr(Stored(1))',
}
def literal(line: str) -> str:
    return '"' + line.replace('"', '""') + '"'
def macro(name: str) -> str:
    code = DECLARATIONS + '\nFunction Check() As String\nDim first As Variant, second As Variant\n' + CASES[name].replace('Probe =', 'Check =') + '\nEnd Function'
    build = '\n'.join('sourceText = sourceText & ' + literal(line) + ' & vbCrLf' for line in code.split('\n'))
    return ('Function Probe() As String\nDim target As Workbook, part As Object\nDim sourceText As String\n'
            'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n'
            'Set part = target.VBProject.VBComponents.Add(1)\npart.Name = "Arrays"\n' + build + '\npart.CodeModule.AddFromString sourceText\n'
            'Probe = CStr(Application.Run("\'" & target.Name & "\'!Arrays.Check"))\ntarget.Close False\nEnd Function')
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
