"""Record native VBA standard-module qualification and argument binding."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/module_namespaces.json'
HELPERS = '''Public Value As Long
Public Values(0 To 1) As Long
Public Const Offset As Long = 7
Public Function Bump(ByRef number As Long, Optional ByVal increment As Long = 1) As Long
number = number + increment
Bump = number
End Function
Private Function Hidden() As Long
Hidden = 13
End Function
Public Function Inside() As Long
Inside = Helpers.Hidden()
End Function
Public Property Get Counter() As Long
Counter = Value
End Property
Public Property Let Counter(ByVal amount As Long)
Value = amount
End Property'''
THING = 'Public Value As Long\nPublic Sub Bump(ByRef number As Long)\nnumber = number + 3\nEnd Sub'
CASES = {
    'qualified_function': 'Probe = CStr(Helpers.Bump(number)) & ":" & CStr(number)',
    'qualified_named': 'Probe = CStr(Helpers.Bump(increment:=4, number:=number)) & ":" & CStr(number)',
    'qualified_expression': 'Probe = CStr(Helpers.Bump((number))) & ":" & CStr(number)',
    'qualified_implicit_parens': 'Helpers.Bump (number)\nProbe = CStr(number)',
    'qualified_explicit_call': 'Call Helpers.Bump(number)\nProbe = CStr(number)',
    'qualified_variable': 'Helpers.Value = 17\nProbe = CStr(Helpers.Value)',
    'qualified_variable_byref': 'Helpers.Value = 17\nHelpers.Bump Helpers.Value\nProbe = CStr(Helpers.Value)',
    'qualified_array': 'Helpers.Values(1) = 19\nProbe = CStr(Helpers.Values(1))',
    'qualified_array_byref': 'Helpers.Values(1) = 19\nHelpers.Bump Helpers.Values(1)\nProbe = CStr(Helpers.Values(1))',
    'local_array_byref': 'Dim values(0 To 1) As Long\nvalues(1) = 23\nHelpers.Bump values(1)\nProbe = CStr(values(1))',
    'qualified_constant': 'Probe = CStr(Helpers.Offset)',
    'qualified_property': 'Helpers.Counter = 23\nProbe = CStr(Helpers.Counter)',
    'qualified_own_private': 'Probe = CStr(Helpers.Inside())',
    'class_byref': 'Dim instance As New Thing\ninstance.Bump number\nProbe = CStr(number)',
    'class_byref_named': 'Dim instance As New Thing\ninstance.Bump number:=number\nProbe = CStr(number)',
    'class_implicit_parens': 'Dim instance As New Thing\ninstance.Bump (number)\nProbe = CStr(number)',
    'class_field_byref': 'Dim instance As New Thing\ninstance.Value = 29\nHelpers.Bump instance.Value\nProbe = CStr(instance.Value)',
    'qualified_shadow': 'Dim Helpers As New Collection\nHelpers.Add "value"\nProbe = CStr(Helpers.Count)',
}


def literal(text: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in text.split('\n'))


def macro(action: str) -> str:
    caller = 'Function Probe() As String\nDim number As Long\nnumber = 5\n' + action + '\nEnd Function'
    source = ('Function Probe() As String\nDim target As Workbook, part As Object\n'
              'Application.EnableEvents = False\nSet target = Workbooks.Add\n')
    for name, kind, code in [('Helpers', 1, HELPERS), ('Caller', 1, caller), ('Thing', 2, THING)]:
        source += f'Set part = target.VBProject.VBComponents.Add({kind})\npart.Name = "{name}"\npart.CodeModule.AddFromString ' + literal(code) + '\n'
    return source + 'Probe = CStr(Application.Run("\'" & target.Name & "\'!Caller.Probe"))\ntarget.Close False\nEnd Function'


def main() -> None:
    rows: list[dict[str, object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name, action in CASES.items():
            excel.new_document()
            source = macro(action)
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            rows.append({'name': name, 'action': action, 'macro': source, 'result': result.value})
            print(name, repr(result.value), flush=True)
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok, version
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel_version_build': version.value,
                               'probes': rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
