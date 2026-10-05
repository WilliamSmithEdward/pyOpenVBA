"""Measure ByRef fields and array elements of VBA user-defined types."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_byref.json'
HELPERS = '''Public Type InnerData
Value As Long
End Type
Public Type RecordData
Value As Long
Nested As InnerData
Values(0 To 1) As Long
End Type
Public Stored As RecordData
Public Sub Bump(ByRef number As Long)
number = number + 10
End Sub'''
CASES = {
    'local_field': ('Dim item As RecordData\nitem.Value = 5', 'Helpers.Bump item.Value', 'item.Value'),
    'nested_field': ('Dim item As RecordData\nitem.Nested.Value = 5', 'Helpers.Bump item.Nested.Value', 'item.Nested.Value'),
    'array_field': ('Dim item As RecordData\nitem.Values(1) = 5', 'Helpers.Bump item.Values(1)', 'item.Values(1)'),
    'named_field': ('Dim item As RecordData\nitem.Value = 5', 'Helpers.Bump number:=item.Value', 'item.Value'),
    'parenthesized_field': ('Dim item As RecordData\nitem.Value = 5', 'Helpers.Bump (item.Value)', 'item.Value'),
    'qualified_field': ('Helpers.Stored.Value = 5', 'Helpers.Bump Helpers.Stored.Value', 'Helpers.Stored.Value'),
}
def literal(source: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in source.split('\n'))
def macro(name: str) -> str:
    setup, action, expression = CASES[name]
    caller = 'Function Probe() As Long\n' + setup + '\n' + action + '\nProbe = ' + expression + '\nEnd Function'
    source = ('Function Probe() As Long\nDim target As Workbook, part As Object\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
    for component, text in [('Helpers', HELPERS), ('Caller', caller)]:
        source += f'Set part = target.VBProject.VBComponents.Add(1)\npart.Name = "{component}"\npart.CodeModule.AddFromString ' + literal(text) + '\n'
    return source + 'Probe = Application.Run("\'" & target.Name & "\'!Caller.Probe")\ntarget.Close False\nEnd Function'
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
