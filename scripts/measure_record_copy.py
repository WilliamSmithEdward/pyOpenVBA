"""Measure value copying and retained object references in VBA records."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_copy.json'
DECLARATIONS = '''Public Type InnerData
Value As Long
End Type
Public Type RecordData
Value As Long
Nested As InnerData
Values(0 To 1) As Long
Reference As Collection
End Type
Private Stored As RecordData
Private Replacement As RecordData
Public Sub ReplaceRecord(ByRef number As Long)
Stored = Replacement
number = number + 10
End Sub
Public Function MakeRecord() As RecordData
MakeRecord = Stored
End Function
Public Sub Change(ByRef item As RecordData)
item.Value = item.Value + 10
End Sub'''
CASES = {
    'assignment': 'first.Value = 5\nsecond = first\nsecond.Value = 15\nProbe = CStr(first.Value) & ":" & CStr(second.Value)',
    'nested_assignment': 'first.Nested.Value = 5\nsecond = first\nsecond.Nested.Value = 15\nProbe = CStr(first.Nested.Value) & ":" & CStr(second.Nested.Value)',
    'array_assignment': 'first.Values(1) = 5\nsecond = first\nsecond.Values(1) = 15\nProbe = CStr(first.Values(1)) & ":" & CStr(second.Values(1))',
    'object_reference': 'Set first.Reference = New Collection\nfirst.Reference.Add 1\nsecond = first\nsecond.Reference.Add 2\nProbe = CStr(first.Reference.Count) & ":" & CStr(second.Reference.Count)',
    'byref_record': 'first.Value = 5\nChange first\nProbe = CStr(first.Value)',
    'scalar_alias_during_copy': 'Stored.Value = 5\nReplacement.Value = 1\nReplaceRecord Stored.Value\nProbe = CStr(Stored.Value)',
    'nested_alias_during_copy': 'Stored.Nested.Value = 5\nReplacement.Nested.Value = 1\nReplaceRecord Stored.Nested.Value\nProbe = CStr(Stored.Nested.Value)',
    'array_alias_during_copy': 'Stored.Values(1) = 5\nReplacement.Values(1) = 1\nReplaceRecord Stored.Values(1)\nProbe = CStr(Stored.Values(1))',
    'rebind_object_field': 'Set first.Reference = New Collection\nfirst.Reference.Add 1\nsecond = first\nSet second.Reference = New Collection\nProbe = CStr(first.Reference.Count) & ":" & CStr(second.Reference.Count)',
    'function_result': 'Stored.Value = 5\nfirst = MakeRecord()\nfirst.Value = 15\nProbe = CStr(Stored.Value) & ":" & CStr(first.Value)',
}
def literal(source: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in source.split('\n'))
def macro(name: str) -> str:
    code = DECLARATIONS + '\nFunction Check() As String\nDim first As RecordData, second As RecordData\n' + CASES[name].replace('Probe =', 'Check =') + '\nEnd Function'
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
