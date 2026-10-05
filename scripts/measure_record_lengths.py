"""Measure native VBA record file lengths and in-memory alignment."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_lengths.json'
CASES = {
    'byte': 'A As Byte',
    'byte_long': 'A As Byte\nB As Long',
    'long_byte': 'A As Long\nB As Byte',
    'byte_integer': 'A As Byte\nB As Integer',
    'integers': 'A As Integer\nB As Integer',
    'byte_double': 'A As Byte\nB As Double',
    'boolean_single': 'A As Boolean\nB As Single',
    'currency_date': 'A As Currency\nB As Date',
    'longptr': 'A As LongPtr',
    'object': 'A As Object',
    'variant': 'A As Variant',
    'variable_string': 'A As String',
    'fixed_string': 'A As String * 3',
    'long_array_byte': 'A(0 To 1) As Long\nB As Byte',
    'nested': 'A As Byte\nB As InnerData',
    'fixed_string_storage': 'Text As String * 3',
}

def macro(name: str) -> str:
    if name == 'fixed_string_storage':
        return '''Private Type RecordData
Text As String * 3
End Type
Function Probe() As String
Dim first As RecordData, second As RecordData
Probe = CStr(AscW(Mid$(first.Text, 1, 1))) & "|"
first.Text = "abcd"
second = first
first.Text = "x"
Probe = Probe & first.Text & "|" & second.Text & "|" & Len(first) & "|" & LenB(first)
End Function'''
    return ('Public Type InnerData\nA As Byte\nB As Long\nEnd Type\nPublic Type RecordData\n' + CASES[name] +
            '\nEnd Type\nFunction Probe() As String\nDim item As RecordData\n'
            'Probe = CStr(Len(item)) & "|" & CStr(LenB(item))\nEnd Function')

def main() -> None:
    probes = []
    version_value = None
    for name in CASES:
        source = macro(name)
        with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
            excel.new_document()
            version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
            assert version.ok and not version.dialogs, version
            version_value = version.value
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            probes.append({'name': name, 'macro': source, 'result': result.value})
            print(name, result.value, flush=True)
    OUT.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')

if __name__ == '__main__':
    main()
