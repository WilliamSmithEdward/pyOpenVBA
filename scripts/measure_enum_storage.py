"""Measure how native VBA stores a variable, parameter, array or field typed as an enum."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/enum_storage.json'
HEAD = ('Public Enum Colour\nRed\nBlue = 5\nEnd Enum\n'
        'Public Type Holder\nB As Byte\nC As Colour\nEnd Type\n')
TAIL = ('\nFunction Pick(ByVal c As Colour) As Colour\nPick = c + 1\nEnd Function'
        '\nSub Bump(ByRef c As Colour)\nc = c + 1\nEnd Sub')
CASES = {
    'default': 'Dim c As Colour\nProbe = TypeName(c) & "|" & VarType(c) & "|" & c',
    'assign_double': 'Dim c As Colour\nc = 3.7\nProbe = TypeName(c) & "|" & c',
    'assign_text': 'Dim c As Colour\nc = "12"\nProbe = TypeName(c) & "|" & c',
    'overflow': 'Dim c As Colour\nOn Error Resume Next\nc = 3000000000#\nProbe = Err.Number & "|" & c',
    'param_return': 'Probe = TypeName(Pick(Blue)) & "|" & Pick(2.6)',
    'byref': 'Dim c As Colour\nc = Blue\nBump c\nProbe = c',
    'byref_long_to_enum': 'Dim n As Long\nn = 4\nBump n\nProbe = n',
    'array': 'Dim a(1) As Colour\na(0) = 2.5\nProbe = TypeName(a) & "|" & TypeName(a(0)) & "|" & a(0)',
    'record_field': 'Dim r As Holder\nr.C = 7.5\nProbe = TypeName(r.C) & "|" & r.C & "|" & Len(r) & "|" & LenB(r)',
    'member_through_enum': 'Probe = TypeName(Colour.Blue) & "|" & Colour.Blue & "|" & Colour.Red',
    'member_through_enum_in_expression': 'Dim c As Colour\nc = Colour.Blue + 1\nProbe = c & "|" & (c = Colour.Blue)',
}

def macro(name: str) -> str:
    return HEAD + 'Function Probe() As String\n' + CASES[name] + '\nEnd Function' + TAIL

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
