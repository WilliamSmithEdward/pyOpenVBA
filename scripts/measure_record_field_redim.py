"""Measure native ReDim of a record's dynamic array field."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_field_redim.json'
HEAD = ('Public Type Inner\nV() As Long\nEnd Type\n'
        'Public Type Outer\nB() As Long\nI As Inner\nS() As String\nEnd Type\n'
        'Dim B As Long\n')
CASES = {
    'field': 'Dim r As Outer\nReDim r.B(0 To 9)\nr.B(9) = 7\nProbe = LBound(r.B) & "|" & UBound(r.B) & "|" & r.B(9) & "|" & B',
    'nested': 'Dim r As Outer\nReDim r.I.V(2 To 4)\nr.I.V(4) = 3\nProbe = LBound(r.I.V) & "|" & UBound(r.I.V) & "|" & r.I.V(4)',
    'preserve': 'Dim r As Outer\nReDim r.B(0 To 1)\nr.B(1) = 5\nReDim Preserve r.B(0 To 3)\nProbe = UBound(r.B) & "|" & r.B(1)',
    'with': 'Dim r As Outer\nWith r\nReDim .S(1)\n.S(1) = "x"\nEnd With\nProbe = UBound(r.S) & "|" & r.S(1)',
    'two': 'Dim r As Outer\nReDim r.B(1), r.S(2)\nProbe = UBound(r.B) & "|" & UBound(r.S) & "|" & TypeName(r.S)',
    'array_element': 'Dim a(1) As Outer\nReDim a(1).B(0 To 4)\nProbe = UBound(a(1).B)',
}

def macro(name: str) -> str:
    return HEAD + 'Function Probe() As String\n' + CASES[name] + '\nEnd Function'

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
