"""Measure initialization of VBA record-valued function and property returns."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_compile_demand import add_source

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_returns.json'
TYPES = ('Public Type InnerData\nValue As Long\nEnd Type\n'
         'Public Type RecordData\nValue As Long\nNested As InnerData\nItems(0 To 1) As Long\nEnd Type\n')
CASES = {
    'default_return': ('Function MakeRecord() As RecordData\nEnd Function', 'item = MakeRecord()\nCheck = item.Value'),
    'field_return': ('Function MakeRecord() As RecordData\nMakeRecord.Value = 7\nEnd Function', 'item = MakeRecord()\nCheck = item.Value'),
    'nested_return': ('Function MakeRecord() As RecordData\nMakeRecord.Nested.Value = 7\nEnd Function', 'item = MakeRecord()\nCheck = item.Nested.Value'),
    'array_field_return': ('Function MakeRecord() As RecordData\nMakeRecord.Items(1) = 7\nEnd Function', 'item = MakeRecord()\nCheck = item.Items(1)'),
    'qualified_return': ('Function MakeRecord() As Helpers.RecordData\nMakeRecord.Value = 7\nEnd Function', 'item = MakeRecord()\nCheck = item.Value'),
    'property_return': ('Property Get MakeRecord() As RecordData\nMakeRecord.Value = 7\nEnd Property', 'item = MakeRecord\nCheck = item.Value'),
}


def macro(name: str) -> str:
    procedure, body = CASES[name]
    caller = procedure + '\nFunction Check() As Long\nDim item As Helpers.RecordData\n' + body + '\nEnd Function'
    source = ('Function Probe() As String\nDim target As Workbook, part As Object\nDim sourceText As String\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
    source += add_source('Helpers', TYPES) + add_source('Caller', caller)
    return source + 'Probe = CStr(Application.Run("\'" & target.Name & "\'!Caller.Check"))\ntarget.Close False\nEnd Function'


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
