"""Measure native nominal record argument and assignment type checks."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_compile_demand import add_source

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_type_checks.json'
HELPERS = ('Public Type FirstData\nValue As Long\nEnd Type\n'
           'Public Type SecondData\nValue As Long\nEnd Type\n'
           'Sub Change(ByRef item As FirstData)\nitem.Value = item.Value + 10\nEnd Sub')
CASES = {
    'matching_argument': ('Dim item As FirstData\nitem.Value = 5\nHelpers.Change item\nCheck = item.Value', ''),
    'mismatched_argument': ('Dim item As SecondData\nHelpers.Change item\nCheck = item.Value', ''),
    'mismatched_array_element': ('Dim items(0 To 1) As SecondData\nHelpers.Change items(1)\nCheck = items(1).Value', ''),
    'mismatched_function_result': ('Helpers.Change MakeRecord()\nCheck = 7', 'Function MakeRecord() As SecondData\nEnd Function'),
    'mismatched_assignment': ('Dim first As FirstData, second As SecondData\nfirst = second\nCheck = first.Value', ''),
    'record_byval_parameter': ('Dim item As FirstData\nChangeByVal item\nCheck = item.Value', 'Sub ChangeByVal(ByVal item As FirstData)\nEnd Sub'),
    'matching_function_result': ('Helpers.Change MakeRecord()\nCheck = 7', 'Function MakeRecord() As FirstData\nEnd Function'),
    'record_byval_direct': ('Dim item As FirstData\nChangeByVal item\nProbe = CStr(item.Value)', 'Sub ChangeByVal(ByVal item As FirstData)\nEnd Sub'),
}


def macro(name: str) -> str:
    body, extra = CASES[name]
    if name == 'record_byval_direct':
        return 'Public Type FirstData\nValue As Long\nEnd Type\n' + extra + '\nFunction Probe() As String\n' + body + '\nEnd Function'
    caller = extra + '\nFunction Check() As Long\n' + body + '\nEnd Function'
    source = ('Function Probe() As String\nDim target As Workbook, part As Object\nDim sourceText As String\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
    source += add_source('Helpers', HELPERS) + add_source('Caller', caller)
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
            if result.ok:
                assert not result.dialogs, result
                probe = {'name': name, 'macro': source, 'result': result.value}
            elif result.error is not None and not result.dialogs:
                probe = {'name': name, 'macro': source, 'runtime_error': str(result.error.number),
                         'error_description': result.error.description}
            else:
                assert result.dialogs and result.dialogs[0].classification == 'compile-error', (name, result)
                probe = {'name': name, 'macro': source, 'compile_error': result.dialogs[0].message}
            probes.append(probe)
            print(name, probe.get('result', probe.get('compile_error', probe.get('runtime_error'))), flush=True)
    OUT.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
