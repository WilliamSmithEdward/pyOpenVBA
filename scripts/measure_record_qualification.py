"""Measure module-qualified record declarations and parenthesized members."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
from measure_compile_demand import HELPERS, add_source
HELPERS = HELPERS.replace('Public Type RecordData', 'Public Type InnerData\nValue As Long\nEnd Type\nPublic Type RecordData').replace('Value As Long\nEnd Type\nPublic Stored', 'Value As Long\nNested As InnerData\nEnd Type\nPublic Stored') + '\nPublic Sub ChangeInner(ByRef item As InnerData)\nitem.Value = item.Value + 10\nEnd Sub'
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_qualification.json'
CASES = {
    'nested_record': 'Dim item As Helpers.RecordData\nitem.Nested.Value = 5\nHelpers.ChangeInner item.Nested\nCheck = item.Nested.Value',
    'nested_record_parentheses': 'Dim item As Helpers.RecordData\nHelpers.ChangeInner (item.Nested)\nCheck = 5',
    'qualified_scalar': 'Dim item As Helpers.RecordData\nitem.Value = 5\nHelpers.Change item\nCheck = item.Value',
    'qualified_fixed_array': 'Dim items(0 To 1) As Helpers.RecordData\nitems(1).Value = 5\nHelpers.Change items(1)\nCheck = items(1).Value',
    'qualified_dynamic_array': 'Dim items() As Helpers.RecordData\nReDim items(0 To 1)\nitems(1).Value = 5\nHelpers.Change items(1)\nCheck = items(1).Value',
    'qualified_scalar_parentheses': 'Dim item As Helpers.RecordData\nHelpers.Change (item)\nCheck = 5',
    'array_element_parentheses': 'Dim items(0 To 1) As RecordData\nHelpers.Change (items(1))\nCheck = 5',
    'qualified_variable_parentheses': 'Helpers.Change (Helpers.Stored)\nCheck = 5',
}
def macro(name: str) -> str:
    caller = 'Function Check() As Long\n' + CASES[name] + '\nEnd Function'
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
            else:
                assert result.dialogs and result.dialogs[0].classification == 'compile-error', (name, result)
                probe = {'name': name, 'macro': source, 'compile_error': result.dialogs[0].message}
            probes.append(probe)
            print(name, probe.get('result', probe.get('compile_error')), flush=True)
    OUT.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')
if __name__ == '__main__':
    main()
