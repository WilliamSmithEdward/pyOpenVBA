"""Measure VBA record visibility and ambiguous unqualified declarations."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_compile_demand import add_source

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/record_scope.json'
RECORD = 'Public Type RecordData\nValue As Long\nEnd Type'
CASES = {
    'private_unqualified_external': [('Helpers', RECORD.replace('Public', 'Private')), ('Caller', 'Function Check() As Long\nDim item As RecordData\nCheck = item.Value\nEnd Function')],
    'private_qualified_external': [('Helpers', RECORD.replace('Public', 'Private')), ('Caller', 'Function Check() As Long\nDim item As Helpers.RecordData\nCheck = item.Value\nEnd Function')],
    'private_qualified_local': [('Helpers', RECORD.replace('Public', 'Private') + '\nFunction Check() As Long\nDim item As Helpers.RecordData\nitem.Value = 7\nCheck = item.Value\nEnd Function')],
    'ambiguous_unqualified': [('Helpers', RECORD), ('Other', RECORD), ('Caller', 'Function Check() As Long\nDim item As RecordData\nCheck = item.Value\nEnd Function')],
    'ambiguous_qualified': [('Helpers', RECORD), ('Other', RECORD), ('Caller', 'Function Check() As Long\nDim item As Other.RecordData\nitem.Value = 7\nCheck = item.Value\nEnd Function')],
    'local_shadows_public': [('Helpers', RECORD), ('Caller', RECORD.replace('Public', 'Private') + '\nFunction Check() As Long\nDim item As RecordData\nitem.Value = 7\nCheck = item.Value\nEnd Function')],
}


def macro(name: str) -> str:
    source = ('Function Probe() As String\nDim target As Workbook, part As Object\nDim sourceText As String\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
    for module, text in CASES[name]:
        source += add_source(module, text)
    owner = 'Helpers' if name == 'private_qualified_local' else 'Caller'
    return source + f'Probe = CStr(Application.Run("\'" & target.Name & "\'!{owner}.Check"))\ntarget.Close False\nEnd Function'


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
