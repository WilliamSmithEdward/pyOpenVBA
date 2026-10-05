"""Measure demand compilation and invalid whole-record call expressions."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/compile_demand.json'
HELPERS = '''Public Type RecordData
Value As Long
End Type
Public Stored As RecordData
Public Sub Change(ByRef item As RecordData)
item.Value = item.Value + 10
End Sub
Public Function MakeRecord() As RecordData
MakeRecord = Stored
End Function'''
BAD = 'Function Bad() As Long\nDim item As RecordData\nHelpers.Change (item)\nBad = 17\nEnd Function'
CASES = {
    'implicit_variable': ('Helpers.Change item', ''),
    'explicit_call': ('Call Helpers.Change(item)', ''),
    'qualified_variable': ('Helpers.Stored.Value = 5\nHelpers.Change Helpers.Stored\nitem = Helpers.Stored', ''),
    'parenthesized_record': ('Helpers.Change (item)', ''),
    'explicit_parenthesized_record': ('Call Helpers.Change((item))', ''),
    'function_result_record': ('Helpers.Change Helpers.MakeRecord()', ''),
    'dead_branch_record': ('If False Then\nHelpers.Change (item)\nEnd If', ''),
    'unused_invalid_procedure': ('', BAD),
    'unused_invalid_module': ('', 'other'),
}
def add_source(name: str, code: str) -> str:
    build = '\n'.join('sourceText = sourceText & "' + line.replace('"', '""') + '" & vbCrLf' for line in code.split('\n'))
    return 'sourceText = ""\nSet part = target.VBProject.VBComponents.Add(1)\npart.Name = "' + name + '"\n' + build + '\npart.CodeModule.AddFromString sourceText\n'
def macro(name: str) -> str:
    action, extra = CASES[name]
    caller = 'Function Check() As Long\nDim item As RecordData\nitem.Value = 5\n' + action + '\nCheck = item.Value\nEnd Function'
    if extra and extra != 'other':
        caller += '\n' + extra
    source = ('Function Probe() As String\nDim target As Workbook, part As Object\nDim sourceText As String\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
    source += add_source('Helpers', HELPERS) + add_source('Caller', caller)
    if extra == 'other':
        source += add_source('Broken', BAD)
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
