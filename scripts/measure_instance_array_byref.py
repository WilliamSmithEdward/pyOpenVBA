"""Measure ByRef array elements inside class/document instances."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/instance_array_byref.json'
HELPERS = 'Public Sub Bump(ByRef number As Long)\nnumber = number + 10\nEnd Sub'
CASES = {
    'fixed_class': ('Private values(0 To 1) As Long', 'values(1) = 5', 'Helpers.Bump values(1)', False),
    'dynamic_class': ('Private values() As Long', 'ReDim values(0 To 1)\nvalues(1) = 5', 'Helpers.Bump values(1)', False),
    'variant_class': ('Private values As Variant', 'values = Array(0, 5)', 'Helpers.Bump values(1)', False),
    'named_class': ('Private values(0 To 1) As Long', 'values(1) = 5', 'Helpers.Bump number:=values(1)', False),
    'local_shadow': ('Private values(0 To 1) As Long', 'Dim values(0 To 1) As Long\nvalues(1) = 5', 'Helpers.Bump values(1)', False),
    'fixed_document': ('Private values(0 To 1) As Long', 'values(1) = 5', 'Helpers.Bump values(1)', True),
}
def literal(source: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in source.split('\n'))
def macro(name: str) -> str:
    declaration, setup, action, document = CASES[name]
    helpers = HELPERS.replace('ByRef number As Long', 'ByRef number As Variant') if name == 'variant_class' else HELPERS
    code = declaration + '\nPublic Function Check() As Long\n' + setup + '\n' + action + '\nCheck = values(1)\nEnd Function'
    if document:
        code += '\nPublic Sub ExportValue()\nRange("A1").Value = Check()\nEnd Sub'
    caller = ('Function Probe() As Long\nDim instance As New Thing\nProbe = instance.Check()\nEnd Function')
    source = ('Function Probe() As Long\nDim target As Workbook, part As Object\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
    for component, kind, text in [('Helpers', 1, helpers), ('Caller', 1, caller)] if not document else [('Helpers', 1, helpers)]:
        source += f'Set part = target.VBProject.VBComponents.Add({kind})\npart.Name = "{component}"\npart.CodeModule.AddFromString ' + literal(text) + '\n'
    if document:
        source += 'target.VBProject.VBComponents("Sheet1").CodeModule.AddFromString ' + literal(code) + '\n'
        source += 'Application.Run "\'" & target.Name & "\'!Sheet1.ExportValue"\nProbe = target.Worksheets(1).Range("A1").Value\n'
    else:
        source += 'Set part = target.VBProject.VBComponents.Add(2)\npart.Name = "Thing"\npart.CodeModule.AddFromString ' + literal(code) + '\n'
        source += 'Probe = Application.Run("\'" & target.Name & "\'!Caller.Probe")\n'
    return source + 'target.Close False\nEnd Function'
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
