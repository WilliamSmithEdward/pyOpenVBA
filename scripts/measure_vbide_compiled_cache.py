"""Measure source replacement versus already compiled module code."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig
OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vbide_compiled_cache.json'
PREFIX = '''Function Probe() As String
Dim target As Workbook, part As Object, other As Object, first As Long
Application.EnableEvents = False
Application.DisplayAlerts = False
Set target = Workbooks.Add
Set part = target.VBProject.VBComponents.Add(1)
part.Name = "Helpers"
part.CodeModule.AddFromString "Function Answer() As Long" & vbCrLf & "Answer = 1" & vbCrLf & "End Function"
'''
SUFFIX = '''Probe = CStr(Application.Run("'" & target.Name & "'!Helpers.Answer"))
target.Close False
End Function'''
def macro(name: str) -> str:
    source = PREFIX
    if name == 'replace_after_other_module':
        source += ('Set other = target.VBProject.VBComponents.Add(1)\nother.Name = "Other"\n'
                   'other.CodeModule.AddFromString "Function Answer() As Long" & vbCrLf & "Answer = 9" & vbCrLf & "End Function"\n'
                   'first = Application.Run("\'" & target.Name & "\'!Other.Answer")\n')
    elif name != 'replace_before_run':
        source += 'first = Application.Run("\'" & target.Name & "\'!Helpers.Answer")\n'
    source += 'part.CodeModule.ReplaceLine 2, "Answer = 3"\n'
    if name == 'replace_then_insert':
        source += 'part.CodeModule.InsertLines 1, "Private Extra As Long"\n'
    return source + SUFFIX

def main() -> None:
    probes = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        for name in ('replace_after_run', 'replace_before_run', 'replace_then_insert', 'replace_after_other_module'):
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
