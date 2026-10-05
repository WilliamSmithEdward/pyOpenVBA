"""Measure worksheet event dispatch after editing VBIDE source."""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vbide_event_edits.json'
HANDLER = 'Private Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nRange("B1").Value = 1\nApplication.EnableEvents = True\nEnd Sub'

def literal(source: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in source.split('\n'))

def macro(name: str) -> str:
    source = ('Function Probe() As String\nDim target As Workbook, code As Object\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\n'
              'Set target = Workbooks.Add\nSet code = target.VBProject.VBComponents("Sheet1").CodeModule\n')
    if name != 'new_handler':
        source += 'code.AddFromString ' + literal(HANDLER) + '\n'
        source += 'Application.EnableEvents = True\ntarget.Worksheets(1).Range("A1").Value = 5\nApplication.EnableEvents = False\n'
    if name in ('new_handler', 'replace_handler'):
        if name == 'replace_handler':
            source += 'code.DeleteLines 1, code.CountOfLines\n'
        source += 'code.AddFromString ' + literal(HANDLER.replace('Value = 1', 'Value = 2')) + '\n'
    elif name in ('replace_line', 'replace_line_then_insert', 'replace_line_then_add'):
        source += 'code.ReplaceLine 3, "Range(""B1"").Value = 3"\n'
        if name == 'replace_line_then_insert':
            source += 'code.InsertLines 1, "Private Extra As Long"\n'
        elif name == 'replace_line_then_add':
            source += 'code.AddFromString "Private Extra As Long"\n'
    elif name == 'remove_handler':
        source += 'code.DeleteLines 1, code.CountOfLines\n'
    elif name == 'rename_document':
        source += 'target.VBProject.VBComponents("Sheet1").Name = "Info"\n'
    source += ('target.Worksheets(1).Range("B1").ClearContents\nApplication.EnableEvents = True\n'
               'target.Worksheets(1).Range("A1").Value = 7\n'
               'Probe = CStr(target.Worksheets(1).Range("B1").Value)\n'
               'Application.EnableEvents = False\ntarget.Close False\nApplication.EnableEvents = True\nEnd Function')
    return source

def main() -> None:
    probes = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        for name in ('unchanged', 'new_handler', 'replace_handler', 'replace_line', 'remove_handler', 'rename_document', 'replace_line_then_insert', 'replace_line_then_add'):
            source = macro(name)
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            print(name, repr(result.value), flush=True)
            probes.append({'name': name, 'macro': source, 'result': result.value})
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
    OUT.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')

if __name__ == '__main__':
    main()
