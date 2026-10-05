"""Measure deferred project compilation and reset effects of VBIDE mutations."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vbide_compile_lifecycle.json'
HELPERS = 'Private Stored As Long\nPublic Sub Mark(ByVal number As Long)\nStored = number\nEnd Sub\nPublic Function Read() As Long\nRead = Stored\nEnd Function'
DOCUMENT = 'Public Stored As Long\nPublic Sub Mark(ByVal number As Long)\nStored = number\nEnd Sub\nPublic Sub ExportValue()\nRange("A1").Value = Stored\nEnd Sub'
ACTIONS = {
    'unchanged': '',
    'edit_standard': 'parts.Item("Helpers").CodeModule.AddFromString "Private Added As Long"',
    'edit_document': 'parts.Item("Sheet1").CodeModule.AddFromString "Private Added As Long"',
    'add_component': 'Set part = parts.Add(1)',
    'remove_component': 'Set part = parts.Add(1)\nparts.Remove part',
    'rename_standard': 'parts.Item("Helpers").Name = "Renamed"\nmoduleName = "Renamed"',
    'rename_document': 'parts.Item("Sheet1").Name = "Info"\ndocumentName = "Info"',
    'rename_project': 'target.VBProject.Name = "OtherProject"',
    'read_catalog': 'number = parts.Count',
    'read_lines': 'text = parts.Item("Helpers").CodeModule.Lines(1, 100)',
    'open_invalid': '',
}


def literal(text: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in text.split('\n'))


def macro(name: str, path: str) -> str:
    source = ('Function Probe() As String\nDim target As Workbook, parts As Object, part As Object\n'
              'Dim moduleName As String, documentName As String, text As String, number As Long\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n'
              'Set parts = target.VBProject.VBComponents\nSet part = parts.Add(1)\npart.Name = "Helpers"\n'
              'part.CodeModule.AddFromString ' + literal(HELPERS) + '\n'
              'parts.Item("Sheet1").CodeModule.AddFromString ' + literal(DOCUMENT) + '\n'
              'moduleName = "Helpers"\ndocumentName = "Sheet1"\n'
              'Application.Run "\'" & target.Name & "\'!Helpers.Mark", 42\n'
              'Application.Run "\'" & target.Name & "\'!Sheet1.Mark", 33\n')
    if name == 'open_invalid':
        source += ('parts.Item("Helpers").CodeModule.AddFromString "this is invalid syntax"\n'
                   'target.SaveAs "' + path.replace('"', '""') + '", 52\ntarget.Close False\n'
                   'Set target = Workbooks.Open("' + path.replace('"', '""') + '")\n'
                   'Probe = CStr(target.VBProject.VBComponents("Helpers").CodeModule.CountOfLines)\n')
    else:
        source += ACTIONS[name] + '\n'
        source += ('Probe = CStr(Application.Run("\'" & target.Name & "\'!" & moduleName & ".Read"))\n'
                   'Application.Run "\'" & target.Name & "\'!" & documentName & ".ExportValue"\n'
                   'Probe = Probe & ":" & CStr(target.Worksheets(1).Range("A1").Value)\n')
    return source + 'target.Close False\nEnd Function'


def main() -> None:
    rows: list[dict[str, object]] = []
    with TemporaryDirectory(prefix='pyopenvba-compile-') as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name in ACTIONS:
            excel.new_document()
            source = macro(name, str(Path(folder) / 'invalid.xlsm'))
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            rows.append({'name': name, 'action': ACTIONS[name], 'macro': macro(name, 'OUTPUT'), 'result': result.value})
            print(name, repr(result.value), flush=True)
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok, version
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel_version_build': version.value,
                               'probes': rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
