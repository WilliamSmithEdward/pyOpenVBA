"""Measure VBIDE import/export, AddFromFile and physical/editor line boundaries."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vbide_source_files.json'
BODY = 'Option Explicit\nPublic Marker As Long\n\nPublic Function Answer() As Long\nAnswer = 42\nEnd Function'
BAS = 'Attribute VB_Name = "Imported"\r\n' + BODY.replace('\n', '\r\n') + '\r\n'
CLS = ('VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n'
       'Attribute VB_Name = "ImportedClass"\r\nAttribute VB_GlobalNameSpace = False\r\n'
       'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\n'
       'Attribute VB_Exposed = False\r\n' + BODY.replace('\n', '\r\n') + '\r\n')
CASES = {
    'export_standard': ('Set part = parts.Add(1)\npart.Name = "Source"', 'part.Export outputPath'),
    'export_class': ('Set part = parts.Add(2)\npart.Name = "Source"', 'part.Export outputPath'),
    'export_document': ('Set part = parts.Item("Sheet1")', 'part.Export outputPath'),
    'export_existing': ('Set part = parts.Add(1)', 'part.Export outputPath\npart.Export outputPath'),
    'export_missing_folder': ('Set part = parts.Add(1)', 'part.Export outputPath & "\\missing\\source.bas"'),
    'import_standard': ('', 'Set part = parts.Import(inputPath)'),
    'import_class': ('', 'Set part = parts.Import(inputPath)'),
    'import_twice': ('Set part = parts.Import(inputPath)', 'Set part = parts.Import(inputPath)'),
    'import_missing': ('', 'Set part = parts.Import(inputPath & ".missing")'),
    'add_file_standard': ('Set part = parts.Add(1)', 'part.CodeModule.AddFromFile inputPath'),
    'add_file_class': ('Set part = parts.Add(1)', 'part.CodeModule.AddFromFile inputPath'),
    'add_file_missing': ('Set part = parts.Add(1)', 'part.CodeModule.AddFromFile inputPath & ".missing"'),
    'insert_empty': ('Set part = parts.Add(1)', 'part.CodeModule.InsertLines 1, ""'),
    'insert_newline': ('Set part = parts.Add(1)', 'part.CodeModule.InsertLines 1, vbCrLf'),
    'insert_double_newline': ('Set part = parts.Add(1)', 'part.CodeModule.InsertLines 1, vbCrLf & vbCrLf'),
    'add_empty': ('Set part = parts.Add(1)', 'part.CodeModule.AddFromString ""'),
    'replace_blank_only': ('Set part = parts.Add(1)\npart.CodeModule.InsertLines 1, "x"', 'part.CodeModule.ReplaceLine 1, ""'),
    'delete_blank_only': ('Set part = parts.Add(1)\npart.CodeModule.InsertLines 1, ""', 'part.CodeModule.DeleteLines 1'),
    'remove_add': ('Set part = parts.Add(1)\nparts.Remove part', 'Set part = parts.Add(1)'),
    'rename_reference_vba': ('Set part = parts.Add(1)', 'part.Name = "VBA"'),
    'rename_reference_stdole': ('Set part = parts.Add(1)', 'part.Name = "stdole"'),
    'rename_reference_office': ('Set part = parts.Add(1)', 'part.Name = "Office"'),
    'file_name': ('Set part = parts.Add(1)', 'extra = project.FileName'),
    'project_type': ('Set part = parts.Add(1)', 'extra = CStr(project.Type)'),
    'project_mode': ('Set part = parts.Add(1)', 'extra = CStr(project.Mode)'),
    'project_protection': ('Set part = parts.Add(1)', 'extra = CStr(project.Protection)'),
    'save_blank': ('Set part = parts.Add(1)\npart.CodeModule.InsertLines 1, ""',
                   'target.SaveAs outputPath & ".xlsm", 52\ntarget.Close False\nSet target = Workbooks.Open(outputPath & ".xlsm")\nSet project = target.VBProject\nSet parts = project.VBComponents\nSet part = parts.Item("Module1")'),
    'save_trailing': ('Set part = parts.Add(1)\npart.CodeModule.AddFromString "Public Value As Long" & vbCrLf',
                      'target.SaveAs outputPath & ".xlsm", 52\ntarget.Close False\nSet target = Workbooks.Open(outputPath & ".xlsm")\nSet project = target.VBProject\nSet parts = project.VBComponents\nSet part = parts.Item("Module1")'),
    'save_leading': ('Set part = parts.Add(1)\npart.CodeModule.InsertLines 1, vbCrLf & "Public Value As Long"',
                     'target.SaveAs outputPath & ".xlsm", 52\ntarget.Close False\nSet target = Workbooks.Open(outputPath & ".xlsm")\nSet project = target.VBProject\nSet parts = project.VBComponents\nSet part = parts.Item("Module1")'),
}


def literal(text: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in text.split('\n'))


def source_for(name: str, input_path: Path, output_path: Path) -> str:
    setup, action = CASES[name]
    source = ('Function Probe() As String\nDim target As Workbook, project As Object, parts As Object\n'
              'Dim part As Object, item As Object, code As Object, number As Long, extra As String\n'
              'Dim inputPath As String, outputPath As String\nApplication.EnableEvents = False\n'
              'Application.DisplayAlerts = False\nSet target = Workbooks.Add\n'
              'Set project = target.VBProject\nSet parts = project.VBComponents\n'
              'inputPath = "' + str(input_path).replace('"', '""') + '"\n'
              'outputPath = "' + str(output_path).replace('"', '""') + '"\n' + setup + '\n')
    if name.startswith('export_'):
        source += 'part.CodeModule.AddFromString ' + literal(BODY) + '\n'
    source += ('On Error Resume Next\n' + action + '\nnumber = Err.Number\nOn Error GoTo 0\n'
               'Probe = "#error:" & CStr(number) & "#extra:" & extra\n'
               'For Each item In parts\nProbe = Probe & "|" & item.Name & ":" & CStr(item.Type) & ":" & CStr(item.CodeModule.CountOfLines)\nNext\n'
               'If Not part Is Nothing Then\nSet code = part.CodeModule\n'
               'Probe = Probe & "#part:" & part.Name & "#decl:" & CStr(code.CountOfDeclarationLines)\n'
               'If code.CountOfLines > 0 Then Probe = Probe & "#source:" & code.Lines(1, code.CountOfLines)\nEnd If\n'
               'Probe = Probe & "#saved:" & CStr(target.Saved)\ntarget.Close False\nEnd Function')
    return source


def main() -> None:
    rows: list[dict[str, object]] = []
    with TemporaryDirectory(prefix='pyopenvba-vbide-') as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name in CASES:
            excel.new_document()
            input_path = Path(folder) / ('input.cls' if name.endswith('class') else 'input.bas')
            input_path.write_bytes((CLS if name.endswith('class') else BAS).encode('cp1252'))
            output_path = Path(folder) / (name + '.txt')
            result = excel.run_vba(source_for(name, input_path, output_path), 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            exported = output_path.read_bytes().decode('cp1252') if output_path.is_file() else None
            rows.append({'name': name, 'setup': CASES[name][0], 'action': CASES[name][1],
                         'macro': source_for(name, Path('INPUT'), Path('OUTPUT')), 'result': result.value, 'exported': exported})
            print(name, repr(result.value), repr(exported), flush=True)
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok, version
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel_version_build': version.value,
                               'body': BODY, 'bas': BAS, 'cls': CLS, 'probes': rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
