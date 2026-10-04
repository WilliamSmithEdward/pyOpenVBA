"""Record native Excel's in-memory VBIDE project and source editing rules."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/vbide_projects.json"
INITIAL = 'Option Explicit\nPublic Marker As Long\n\nPublic Function ReadMarker() As Long\nReadMarker = Marker\nEnd Function'
ACTIONS = {
    "empty": "",
    "add_standard": "Set part = parts.Add(1)",
    "add_class": "Set part = parts.Add(2)",
    "add_document": "Set part = parts.Add(100)",
    "add_unknown": "Set part = parts.Add(999)",
    "lookup_zero": "Set part = parts.Item(0)",
    "lookup_fraction": "Set part = parts.Item(1.5)",
    "lookup_missing": 'Set part = parts.Item("missing")',
    "remove_document": "parts.Remove parts.Item(1)",
    "remove_standard": "Set part = parts.Add(1)\nparts.Remove part\nSet part = Nothing",
    "rename_document": 'parts.Item("Sheet1").Name = "Info"',
    "rename_invalid": 'parts.Item("Sheet1").Name = "bad name"',
    "rename_collision": 'parts.Item("Sheet1").Name = "ThisWorkbook"',
    "rename_case": 'parts.Item("Sheet1").Name = "sHeEt1"',
    "project_rename": 'project.Name = "OtherProject"',
    "project_invalid": 'project.Name = "bad name"',
    "source_initial": "",
    "source_add": 'code.AddFromString "Public Added As Long"',
    "source_add_empty": 'code.AddFromString ""',
    "source_add_invalid": 'code.AddFromString "this is invalid syntax"',
    "source_add_missing_end": 'code.AddFromString "Public Sub Broken()"',
    "source_add_attribute": 'code.AddFromString "Attribute VB_Name = ""Other"""',
    "source_insert_blank": 'code.InsertLines 3, ""',
    "source_insert_first": 'code.InsertLines 1, "Public Added As Long"',
    "source_insert_zero": 'code.InsertLines 0, "Public Added As Long"',
    "source_insert_negative": 'code.InsertLines -1, "Public Added As Long"',
    "source_insert_after": 'code.InsertLines 999, "Public Added As Long"',
    "source_replace_first": 'code.ReplaceLine 1, "Option Base 1"',
    "source_replace_zero": 'code.ReplaceLine 0, "Option Base 1"',
    "source_replace_after": 'code.ReplaceLine 999, "Option Base 1"',
    "source_delete_first": "code.DeleteLines 1",
    "source_delete_zero": "code.DeleteLines 0",
    "source_delete_after": "code.DeleteLines 999",
    "source_delete_oversized": "code.DeleteLines 2, 999",
    "source_delete_count_zero": "code.DeleteLines 2, 0",
    "source_delete_count_negative": "code.DeleteLines 2, -1",
    "source_lines_zero": 'extra = code.Lines(0, 1)',
    "source_lines_after": 'extra = code.Lines(999, 1)',
    "source_lines_oversized": 'extra = code.Lines(2, 999)',
    "source_lines_count_zero": 'extra = code.Lines(2, 0)',
    "source_lines_count_negative": 'extra = code.Lines(2, -1)',
    "source_proc_start": 'extra = CStr(code.ProcStartLine("ReadMarker", 0))',
    "source_proc_body": 'extra = CStr(code.ProcBodyLine("ReadMarker", 0))',
    "source_proc_count": 'extra = CStr(code.ProcCountLines("ReadMarker", 0))',
    "source_proc_of_decl": 'extra = code.ProcOfLine(2, 0)',
    "source_proc_of_body": 'extra = code.ProcOfLine(5, 0)',
    "source_proc_missing": 'extra = CStr(code.ProcStartLine("Missing", 0))',
}


def literal(text: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in text.split('\n'))


def main() -> None:
    rows: list[dict[str, object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name, action in ACTIONS.items():
            excel.new_document()
            source = ('Function Probe() As String\nDim target As Workbook, project As Object, parts As Object\n'
                      'Dim part As Object, code As Object, item As Object, number As Long, extra As String\n'
                      'Application.EnableEvents = False\nApplication.DisplayAlerts = False\n'
                      'Set target = Workbooks.Add\nSet project = target.VBProject\nSet parts = project.VBComponents\n')
            if name.startswith("source_"):
                source += 'Set part = parts.Add(1)\nSet code = part.CodeModule\ncode.AddFromString ' + literal(INITIAL) + '\n'
            source += ('On Error Resume Next\n' + action + '\nnumber = Err.Number\nOn Error GoTo 0\n'
                       'Probe = "#error:" & CStr(number) & "#extra:" & extra & "#project:" & project.Name\n'
                       'For Each item In parts\nProbe = Probe & "|" & item.Name & ":" & CStr(item.Type) & ":" & CStr(item.CodeModule.CountOfLines)\nNext\n'
                       'If Not code Is Nothing Then\nProbe = Probe & "#decl:" & CStr(code.CountOfDeclarationLines)\n'
                       'If code.CountOfLines > 0 Then Probe = Probe & "#source:" & code.Lines(1, code.CountOfLines)\nEnd If\n'
                       'Probe = Probe & "#sheet:" & target.Worksheets(1).CodeName & "#saved:" & CStr(target.Saved)\n'
                       'target.Close False\nEnd Function')
            result = excel.run_vba(source, "Probe", timeout=30.0)
            assert result.ok and not result.dialogs, (name, result)
            rows.append({"name": name, "action": action, "result": result.value})
            print(name, repr(result.value), flush=True)
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok, version
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(),
                               "excel_version_build": version.value, "initial_source": INITIAL,
                               "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == "__main__":
    main()
