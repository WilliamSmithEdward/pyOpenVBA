"""Measure worksheet names, code names and document-module lifecycle."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/worksheet_documents.json"
DOCUMENT = '''Public Marker As Long
Public Sub Mark(ByVal value As Long)
Marker = value
End Sub
Public Sub ExportMarker()
Range("B1").Value = Marker
End Sub
'''
ACTIONS = {
    "unmaterialized": '',
    "unmaterialized_add": 'Set added = target.Worksheets.Add',
    "initial": '',
    "add_one": 'Set added = target.Worksheets.Add',
    "add_two": 'Set added = target.Worksheets.Add(Count:=2)',
    "add_fraction": 'Set added = target.Worksheets.Add(Count:=1.5)',
    "add_fraction_two": 'Set added = target.Worksheets.Add(Count:=2.9)',
    "add_zero": 'On Error Resume Next\nSet added = target.Worksheets.Add(Count:=0)\nnumber = Err.Number\nOn Error GoTo 0',
    "add_negative": 'On Error Resume Next\nSet added = target.Worksheets.Add(Count:=-1)\nnumber = Err.Number\nOn Error GoTo 0',
    "rename_visible_then_add": 'target.Worksheets(1).Name = "Sheet7"\nSet added = target.Worksheets.Add',
    "delete_latest_then_add": 'Set added = target.Worksheets.Add\nadded.Delete\nSet added = target.Worksheets.Add',
    "delete_first_then_add": 'Set added = target.Worksheets.Add\ntarget.Worksheets("Sheet1").Delete\nSet added = target.Worksheets.Add',
    "rename_code_then_add": 'target.VBProject.VBComponents("Sheet1").Name = "Info"\nSet added = target.Worksheets.Add',
    "copy_same_book": 'target.Worksheets(1).Copy After:=target.Worksheets(1)',
    "copy_custom_code": 'target.VBProject.VBComponents("Sheet1").Name = "Info"\ntarget.Worksheets(1).Copy After:=target.Worksheets(1)',
    "copy_marker_reset": 'Application.Run "\'" & target.Name & "\'!Sheet1.Mark", 42\nApplication.Run "\'" & target.Name & "\'!Sheet1.ExportMarker"\ntarget.Worksheets(1).Copy After:=target.Worksheets(1)\nApplication.Run "\'" & target.Name & "\'!" & ActiveSheet.CodeName & ".ExportMarker"',
    "copy_then_add": 'target.Worksheets(1).Copy After:=target.Worksheets(1)\nSet added = target.Worksheets.Add',
    "copy_new_book": 'target.Worksheets(1).Copy\nSet other = target\nSet target = ActiveWorkbook',
    "copy_other_book": 'Set other = Workbooks.Add\ntarget.Worksheets(1).Copy After:=other.Worksheets(1)\nSet original = target\nSet target = other\nSet other = original',
    "move_other_book": 'Set other = Workbooks.Add\ntarget.Worksheets(1).Move After:=other.Worksheets(1)\nSet target = other\nSet other = Nothing',
    "move_marker_reset": 'Application.Run "\'" & target.Name & "\'!Sheet1.Mark", 42\nApplication.Run "\'" & target.Name & "\'!Sheet1.ExportMarker"\nSet other = Workbooks.Add\ntarget.Worksheets(1).Move After:=other.Worksheets(1)\nSet target = other\nSet other = Nothing\nApplication.Run "\'" & target.Name & "\'!" & ActiveSheet.CodeName & ".ExportMarker"',
    "delete_document": 'Set added = target.Worksheets.Add\ntarget.Worksheets("Sheet1").Delete',
}


def main() -> None:
    rows: list[dict[str, object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name, action in ACTIONS.items():
            excel.new_document()
            source = ('Function Probe() As String\nDim target As Workbook, other As Workbook, original As Workbook\n'
                      'Dim added As Worksheet, sheet As Worksheet, part As Object, text As String, number As Long\n'
                      'Application.EnableEvents = False\nApplication.DisplayAlerts = False\nSet target = Workbooks.Add\n')
            if not name.startswith("unmaterialized"):
                source += 'Set part = target.VBProject.VBComponents("Sheet1")\ntext = ""\n'
                source += '\n'.join('text = text & "' + line.replace('"', '""') + '" & vbLf' for line in DOCUMENT.splitlines())
                source += '\npart.CodeModule.AddFromString text\n'
            source += action
            source += ('\nProbe = "#book:" & target.CodeName & "|"\nFor Each sheet In target.Worksheets\nProbe = Probe & sheet.Name & ":" & sheet.CodeName & "|"\nNext\n'
                       'For Each part In target.VBProject.VBComponents\nProbe = Probe & "#component:" & part.Name & ":"'
                       ' & CStr(part.Type) & ":" & CStr(part.CodeModule.CountOfLines) & "|"\nNext\n'
                       'For Each sheet In target.Worksheets\nProbe = Probe & "#after:" & sheet.Name & ":" & sheet.CodeName'
                       ' & ":" & CStr(sheet.Range("B1").Value) & "|"\nNext\n'
                       'If Not added Is Nothing Then Probe = Probe & "#added:" & added.Name & ":" & added.CodeName\n'
                       'Probe = Probe & "#modules:" & CStr(target.VBProject.VBComponents.Count) & "#error:" & CStr(number)\n'
                       'target.Close False\nIf Not other Is Nothing Then other.Close False\nEnd Function')
            result = excel.run_vba(source, 'Probe', timeout=30.0)
            assert result.ok, (name, result)
            rows.append({"name": name, "action": action, "result": result.value})
            print(name, result.value, flush=True)
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok, version
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "excel_version_build": version.value,
                               "document": DOCUMENT, "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
