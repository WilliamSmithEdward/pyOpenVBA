"""Measure Forms control bindings through workbook defined names."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "plain": 'sh.ControlFormat.LinkedCell = "Target"',
    "equals": 'sh.ControlFormat.LinkedCell = "=Target"',
    "case": 'sh.ControlFormat.LinkedCell = "target"',
    "missing": 'sh.ControlFormat.LinkedCell = "MissingName"',
    "multi_cell": 'sh.ControlFormat.LinkedCell = "Choices"',
    "list_name": 'sh.ControlFormat.ListFillRange = "Choices"',
    "list_equals": 'sh.ControlFormat.ListFillRange = "=Choices"',
    "list_missing": 'sh.ControlFormat.ListFillRange = "MissingName"',
    "list_cell": 'sh.ControlFormat.ListFillRange = "Target"',
    "value_write": 'sh.ControlFormat.LinkedCell = "Target"\nsh.ControlFormat.Value = 2',
    "cell_write": 'sh.ControlFormat.LinkedCell = "Target"\nRange("H1").Value = 2',
    "retarget": 'sh.ControlFormat.LinkedCell = "Target"\nActiveWorkbook.Names("Target").RefersTo = "=Sheet1!$H$2"\nsh.ControlFormat.Value = 2',
    "missing_write": 'sh.ControlFormat.LinkedCell = "MissingName"\nsh.ControlFormat.Value = 2',
    "multi_write": 'sh.ControlFormat.LinkedCell = "Choices"\nsh.ControlFormat.Value = 2',
    "list_case": 'sh.ControlFormat.ListFillRange = "choices"',
    "list_retarget": 'sh.ControlFormat.ListFillRange = "Choices"\nActiveWorkbook.Names("Choices").RefersTo = "=Sheet1!$J$1:$J$2"',
    "list_delete": 'sh.ControlFormat.ListFillRange = "Choices"\nActiveWorkbook.Names("Choices").Delete',
    "link_delete": 'sh.ControlFormat.LinkedCell = "Target"\nActiveWorkbook.Names("Target").Delete\nsh.ControlFormat.Value = 2',
    "alias": 'ActiveWorkbook.Names.Add "Alias", "=Target"\nsh.ControlFormat.LinkedCell = "Alias"\nsh.ControlFormat.Value = 2',
    "local": 'ActiveWorkbook.Names.Add "Sheet1!LocalTarget", "=Sheet1!$H$2"\nsh.ControlFormat.LinkedCell = "LocalTarget"\nsh.ControlFormat.Value = 2',
    "qualified_local": 'ActiveWorkbook.Names.Add "Sheet1!LocalTarget", "=Sheet1!$H$2"\nsh.ControlFormat.LinkedCell = "Sheet1!LocalTarget"\nsh.ControlFormat.Value = 2',
    "local_shadows": 'ActiveWorkbook.Names.Add "Sheet1!Target", "=Sheet1!$H$2"\nsh.ControlFormat.LinkedCell = "Target"\nsh.ControlFormat.Value = 2',
    "list_shrink_selected": 'sh.ControlFormat.ListFillRange = "Choices"\nsh.ControlFormat.Value = 3\nActiveWorkbook.Names("Choices").RefersTo = "=Sheet1!$J$1:$J$2"',
    "list_delete_selected": 'sh.ControlFormat.ListFillRange = "Choices"\nsh.ControlFormat.Value = 2\nActiveWorkbook.Names("Choices").Delete',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, statement in PROBES.items():
            body = ('Range("H1:H2").ClearContents\nRange("J1").Value = "a"\nRange("J2").Value = "b"\nRange("J3").Value = "c"\n'
                    'ActiveWorkbook.Names.Add "Target", "=Sheet1!$H$1"\n'
                    'ActiveWorkbook.Names.Add "Choices", "=Sheet1!$J$1:$J$3"\n'
                    'Set sh = ActiveSheet.Shapes.AddFormControl(6, 0, 0, 90, 60)\n'
                    'sh.ControlFormat.ListFillRange = "$J$1:$J$3"\nOn Error Resume Next\n'
                    + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|" & sh.ControlFormat.LinkedCell & "|" & sh.ControlFormat.ListFillRange & "|" & _\n'
                    'CStr(sh.ControlFormat.Value) & "|" & CStr(sh.ControlFormat.ListCount) & "|" & _\n'
                    'CStr(Range("H1").Value) & "|" & CStr(Range("H2").Value) & "|" & _\n'
                    'CStr(Range("J1").Value) & ":" & CStr(Range("J2").Value) & ":" & CStr(Range("J3").Value)\n')
            code = ('Public Function Report() As String\nDim sh As Object, n As Long\n'
                    'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                    + body + 'End Function\n')
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, result.message
            records.append({"name": name, "body": body, "reported": str(result.value)})
            print(name, result.value)
    (ROOT / "tests/fixtures/shapes/control_names.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
