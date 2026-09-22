"""Record inline list editing and selection modes from live Excel."""
import json
from pathlib import Path
from zipfile import ZipFile

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "append": 'sh.ControlFormat.AddItem "d"',
    "insert": 'sh.ControlFormat.AddItem "x", 2',
    "insert_zero": 'sh.ControlFormat.AddItem "x", 0',
    "insert_end": 'sh.ControlFormat.AddItem "x", 4',
    "insert_large": 'sh.ControlFormat.AddItem "x", 5',
    "remove": 'sh.ControlFormat.RemoveItem 2',
    "remove_zero": 'sh.ControlFormat.RemoveItem 0',
    "remove_many": 'sh.ControlFormat.RemoveItem 2, 2',
    "clear": 'sh.ControlFormat.RemoveAllItems',
    "rename": 'sh.ControlFormat.List(2) = "changed"',
    "selected_insert": 'sh.ControlFormat.Value = 2\nsh.ControlFormat.AddItem "x", 1',
    "selected_remove": 'sh.ControlFormat.Value = 2\nsh.ControlFormat.RemoveItem 2',
    "multi": 'sh.ControlFormat.MultiSelect = 2\nsh.ControlFormat.Value = 2',
    "extended": 'sh.ControlFormat.MultiSelect = 3\nsh.ControlFormat.Value = 2',
    "multi_add": 'sh.ControlFormat.MultiSelect = 2\nsh.ControlFormat.Value = 1\nsh.ControlFormat.Value = 3',
    "multi_zero": 'sh.ControlFormat.MultiSelect = 2\nsh.ControlFormat.Value = 2\nsh.ControlFormat.Value = 0',
    "multi_single": 'sh.ControlFormat.MultiSelect = 2\nsh.ControlFormat.Value = 2\nsh.ControlFormat.MultiSelect = 1',
    "single_multi": 'sh.ControlFormat.Value = 2\nsh.ControlFormat.MultiSelect = 2',
    "invalid_mode": 'sh.ControlFormat.MultiSelect = 4',
    "remove_past": 'sh.ControlFormat.RemoveItem 2, 9',
    "remove_none": 'sh.ControlFormat.RemoveItem 2, 0',
    "selected_many": 'sh.ControlFormat.MultiSelect = 2\nsh.DrawingObject.Selected(1) = True\nsh.DrawingObject.Selected(3) = True',
    "source_from_inline": 'Range("J1").Value = "x"\nRange("J2").Value = "y"\nsh.ControlFormat.ListFillRange = "$J$1:$J$2"',
    "clear_inline_source": 'sh.ControlFormat.ListFillRange = ""',
    "deselect": 'sh.ControlFormat.Value = 2\nsh.DrawingObject.Selected(2) = False',
    "enum_none": 'sh.ControlFormat.MultiSelect = -4142',
    "enum_simple": 'sh.ControlFormat.MultiSelect = -4154\nsh.ControlFormat.Value = 2',
}


def main() -> None:
    records = []
    folder = ROOT / ".cache/project-review/inline-controls"
    folder.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for kind in (2, 6):
            for name, statement in PROBES.items():
                body = (
                    'Range("H1").ClearContents\n'
                    f'Set sh = ActiveSheet.Shapes.AddFormControl({kind}, 0, 0, 90, 60)\n'
                    'sh.ControlFormat.AddItem "a"\nsh.ControlFormat.AddItem "b"\nsh.ControlFormat.AddItem "c"\n'
                    'sh.ControlFormat.LinkedCell = "$H$1"\nOn Error Resume Next\n'
                    + statement + '\nn = Err.Number\nErr.Clear\nv = CStr(sh.ControlFormat.Value)\n'
                    'If Err.Number <> 0 Then v = "error" & CStr(Err.Number)\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|" & v & "|" & _\n'
                    'CStr(sh.ControlFormat.ListCount) & "|" & _\n'
                    'CStr(Range("H1").Value) & "|"\n'
                    'For i = 1 To sh.ControlFormat.ListCount\n'
                    'Report = Report & sh.ControlFormat.List(i) & ";"\nNext i\n'
                )
                target = folder / "state.xlsm"
                code = ('Public Function Report() As String\nDim sh As Object, n As Long, i As Long, v As String\n'
                        'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                        + body + 'Application.DisplayAlerts = False\n'
                        + f'ActiveWorkbook.SaveAs "{target}", 52\nEnd Function\n')
                result = excel.run_vba(code, "Report", timeout=120.0)
                assert result.ok, f"{result.outcome}: {result.message}"
                with ZipFile(target) as package:
                    parts = {part: package.read(part).decode() for part in package.namelist()
                             if part.startswith("xl/ctrlProps/") or part.endswith(".vml")}
                records.append({"kind": kind, "name": name, "body": body,
                                "reported": str(result.value), "parts": parts})
                print(kind, name, result.value)
    (ROOT / "tests/fixtures/shapes/inline_controls.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8",
    )


if __name__ == "__main__":
    main()
