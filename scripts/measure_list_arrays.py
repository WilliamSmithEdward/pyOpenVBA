"""Measure whole-list arrays and edits to range-backed lists in Excel."""
import json
from pathlib import Path
from zipfile import ZipFile

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "read": "",
    "replace": 'sh.ControlFormat.List = Array("x", "y")',
    "replace_mixed": 'sh.ControlFormat.List = Array(5, True, Empty)',
    "replace_empty": 'sh.ControlFormat.List = Array()',
    "replace_scalar": 'sh.ControlFormat.List = "x"',
    "append": 'sh.ControlFormat.AddItem "x"',
    "insert": 'sh.ControlFormat.AddItem "x", 2',
    "remove": 'sh.ControlFormat.RemoveItem 2',
    "clear": 'sh.ControlFormat.RemoveAllItems',
    "rename": 'sh.ControlFormat.List(2) = "changed"',
    "invalid_rename": 'sh.ControlFormat.List(8) = "changed"',
    "zero_rename": 'sh.ControlFormat.List(0) = "changed"',
    "negative_rename": 'sh.ControlFormat.List(-1) = "changed"',
    "replace_numbers": 'sh.ControlFormat.List = Array(5, 6)',
    "replace_boolean": 'sh.ControlFormat.List = Array(True, False)',
    "replace_blank": 'sh.ControlFormat.List = Array("x", Empty, "z")',
    "replace_null": 'sh.ControlFormat.List = Null',
    "replace_number": 'sh.ControlFormat.List = 5',
    "replace_bounds": 'Dim b(-2 To -1) As String\nb(-2) = "x"\nb(-1) = "y"\nsh.ControlFormat.List = b',
    "replace_matrix": 'Dim b(1 To 2, 1 To 2) As Variant\nb(1, 1) = "x"\nb(2, 1) = "y"\nb(1, 2) = "z"\nb(2, 2) = "w"\nsh.ControlFormat.List = b',
    "replace_empty_scalar": 'sh.ControlFormat.List = Empty',
    "replace_empty_string": 'sh.ControlFormat.List = Array("x", "", "z")',
    "read_detached": 'v = sh.ControlFormat.List\nv(1) = "local edit"',
}


def main() -> None:
    records = []
    folder = ROOT / ".cache/project-review/list-arrays"
    folder.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for kind in (2, 6):
            for bound in (False, True):
                for name, statement in PROBES.items():
                    setup = ('sh.ControlFormat.ListFillRange = "$J$1:$J$3"\n' if bound else
                             'sh.ControlFormat.AddItem "a"\nsh.ControlFormat.AddItem "b"\nsh.ControlFormat.AddItem "c"\n')
                    body = (
                        'Range("H1").ClearContents\nRange("J1").Value = "a"\n'
                        'Range("J2").Value = "b"\nRange("J3").Value = "c"\n'
                        f'Set sh = ActiveSheet.Shapes.AddFormControl({kind}, 0, 0, 90, 60)\n' + setup
                        + 'sh.ControlFormat.LinkedCell = "$H$1"\nsh.ControlFormat.Value = 2\n'
                        'On Error Resume Next\n' + statement + '\nn = Err.Number\nErr.Clear\n'
                        'v = sh.ControlFormat.List\na = TypeName(v)\n'
                        'a = a & ":" & CStr(LBound(v)) & ":" & CStr(UBound(v))\n'
                        'e = Err.Number\nOn Error GoTo 0\n'
                        'Report = CStr(n) & "|" & CStr(sh.ControlFormat.Value) & "|" & _\n'
                        'sh.ControlFormat.ListFillRange & "|" & CStr(Range("H1").Value) & "|" & a & ":" & CStr(e) & "|"\n'
                        'For i = 1 To sh.ControlFormat.ListCount\nReport = Report & sh.ControlFormat.List(i) & ";"\nNext i\n'
                        'Report = Report & "|" & CStr(Range("J1").Value) & ";" & CStr(Range("J2").Value) & ";" & CStr(Range("J3").Value)\n'
                    )
                    target = folder / "state.xlsm"
                    code = ('Public Function Report() As String\nDim sh As Object, n As Long, i As Long, v As Variant, a As String, e As Long\n'
                            'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                            + body + 'Application.DisplayAlerts = False\n'
                            + f'ActiveWorkbook.SaveAs "{target}", 52\nEnd Function\n')
                    result = excel.run_vba(code, "Report", timeout=120.0)
                    assert result.ok, f"{kind} {bound} {name}: {result.outcome}: {result.message}"
                    with ZipFile(target) as package:
                        parts = {part: package.read(part).decode() for part in package.namelist()
                                 if part.startswith("xl/ctrlProps/") or part.endswith(".vml")}
                    records.append({"kind": kind, "bound": bound, "name": name, "body": body,
                                    "reported": str(result.value), "parts": parts})
                    print(kind, bound, name, result.value)
    (ROOT / "tests/fixtures/shapes/list_arrays.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
