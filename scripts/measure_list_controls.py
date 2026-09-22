"""Record Excel dropdown/list-box values and their saved control properties."""
import json
from pathlib import Path
from zipfile import ZipFile

from pyvbaharness import ExcelSession

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "zero": "sh.ControlFormat.Value = 0",
    "two": "sh.ControlFormat.Value = 2",
    "too_large": "sh.ControlFormat.Value = 4",
    "negative": "sh.ControlFormat.Value = -1",
    "fraction": "sh.ControlFormat.Value = 1.5",
    "true": "sh.ControlFormat.Value = True",
    "clear_selection": "sh.ControlFormat.Value = 2\nsh.ControlFormat.Value = 0",
    "cell_two": 'Range("H1").Value = 2',
    "cell_large": 'Range("H1").Value = 4',
    "cell_negative": 'Range("H1").Value = -1',
    "cell_fraction": 'Range("H1").Value = 1.5',
    "cell_text": 'sh.ControlFormat.Value = 2\nRange("H1").Value = "hello"',
    "cell_clear": 'sh.ControlFormat.Value = 2\nRange("H1").ClearContents',
    "cell_error": 'sh.ControlFormat.Value = 2\nRange("H1").Value = CVErr(2042)',
    "formula": 'Range("H1").Formula = "=G1+1"\nRange("G1").Value = 1',
    "rebind": 'sh.ControlFormat.LinkedCell = ""\nRange("H1").Value = 2\nsh.ControlFormat.LinkedCell = "$H$1"',
    "shrink": 'sh.ControlFormat.Value = 3\nsh.ControlFormat.ListFillRange = "$J$1:$J$2"',
    "clear_source": 'sh.ControlFormat.Value = 2\nsh.ControlFormat.ListFillRange = ""',
    "horizontal": 'sh.ControlFormat.ListFillRange = "$J$1:$L$1"',
    "rectangle": 'sh.ControlFormat.ListFillRange = "$J$1:$K$3"',
    "blank_source": 'sh.ControlFormat.ListFillRange = "$L$1:$L$3"',
    "shrink_expand": 'sh.ControlFormat.Value = 3\nsh.ControlFormat.ListFillRange = "$J$1:$J$2"\nsh.ControlFormat.ListFillRange = "$J$1:$J$3"',
    "clear_restore": 'sh.ControlFormat.Value = 2\nsh.ControlFormat.ListFillRange = ""\nsh.ControlFormat.ListFillRange = "$J$1:$J$3"',
    "unbound_shrink_expand": 'sh.ControlFormat.LinkedCell = ""\nsh.ControlFormat.Value = 3\nsh.ControlFormat.ListFillRange = "$J$1:$J$2"\nsh.ControlFormat.ListFillRange = "$J$1:$J$3"',
    "unbound_clear_restore": 'sh.ControlFormat.LinkedCell = ""\nsh.ControlFormat.Value = 2\nsh.ControlFormat.ListFillRange = ""\nsh.ControlFormat.ListFillRange = "$J$1:$J$3"',
}


def main() -> None:
    records = []
    folder = ROOT / ".cache/project-review/list-controls"
    folder.mkdir(parents=True, exist_ok=True)
    with ExcelSession() as excel:
        excel.new_document()
        for kind in (2, 6):
            for name, statement in PROBES.items():
                setup = (
                    'Range("G1:H1").ClearContents\nRange("J1").Value = "a"\n'
                    'Range("J2").Value = "b"\nRange("J3").Value = "c"\n'
                    f'Set sh = ActiveSheet.Shapes.AddFormControl({kind}, 0, 0, 90, 60)\n'
                    'sh.ControlFormat.ListFillRange = "$J$1:$J$3"\n'
                    'sh.ControlFormat.LinkedCell = "$H$1"\n'
                )
                body = (
                    setup + 'On Error Resume Next\n' + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|" & CStr(sh.ControlFormat.Value) & "|" & _\n'
                    'CStr(sh.ControlFormat.ListCount) & "|" & TypeName(Range("H1").Value) & "|" & _\n'
                    'CStr(Range("H1").Value)\n'
                )
                target = folder / "state.xlsm"
                code = ('Public Function Report() As String\nDim sh As Object, n As Long\n'
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
    (ROOT / "tests/fixtures/shapes/list_controls.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8",
    )


if __name__ == "__main__":
    main()
