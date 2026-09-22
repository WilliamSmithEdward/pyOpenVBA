"""Measure Forms option-button grouping, values and linked cells in Excel."""
import json
from pathlib import Path
from shutil import copyfile
from zipfile import ZipFile

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "defaults": "",
    "off_other": "a.ControlFormat.Value = 1\nb.ControlFormat.Value = 0",
    "unchanged": 'a.ControlFormat.Value = 1\nRange("H1").Value = "text"\na.ControlFormat.Value = 1',
    "cell_fraction": 'Range("H1").Value = 1.8',
    "cell_negative": 'a.ControlFormat.Value = 1\nRange("H1").Value = -1',
    "cell_boolean": 'Range("H1").Value = True',
    "partial_overlap": 'c.Left = 100\na.ControlFormat.Value = 1\nc.ControlFormat.Value = 1',
    "move_out": 'b.Top = 110\na.ControlFormat.Value = 1\nb.ControlFormat.Value = 1',
    "late_box": 'ActiveSheet.Shapes.AddFormControl 4, 0, 0, 150, 100\na.ControlFormat.Value = 1\nc.ControlFormat.Value = 1',
    "select_first": "a.ControlFormat.Value = 1",
    "select_second": "b.ControlFormat.Value = 1",
    "switch": "a.ControlFormat.Value = 1\nb.ControlFormat.Value = 1",
    "outside": "a.ControlFormat.Value = 1\nc.ControlFormat.Value = 1",
    "off": "b.ControlFormat.Value = 1\nb.ControlFormat.Value = -4146",
    "zero": "b.ControlFormat.Value = 1\nb.ControlFormat.Value = 0",
    "mixed": "b.ControlFormat.Value = 2",
    "invalid": "b.ControlFormat.Value = 3",
    "cell_two": 'Range("H1").Value = 2',
    "cell_three": 'Range("H1").Value = 3',
    "cell_large": 'a.ControlFormat.Value = 1\nRange("H1").Value = 9',
    "cell_zero": 'a.ControlFormat.Value = 1\nRange("H1").Value = 0',
    "cell_text": 'a.ControlFormat.Value = 1\nRange("H1").Value = "text"',
    "cell_error": 'a.ControlFormat.Value = 1\nRange("H1").Value = CVErr(2042)',
    "cell_clear": 'a.ControlFormat.Value = 1\nRange("H1").ClearContents',
    "rebind": 'b.ControlFormat.Value = 1\nRange("H2").Value = 9\nb.ControlFormat.LinkedCell = "$H$2"',
    "unlink": 'b.ControlFormat.LinkedCell = ""\nb.ControlFormat.Value = 1',
}


def main() -> None:
    folder = ROOT / ".cache/project-review/radio-controls"
    folder.mkdir(parents=True, exist_ok=True)
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for boxed in (False, True):
            for name, statement in PROBES.items():
                body = ('Range("H1:H2").ClearContents\n'
                        + ('ActiveSheet.Shapes.AddFormControl 4, 0, 0, 150, 100\n' if boxed else '')
                        + 'Set a = ActiveSheet.Shapes.AddFormControl(7, 10, 10, 90, 20)\n'
                        'Set b = ActiveSheet.Shapes.AddFormControl(7, 10, 40, 90, 20)\n'
                        'Set c = ActiveSheet.Shapes.AddFormControl(7, 200, 10, 90, 20)\n'
                        'a.Name = "First"\nb.Name = "Second"\nc.Name = "Third"\n'
                        'a.ControlFormat.LinkedCell = "$H$1"\nOn Error Resume Next\n'
                        + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                        'Report = CStr(n) & "|" & CStr(a.ControlFormat.Value) & "|" & CStr(b.ControlFormat.Value) & "|" & _\n'
                        'CStr(c.ControlFormat.Value) & "|" & CStr(Range("H1").Value) & "|" & CStr(Range("H2").Value) & "|" & _\n'
                        'a.ControlFormat.LinkedCell & "|" & b.ControlFormat.LinkedCell & "|" & c.ControlFormat.LinkedCell\n')
                target = folder / "state.xlsm"
                code = ('Public Function Report() As String\nDim a As Object, b As Object, c As Object, n As Long\n'
                        'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                        + body + 'Application.DisplayAlerts = False\n'
                        + f'ActiveWorkbook.SaveAs "{target}", 52\nEnd Function\n')
                result = excel.run_vba(code, "Report", timeout=120.0)
                assert result.ok, f"{boxed} {name}: {result.message}"
                if boxed and name == "move_out":
                    copyfile(target, ROOT / "tests/fixtures/shapes/radios_moved.xlsm")
                with ZipFile(target) as package:
                    parts = {part: package.read(part).decode() for part in package.namelist()
                             if part.startswith("xl/ctrlProps/") or part.endswith(".vml")}
                records.append({"boxed": boxed, "name": name, "body": body,
                                "reported": str(result.value), "parts": parts})
                print(boxed, name, result.value)
    (ROOT / "tests/fixtures/shapes/radio_controls.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
