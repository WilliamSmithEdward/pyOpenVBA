"""Measure checkbox assignment and linked-cell behavior in an isolated Excel instance."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "set_off": "sh.ControlFormat.Value = -4146",
    "set_on": "sh.ControlFormat.Value = 1",
    "set_mixed": "sh.ControlFormat.Value = 2",
    "set_zero": "sh.ControlFormat.Value = 0",
    "set_three": "sh.ControlFormat.Value = 3",
    "set_minus_one": "sh.ControlFormat.Value = -1",
    "set_true": "sh.ControlFormat.Value = True",
    "set_false": "sh.ControlFormat.Value = False",
    "set_fraction": "sh.ControlFormat.Value = 1.5",
    "set_string": 'sh.ControlFormat.Value = "1"',
    "on_then_off": "sh.ControlFormat.Value = 1\nsh.ControlFormat.Value = -4146",
    "on_then_text": 'sh.ControlFormat.Value = 1\nRange("H1").Value = "hello"',
    "on_then_clear": 'sh.ControlFormat.Value = 1\nRange("H1").ClearContents',
    "cell_string_true": 'Range("H1").Value = "TRUE"',
    "cell_string_number": 'Range("H1").Value = "2"',
    "cell_divzero": 'Range("H1").Value = CVErr(2007)',
    "formula_true": 'Range("H1").Formula = "=1=1"',
    "formula_dependency": 'Range("H1").Formula = "=G1>0"\nRange("G1").Value = 1',
    "cell_true": 'Range("H1").Value = True',
    "cell_false": 'Range("H1").Value = False',
    "cell_zero": 'Range("H1").Value = 0',
    "cell_one": 'Range("H1").Value = 1',
    "cell_two": 'Range("H1").Value = 2',
    "cell_text": 'Range("H1").Value = "hello"',
    "cell_clear": 'Range("H1").ClearContents',
    "cell_na": 'Range("H1").Value = CVErr(2042)',
    "link_true": 'sh.ControlFormat.LinkedCell = ""\nRange("H1").Value = True\nsh.ControlFormat.LinkedCell = "$H$1"',
    "link_on_empty": 'sh.ControlFormat.LinkedCell = ""\nsh.ControlFormat.Value = 1\nsh.ControlFormat.LinkedCell = "$H$1"',
    "link_off_number": 'sh.ControlFormat.LinkedCell = ""\nRange("H1").Value = 0\nsh.ControlFormat.LinkedCell = "$H$1"',
    "link_on_number": 'sh.ControlFormat.LinkedCell = ""\nsh.ControlFormat.Value = 1\nRange("H1").Value = 2\nsh.ControlFormat.LinkedCell = "$H$1"',
}


def main() -> None:
    records = []
    with ExcelSession() as excel:
        excel.new_document()
        for name, statement in PROBES.items():
            code = (
                'Public Function Report() As String\nDim sh As Object, n As Long\n'
                'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                'Range("H1").ClearContents\n'
                'Set sh = ActiveSheet.Shapes.AddFormControl(1, 0, 0, 90, 20)\n'
                'sh.ControlFormat.LinkedCell = "$H$1"\nOn Error Resume Next\n'
                + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                'Report = CStr(n) & "|" & CStr(sh.ControlFormat.Value) & "|" & _\n'
                'TypeName(Range("H1").Value) & "|" & CStr(Range("H1").Value)\nEnd Function\n'
            )
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, f"{result.outcome}: {result.message}"
            records.append({"name": name, "statement": statement, "reported": str(result.value)})
    target = ROOT / "tests/fixtures/shapes/checkbox_links.json"
    target.write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")
    print(target.read_text(encoding="utf-8"))


if __name__ == "__main__":
    main()
