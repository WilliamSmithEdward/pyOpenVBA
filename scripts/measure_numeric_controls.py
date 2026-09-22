"""Measure spinner/scroll-bar values and links in live Excel."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "defaults": "",
    "min_negative": "sh.ControlFormat.Min = -1",
    "min_above_max": "sh.ControlFormat.Min = 30001",
    "max_negative": "sh.ControlFormat.Max = -1",
    "max_above_limit": "sh.ControlFormat.Max = 30001",
    "max_below_min": "sh.ControlFormat.Min = 10\nsh.ControlFormat.Max = 5",
    "min_above_max_valid": "sh.ControlFormat.Max = 5\nsh.ControlFormat.Min = 10",
    "small_zero": "sh.ControlFormat.SmallChange = 0",
    "small_negative": "sh.ControlFormat.SmallChange = -1",
    "small_large": "sh.ControlFormat.SmallChange = 30001",
    "large_zero": "sh.ControlFormat.LargeChange = 0",
    "large_negative": "sh.ControlFormat.LargeChange = -1",
    "large_large": "sh.ControlFormat.LargeChange = 30001",
    "min_cell_error": 'sh.ControlFormat.Min = 10\nRange("H1").Value = CVErr(2042)',
    "value": "sh.ControlFormat.Value = 7",
    "negative": "sh.ControlFormat.Value = -1",
    "too_large": "sh.ControlFormat.Value = 30001",
    "cell_value": 'Range("H1").Value = 7',
    "cell_negative": 'Range("H1").Value = -1',
    "cell_large": 'Range("H1").Value = 30001',
    "cell_text": 'sh.ControlFormat.Value = 7\nRange("H1").Value = "text"',
    "cell_error": 'sh.ControlFormat.Value = 7\nRange("H1").Value = CVErr(2042)',
    "cell_clear": 'sh.ControlFormat.Value = 7\nRange("H1").ClearContents',
    "minimum": 'sh.ControlFormat.Min = 10',
    "maximum": 'sh.ControlFormat.Value = 7\nsh.ControlFormat.Max = 5',
    "rebind": 'sh.ControlFormat.LinkedCell = ""\nRange("H1").Value = 7\nsh.ControlFormat.LinkedCell = "$H$1"',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for kind in (8, 9):
            for name, statement in PROBES.items():
                body = (
                    'Range("H1").ClearContents\n'
                    f'Set sh = ActiveSheet.Shapes.AddFormControl({kind}, 0, 0, 90, 30)\n'
                    'sh.ControlFormat.LinkedCell = "$H$1"\nOn Error Resume Next\n'
                    + statement + '\nn = Err.Number\nErr.Clear\np = CStr(sh.ControlFormat.LargeChange)\n'
                    'If Err.Number <> 0 Then p = "error" & CStr(Err.Number)\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|" & CStr(sh.ControlFormat.Value) & "|" & _\n'
                    'CStr(sh.ControlFormat.Min) & "|" & CStr(sh.ControlFormat.Max) & "|" & _\n'
                    'CStr(sh.ControlFormat.SmallChange) & "|" & p & "|" & _\n'
                    'CStr(Range("H1").Value)\n'
                )
                code = ('Public Function Report() As String\nDim sh As Object, n As Long, p As String\n'
                        'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                        + body + 'End Function\n')
                result = excel.run_vba(code, "Report", timeout=120.0)
                assert result.ok, f"{kind} {name}: {result.outcome}: {result.message}"
                records.append({"kind": kind, "name": name, "body": body, "reported": str(result.value)})
                print(kind, name, result.value)
    (ROOT / "tests/fixtures/shapes/numeric_controls.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
