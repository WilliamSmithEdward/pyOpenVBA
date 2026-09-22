"""Measure native FormulaR1C1 array assignment, including mismatched sizes."""
import argparse
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    "matrix": 'Dim a(1 To 2, 1 To 2) As Variant\na(1, 1) = "=RC[-1]"\na(1, 2) = "=R1C1"\na(2, 1) = 12\na(2, 2) = "text"',
    "lower_bounds": 'Dim a(-2 To -1, 4 To 5) As Variant\na(-2, 4) = "=RC[-1]"\na(-2, 5) = True\na(-1, 4) = Empty\na(-1, 5) = "12"',
    "short_matrix": 'Dim a(1 To 1, 1 To 1) As Variant\na(1, 1) = "=RC[-1]"',
    "wide_matrix": 'Dim a(1 To 3, 1 To 3) As Variant\na(1, 1) = "=RC[-1]"\na(2, 2) = "=RC[-1]+1"',
    "flat": 'Dim a As Variant\na = Array("=RC[-1]", "=R1C1")',
    "short_flat": 'Dim a As Variant\na = Array("=RC[-1]")',
    "empty_flat": 'Dim a As Variant\na = Array()',
    "three_dimensions": 'Dim a(1 To 2, 1 To 2, 1 To 2) As Variant\na(1, 1, 1) = "=RC[-1]"\na(2, 2, 1) = 9\na(1, 1, 2) = 99',
    "errors": 'Dim a As Variant\na = Array(CVErr(2042), Null)',
    "merged": 'Dim a As Variant\na = Array("=RC[-1]", 99)\nRange("C3:D4").Merge',
    "single": 'Dim a As Variant\na = Array("=RC[-1]", 99)',
    "row_matrix": 'Dim a(1 To 1, 1 To 2) As Variant\na(1, 1) = "=RC[-1]"\na(1, 2) = "=RC[-2]"',
    "column_matrix": 'Dim a(1 To 2, 1 To 1) As Variant\na(1, 1) = "=RC[-1]"\na(2, 1) = "=R[-1]C[-1]"',
    "tiled": 'Dim a(1 To 2, 1 To 2) As Variant\na(1, 1) = "=RC[-1]"\na(1, 2) = 20\na(2, 1) = 30\na(2, 2) = "=RC[-1]"',
    "multi_area": 'Dim a As Variant\na = Array("=RC[-1]", 99)',
}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--a1", action="store_true", help="Measure Formula instead of FormulaR1C1")
    args = parser.parse_args()
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, setup in PROBES.items():
            if args.a1:
                for old, new in (("R[-1]C[-1]", "B2"), ("RC[-2]", "A3"), ("RC[-1]", "B3"), ("R1C1", "$A$1")):
                    setup = setup.replace(old, new)
            target = "C3,D4" if name == "multi_area" else "C3" if name == "single" else "C3:E5" if name == "tiled" else "C3:D4"
            body = ('Application.DisplayAlerts = False\nRange("A1:F6").UnMerge\nRange("A1:F6").Clear\n'
                    'Range("A1:B6").Value = 7\nRange("C3:D4").Value = "old"\n' + setup + '\n'
                    'On Error Resume Next\nRange("' + target + '").FormulaR1C1 = a\nn = Err.Number\n'
                    'On Error GoTo 0\nReport = CStr(n)\nFor Each cell In Range("' + ("C3:E5" if name == "tiled" else "C3:D4") + '")\n'
                    'Report = Report & "|" & CStr(cell.FormulaR1C1) & ":" & TypeName(cell.Value) & ":" & CStr(cell.Value)\nNext cell\n')
            if args.a1:
                body = body.replace("FormulaR1C1", "Formula")
            code = 'Public Function Report() As String\nDim cell As Object, n As Long\n' + body + 'End Function'
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, result.message
            records.append(dict(name=name, body=body, reported=str(result.value)))
            print(name, result.value, flush=True)
    filename = "formula_a1_arrays.json" if args.a1 else "formula_arrays.json"
    (Path(__file__).resolve().parents[1] / "tests/fixtures" / filename).write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
