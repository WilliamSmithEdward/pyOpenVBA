"""Measure native Range.Copy destination geometry and overlapping copies."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    "overlap_down": 'Range("A1:B2").Copy Range("A2")',
    "overlap_right": 'Range("A1:B2").Copy Range("B1")',
    "overlap_up": 'Range("A2:B3").Copy Range("A1")',
    "self": 'Range("A1:B2").Copy Range("A1")',
    "tile": 'Range("A1:B2").Copy Range("C3:F6")',
    "incompatible": 'Range("A1:B2").Copy Range("C3:E5")',
    "partial_rows": 'Range("A1:B2").Copy Range("C3:F5")',
    "partial_columns": 'Range("A1:B2").Copy Range("C3:E6")',
    "short_destination": 'Range("A1:B2").Copy Range("D4:D5")',
    "blank_tail": 'Range("D4:E5").Value = "old"\nRange("A1:B3").Copy Range("D3")',
    "all_blank": 'Range("E5:F6").Value = "old"\nRange("C1:D2").Copy Range("E5")',
    "single_repeat": 'Range("B1").Copy Range("D3:E4")',
    "formula": 'Range("B2").Formula = "=A1+$A$1"\nRange("A1:B2").Copy Range("D3")',
    "format": 'Range("B2").Font.Bold = True\nRange("B2").NumberFormat = "0.00"\nRange("A1:B2").Copy Range("D3")',
    "edge": 'Range("A1:B2").Copy Range("XFD1048576")',
    "tiled_formula": 'Range("B2").Formula = "=A1+$A$1"\nRange("A1:B2").Copy Range("C3:F6")',
    "cross_sheet": 'Worksheets.Add.Name = "CopyTarget"\nWorksheets("Sheet1").Range("A1:B2").Copy Range("D3")',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, statement in PROBES.items():
            body = ('Range("A1:F6").Clear\nRange("A1").Value = "a"\nRange("B1").Value = "b"\n'
                    'Range("A2").Value = "c"\nRange("B2").Value = "d"\nOn Error Resume Next\n'
                    + statement + '\nn = Err.Number\nOn Error GoTo 0\nReport = CStr(n)\n'
                    'For Each cell In Range("A1:F6")\nReport = Report & "|" & cell.Formula & ":" & CStr(cell.Font.Bold) & ":" & cell.NumberFormat\nNext cell\n')
            code = 'Function Report() As String\nDim cell As Object, n As Long\n' + body + 'End Function'
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, result.message
            records.append(dict(name=name, body=body, reported=str(result.value)))
            print(name, str(result.value)[:90], flush=True)
    (Path(__file__).resolve().parents[1] / "tests/fixtures/range_copy.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
