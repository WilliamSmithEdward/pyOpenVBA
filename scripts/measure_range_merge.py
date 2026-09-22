"""Native merged-range semantics with alerts disabled, as in the headless host."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "clear_anchor": 'Range("A1:B2").Merge\nRange("A1").ClearContents',
    "clear_interior_all": 'Range("A1:B2").Merge\nRange("B2").Clear',
    "contained_flag": 'Range("A1:C3").Merge\nanswer = CStr(Range("B1:B2").MergeCells)',
    "across_flag": 'Range("A1:C3").Merge Across:=True\nIf IsNull(Range("A1:C3").MergeCells) Then answer = "Null" Else answer = CStr(Range("A1:C3").MergeCells)',
    "formula_moved": 'Range("A1").ClearContents\nRange("B1").Formula = "=C2"\nRange("A1:B2").Merge\nanswer = Range("A1").Formula',
    "rectangle": 'Range("A1:B2").Merge',
    "across": 'Range("A1:C2").Merge Across:=True',
    "single": 'Range("A1").Merge',
    "empty_anchor": 'Range("A1").ClearContents\nRange("A1:B2").Merge',
    "repeat": 'Range("A1:B2").Merge\nRange("A1:B2").Merge',
    "unmerge_corner": 'Range("A1:B2").Merge\nRange("B2").UnMerge',
    "unmerge_all": 'Range("A1:C2").Merge Across:=True\nRange("A1:C3").UnMerge',
    "set_true": 'Range("A1:B2").MergeCells = True',
    "set_false": 'Range("A1:B2").Merge\nRange("A1:B2").MergeCells = False',
    "write_interior": 'Range("A1:B2").Merge\nRange("B2").Value = "inside"',
    "clear_interior": 'Range("A1:B2").Merge\nRange("B2").ClearContents',
    "expand": 'Range("A1:B2").Merge\nRange("A1:C3").Merge',
    "partial_overlap": 'Range("A1:B2").Merge\nRange("B2:C3").Merge',
    "merge_area_multiple": 'Range("A1:B2").Merge\nanswer = Range("A1:B2").MergeArea.Address',
    "shrink": 'Range("A1:C3").Merge\nRange("A1:B2").Merge',
    "clear_all": 'Range("A1:B2").Merge\nRange("A1:B2").Clear',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, statement in PROBES.items():
            body = ('Application.DisplayAlerts = False\nRange("A1:C3").UnMerge\nRange("A1:C3").Clear\n'
                    'Range("A1").Value = "a"\nRange("B1").Value = "b"\nRange("C1").Value = "c"\n'
                    'Range("A2").Value = "d"\nRange("B2").Value = "e"\nRange("C2").Value = "f"\n'
                    'On Error Resume Next\n' + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|"\n'
                    'For Each cell In Range("A1:C3")\n'
                    'Report = Report & CStr(cell.Value) & ":" & cell.MergeArea.Address & ";"\nNext cell\n'
                    'If IsNull(Range("A1:C3").MergeCells) Then\nReport = Report & "Null"\nElse\n'
                    'Report = Report & CStr(Range("A1:C3").MergeCells)\nEnd If\nReport = Report & "|" & answer\n')
            code = 'Public Function Report() As String\nDim cell As Object, n As Long, answer As String\n' + body + 'End Function\n'
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, result.message
            records.append({"name": name, "body": body, "reported": str(result.value)})
            print(name, result.value, flush=True)
    (ROOT / "tests/fixtures/range_merge.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
