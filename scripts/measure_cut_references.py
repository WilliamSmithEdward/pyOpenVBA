"""What Range.Cut does to references that only overlap the cells it moves.

Every layout gets a new workbook: numbers in B2:C7 of a sheet named Src,
formulas in F1:F9 reading blocks, cells and columns of them, a formula
on a second sheet, Dst, and a name; then one cut, on Src or to Dst. The
dump lists Src's formulas in F1:F9, Dst's in H1, the name, and every
cell of Src!A1:E12 and Dst!A1:E8 as it stands.

    python scripts/measure_cut_references.py

writes tests/fixtures/cut_references.json, which tests/test_excel_cut_references.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "cut_references.json"

HELPER = '''
Private Function Dump(src As Object, dst As Object) As String
    Dim out As String, c As Object
    For Each c In src.Range("F1:F9").Cells
        out = out & c.Formula & ";"
    Next
    out = out & " dst=" & dst.Range("H1").Formula & " name=" & src.Parent.Names("Tracked").RefersTo & " src="
    For Each c In src.Range("A1:E12").Cells
        out = out & c.Formula & ","
    Next
    out = out & " dstcells="
    For Each c In dst.Range("A1:E8").Cells
        out = out & c.Formula & ","
    Next
    Dump = out
End Function
'''

SETUP = '''For r = 2 To 7
    src.Cells(r, 2).Value = r - 1
    src.Cells(r, 3).Value = (r - 1) * 10
Next
src.Range("F1").Formula = "=SUM(B2:B7)"
src.Range("F2").Formula = "=SUM(B2:C6)"
src.Range("F3").Formula = "=B7"
src.Range("F4").Formula = "=SUM($B$2:$B$7)"
src.Range("F5").Formula = "=SUM(B:B)"
src.Range("F6").Formula = "=SUM(D2:D6)"
src.Range("F7").Formula = "=SUM(B3:B5)"
src.Range("F8").Formula = "=SUM(A1:C8)"
src.Range("F9").Formula = "=SUM(C2:C7)"
dst.Range("H1").Formula = "=SUM(Src!B2:B7)"
src.Parent.Names.Add Name:="Tracked", RefersTo:="=Src!$B$2:$B$7"
'''

#: name -> the cut.
LAYOUTS: dict[str, str] = {
    "same_bottom_cell": 'src.Range("B7").Cut src.Range("B9")',
    "same_top_cell": 'src.Range("B2").Cut src.Range("B1")',
    "same_top_cell_far": 'src.Range("B2").Cut src.Range("D10")',
    "same_bottom_strip": 'src.Range("B6:B7").Cut src.Range("B10")',
    "same_top_strip_far": 'src.Range("A2:C4").Cut src.Range("A10")',
    "same_top_strip_near": 'src.Range("B2:B3").Cut src.Range("B8")',
    "same_sideways": 'src.Range("B2:B4").Cut src.Range("D2")',
    "same_whole": 'src.Range("B2:B7").Cut src.Range("D1")',
    "same_right_column": 'src.Range("C2:C6").Cut src.Range("E2")',
    "same_down_one": 'src.Range("B2:B7").Cut src.Range("B3")',
    "same_onto_ref": 'src.Range("B2:B3").Cut src.Range("D2")',
    "same_onto_whole_ref": 'src.Range("B2:B6").Cut src.Range("D2")',
    "same_block_down": 'src.Range("B2:C7").Cut src.Range("B4")',
    "other_top_cell": 'src.Range("B2").Cut dst.Range("A1")',
    "other_bottom_cell": 'src.Range("B7").Cut dst.Range("A1")',
    "other_middle_cell": 'src.Range("B4").Cut dst.Range("A1")',
    "other_top_strip": 'src.Range("B2:B4").Cut dst.Range("A1")',
    "other_bottom_strip": 'src.Range("B5:B7").Cut dst.Range("A1")',
    "other_right_column": 'src.Range("C2:C6").Cut dst.Range("A1")',
    "other_left_column": 'src.Range("B2:B6").Cut dst.Range("A1")',
    "other_wide_top": 'src.Range("A2:E3").Cut dst.Range("A1")',
    "other_corner": 'src.Range("B2:B3").Cut dst.Range("A1")',
    "other_whole": 'src.Range("B2:B7").Cut dst.Range("A1")',
    "other_more_than_whole": 'src.Range("B1:B8").Cut dst.Range("A1")',
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(src As Object, dst As Object) As String",
             "Dim failed As String, r As Long", *SETUP.splitlines(), "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f"Case{index} = failed & Dump(src, dst)", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, src As Object, dst As Object",
             "Dim out As String", "Application.DisplayAlerts = False"]
    bodies: list[str] = []
    for index, setup in enumerate(LAYOUTS.values()):
        build += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set src = wb.Worksheets(1)", 'src.Name = "Src"',
                  "Set dst = wb.Worksheets.Add(After:=src)", 'dst.Name = "Dst"', "src.Activate",
                  f'out = out & Case{index}(src, dst) & "|"', "wb.Close False"]
        bodies.append(case_code(index, setup))
    build += ["Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "setup": SETUP,
              "layouts": [{"name": name, "cut": cut, "answers": answer}
                          for (name, cut), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        print(f"{name}: {answer.split(' src=')[0]}")


if __name__ == "__main__":
    main()
