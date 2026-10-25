"""Range.Precedents, DirectPrecedents, Dependents and DirectDependents in live Excel.

Each layout fills a new workbook's first sheet, and a second sheet named Other, with constants and formulas; then each
listed range is asked for each of the four, the answer's Address and its count of areas noted, or the error's number
and description where Excel raises one.

    python scripts/measure_precedents.py

writes tests/fixtures/precedents.json, which tests/test_excel_precedents.py replays.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "precedents.json"
MEMBERS = ("Precedents", "DirectPrecedents", "Dependents", "DirectDependents")

#: name -> (cells as (address, formula or value) on the first sheet, cells on Other, names, the ranges asked).
LAYOUTS: dict[str, tuple[list[tuple[str, str]], list[tuple[str, str]], list[tuple[str, str]], list[str]]] = {
    "chain": ([("A1", "1"), ("A2", "2"), ("A3", "=A1+A2"), ("A4", "=A3*2"), ("B1", "=SUM(A1:A3)"),
               ("B2", "=A4+B1"), ("C2", "5")], [], [],
              ["A1", "A2", "A3", "A4", "B1", "B2", "C2", "H9", "A3:A4", "A1:A2", "A1:B2"]),
    "sheets": ([("A1", "1"), ("C1", "=Other!A1+A1"), ("C2", "=Other!A1")], [("A1", "3"), ("B1", "=Sheet1!A1*2")], [],
               ["A1", "C1", "C2"]),
    "functions": ([("A1", "1"), ("A2", "2"), ("D1", '=INDIRECT("A1")'), ("D2", "=OFFSET(A1,1,0)"),
                   ("D3", "=INDEX(A1:A2,2)"), ("D4", "=IF(A1>0,A2,B9)"), ("D5", '=COUNTIF(A1:A2,">0")'),
                   ("D6", "=A1:A2"), ("D7", "=SUM(A:A)"), ("D8", "=SUM(1:1)"), ("D9", "=ROW()")], [], [],
                  ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "A1", "A2", "B9"]),
    "names": ([("A1", "1"), ("A2", "2"), ("E1", "=Pair"), ("E2", "=SUM(Pair)+One"), ("E3", "=Spread")], [],
              [("Pair", "=Sheet1!$A$1:$A$2"), ("One", "=Sheet1!$A$1"), ("Spread", "=Sheet1!$A$1*2")],
              ["E1", "E2", "E3", "A1", "A2"]),
    "repeats": ([("A1", "1"), ("A2", "2"), ("F1", "=A1+A1"), ("F2", "=$A$1+A$1"), ("F3", "=A1+A2+A1"),
                 ("F4", "=SUM(A1:A2,A1)"), ("F5", "=A2+A1")], [], [], ["F1", "F2", "F3", "F4", "F5", "A1"]),
    "order": ([("C3", "1"), ("A1", "2"), ("B2", "3"), ("A5", "4"), ("D1", "=C3+A1+B2+A5"), ("D2", "=A5+B2+A1+C3"),
               ("D3", "=SUM(C3,A1)+SUM(B2:B3)")], [], [], ["D1", "D2", "D3", "C3", "A1"]),
    "loops": ([("A1", "1"), ("G1", "=G2+A1"), ("G2", "=G1+1")], [], [], ["G1", "G2"]),
    "wide": ([("A1", "1"), ("A2", "2"), ("A3", "3"), ("B1", "=A1"), ("B2", "=A2"), ("B3", "=A3"), ("C1", "=B1+B2+B3"),
              ("C2", "=C1")], [], [], ["C2", "C1", "A1", "A2:A3", "B1:B3"]),
    "intersections": ([("A1", "1"), ("A2", "2"), ("X1", "=A1:A2"), ("X2", "=Pair"), ("X6", "=Pair"),
                       ("X7", "=A1:A2+1"), ("Y1", "=SUM(A1:A2)"), ("Z1", "@2=A1:A2"), ("Z4", "@2=Pair"),
                       ("W1", "=@A1:A2"), ("W8", "=@Pair")], [], [("Pair", "=Sheet1!$A$1:$A$2")],
                      ["X1", "X2", "X6", "X7", "Y1", "Z1", "Z4", "W1", "W8", "A1", "A2"]),
    "entry_order": ([("A1", "1"), ("C3", "=A1"), ("B2", "=A1"), ("D1", "=A1"), ("A5", "=A1"), ("C3", "=A1*2")], [], [],
                    ["A1"]),
}
#: A layout saved and opened again, then asked again: the entry order a file gives its formulas.
REOPENED = ("chain", ["A1", "A3", "A1:A2", "A3:A4", "A1:B2"])

HELPER = '''Private Function Asked(target As Range, member As String) As String
    Dim found As Range
    On Error Resume Next
    Select Case member
        Case "Precedents": Set found = target.Precedents
        Case "DirectPrecedents": Set found = target.DirectPrecedents
        Case "Dependents": Set found = target.Dependents
        Case "DirectDependents": Set found = target.DirectDependents
    End Select
    If Err.Number <> 0 Then
        Asked = "E" & Err.Number & ":" & Err.Description
    Else
        Asked = found.Parent.Name & "!" & found.Address & "#" & found.Areas.Count
    End If
    Err.Clear
End Function
'''


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def layout_code(index: int, name: str, cells: list[tuple[str, str]], other: list[tuple[str, str]],
                names: list[tuple[str, str]], asked: list[str]) -> str:
    lines = [f"Private Function Layout{index}() As String", "    Dim wb As Workbook, ws As Worksheet, out As String",
             "    Set wb = Workbooks.Add(xlWBATWorksheet)", "    Set ws = wb.Worksheets(1)",
             '    wb.Worksheets.Add(After:=ws).Name = "Other"', "    ws.Activate"]
    for name_, refers in names:
        lines.append(f"    wb.Names.Add {vba_text(name_)}, {vba_text(refers)}")
    for address, content in other:
        lines.append(f'    wb.Worksheets("Other").Range("{address}").Formula = {vba_text(content)}')
    for address, content in cells:
        # A formula marked @2 is written through Formula2, a dynamic-array formula.
        member, written = ("Formula2", content[2:]) if content.startswith("@2") else ("Formula", content)
        lines.append(f'    ws.Range("{address}").{member} = {vba_text(written)}')
    lines.append(f'    out = "{name}"')
    lines += asking(asked)
    lines += ["    wb.Close False", f'    Layout{index} = out & "~"', "End Function"]
    return "\n".join(lines) + "\n"


def asking(asked: list[str]) -> list[str]:
    return [f'    out = out & "^{address} {member}=" & Asked(ws.Range("{address}"), "{member}")'
            for address in asked for member in MEMBERS]


def reopened_code() -> str:
    """The layout REOPENED names, made, saved, closed, opened and asked again."""
    name, asked = REOPENED
    cells, _other, _names, _asked = LAYOUTS[name]
    lines = ["Public Function Reopened(folder As String) As String", "    Dim wb As Workbook, ws As Worksheet, out As String",
             "    Application.DisplayAlerts = False", "    Set wb = Workbooks.Add(xlWBATWorksheet)",
             "    Set ws = wb.Worksheets(1)"]
    lines += [f'    ws.Range("{address}").Formula = {vba_text(content)}' for address, content in cells]
    lines += ['    wb.SaveAs folder & "\\reopened.xlsx", 51', "    wb.Close False",
              '    Set wb = Workbooks.Open(folder & "\\reopened.xlsx")', "    Set ws = wb.Worksheets(1)",
              f'    out = "reopened_{name}"', *asking(asked), "    wb.Close False", '    Reopened = out & "~"',
              "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    code = HELPER + "".join(layout_code(index, name, *parts) for index, (name, parts) in enumerate(LAYOUTS.items()))
    calls = "".join(f"    out = out & Layout{index}()\n" for index in range(len(LAYOUTS)))
    code += ("Public Function Probe() As String\n    Dim out As String\n    Application.DisplayAlerts = False\n"
             + calls + "    Probe = out\nEnd Function\n")
    reopened = HELPER + reopened_code()
    with tempfile.TemporaryDirectory(prefix="pyopenvba-precedents-") as name:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(code, "Probe", timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            again = excel.run_vba(reopened, "Reopened", [str(Path(name).resolve())], timeout=300.0)
            assert again.ok, f"{again.outcome}: {again.message} {again.error}"
    cases = [entry.split("^") for entry in (str(result.value) + str(again.value)).split("~") if entry]
    record = {"probe": code, "reopened": reopened, "layouts": {key: list(parts) for key, parts in LAYOUTS.items()},
              "cases": cases}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in cases:
        for item in case:
            print(item)


if __name__ == "__main__":
    main()
