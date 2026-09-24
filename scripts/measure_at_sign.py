"""What Excel makes of an @ written in a formula, and of SINGLE, which it writes as @.

Formula2 writes =@A1:A3 as a legacy formula that cuts A1:A3 to one value;
here each of a list of formulas with an @ or a SINGLE is written through
Formula, Formula2, FormulaR1C1 and Value to C8 of a sheet called Data
with 1 to 3 in A1:A3 and =SEQUENCE(3) spilling from E1, and its Formula,
Formula2, FormulaR1C1 and Text read back, or ! and the error number Excel
raised writing it. Then a workbook of such formulas, each written through
Formula in column C and through Formula2 in column G, is saved, and what
each cell reads and what the file keeps for it are recorded.

    python scripts/measure_at_sign.py

writes tests/fixtures/at_sign.json and at_sign.xlsx, which
tests/test_excel_at_sign.py replays.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "at_sign.json"
WORKBOOK = ROOT / "tests" / "fixtures" / "at_sign.xlsx"

SETUP = ('ws.Name = "Data"\nws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n'
         'ws.Range("E1").Formula2 = "=SEQUENCE(3)"')
FORMULAS = [
    "=@A1:A3", "=@A1:A3*2", "=SUM(@A1:A3)", "=@SUM(A1:A3)", "=@A1", "=@A1:A3+@A1:A3", "=@(A1:A3)",
    "=@INDEX(A1:A3,0)", "=@SEQUENCE(3)", "=@@A1:A3", "=@1", '=@"a"', "=@{1,2,3}", "=-@A1:A3", "=@-A1:A3",
    "=SINGLE(A1:A3)", "=SINGLE(A1)", "=SINGLE(A1:A3)*2", "=SUM(SINGLE(A1:A3))", "=single(A1:A3)",
    "=SINGLE(SEQUENCE(3))", "=SINGLE(A1:A3,1)", "=SINGLE()", "=LET(x,A1:A3,@x)", "=@Data!A1:A3",
    "=@ A1:A3", "=IF(@A1:A3>1,1,0)",
    # An @ beside a place a legacy formula cuts on its own.
    "=@A1+A1:A3", "=A1:A3+@A1", "=SUM(@A1:A3)+A1:A3", "=@A1*SUM(A1:A3*2)", "=@(A1:A3*2)", "=@(A1:A3)*A1:A3",
    "=IF(@A1,A1:A3,0)", "=@SUM(A1:A3)+A1:A3", "=@1+A1:A3", "=(A1:A3)",
    # How tightly an @ holds: against a space, a percent sign, a power and a sign.
    "=@A1:A3 A2:B2", "=@A1:A3%", "=@-A1:A3*2", "=@A1:A3^2", "=@INDEX(A1:A3,2)", "=@A1:A1", "=@A:A",
    "=@Data!A1", "=@E1#", "=SUM(@E1#)", "=@E1#+1",
    # How SINGLE is spelled as @.
    "=SINGLE(A1+1)", "=SINGLE(-A1:A3)", "=SINGLE(A1:A3%)", "=SINGLE(A1:A3 A2:B2)", "=SINGLE((A1:A3))",
    "=SINGLE(A1)*SINGLE(A2)", "=SINGLE( A1)", "=SINGLE(A1:A3)^2", "=SINGLE(SINGLE(A1:A3))", "=SINGLE(E1#)",
    "=SINGLE(A1,)", "=SINGLE(A1:A3)%", "=-SINGLE(A1:A3)", "=SINGLE(A1:A3):A5",
    # A name the workbook does not define, which Formula2 cannot know gives one value.
    "=nosuch", "=nosuch+1", "=SUM(nosuch)", "=@nosuch", "=ABS(nosuch)",
]
MEMBERS = ("Formula", "Formula2", "FormulaR1C1", "Value")
R1C1 = ["=@R1C1:R3C1", "=SINGLE(R1C1:R3C1)", "=@R[-7]C[-2]:R[-5]C[-2]"]
#: How many cases one procedure holds; many more and Excel shows it as blocked.
BATCH = 80

#: The formulas of the saved workbook, each written through Formula in column C and Formula2 in column G, from row
#: 10 down four rows apart, so what spills has room.
FILE_FORMULAS = [
    "=@A1", "=@A1:A3", "=@SUM(A1:A3)", "=SUM(@A1:A3)", "=@(A1:A3)", "=@@A1:A3", "=@1", "=@-A1:A3", "=@A1+A1:A3",
    "=@(A1:A3*2)", "=@E1#", "=SUM(@E1#)", "=LET(x,A1:A3,@x)", "=@INDEX(A1:A3,0)", "=@A1:A3%",
    "=SEQUENCE(3)+@A1:A3", "=SINGLE(A1+1)", "=IF(@A1,A1:A3,0)", "=nosuch+1", "=SINGLE(A1:A3%)",
    "=SINGLE(A1:A3):A5",
]


def _module(cases: list[tuple[str, str]]) -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String",
             "Set ws = ActiveWorkbook.Worksheets.Add", SETUP, "On Error Resume Next"]
    for member, formula in cases:
        quoted = formula.replace('"', '""')
        lines += ["Err.Clear", 'ws.Range("C8").ClearContents', f'ws.Range("C8").{member} = "{quoted}"',
                  'If Err.Number <> 0 Then out = out & "!" & Err.Number Else out = out & ws.Range("C8").Formula & '
                  '"~" & ws.Range("C8").Formula2 & "~" & ws.Range("C8").FormulaR1C1 & "~" & ws.Range("C8").Text',
                  'out = out & "|"']
    lines += ["On Error GoTo 0", "Application.DisplayAlerts = False", "ws.Delete", "Application.DisplayAlerts = True",
              "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def file_writes() -> list[tuple[str, str, str]]:
    """Each cell of the saved workbook with the member that wrote it and its formula."""
    return [(f"{column}{10 + 4 * index}", member, formula) for index, formula in enumerate(FILE_FORMULAS)
            for column, member in (("C", "Formula"), ("G", "Formula2"))]


def _saved(excel: ExcelSession, folder: Path) -> tuple[dict[str, str], dict[str, str]]:
    """What each cell of the saved workbook reads, Formula~Formula2~Text, and the parts of the file."""
    path = folder / "at_sign.xlsx"
    writes = file_writes()
    lines = ["Public Function Report() As String", "Dim wb As Object, ws As Object, out As String, one As Variant",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", SETUP, "On Error Resume Next"]
    for cell, member, formula in writes:
        quoted = formula.replace('"', '""')
        lines += ["Err.Clear", f'ws.Range("{cell}").{member} = "{quoted}"',
                  f'If Err.Number <> 0 Then out = out & "{cell}!" & Err.Number & "|"']
    lines += ["On Error GoTo 0", "For Each one In Array(" + ", ".join(f'"{cell}"' for cell, _, _ in writes) + ")",
              '    out = out & one & "~" & ws.Range(one).Formula & "~" & ws.Range(one).Formula2 & "~" & '
              'ws.Range(one).Text & "|"', "Next",
              f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False", "Report = out", "End Function"]
    result = excel.run_vba("\n".join(lines) + "\n", "Report", timeout=120.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    reads: dict[str, str] = {}
    for answer in str(result.value).split("|")[:-1]:
        if "~" not in answer:
            # A write Excel refused, C10!1004, read afterwards as the empty cell it left.
            cell, _, number = answer.partition("!")
            reads[f"{cell}!"] = number
            continue
        cell, _, read = answer.partition("~")
        reads[cell] = read
    # A workbook of its own, not the probe's, so the file carries no macro; the model opens it.
    WORKBOOK.write_bytes(path.read_bytes())
    parts: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name in ("xl/worksheets/sheet1.xml", "xl/metadata.xml", "[Content_Types].xml"):
                parts[name] = package.read(name).decode("utf-8")
    return reads, parts


def main() -> None:
    cases = [(member, formula) for formula in FORMULAS for member in MEMBERS]
    cases += [(member, formula) for formula in R1C1 for member in ("FormulaR1C1", "Formula2R1C1")]
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for start in range(0, len(cases), BATCH):
            batch = cases[start:start + BATCH]
            result = excel.run_vba(_module(batch), "Probe", timeout=300.0)
            assert result.ok, f"{result.outcome}: {result.message}"
            found = str(result.value).split("|")[:-1]
            assert len(found) == len(batch), (len(found), len(batch))
            answers += found
        with tempfile.TemporaryDirectory() as folder:
            reads, parts = _saved(excel, Path(folder))
    record = [{"member": member, "formula": formula, "answer": answer}
              for (member, formula), answer in zip(cases, answers, strict=True)]
    saved = {"writes": [list(write) for write in file_writes()], "reads": reads, "parts": parts}
    OUT.write_text(json.dumps({"setup": SETUP, "cases": record, "file": saved}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
