"""The names a LET or a LAMBDA binds, in Excel: one that looks like a cell, R or C, and LAMBDA's [optional] ones.

Excel takes LET(x1,5,x1) with x1 a name, not the cell, and keeps it as
written; LAMBDA(x,[y],x)(1) leaves out y, and LAMBDA(x,x)() is #VALUE!.
Each formula is written through Range.Formula to its own cell of column
C, with 1 to 4 in A1:A4, and its Formula, Formula2 and Text read; the
workbook is saved, each formula's <f> element kept, and it is opened
again and each cell's Formula and Text read.

    python scripts/measure_bound_names.py

writes tests/fixtures/bound_names.json, which
tests/test_excel_bound_names.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "bound_names.json"

FORMULAS = [
    "=LET(x1,5,x1)", "=LET(A1,5,A1)", "=LET(A1,5,A1+1)", "=LET(a1,5,A1)", "=LET(x1,5,y1,x1+1,y1)",
    "=LET(x1,5,SUM(x1,A1))", "=LET(R1C1,5,R1C1)", "=LET(r,1,r)", "=LET(C,2,C)", "=LET(RC,3,RC)",
    "=LET(x,5,LET(x1,x+1,x1))", "=LAMBDA(A1,A1)(5)", "=LAMBDA(x1,x1*2)(5)", "=MAP(A1:A2,LAMBDA(a1,a1*2))",
    "=LAMBDA(x,[y],x)(1)", "=LAMBDA([x],1)()", "=LAMBDA(x,x)()", "=LAMBDA(x,[y],IF(ISOMITTED(y),x,y))(1)",
    "=LAMBDA(x,[y],IF(ISOMITTED(y),x,y))(1,2)", "=LAMBDA(x,[y],ISOMITTED(x))(1)", "=LAMBDA(x,y,x)(1,)",
    "=LAMBDA(x,[y],y)(1)", "=LAMBDA([x],[y],x+y)(,)", "=LET(f,LAMBDA(a,[b],a),f(3))",
]
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>(.*?)</c>)', re.DOTALL)
_FORMULA = re.compile(r"<f\b[^>]*>.*?</f>|<f\b[^>]*/>", re.DOTALL)


def _module(path: Path) -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String, cell As Object",
             "Dim row As Long", "Application.DisplayAlerts = False",
             "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
             'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))', "On Error Resume Next"]
    for row, formula in enumerate(FORMULAS, start=1):
        quoted = formula.replace('"', '""')
        lines += ["Err.Clear", f'ws.Range("C{row}").Formula = "{quoted}"',
                  f'If Err.Number <> 0 Then out = out & "!" & Err.Number Else out = out & ws.Range("C{row}").Formula'
                  f' & "~" & ws.Range("C{row}").Formula2 & "~" & ws.Range("C{row}").Text',
                  'out = out & "|"']
    lines += ["On Error GoTo 0", f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False",
              f'Set wb = Workbooks.Open("{path}")', "Set ws = wb.Worksheets(1)",
              f"For row = 1 To {len(FORMULAS)}",
              '    Set cell = ws.Range("C" & row)',
              '    out = out & cell.Formula & "~" & cell.Text & "|"',
              "Next", "wb.Close False", "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "bound.xlsx"
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(_module(path), "Probe", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        with zipfile.ZipFile(path) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    saved: dict[str, str] = {}
    for address, body in _CELL.findall(xml):
        element = _FORMULA.search(body or "")
        if element is not None:
            saved[address] = element.group()
    answers = str(result.value).split("|")[:-1]
    written, opened = answers[:len(FORMULAS)], answers[len(FORMULAS):]
    record = [{"cell": f"C{row}", "written": formula, "read": read, "saved": saved.get(f"C{row}", ""),
               "opened": again}
              for row, (formula, read, again) in enumerate(zip(FORMULAS, written, opened, strict=True), start=1)]
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
