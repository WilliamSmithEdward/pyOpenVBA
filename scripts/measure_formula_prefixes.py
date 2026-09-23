"""How Excel spells a function newer than Excel 2007 in the file, and what it reads back.

A file names a function Excel 2007 did not have with a prefix, _xlfn.,
and the dynamic-array functions and a LAMBDA's parameters with more:
_xlfn._xlws.FILTER, _xlpm.x. Each formula below is written through
Range.Formula into its own cell of column C, with 1 to 4 in A1:A4; the
workbook is saved, and each formula cell Excel wrote is kept, then the
workbook is opened again and each cell's Formula, Formula2 and Text read.
A second workbook takes a call of every one of Excel's functions, one to
a row, as scripts/measure_cells_functions.py calls it, and keeps the <f>
element Excel saved for each. A third, the probe's own workbook, takes
formulas Excel works out whenever anything changes, beside a macro's
function and a named LAMBDA, and is saved with its macros apart from the
fixtures: its <f> elements and defined names are kept.

    python scripts/measure_formula_prefixes.py

writes tests/fixtures/formula_prefixes/prefixes.xlsx, functions.xlsx and
prefixes.json, which tests/test_excel_formula_prefixes.py replays.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_cells_functions import FUNCTIONS, call
from pyopenvba.formula._calc.catalog import EXCEL_FUNCTIONS

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "formula_prefixes"

#: One formula to a cell, in Range.Formula's spelling: Excel 2007's functions, and newer ones of every kind.
FORMULAS = [
    "=SUMIFS(A1:A4,A1:A4,\">1\")",
    "=IFERROR(1/0,2)",
    "=AGGREGATE(9,6,A1:A4)",
    "=NORM.DIST(1,0,1,TRUE)",
    "=PERCENTILE.INC(A1:A4,0.5)",
    "=NETWORKDAYS.INTL(1,10)",
    "=CEILING.MATH(2.5)",
    "=DAYS(10,1)",
    "=ISFORMULA(A1)",
    "=SHEET()",
    "=ENCODEURL(\"a b\")",
    "=CONCAT(A1,A2)",
    "=TEXTJOIN(\",\",TRUE,A1:A4)",
    "=IFS(A1>0,1)",
    "=SWITCH(A1,1,\"one\",\"other\")",
    "=MAXIFS(A1:A4,A1:A4,\">1\")",
    "=STDEV.S(A1:A4)",
    "=FORECAST.LINEAR(5,A1:A4,A1:A4)",
    "=XLOOKUP(2,A1:A4,A1:A4)",
    "=XMATCH(2,A1:A4)",
    "=FILTER(A1:A4,A1:A4>2)",
    "=SORT(A1:A4)",
    "=SORTBY(A1:A4,A1:A4)",
    "=UNIQUE(A1:A4)",
    "=SEQUENCE(2)",
    "=LET(x,A1,x*2)",
    "=LET(a,1,b,2,a+b)",
    "=LAMBDA(x,x*2)(A2)",
    "=MAP(A1:A2,LAMBDA(v,v+1))",
    "=BYROW(A1:A2,LAMBDA(r,SUM(r)))",
    "=REDUCE(0,A1:A4,LAMBDA(a,v,a+v))",
    "=TAKE(A1:A4,2)",
    "=CHOOSECOLS(A1:A4,1)",
    "=VSTACK(A1,A2)",
    "=HSTACK(A1,A2)",
    "=TOCOL(A1:A2)",
    "=TEXTBEFORE(\"a-b\",\"-\")",
    "=ARRAYTOTEXT(A1:A2)",
    "=TRIMRANGE(A1:A4)",
    "=REGEXTEST(\"abc\",\"b\")",
    "=PERCENTOF(A1,A1:A4)",
    "=SINGLE(A1)",
    "=NOTAFUNCTION(1)",
    "=SUM(A1:A4)+XLOOKUP(2,A1:A4,A1:A4)",
]


def module(path: Path) -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String, cell As Object",
             "Dim row As Long", "Application.DisplayAlerts = False",
             "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
             'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))']
    for row, formula in enumerate(FORMULAS, start=1):
        quoted = formula.replace('"', '""')
        lines.append(f'ws.Range("C{row}").Formula = "{quoted}"')
    lines += [f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False",
              f'Set wb = Workbooks.Open("{path}")', "Set ws = wb.Worksheets(1)",
              f"For row = 1 To {len(FORMULAS)}",
              '    Set cell = ws.Range("C" & row)',
              '    out = out & cell.Formula & "~" & cell.Formula2 & "~" & cell.Text & "|"',
              "Next", "wb.Close False", "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


#: A macro's own function, which a formula in the workbook holding it can call.
UDF = """Public Function MyUdf(x As Variant) As Variant
    MyUdf = x
End Function
"""

#: Formulas Excel works out whenever anything changes, and some it does not, written to the macro's own workbook:
#: a shared formula, array formulas of one cell and of several, a macro's function and one no one has.
VOLATILE = """ws.Range("A1:B3").Value = 1
ws.Range("C1:C3").Formula = "=NOW()+A1"
ws.Range("D1:D3").Formula = "=A1*2"
ws.Range("E1:E2").FormulaArray = "=RAND()+A1:A2"
ws.Range("F1").FormulaArray = "=SUM(OFFSET(A1,0,0,2,1))"
ws.Range("G1:G2").FormulaArray = "=OFFSET(A1,0,0,2,1)"
ws.Range("H1:H2").FormulaArray = "=RAND()+NOW()"
ws.Range("I1").Formula = "=MyUdf(1)"
ws.Range("I2").Formula = "=SUM(A1,MyUdf(2))"
ws.Range("I3").Formula = "=SUM(A1:A3,INDIRECT(""A1""))"
ws.Range("J1:J3").Formula = "=MyUdf(A1)"
ws.Range("K1").Formula = "=NOTAFUNCTION(1)"
ws.Range("K2").Formula = "=INFO(""numfile"")"
ws.Range("K3").FormulaArray = "=RANDBETWEEN(1,9)"
ws.Range("L1").Formula = "=LET(f,LAMBDA(x,x+1),f(2))"
ws.Range("L2").Formula = "=LET(x,1,LET(y,x+1,x+y))"
ThisWorkbook.Names.Add "Twice", "=LAMBDA(x,x*2)"
ws.Range("L3").Formula = "=Twice(5)"
ws.Range("M1").Formula = "=SUM(A1:A3 (A1:B1))"
ws.Range("M2").Formula = "=LET(x,{1,2},SUM(x))"
ws.Range("M3").Formula = "=MAP(A1:A2,LAMBDA(a,LET(b,a*2,b+a)))"
"""


def volatile_module(path: Path) -> str:
    """The volatile formulas written to the probe's own workbook, beside MyUdf, and the workbook saved with its
    macros."""
    return (UDF + "\nPublic Function Probe() As String\nDim ws As Object\nApplication.DisplayAlerts = False\n"
            "Set ws = ThisWorkbook.Worksheets(1)\n" + VOLATILE
            + f'ThisWorkbook.SaveAs Filename:="{path}", FileFormat:=52\nProbe = "done"\nEnd Function\n')


def every_function(path: Path, names: list[str]) -> str:
    """A workbook with a call of each function in column C, one to a row; a call Excel refuses leaves its row.

    A function the engine has not got is tried with none to six arguments, each A1, until Excel takes one. The
    answer is what each such call was written as, a ; after each."""
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, n As Long, i As Long",
             "Dim args As String, out As String", "Application.DisplayAlerts = False",
             "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
             'ws.Range("A1").Value = 1', "On Error Resume Next"]
    for row, name in enumerate(names, start=1):
        if name in FUNCTIONS:
            text = call(name).replace('"', '""')
            lines.append(f'ws.Range("C{row}").Formula = "={text}"')
            continue
        lines += ["For n = 0 To 6", '    args = ""', "    For i = 1 To n",
                  '        args = args & IIf(i > 1, ",", "") & "A1"', "    Next",
                  f'    ws.Range("C{row}").Formula = "={name}(" & args & ")"',
                  f'    If ws.Range("C{row}").HasFormula Then Exit For', "Next",
                  f'out = out & "{name}=" & ws.Range("C{row}").Formula & ";"']
    lines += ["On Error GoTo 0", f'wb.SaveAs Filename:="{path}", FileFormat:=51', "wb.Close False",
              "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def _formula_element(cell: str) -> str:
    """A saved cell's <f> element alone, its value left out."""
    found = re.search(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", cell, re.DOTALL)
    return found.group() if found else ""


def formula_cells(path: Path) -> dict[str, str]:
    """Each formula cell as Excel wrote it, by address."""
    with zipfile.ZipFile(path) as package:
        xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    return {match[1]: match[0] for match in re.finditer(r'<c r="([A-Z]+\d+)"(?:(?!</c>).)*<f\b(?:(?!</c>).)*</c>',
                                                           xml, re.DOTALL)}


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    path = FOLDER / "prefixes.xlsx"
    every = FOLDER / "functions.xlsx"
    names = sorted(EXCEL_FUNCTIONS)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(path), "Probe", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        written = excel.run_vba(every_function(every, names), "Probe", timeout=600.0)
        assert written.ok, f"{written.outcome}: {written.message} {written.error}"
        # Saved apart from the fixtures: the probe's own workbook carries the harness's modules.
        volatile = Path(tempfile.mkdtemp()) / "volatile.xlsm"
        result_volatile = excel.run_vba(volatile_module(volatile), "Probe", timeout=300.0)
        assert result_volatile.ok, f"{result_volatile.outcome}: {result_volatile.message} {result_volatile.error}"
    read = [part.split("~") for part in str(result.value).split("|")[:-1]]
    cells = formula_cells(path)
    records = [{"cell": f"C{row}", "written": formula, "saved": cells.get(f"C{row}", ""),
                "formula": one[0], "formula2": one[1], "text": one[2]}
               for row, (formula, one) in enumerate(zip(FORMULAS, read, strict=True), start=1)]
    saved = formula_cells(every)
    tried = dict(part.split("=", 1) for part in str(written.value).split(";") if part)
    functions = {name: {"call": call(name) if name in FUNCTIONS else tried.get(name, "").removeprefix("="),
                        "saved": _formula_element(saved.get(f"C{row}", ""))}
                 for row, name in enumerate(names, start=1)}
    volatile_cells = {address: _formula_element(cell) for address, cell in formula_cells(volatile).items()}
    with zipfile.ZipFile(volatile) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
    names = dict(re.findall(r'<definedName name="([^"]+)"[^>]*>(.*?)</definedName>', book))
    (FOLDER / "prefixes.json").write_text(json.dumps({"formulas": records, "functions": functions,
                                                      "volatile": {"udf": UDF, "writes": VOLATILE,
                                                                   "cells": volatile_cells, "names": names}},
                                                     indent=1) + "\n", encoding="utf-8")
    for record in records:
        print(record["written"], "|", re.sub(r"<v>.*?</v>|<v/>", "", record["saved"]), "|", record["formula"], "|",
              record["formula2"], "|", record["text"])


if __name__ == "__main__":
    main()
