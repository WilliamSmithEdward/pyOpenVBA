"""Where Formula2 puts the @ of a legacy formula: the implicit intersection Range.Formula leaves unwritten.

A formula written through Range.Formula is worked out as Excel worked
formulas out before dynamic arrays: where one value is wanted, a range is
cut to the formula's row or column. Formula2 reads such a formula back
with an @ wherever that can happen. Each of Excel's functions is written
through Range.Formula twice, each argument A1 and then each A1:A2, and
read back through Formula2; so is each of a list of formulas exercising
operators, names, LET and LAMBDA and the rest, on a sheet called Data
with 1 to 3 in A1:A3 and names over it.

    python scripts/measure_implicit_intersection.py

writes tests/fixtures/implicit_intersection.json, which
tests/test_excel_formula2.py replays.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "implicit_intersection.json"
CALLS = ROOT / "tests" / "fixtures" / "formula_prefixes" / "prefixes.json"
BATCH = 100

#: What the sheet and the workbook hold before each formula is written.
SETUP = """ws.Name = "Data"
ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
ws.Range("B1").Value = 5
ws.Range("E1").Formula2 = "=SEQUENCE(3)"
ActiveWorkbook.Names.Add "Cells3", "=Data!$A$1:$A$3"
ActiveWorkbook.Names.Add "One", "=Data!$B$1"
ActiveWorkbook.Names.Add "Const", "=5"
ActiveWorkbook.Names.Add "Arr", "={1,2,3}"
ActiveWorkbook.Names.Add "Calc", "=Data!$A$1:$A$3*2"
ActiveWorkbook.Names.Add "Twice", "=LAMBDA(x,x*2)"
"""

#: Formulas written through Range.Formula to a cell of row 8, in the order written.
FORMULAS = [
    "=A1:A3", "=A1", "=A1:A1", "=A:A", "=1:1", "=A1:B1", "=Data!A1:A3", "=Cells3", "=One", "=Const", "=Arr",
    "=Calc", "=Twice(A1:A3)", "=Twice(A1)",
    "=A1:A3*2", "=-A1:A3", "=A1:A3%", '=A1:A3&"x"', "=A1:A3=1", "=(A1:A3)", "=+A1:A3", "=A1:A3+A1:A3",
    "=SUM(A1:A3)", "=SUM(A1:A3*2)", "=SUM(-A1:A3)", "=SUMPRODUCT(A1:A3*2)", "=ABS(A1:A3)", "=ABS(A1)",
    "=ABS(Cells3)", "=LEN(A1:A3)", "=IF(A1:A3>1,1,0)", "=IF(A1>1,A1:A3,0)", "=IF(A1>1,1,A1:A3)",
    "=INDEX(A1:A3,2)", "=INDEX(A1:A3,0)", "=INDEX(A1:A3,A1:A3)", "=VLOOKUP(A1:A3,A1:B3,2,FALSE)",
    "=MATCH(A1:A3,A1:A3,0)", "=COUNTIF(A1:A3,A1:A3)", '=SUMIF(A1:A3,">1",A1:A3)', "=CHOOSE(1,A1:A3)",
    "=CHOOSE(A1:A3,1,2,3)", "=OFFSET(A1,0,0,3,1)", "=SUM(OFFSET(A1,0,0,3,1))", '=INDIRECT("A1:A3")',
    "=ROW(A1:A3)", "=ROWS(A1:A3)", "=ROW()", "=COLUMN(A1:B1)", "=TRANSPOSE(A1:A3)",
    "=MMULT(A1:A3,TRANSPOSE(A1:A3))", "=N(A1:A3)", "=T(A1:A3)", "=IFERROR(A1:A3,0)", "=IFERROR(1/A1:A3,0)",
    "=MAX(A1:A3)*A1:A3", "=SUM(A1:A3)+A1:A3", "={1,2,3}", "={1,2,3}*A1:A3", '=TEXT(A1:A3,"0")',
    "=DATE(2020,A1:A3,1)", "=LET(x,A1:A3,x)", "=LET(x,A1:A3,SUM(x))", "=LET(x,A1:A3,x*2)", "=LET(x,A1,x)",
    "=LAMBDA(x,x*2)(A1:A3)", "=MAP(A1:A3,LAMBDA(v,v*2))", "=SEQUENCE(A1:A3)", "=XLOOKUP(A1:A3,A1:A3,A1:A3)",
    "=XLOOKUP(2,A1:A3,A1:B3)", "=FILTER(A1:A3,A1:A3>1)", "=SORT(A1:A3)", "=A1:A3 A2:A3", "=(A1:A3,B1:B3)",
    "=SUM((A1:A3,B1:B3))", "=E1#", "=SUM(E1#)", "=E1#*2", "=Data!A1:Data!A3", "=SUBTOTAL(9,A1:A3)",
    "=AGGREGATE(9,6,A1:A3)", '=HYPERLINK("x",A1:A3)', '=CELL("row",A1:A3)', "=ISBLANK(A1:A3)", "=AND(A1:A3>1)",
    "=OR(A1:A3)", "=SUM(IF(A1:A3>1,1,0))", "=CONCAT(A1:A3)", '=TEXTJOIN(",",1,A1:A3)', "=UNIQUE(A1:A3)",
    "=XMATCH(A1:A3,A1:A3)", "=A1:A3>A2:A4", "=MAX(A1:A3)", "=MIN(A1:A3,A1:A3*2)", "=NOTAFUNCTION(A1:A3)",
    "=IFS(A1>0,A1:A3)", "=SWITCH(A1,1,A1:A3)", "=LOOKUP(2,A1:A3)", "=HLOOKUP(1,A1:B3,A1:A3)",
    '=SUMIFS(A1:A3,A1:A3,">1")', "=COUNTA(A1:A3)", "=RANK(A1,A1:A3)", "=LARGE(A1:A3,A1:A3)",
    "=@A1:A3", "=LET(x,A1:A2,y,A1:A2,x+y)", "=LET(x,A1:A3,y,x*2,SUM(y))", "=LAMBDA(a,b,a+b)(A1:A2,A1:A2)",
    "=SUM(A1:A3)*A1", "=A1*B1:B2", "=INDEX(A1:B3,2,2)", "=INDEX(A1:B3,2)", "=INDEX(A1:B3,0,1)",
    "=INDEX(A1:B3,1,0)", "=OFFSET(A1,1,1)", "=OFFSET(A1,0,0,1,1)", "=OFFSET(A1:B2,0,0)", "=ROW(A1)",
    "=ROW(A1:A1)", "=COLUMN(A1:A1)", "=XLOOKUP(2,A1:A3,B1:B3)", "=XLOOKUP(2,A1:C1,A2:C3)", "=IF(TRUE,A1,A2)",
    "=IF(TRUE,A1:A2)", "=CHOOSE(2,A1,A2:A3)", "=IFERROR(A1,A2:A3)", "=VLOOKUP(1,A1:B3,{1,2})", "=SUM(A1:A3,A1)",
    "=ISREF(A1:A3)", "=ISFORMULA(A1)", "=ISFORMULA(A1:A1)", '=CELL("row",A1)', '=INDIRECT("A1")',
    '=HYPERLINK(A1:A2,"x")', "=Cells3*2", "=SUM(Cells3)", "=One+Cells3", "=Arr*2", "=SUM(Arr)", "=Twice(Cells3)",
    "=LET(x,Cells3,x)", "=E1#+1", "=SUM(E1#*2)", "=TRANSPOSE(E1#)", "=IF(A1:A3>1,A1:A3)", "=SUM(A1:A3>1)",
    "=COUNT(A1:A3*2)", "=SUMPRODUCT((A1:A3>1)*A1:A3)", "=INDEX(A1:A3*2,2)", "=INDEX(A1:A3*2,0)",
    "=MATCH(2,A1:A3*1,0)", "=VLOOKUP(2,A1:B3*1,2,FALSE)", "=SUM(INDEX(A1:A3,0))", "=MAX(INDEX(A1:A3*2,0))",
]


def _calls() -> dict[str, str]:
    """Each function Excel took a call of, and the call, from the prefixes fixture."""
    record = json.loads(CALLS.read_text(encoding="utf-8"))
    return {name: one["call"] for name, one in record["functions"].items() if one["saved"]}


def _batch(formulas: list[str]) -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String",
             "Set ws = ActiveWorkbook.Worksheets.Add", SETUP, "On Error Resume Next"]
    for formula in formulas:
        quoted = formula.replace('"', '""')
        lines += ["Err.Clear", 'ws.Range("C8").ClearContents', f'ws.Range("C8").Formula = "{quoted}"',
                  'If Err.Number <> 0 Then out = out & "!" & Err.Number Else out = out & ws.Range("C8").Formula2',
                  'out = out & "~|~"']
    lines += ["On Error GoTo 0", "Application.DisplayAlerts = False", "ws.Delete", "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def _wide() -> dict[str, str]:
    """A call of each function the engine has with more arguments than it needs, each A1:A2: up to two more of its
    repeating arguments, or all of its optional ones."""
    from pyopenvba.formula._calc import functions as functions  # imported to register every function
    from pyopenvba.formula._calc.registry import FUNCTIONS

    special = {"LET": "LET(x,A1:A2,y,A1:A2,x+y)", "LAMBDA": "LAMBDA(x,y,x+y)(A1:A2,A1:A2)"}
    found: dict[str, str] = {}
    for name, entry in sorted(FUNCTIONS.items()):
        count = min(entry.maximum, entry.minimum + 2 * entry.repeat)
        if name in special:
            found[name] = "=" + special[name]
        elif count > entry.minimum and count > 0:
            found[name] = f"={name}({','.join(['A1:A2'] * count)})"
    return found


def main() -> None:
    calls = _calls()
    written = [f"={call}" for call in calls.values()]
    ranged = [re.sub(r"\bA1\b", "A1:A2", formula) for formula in written]
    # Each argument an operation on cells: worked out whole, or cut to one value inside a function that reads cells.
    worked = [re.sub(r"\bA1\b", "A1:A2*1", formula) for formula in written]
    wide = _wide()
    wide_worked = {name: formula.replace("A1:A2", "A1:A2*1") for name, formula in wide.items()
                   if name not in ("LET", "LAMBDA")}
    everything = written + ranged + worked + list(wide.values()) + list(wide_worked.values()) + FORMULAS
    read: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for start in range(0, len(everything), BATCH):
            batch = everything[start:start + BATCH]
            result = excel.run_vba(_batch(batch), "Probe", timeout=300.0)
            assert result.ok, f"{result.outcome}: {result.message}"
            answers = str(result.value).split("~|~")[:-1]
            assert len(answers) == len(batch), (len(answers), len(batch))
            read.extend(answers)
            print(f"measured {start + len(batch)} of {len(everything)}", flush=True)
    count = len(calls)
    functions: dict[str, dict[str, list[str]]] = {
        name: {"cell": [written[index], read[index]], "range": [ranged[index], read[count + index]],
               "worked": [worked[index], read[2 * count + index]]}
        for index, name in enumerate(calls)}
    for index, (name, formula) in enumerate(wide.items()):
        functions.setdefault(name, {})["wide"] = [formula, read[3 * count + index]]
    start = 3 * count + len(wide)
    for index, (name, formula) in enumerate(wide_worked.items()):
        functions[name]["wide_worked"] = [formula, read[start + index]]
    formulas = dict(zip(FORMULAS, read[start + len(wide_worked):], strict=True))
    OUT.write_text(json.dumps({"setup": SETUP, "functions": functions, "formulas": formulas}, indent=1) + "\n",
                   encoding="utf-8")


if __name__ == "__main__":
    main()
