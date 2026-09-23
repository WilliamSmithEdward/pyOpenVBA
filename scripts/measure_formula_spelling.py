"""How Excel spells a formula a macro writes, as Range.Formula reads it back.

Excel does not keep the text it is given: it reads the formula and
writes it out again. Each formula here is written through Range.Formula
into A1 of a sheet called Data, beside a sheet called My Sheet, with a
workbook name MyName and a function MyUdf in the probe's own module, and
read back through Formula and FormulaR1C1. The formulas ask about the
case of references, sheet names, functions known and unknown, names and
constants; how numbers are spelled again; what happens to spaces, line
breaks and brackets; ranges written corner to corner the wrong way; array
constants; and text that starts with + or - rather than =. A formula
Excel refuses records E and the error number.

    python scripts/measure_formula_spelling.py

writes tests/fixtures/formula_spelling.json, which
tests/test_excel_formula_spelling.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "formula_spelling.json"

HELPER = '''Public Function MyUdf(x As Variant) As Variant
    MyUdf = x
End Function

Private Function Spelled(ws As Object, text As String) As String
    Dim c As Object
    Set c = ws.Range("A1")
    c.Clear
    On Error Resume Next
    Err.Clear
    c.Formula = text
    If Err.Number <> 0 Then
        Spelled = "E" & Err.Number
    Else
        Spelled = c.Formula & "<~>" & c.FormulaR1C1 & "<~>" & c.HasFormula
    End If
    On Error GoTo 0
End Function
'''

#: The formulas written, as VBA expressions.
FORMULAS = [
    # References and ranges.
    '"=a1+b2"', '"=$a$1+a$1+$a1"', '"=sum(a1:b2)"', '"=SUM(B2:A1)"', '"=SUM(A2:B1)"', '"=SUM($B$2:A1)"',
    '"=a1:a1"', '"=SUM(a:a)"', '"=SUM(1:1)"', '"=SUM(b:a)"', '"=SUM(2:1)"', '"=SUM($b:a)"', '"=xfd1048576"',
    '"=SUM(A1:B2 B1:C3)"', '"=SUM((A1,B1))"',
    # Sheets.
    '"=data!a1"', '"=DATA!A1"', '"=\'data\'!A1"', '"=\'Data\'!A1"', '"=\'My Sheet\'!a1"', '"=\'my sheet\'!A1"',
    '"=Data!A1:B2"', '"=SUM(Data!B2:A1)"', '"=SUM(data!a1:data!b2)"', '"=My Sheet!A1"',
    # Functions and names.
    '"=sum(1,2)"', '"=Sum(1, 2)"', '"=if(true,1,0)"', '"=foo(1)"', '"=Foo(1)+foo(2)"', '"=myudf(1)"',
    '"=MYUDF(1)"', '"=na()"', '"=pi()"', '"=true"', '"=false"', '"=True+1"', '"=true()"', '"=myname"',
    '"=MYNAME"', '"=undefinedname"', '"=sum(myname)"', '"=Undefined+undefined"', '"=len(""x"")"',
    '"=iferror(1/0,0)"', '"=xlookup(1,a2:a3,b2:b3)"', '"=textjoin("","",true,a2:a3)"',
    # Functions a sheet has that VBA's WorksheetFunction does not, or spells with _ for the dot.
    '"=sin(1)"', '"=address(1,1)"', '"=areas(a1)"', '"=datedif(1,2,""d"")"', '"=norm.dist(1,0,1,true)"',
    '"=Stdev.S(1,2)"', '"=forecast.linear(1,{1,2},{1,2})"', '"=webservice(""x"")"', '"=sequence(2)"',
    '"=let(x,1,X+1)"', '"=Let(Total,1,total*2)"',
    # Numbers.
    '"=1.50"', '"=1e3"', '"=1E+3"', '"=1e-3"', '"=.5"', '"=5."', '"=0005"', '"=1.0E-10"',
    '"=123456789012345678"', '"=1e308"', '"=0.1+0.2"', '"=1234567890.12345"', '"=100000000000000000000"',
    '"=1E+21"', '"=0.000001"', '"=0.0000001"', '"=1e-20"', '"=3.14159265358979323846"', '"=-0"', '"=0.0"',
    # Errors and text.
    '"=#n/a"', '"=#div/0!"', '"=#REF!"', '"=#value!"', '"=#name?"', '"=#num!"', '"=#null!"',
    '"=iserror(#n/a)"', '"=""abc"""', '"=""a""&""b"""', '"=""A1"""', '"="""""', '"=""a b"""',
    # Spaces, line breaks and brackets.
    '"= 1 + 2"', '"=1  +  2"', '"=SUM( 1 , 2 )"', '"=  1"', '"=1 "', '"= sum (1)"', '"=1+" & vbLf & "2"',
    '"=(1)"', '"=((1))"', '"=( 1 )"', '"=SUM(1 ,2)"', '"=-( 1)"', '"=A1 "', '"=A1  +B1"',
    # Operators.
    '"=1<>2"', '"=1>=2"', '"=1 <> 2"', '"=+1"', '"=--1"', '"=1%"', '"=2^3"', '"=-a1"', '"=1&2"', '"=a1=b1"',
    '"=1=1"',
    # Array constants.
    '"={1,2;3,4}"', '"={1, 2}"', '"={1.50,""a"",true}"', '"=SUM({1,2,3})"', '"={1;2}"', '"={-1,+2}"',
    # Not a formula, or not one Excel can read.
    '"+1+2"', '"-1+2"', '"+a1"', '"=+a1"', '"=1+"', '"=sum("', '"=a1:"', '"=)"', '"="',
]


def main() -> None:
    # The probe's own workbook, so that MyUdf is a function its formulas can call.
    lines = [HELPER, "Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = ThisWorkbook", "Set ws = wb.Worksheets(1)",
             'ws.Name = "Data"', 'wb.Worksheets.Add(After:=ws).Name = "My Sheet"',
             'wb.Names.Add Name:="MyName", RefersTo:="=Data!$B$1"']
    lines += [f'out = out & Spelled(ws, {formula}) & "|"' for formula in FORMULAS]
    lines += ["Probe = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(lines) + "\n", "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(FORMULAS)]
    record = {"helper": HELPER, "cases": [{"formula": formula, "answer": answer}
                                          for formula, answer in zip(FORMULAS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for formula, answer in zip(FORMULAS, answers, strict=True):
        print(f"{formula:40} {answer}")


if __name__ == "__main__":
    main()
