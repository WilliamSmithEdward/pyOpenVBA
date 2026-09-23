"""When Excel sets a formula's last sum or difference to zero, bit for bit.

Excel answers =0.5-0.4-0.1 with 0, where the doubles give -2.8E-17: an
addition or subtraction that ends a formula and all but cancels comes
out as zero. This probe finds how near is near. Each row holds two
doubles written through LSet, so they are exactly the ones meant: a
base, and the base moved a number of steps of its last bit up or down.
Formulas over the pair read the answer's eight bytes the same way.

The first block sweeps bases and steps with the subtraction both ways
round, the sum of a negated operand, the difference in parentheses and
the comparison. The second tries the places the difference can stand:
inside a product, a sum, a function, before a +0. The third has the
difference reached through an array formula, a defined name and
Evaluate. The fourth holds a beside -b and adds them with SUM, AVERAGE,
SUMPRODUCT, SUMIF and SUBTOTAL, with a tiny third term first or last to
show which addition is the one set to zero. The fifth compares the pair
with each operator, the criteria functions and the lookups, and reads
the text of b.

    python scripts/measure_zero_snap.py

writes tests/fixtures/zero_snap.json, which tests/test_formula_zero_snap.py replays.
"""

from __future__ import annotations

import json
import math
import random
import struct
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "zero_snap.json"

HELPER = '''Private Type Bits
    Value As Double
End Type

Private Type Halves
    Low As Long
    High As Long
End Type

Private Function Hex64(v As Variant) As String
    Dim b As Bits, h As Halves
    If IsError(v) Then
        Hex64 = "E" & CStr(CLng(v))
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        Hex64 = IIf(v, "TRUE", "FALSE")
        Exit Function
    End If
    If VarType(v) = vbString Then
        Hex64 = "S:" & v
        Exit Function
    End If
    b.Value = CDbl(v)
    LSet h = b
    Hex64 = Right$("00000000" & Hex$(h.High), 8) & Right$("00000000" & Hex$(h.Low), 8)
End Function

Private Function FromHex(text As String) As Double
    Dim b As Bits, h As Halves
    h.High = CLng("&H" & Left$(text, 8))
    h.Low = CLng("&H" & Right$(text, 8))
    LSet b = h
    FromHex = b.Value
End Function
'''

#: The sweep's formulas, R1C1 over the base in column 1 and the moved value in column 2.
SWEEP = {"a-b": "=RC1-RC2", "b-a": "=RC2-RC1", "a+-b": "=RC1+-RC2", "(a-b)": "=(RC1-RC2)", "a=b": "=RC1=RC2"}
#: Where else the difference can stand.
PLACES = {"a-b": "=RC1-RC2", "1*(a-b)": "=1*(RC1-RC2)", "a-b+0": "=RC1-RC2+0", "0+a-b": "=0+RC1-RC2",
          "a-b-0": "=RC1-RC2-0", "-b+a": "=-RC2+RC1", "a*1-b": "=RC1*1-RC2", "SUM(a,-b)": "=SUM(RC1,-RC2)",
          "IF(TRUE,a-b)": "=IF(TRUE,RC1-RC2)", "(a)-b": "=(RC1)-RC2", "a-(b)": "=RC1-(RC2)",
          "a-b&\"\"": "=RC1-RC2&\"\"", "--(a-b)": "=--(RC1-RC2)"}
#: Sums that run across a pair written as a and -b, and whether each snaps, and at which step.
SUMS = {"a+c": "=RC1+RC2", "SUM(a:c)": "=SUM(RC1:RC2)", "SUM(a,c)": "=SUM(RC1,RC2)",
        "SUM(a:c,tiny)": "=SUM(RC1:RC2,1E-20)", "SUM(tiny,a:c)": "=SUM(1E-20,RC1:RC2)", "SUM(a:c)*1": "=SUM(RC1:RC2)*1",
        "AVERAGE(a:c)": "=AVERAGE(RC1:RC2)", "SUMPRODUCT(a:c)": "=SUMPRODUCT(RC1:RC2)",
        "SUMIF(a:c)": "=SUMIF(RC1:RC2,\"<>x\")", "SUBTOTAL(9,a:c)": "=SUBTOTAL(9,RC1:RC2)",
        "a+c+tiny": "=RC1+RC2+1E-20", "tiny+a+c": "=1E-20+RC1+RC2"}
#: Comparisons of the pair, by operator and by the functions that match values.
COMPARE = {"a=b": "=RC1=RC2", "a<>b": "=RC1<>RC2", "a<b": "=RC1<RC2", "a>b": "=RC1>RC2", "a<=b": "=RC1<=RC2",
           "a>=b": "=RC1>=RC2", "IF(a=b)": "=IF(RC1=RC2,1,0)", "MATCH": "=MATCH(RC1,RC2,0)",
           "COUNTIF": "=COUNTIF(RC2,RC1)", "VLOOKUP": "=VLOOKUP(RC1,RC2,1,FALSE)", "EXACT": "=EXACT(RC1,RC2)",
           "a-b=0": "=RC1-RC2=0", "HLOOKUP": "=HLOOKUP(RC1,RC2,1,FALSE)", "XLOOKUP": "=XLOOKUP(RC1,RC2,RC2)",
           "SWITCH": "=SWITCH(RC1,RC2,1,0)", "MATCH(1)": "=MATCH(RC1,RC2,1)", "MATCH(-1)": "=MATCH(RC1,RC2,-1)",
           "SUMIF": "=SUMIF(RC2,RC1)", "COUNTIFS": "=COUNTIFS(RC2,RC1)", "COUNTIF(<)": "=COUNTIF(RC2,\"<\"&RC1)",
           "b&\"\"": "=RC2&\"\"", "VLOOKUP(TRUE)": "=VLOOKUP(RC1,RC2,1,TRUE)",
           "HLOOKUP(TRUE)": "=HLOOKUP(RC1,RC2,1,TRUE)"}
#: The steps of the last bit the second value is moved by.
STEPS = [*range(1, 17), 20, 24, 32, 64]
PLACE_STEPS = list(range(1, 13))
CONTEXT_STEPS = [2, 6, 7, 8, 9, 10, 14]
COMPARE_STEPS = [*range(1, 9), 20, 24]


def bases() -> list[float]:
    fixed = [1.0, 1.5, 2.0 - 2.0 ** -52, 1.0 + 2.0 ** -52, 0.75, 1.25, 1.75, 7.0, 0.01, 0.1, 0.3,
             0.8066666666666666, 3.109126351029605, 1000.1, 12345.678, 1e15, 1e-10, 2.0 ** 52 + 1, 1e300,
             2.0 ** -1000, -0.1, -1.0, -1000.1]
    rng = random.Random(20260923)
    for _ in range(12):
        fixed.append(rng.choice((1.0, -1.0)) * rng.uniform(1.0, 2.0) * 2.0 ** rng.randint(-40, 40))
    return fixed


def moved(base: float, steps: int, up: bool) -> float:
    value = base
    for _ in range(steps):
        value = math.nextafter(value, math.inf if up else -math.inf)
    return value


def bits(value: float) -> str:
    return struct.pack(">d", value).hex().upper()


def pairs(values: list[float], steps: list[int]) -> list[tuple[float, float]]:
    return [(base, moved(base, step, up)) for base in values for step in steps for up in (True, False)]


#: Pairs written per procedure: VBA refuses to compile a procedure past its size limit.
BATCH = 100


def fills(name: str, first: int, rows: list[tuple[float, float]]) -> tuple[list[str], list[str]]:
    """Procedures writing the pairs from row ``first``, and the calls to them."""
    procedures: list[str] = []
    calls: list[str] = []
    for start in range(0, len(rows), BATCH):
        procedure = f"{name}Fill{start // BATCH}"
        procedures.append(f"Private Sub {procedure}(ws As Object)")
        for row, (left, right) in enumerate(rows[start:start + BATCH], first + start):
            procedures.append(f'ws.Cells({row}, 1).Value = FromHex("{bits(left)}"): '
                              f'ws.Cells({row}, 2).Value = FromHex("{bits(right)}")')
        procedures.append("End Sub")
        calls.append(f"{procedure} ws")
    return procedures, calls


def block(name: str, first: int, rows: list[tuple[float, float]], formulas: dict[str, str]) -> list[str]:
    """A procedure writing one block's pairs and formulas and reading every answer back."""
    last = first + len(rows) - 1
    procedures, calls = fills(name, first, rows)
    lines = [*procedures, f"Private Function {name}(ws As Object) As String", "Dim out As String, r As Long, c As Long",
             *calls]
    for column, formula in enumerate(formulas.values(), 3):
        lines.append(f'ws.Range(ws.Cells({first}, {column}), ws.Cells({last}, {column})).FormulaR1C1 = '
                     f'"{formula.replace(chr(34), chr(34) * 2)}"')
    lines += [f"For r = {first} To {last}",
              'out = out & Hex64(ws.Cells(r, 1).Value) & "," & Hex64(ws.Cells(r, 2).Value) & ";"',
              f"For c = 3 To {2 + len(formulas)}",
              'out = out & Hex64(ws.Cells(r, c).Value) & ","', "Next c", 'out = out & "|"', "Next r",
              f"{name} = out", "End Function"]
    return lines


def contexts(first: int, rows: list[tuple[float, float]]) -> list[str]:
    """The difference through an array formula, a defined name and Evaluate, on rows from ``first``."""
    last = first + len(rows) - 1
    lines = ["Private Function Contexts(ws As Object) As String", "Dim out As String, r As Long"]
    for row, (left, right) in enumerate(rows, first):
        lines.append(f'ws.Cells({row}, 1).Value = FromHex("{bits(left)}"): '
                     f'ws.Cells({row}, 2).Value = FromHex("{bits(right)}")')
        # An underscore: Gap3001 would be a cell, column GAP.
        refers = f'"=" & ws.Name & "!$A${row}-" & ws.Name & "!$B${row}"'
        lines.append(f'ws.Parent.Names.Add Name:="Gap_{row}", RefersTo:={refers}')
        lines.append(f'ws.Cells({row}, 4).Formula = "=Gap_{row}"')
    lines.append(f'ws.Range("C{first}:C{last}").FormulaArray = "=A{first}:A{last}-B{first}:B{last}"')
    lines += [f"For r = {first} To {last}",
              'out = out & Hex64(ws.Cells(r, 1).Value) & "," & Hex64(ws.Cells(r, 2).Value) & ";"',
              'out = out & Hex64(ws.Cells(r, 3).Value) & "," & Hex64(ws.Cells(r, 4).Value) & ","',
              'out = out & Hex64(ws.Evaluate("A" & r & "-B" & r)) & ","',
              'out = out & Hex64(Application.Evaluate("=" & ws.Name & "!A" & r & "-" & ws.Name & "!B" & r)) & ","',
              'out = out & "|"', "Next r", "Contexts = out", "End Function"]
    return lines


CONTEXTS = ["array", "name", "ws.Evaluate", "Application.Evaluate"]


def read(answer: str, names: list[str]) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    for part in answer.split("|"):
        if not part:
            continue
        cells, answers = part.split(";")
        left, right = cells.split(",")
        found = dict(zip(names, answers.rstrip(",").split(","), strict=True))
        records.append({"a": left, "b": right, "answers": found})
    return records


def main() -> None:
    values = bases()
    sweep = pairs(values, STEPS)
    places = pairs([1.0, 1.5, 0.1, 1000.1, 2.0 - 2.0 ** -52], PLACE_STEPS)
    reached = pairs([0.1, 1.5], CONTEXT_STEPS)
    # a beside -b, so a range can run across the pair.
    sums = [(left, -right) for left, right in pairs([1.0, 1.5, 0.1, 1000.1, 2.0 - 2.0 ** -52], PLACE_STEPS)]
    compared = pairs([3.109126351029605, 0.8066666666666666, 1.0, 0.1, 1000.1, 2.0 ** 52 + 1], COMPARE_STEPS)
    code = [HELPER, *block("Sweep", 1, sweep, SWEEP), *block("Places", 2001, places, PLACES),
            *contexts(3001, reached), *block("Sums", 4001, sums, SUMS), *block("Compare", 5001, compared, COMPARE),
            "Public Function Probe() As String", "Dim ws As Object, out As String",
            "Set ws = ActiveWorkbook.Worksheets(1)",
            'out = Sweep(ws) & "#" & Places(ws) & "#" & Contexts(ws) & "#" & Sums(ws) & "#" & Compare(ws)',
            "Probe = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(code) + "\n", "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    first, second, third, fourth, fifth = str(result.value).split("#")
    record = {"sweep": {"formulas": SWEEP, "rows": read(first, list(SWEEP))},
              "places": {"formulas": PLACES, "rows": read(second, list(PLACES))},
              "contexts": {"formulas": {name: name for name in CONTEXTS}, "rows": read(third, CONTEXTS)},
              "sums": {"formulas": SUMS, "rows": read(fourth, list(SUMS))},
              "compare": {"formulas": COMPARE, "rows": read(fifth, list(COMPARE))}}
    for name, rows in (("sweep", sweep), ("places", places), ("contexts", reached), ("sums", sums),
                       ("compare", compared)):
        got = record[name]["rows"]
        assert len(got) == len(rows), (name, len(got), len(rows))
        for one, (left, right) in zip(got, rows, strict=True):
            assert (one["a"], one["b"]) == (bits(left), bits(right)), (name, one, left, right)
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(len(sweep), "sweep rows,", len(places), "place rows,", len(reached), "context rows,", len(sums), "sum rows,",
          len(compared), "comparison rows")


if __name__ == "__main__":
    main()
