"""What Excel refuses when a formula is written, past its syntax.

Excel will not take a formula through Range.Formula (error 1004) that
gives a function too few or too many arguments, that puts an operation
where a function reads cells, or that names LET's and LAMBDA's names
badly; one too long is error 7. Measured here, on a sheet called Data
with 1 to 3 in A1:A3 and names over it, as
scripts/measure_implicit_intersection.py has it:

- ``counts``: each of Excel's functions written with 0 to 16 and 249 to
  256 arguments, each A1;
- ``places``: each function Excel refused with an operation in every
  argument (tests/fixtures/implicit_intersection.json), written with
  each count it took up to 8 and A1:A2*1 in one argument at a time;
- ``forms``: what such a function takes in the first argument that has
  to be cells: literals, names, references, calls and operations;
- ``values``: LET and LAMBDA, and calls at the edges of what their
  functions take, what each formula taken reads back as and shows;
- ``limits``: long text, long formulas, deep nesting and runs of signs,
  through Formula, Formula2, FormulaR1C1, FormulaArray and Value.

    python scripts/measure_formula_refusals.py

writes tests/fixtures/formula_refusals.json, which
tests/test_excel_formula_refusals.py replays. Calculation is manual
while formulas are written, so none is worked out, except for the
values, which are wanted.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.formula._calc.catalog import EXCEL_FUNCTIONS

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "formula_refusals.json"
INTERSECTION = ROOT / "tests" / "fixtures" / "implicit_intersection.json"
#: The most cases, and the most characters of formula, a probe module writes.
BATCH = 100
BATCH_CHARACTERS = 20_000
#: The longest piece of a formula one line of VBA spells.
PIECE = 400

#: What the sheet and the workbook hold before each formula is written.
SETUP = json.loads(INTERSECTION.read_text(encoding="utf-8"))["setup"]
#: How many arguments each function is written with.
COUNTS = (*range(17), *range(249, 257))
#: Functions whose arguments are not a list of values, written in ``bindings`` instead; ANCHORARRAY is the file's
#: spelling of A1#, which Excel does not take written as a call.
SPECIAL = frozenset({"LET", "LAMBDA", "ANCHORARRAY"})
#: What goes where a function reads cells.
FORMS = (
    "1", '"a"', "TRUE", "#N/A", "{1}", "One", "Const", "Arr", "Calc", "Cells3", "A1:A2", "A:A", "1:1", "Data!A1",
    "E1#", "(A1)", "(A1,A2)", "A1:A2 A1", "A1:INDEX(A1:A2,2)", "INDEX(A1:A2,1)", "LEN(A1)", "IF(1,A1)",
    "NOTAFUNCTION(A1)", "Twice(A1)", "LAMBDA(x,x)(A1)", "A1+0", "-A1", "A1%",
)
#: Formulas whose value is wanted as well: LET and LAMBDA, whose arguments are names and values, and calls at the
#: edges of what their functions take.
VALUES = (
    "=LET()", "=LET(x)", "=LET(x,1)", "=LET(x,1,x)", "=LET(x,1,y,2)", "=LET(x,1,y,2,x+y)", "=LET(1,1,1)",
    "=LET(A1,1,A1)", "=LET(A1,5,A1)", "=LET(A1,5,A1+1)", "=LET(x1,5,x1)", "=LET(R1C1,5,R1C1)", "=LET(A1:A2,5,1)",
    "=LET($A$1,5,1)", "=LET(Data!A1,5,1)", "=LET(x,1,x,2,x)", "=LET(x.y,1,x.y)", "=LET(_x,1,_x)", "=LET(TRUE,1,1)",
    "=LET(One,5,One)", "=LET(SUM,5,SUM)", "=LET(x,A1:A3,SUM(x))", "=LET(x,5,LET(x,6,x))", "=LET(x,5,y,x+1,y)",
    "=LET(\\x,1,1)", "=LET(x\\,1,1)", "=LET(x?,1,1)", "=LET(_xlpm.x,1,1)", "=LET(R,1,R)", "=LET(C,1,C)",
    "=LET(RC,1,RC)", "=LET(A,1,A)",
    "=LAMBDA()", "=LAMBDA(1)", "=LAMBDA(x)", "=LAMBDA(x,1)", "=LAMBDA(1,1)", "=LAMBDA(x,x,1)", "=LAMBDA(A1,1)",
    "=LAMBDA(A1,A1)(5)", "=LAMBDA(x1,x1)(5)", "=LAMBDA(x,y,x+y)(1,2)", "=LAMBDA(x,x)()", "=LAMBDA(x,x)(1,2)",
    "=LAMBDA(x,[y],x)(1)", "=LAMBDA([x],1)()", "=LAMBDA(x,[x],1)", "=LAMBDA(TRUE,1)", "=LAMBDA(x.y,1)",
    "=TEXTJOIN(\",\",TRUE," + ",".join(["A1"] * 251) + ")", "=TEXTJOIN(\",\",TRUE," + ",".join(["A1"] * 252) + ")",
    "=BYROW(A1:A3)", "=BYCOL(A1:A3)", "=INDEX(A1:A3,2,1,1)", "=SWITCH(1,1,5,6)", "=SWITCH(2,1,5,6)",
    "=SORTBY(A1:A3,A1:A3,-1)", "=GETPIVOTDATA(A1,A1,A1)",
)


def _call(name: str, arguments: list[str]) -> str:
    return f"={name}({','.join(arguments)})"


def _long(size: int) -> str:
    """A formula of ``size`` characters joining texts of up to 255 a's, "aa…"&"aa…": few tokens, so only its length
    is at a limit."""
    chunks = -(-size // 258)
    letters = size - 3 * chunks
    return "=" + "&".join('"' + "a" * (letters // chunks + (index < letters % chunks)) + '"' for index in range(chunks))


def _limits() -> list[tuple[str, str, str]]:
    """Formulas at the limits and past them: what each is, the property it is written through, and the formula."""
    cases = [(f"text of {size}", "Formula", '=LEN("' + "a" * size + '")')
             for size in (255, 256, 1000, 4095, 4096, 8000)]
    cases += [(f"text of {size} with quotes", "Formula", '=LEN("' + '""' * 10 + "a" * (size - 10) + '")')
              for size in (4095, 4096)]
    cases += [(f"length {size}", "Formula", _long(size)) for size in (8191, 8192, 8193, 8194)]
    cases += [(f"length {size} through {name}", name, _long(size)) for name in ("Formula2", "FormulaR1C1", "Value")
              for size in (8192, 8193)]
    cases += [(f"length {size} through FormulaArray", "FormulaArray", _long(size))
              for size in (255, 256, 257, 8192, 8193)]
    cases += [(f"{depth} functions deep", "Formula", "=" + "ABS(" * depth + "1" + ")" * depth)
              for depth in (64, 65, 66, 128, 255, 256, 512, 1000, 1600)]
    cases += [(f"{depth} brackets deep", "Formula", "=" + "(" * depth + "1" + ")" * depth)
              for depth in (255, 256, 257, 1000, 4000)]
    cases += [(f"65 functions in {depth} brackets", "Formula", "=" + "(" * depth + "ABS(" * 65 + "1" + ")" * (depth + 65))
              for depth in (256, 257)]
    for sign in "-+":
        cases += [(f"{count} {sign} signs", "Formula", "=" + sign * count + "1")
                  for count in (1000, 1022, 1023, 1024, 1025, 2047, 2048, 2049, 4095, 4096)]
    cases += [(f"{count} signs after a term", "Formula", "=1+" + "-" * count + "1") for count in (1023, 1024, 1025)]
    cases += [(f"{count} signs in a function", "Formula", "=ABS(" + "-" * count + "1)") for count in (1023, 1024, 1025)]
    cases += [(f"{count} signs and brackets", "Formula", "=" + "-(" * count + "1" + ")" * count)
              for count in (256, 257)]
    return cases


def _assign(member: str, formula: str) -> list[str]:
    """VBA lines writing ``formula`` to C8 through ``member``, in pieces a line of VBA can hold."""
    pieces = [formula[start:start + PIECE].replace('"', '""') for start in range(0, len(formula), PIECE)]
    return ['f = ""', *[f'f = f & "{piece}"' for piece in pieces], f'ws.Range("C8").{member} = f']


#: What a probe records of each formula: whether Excel took it, or what it read back and showed.
_TAKEN = 'If Err.Number <> 0 Then out = out & "!" & Err.Number Else out = out & "ok"'
_SHOWN = 'If Err.Number <> 0 Then out = out & "!" & Err.Number Else out = out & ws.Range("C8").Formula & "~" & ' \
         'ws.Range("C8").Text'


def _probe(body: list[str], *, values: bool = False) -> str:
    """A probe module: the sheet set up, ``body`` run with errors trapped, the sheet gone again."""
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String, f As String",
             "Dim functions As Variant, counts As Variant, fn As Variant, n As Variant",
             "Set ws = ActiveWorkbook.Worksheets.Add", SETUP]
    if not values:
        lines.append("Application.Calculation = -4135")
    lines += ["On Error Resume Next", *body, "On Error GoTo 0", "Application.Calculation = -4105",
              "Application.DisplayAlerts = False", "ws.Delete", "Application.DisplayAlerts = True", "Probe = out",
              "End Function",
              "Private Function Several(ByVal count As Long) As String",
              'If count > 0 Then Several = Mid$(Replace(Space$(count), " ", ",A1"), 2)', "End Function"]
    return "\n".join(lines) + "\n"


def _module(cases: list[tuple[str, str]], *, values: bool) -> str:
    body: list[str] = []
    for member, formula in cases:
        body += ["Err.Clear", 'ws.Range("C8").ClearContents', *_assign(member, formula), _SHOWN if values else _TAKEN,
                 'out = out & "|"']
    return _probe(body, values=values)


def _counts(excel: ExcelSession, names: list[str]) -> dict[str, dict[str, str]]:
    """Whether Excel took each function with each of COUNTS arguments, each A1; a loop in VBA, the calls being
    many."""
    counts: dict[str, dict[str, str]] = {}
    for start in range(0, len(names), 40):
        batch = names[start:start + 40]
        body = [f"functions = Array({', '.join(chr(34) + name + chr(34) for name in batch)})",
                f"counts = Array({', '.join(map(str, COUNTS))})",
                "For Each fn In functions", "For Each n In counts", "Err.Clear",
                'ws.Range("C8").Formula = "=" & fn & "(" & Several(n) & ")"', _TAKEN, 'out = out & "|"',
                "Next", "Next"]
        result = excel.run_vba(_probe(body), "Probe", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        found = str(result.value).split("|")[:-1]
        assert len(found) == len(batch) * len(COUNTS), (len(found), len(batch))
        for index, name in enumerate(batch):
            answers = found[index * len(COUNTS):(index + 1) * len(COUNTS)]
            counts[name] = {str(count): answer for count, answer in zip(COUNTS, answers, strict=True)}
        print(f"counted {start + len(batch)} of {len(names)}", flush=True)
    return counts


def _batches(cases: list[tuple[str, str]]) -> list[list[tuple[str, str]]]:
    """The cases in batches of at most BATCH cases and BATCH_CHARACTERS characters of formula, one case at least."""
    batches: list[list[tuple[str, str]]] = []
    size = 0
    for case in cases:
        if not batches or len(batches[-1]) == BATCH or size + len(case[1]) > BATCH_CHARACTERS:
            batches.append([])
            size = 0
        batches[-1].append(case)
        size += len(case[1])
    return batches


def _measure(excel: ExcelSession, cases: list[tuple[str, str]], *, values: bool = False) -> list[str]:
    """What Excel did with each formula written through its member: ok, or what it read back and showed with
    ``values``; or ! and the error number."""
    answers: list[str] = []
    for batch in _batches(cases):
        result = excel.run_vba(_module(batch, values=values), "Probe", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        found = str(result.value).split("|")[:-1]
        assert len(found) == len(batch), (len(found), len(batch))
        answers.extend(found)
        print(f"measured {len(answers)} of {len(cases)}", flush=True)
    return answers


def _refused() -> list[str]:
    """The functions Excel refused with an operation in every argument."""
    record = json.loads(INTERSECTION.read_text(encoding="utf-8"))["functions"]
    return sorted(name for name, variants in record.items()
                  if any(variant.endswith("worked") and read.startswith("!")
                         for variant, (_, read) in variants.items()))


def _formulas(excel: ExcelSession, written: list[tuple[str, str]]) -> dict[str, dict[str, str]]:
    """Whether Excel took each of a function's formulas, by function and formula."""
    found: dict[str, dict[str, str]] = {}
    answers = _measure(excel, [("Formula", formula) for _, formula in written])
    for (name, formula), answer in zip(written, answers, strict=True):
        found.setdefault(name, {})[formula] = answer
    return found


def main() -> None:
    names = sorted(EXCEL_FUNCTIONS - SPECIAL)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        counts = _counts(excel, names)

        written: list[tuple[str, str]] = []
        for name in _refused():
            for count in (count for count in range(1, 9) if counts[name][str(count)] == "ok"):
                for place in range(count):
                    arguments = ["A1"] * count
                    arguments[place] = "A1:A2*1"
                    written.append((name, _call(name, arguments)))
        places = _formulas(excel, written)

        written = []
        for name, measured in places.items():
            refusing = [formula for formula, answer in measured.items() if answer != "ok"]
            if refusing:
                arguments = refusing[0][len(name) + 2:-1].split(",")
                place = arguments.index("A1:A2*1")
                for form in FORMS:
                    written.append((name, _call(name, [*arguments[:place], form, *arguments[place + 1:]])))
        forms = _formulas(excel, written)

        values = dict(zip(VALUES, _measure(excel, [("Formula", formula) for formula in VALUES], values=True),
                          strict=True))
        limits = _limits()
        answers = _measure(excel, [(member, formula) for _, member, formula in limits])
    record = {"setup": SETUP, "counts": counts, "places": places, "forms": forms, "values": values,
              "limits": {label: [member, formula, answer]
                         for (label, member, formula), answer in zip(limits, answers, strict=True)}}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
