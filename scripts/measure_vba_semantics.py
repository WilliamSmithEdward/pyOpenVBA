"""Measure what real VBA answers for every probe, and record it.

Reads ``tests/fixtures/vba_semantics/probes.txt``, evaluates each
expression inside live Excel through pyvbaharness, and writes
``measured.json`` beside it.  ``tests/test_vba_semantics.py`` then holds
the interpreter to that file without needing Office.

Run it on a Windows machine with Excel installed:

    python scripts/measure_vba_semantics.py

Each answer is ``TypeName|CStr`` -- the type VBA gave the value and the
string it turns into -- or ``!<number>`` for the error VBA raised.  Both
halves matter: an interpreter that answers 1 where VBA answers 1& is
wrong in a way that shows up later as an overflow that never happened.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PROBES = ROOT / "tests" / "fixtures" / "vba_semantics" / "probes.txt"
MEASURED = ROOT / "tests" / "fixtures" / "vba_semantics" / "measured.json"

#: Excel's own limit on one procedure, which is why probes run in batches.
BATCH = 40

_DESCRIBE = """
Private Function Describe(v As Variant) As String
    On Error GoTo Bad
    If IsNull(v) Then
        Describe = "Null|Null"
    ElseIf IsObject(v) Then
        Describe = TypeName(v) & "|<object>"
    ElseIf IsArray(v) Then
        Describe = TypeName(v) & "|" & LBound(v) & ".." & UBound(v)
    ElseIf IsError(v) Then
        Describe = "Error|<error>"
    Else
        Describe = TypeName(v) & "|" & CStr(v)
    End If
    Exit Function
Bad:
    Describe = "!" & Err.Number
End Function
"""


def split_probe(line: str) -> tuple[str, str]:
    """A probe is an expression, or setup and an expression around ``;;``."""
    setup, marker, expression = line.partition(";;")
    if not marker:
        return "", line.strip()
    return setup.strip(), expression.strip()


#: Given to every probe when the sheet has to start empty, which is how
#: a workbook probe gets a grid nobody else has written on.
_FRESH_SHEET = (
    "    Dim probeSheet As Worksheet\n"
    "    Set probeSheet = ActiveWorkbook.Worksheets.Add\n"
    "    probeSheet.Activate\n"
)


def _probe_function(index: int, line: str, *, fresh_sheet: bool = False) -> str:
    setup, expression = split_probe(line)
    return (
        f"Private Function P{index}() As String\n"
        f"    On Error GoTo Bad\n"
        f"    Dim v As Variant\n"
        + (_FRESH_SHEET if fresh_sheet else "")
        + (f"    {setup}\n" if setup else "")
        + f"    v = ({expression})\n"
        f"    P{index} = Describe(v)\n"
        f"    Exit Function\n"
        f"Bad:\n"
        f'    P{index} = "!" & Err.Number\n'
        f"End Function\n"
    )


def build_module(expressions: list[tuple[int, str]], *, fresh_sheet: bool = False) -> str:
    """A module whose Main returns one line per probe."""
    body = ["Public Function Main() As String", "    Dim out As String"]
    for index, _ in expressions:
        body.append(f'    out = out & "{index}=" & P{index}() & vbLf')
    body.append("    Main = out")
    body.append("End Function")
    parts = ["\n".join(body), _DESCRIBE]
    parts.extend(
        _probe_function(index, expression, fresh_sheet=fresh_sheet) for index, expression in expressions
    )
    return "\n".join(parts)


def _collect(value: object, expressions: list[str], answers: dict[str, str]) -> None:
    for line in str(value or "").split("\n"):
        index, _, answer = line.partition("=")
        if index.isdigit():
            answers[expressions[int(index)]] = answer.rstrip("\r")


def read_probes(path: Path = PROBES) -> list[str]:
    out: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            out.append(stripped)
    return out


def measure(probes: Path, measured: Path, *, fresh_sheet: bool) -> int:
    try:
        from pyvbaharness import ExcelSession
    except ImportError:
        print("pyvbaharness is not installed; this script needs it and live Excel")
        return 2

    expressions = read_probes(probes)
    print(f"{len(expressions)} probes")
    answers: dict[str, str] = {}
    with ExcelSession() as excel:
        excel.new_document()
        for start in range(0, len(expressions), BATCH):
            batch = list(enumerate(expressions))[start : start + BATCH]
            result = excel.run_vba(build_module(batch, fresh_sheet=fresh_sheet), "Main", timeout=180.0)
            if result.ok:
                _collect(result.value, expressions, answers)
            else:
                # One probe that will not compile blocks its whole batch,
                # so fall back to one at a time and record which it was.
                print(f"  batch at {start} blocked ({result.outcome}); running singly")
                for one in batch:
                    single = excel.run_vba(build_module([one], fresh_sheet=fresh_sheet), "Main", timeout=60.0)
                    if single.ok:
                        _collect(single.value, expressions, answers)
                    else:
                        answers[expressions[one[0]]] = "!compile"
                        print(f"    {one[1]!r} does not compile in VBA")
            print(f"  measured {start + len(batch)} of {len(expressions)}")

    missing = [expression for expression in expressions if expression not in answers]
    if missing:
        print(f"no answer for {len(missing)} probes, first: {missing[0]!r}")
        return 1
    measured.write_text(
        json.dumps({"probes": answers}, indent=1, ensure_ascii=False, sort_keys=False) + "\n",
        encoding="utf-8",
    )
    print(f"wrote {measured.relative_to(ROOT)}")
    return 0


def main() -> int:
    which = sys.argv[1] if len(sys.argv) > 1 else "language"
    if which == "excel":
        folder = ROOT / "tests" / "fixtures" / "excel_model"
        return measure(folder / "probes.txt", folder / "measured.json", fresh_sheet=True)
    return measure(PROBES, MEASURED, fresh_sheet=False)


if __name__ == "__main__":
    sys.exit(main())
