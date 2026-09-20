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

#: Where a host wants smaller batches than the default.
BATCH_FOR: dict[str, int] = {}

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

#: The same for a slide: each probe gets one nobody has drawn on.  Layout
#: 12 is ppLayoutBlank, so nothing is on it but what the probe adds.
_FRESH_SLIDE = (
    "    Dim probeSlide As Object\n"
    "    Set probeSlide = ActivePresentation.Slides.Add("
    "ActivePresentation.Slides.Count + 1, 12)\n"
    "    probeSlide.Select\n"
)

#: And for a document: everything a probe left behind is cleared out.
#: A whole new document rather than an emptied one: Word numbers a new
#: shape from a counter that does not go back down when shapes are
#: deleted, so only a new document gives the name a fresh model would.
#: The document is emptied rather than replaced: opening a new one
#: moves the document the harness put its module in, and everything
#: after that fails to run.  Word's own counter for shape names does
#: not reset when shapes are deleted, so a probe that wants a name
#: measures the part of it that does not depend on what ran before.
_FRESH_DOCUMENT = (
    "    ActiveDocument.Content.Delete\n"
    "    Dim probeShape As Object\n"
    "    For Each probeShape In ActiveDocument.Shapes\n"
    "        probeShape.Delete\n"
    "    Next probeShape\n"
)

#: What each host's probes get before the probe itself runs.
PREAMBLE = {
    "excel": _FRESH_SHEET,
    "powerpoint": _FRESH_SLIDE,
    "word": _FRESH_DOCUMENT,
}


def _probe_function(
    index: int, line: str, *, fresh_sheet: bool = False, host: str = "excel"
) -> str:
    setup, expression = split_probe(line)
    return (
        f"Private Function P{index}() As String\n"
        f"    On Error GoTo Bad\n"
        f"    Dim v As Variant\n"
        + (PREAMBLE.get(host, _FRESH_SHEET) if fresh_sheet else "")
        + (f"    {setup}\n" if setup else "")
        + f"    v = ({expression})\n"
        f"    P{index} = Describe(v)\n"
        f"    Exit Function\n"
        f"Bad:\n"
        f'    P{index} = "!" & Err.Number\n'
        f"End Function\n"
    )


#: Where a formula probe's own formula is written, clear of the cells a
#: setup writes into.
FORMULA_CELL = "H1"


def _formula_function(index: int, line: str) -> str:
    """A probe that puts a formula in a cell and reports what it shows."""
    setup, formula = split_probe(line)
    quoted = formula.replace('"', '""')
    return (
        f"Private Function P{index}() As String\n"
        f"    On Error GoTo Bad\n"
        f"    Dim v As Variant\n"
        + _FRESH_SHEET
        + (f"    {setup}\n" if setup else "")
        + f'    Range("{FORMULA_CELL}").Formula = "{quoted}"\n'
        f'    v = Range("{FORMULA_CELL}").Value\n'
        f"    If IsError(v) Then\n"
        f'        P{index} = "Error|" & Range("{FORMULA_CELL}").Text\n'
        f"    Else\n"
        f"        P{index} = Describe(v)\n"
        f"    End If\n"
        f"    Exit Function\n"
        f"Bad:\n"
        f'    P{index} = "!" & Err.Number\n'
        f"End Function\n"
    )


def build_module(
    expressions: list[tuple[int, str]],
    *,
    fresh_sheet: bool = False,
    formulas: bool = False,
    host: str = "excel",
) -> str:
    """A module whose Main returns one line per probe."""
    body = ["Public Function Main() As String", "    Dim out As String"]
    for index, _ in expressions:
        body.append(f'    out = out & "{index}=" & P{index}() & vbLf')
    body.append("    Main = out")
    body.append("End Function")
    parts = ["\n".join(body), _DESCRIBE]
    if formulas:
        parts.extend(_formula_function(index, expression) for index, expression in expressions)
    else:
        parts.extend(
            _probe_function(index, expression, fresh_sheet=fresh_sheet, host=host)
            for index, expression in expressions
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


def measure(
    probes: Path,
    measured: Path,
    *,
    fresh_sheet: bool,
    formulas: bool = False,
    host: str = "excel",
) -> int:
    try:
        import pyvbaharness
    except ImportError:
        print("pyvbaharness is not installed; this script needs it and live Office")
        return 2

    session_for = {
        "excel": "ExcelSession",
        "powerpoint": "PowerPointSession",
        "word": "WordSession",
    }[host]
    opener = getattr(pyvbaharness, session_for)

    expressions = read_probes(probes)
    print(f"{len(expressions)} probes")
    answers: dict[str, str] = {}
    with opener() as office:
        office.new_document()
        size = BATCH_FOR.get(host, BATCH)
        for start in range(0, len(expressions), size):
            batch = list(enumerate(expressions))[start : start + size]
            result = office.run_vba(
                build_module(batch, fresh_sheet=fresh_sheet, formulas=formulas, host=host),
                "Main",
                timeout=180.0,
            )
            if result.ok:
                _collect(result.value, expressions, answers)
            else:
                # One probe that will not compile blocks its whole batch,
                # so fall back to one at a time and record which it was.
                print(f"  batch at {start} blocked ({result.outcome}); running singly")
                for one in batch:
                    single = office.run_vba(
                        build_module(
                            [one], fresh_sheet=fresh_sheet, formulas=formulas, host=host
                        ),
                        "Main",
                        timeout=60.0,
                    )
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
    if which in ("powerpoint", "word"):
        folder = ROOT / "tests" / "fixtures" / f"{which}_model"
        return measure(
            folder / "probes.txt", folder / "measured.json", fresh_sheet=True, host=which
        )
    if which == "formula":
        folder = ROOT / "tests" / "fixtures" / "formula"
        return measure(
            folder / "probes.txt", folder / "measured.json", fresh_sheet=True, formulas=True
        )
    return measure(PROBES, MEASURED, fresh_sheet=False)


if __name__ == "__main__":
    sys.exit(main())
