"""Measure what Power Query answers for every M probe, and record it.

Reads ``tests/fixtures/mlang/probes.txt``, builds a workbook whose
query evaluates all of them, has live Excel refresh it, and writes what
came back to ``measured.json`` beside the probes.
``tests/test_mlang_semantics.py`` then holds the evaluator to that file
without needing Office.

Every probe is described the same way on both sides: the type M gave the
value and a short rendering of it, or ``!Reason`` for the error M
raised.  One refresh covers a whole batch, which is what makes measuring
a hundred of them take a minute rather than an hour.

Run it on a Windows machine with Excel installed:

    python scripts/measure_m.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "mlang"
PROBES = FOLDER / "probes.txt"
MEASURED = FOLDER / "measured.json"

#: How many probes go into one query.  One that will not parse takes its
#: whole batch down, so the batch is small enough to bisect cheaply.
BATCH = 25

#: The same description both sides write, in M.
_DESCRIBE = """    Describe = (v) =>
        if v = null then "null|null"
        else if Value.Is(v, type table) then
            "table|" & Text.From(Table.RowCount(v)) & "x" & Text.From(Table.ColumnCount(v))
                & "|" & Text.Combine(Table.ColumnNames(v), ",")
                & "|" & Text.Combine(List.Transform(Table.ToRows(v), (r) =>
                    Text.Combine(List.Transform(r, (c) => if c = null then "" else Text.From(c)), ",")), ";")
        else if Value.Is(v, type record) then
            "record|" & Text.Combine(Record.FieldNames(v), ",")
                & "|" & Text.Combine(List.Transform(Record.FieldValues(v), (c) =>
                    if c = null then "" else Text.From(c)), ",")
        else if Value.Is(v, type list) then
            "list|" & Text.From(List.Count(v))
                & "|" & Text.Combine(List.Transform(v, (c) => if c = null then "" else Text.From(c)), ",")
        else if Value.Is(v, type function) then "function|"
        else if Value.Is(v, type logical) then "logical|" & Text.From(v)
        else if Value.Is(v, type number) then "number|" & Text.From(v)
        else if Value.Is(v, type text) then "text|" & v
        else if Value.Is(v, type date) then "date|" & Date.ToText(v, "yyyy-MM-dd")
        else if Value.Is(v, type datetime) then "datetime|" & DateTime.ToText(v, "yyyy-MM-dd HH:mm:ss")
        else if Value.Is(v, type time) then "time|" & Time.ToText(v)
        else if Value.Is(v, type duration) then "duration|" & Duration.ToText(v)
        else "other|",
    Show = (r) => if r[HasError] then "!" & r[Error][Reason] else Describe(r[Value]),
"""


def read_probes(path: Path = PROBES) -> list[str]:
    out: list[str] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        # M's own comment marker, since a probe may open with #table.
        if stripped and not stripped.startswith("//"):
            out.append(stripped)
    return out


def build_query(batch: list[tuple[int, str]]) -> str:
    """One query that evaluates every probe in the batch."""
    steps = [f"    Probe{index} = try ({source})," for index, source in batch]
    rows = ", ".join(f'{{{index}, Show(Probe{index})}}' for index, _ in batch)
    return (
        "let\n"
        + _DESCRIBE
        + "\n".join(steps)
        + f'\n    Result = #table({{"I", "V"}}, {{{rows}}})\n'
        "in\n"
        "    Result"
    )


READER = """
Public Function Report() As String
    On Error GoTo Bad
    Dim c As WorkbookConnection
    For Each c In ActiveWorkbook.Connections
        On Error Resume Next
        c.OLEDBConnection.BackgroundQuery = False
        On Error GoTo Bad
    Next c
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Dim out As String, r As Long
    Dim last As Long
    last = ActiveSheet.UsedRange.Rows.Count
    For r = 2 To last
        out = out & CStr(ActiveSheet.Cells(r, 1).Value) & "=" & CStr(ActiveSheet.Cells(r, 2).Value) & vbLf
    Next r
    Report = out
    Exit Function
Bad:
    Report = "!err " & Err.Number & " " & Err.Description
End Function
"""


def measure() -> int:
    try:
        from pyvbaharness import ExcelSession
    except ImportError:
        print("pyvbaharness is not installed; this script needs it and live Excel")
        return 2
    from pyopenvba import PowerQueryWorkbook
    from pyopenvba.excel import ExcelFile

    probes = read_probes()
    print(f"{len(probes)} probes")
    answers: dict[str, str] = {}
    scratch = FOLDER / "_scratch"
    scratch.mkdir(exist_ok=True)
    with ExcelSession() as excel:
        for start in range(0, len(probes), BATCH):
            batch = list(enumerate(probes))[start : start + BATCH]
            path = scratch / f"batch_{start}.xlsm"
            _build_workbook(ExcelFile, PowerQueryWorkbook, path, build_query(batch))
            excel.open_document(path)
            result = excel.run_vba(READER, "Report", timeout=240.0)
            if not result.ok:
                print(f"  batch at {start} failed: {result.outcome} {result.message}")
                return 1
            body = str(result.value or "")
            if body.startswith("!err"):
                print(f"  batch at {start}: {body}")
                return 1
            for line in body.split("\n"):
                index, _, answer = line.partition("=")
                if index.strip().isdigit():
                    answers[probes[int(float(index))]] = answer.strip()
            print(f"  measured {start + len(batch)} of {len(probes)}")

    missing = [one for one in probes if one not in answers]
    if missing:
        print(f"no answer for {len(missing)} probes, first: {missing[0]!r}")
        for one in missing:
            answers[one] = "!unmeasured"
    MEASURED.write_text(
        json.dumps({"probes": answers}, indent=1, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(f"wrote {MEASURED.relative_to(ROOT)}")
    return 0


def _build_workbook(excel_file: object, power_query: object, path: Path, formula: str) -> None:
    """A workbook holding one query, loaded to a sheet so it refreshes."""
    seed = path.with_name("seed.xlsm")
    excel_file.create_new(seed)  # type: ignore[attr-defined]
    book = power_query(seed)  # type: ignore[operator]
    book.add_query("Probes", formula)
    book.load_to_sheet("Probes", ["I", "V"])
    book.save(path)


if __name__ == "__main__":
    sys.exit(measure())
