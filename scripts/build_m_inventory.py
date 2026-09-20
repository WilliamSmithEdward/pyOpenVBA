"""Ask Power Query for the names it knows, and commit the list.

``#shared`` is a record of everything in scope in an M expression, which
is the library itself.  This asks live Excel for it and writes
``src/pyopenvba/mlang/_inventory_data.py``, so that a name this does not
implement can be told apart from a name M has never had without either
being guessed at.

Run it on a Windows machine with Excel installed:

    python scripts/build_m_inventory.py
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "src" / "pyopenvba" / "mlang" / "_inventory_data.py"
SCRATCH = ROOT / "tests" / "fixtures" / "mlang" / "_scratch"

QUERY = """let
    Names = Record.FieldNames(#shared),
    Sorted = List.Sort(Names),
    Result = #table({"Name"}, List.Transform(Sorted, (one) => {one}))
in
    Result"""

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
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        out = out & CStr(ActiveSheet.Cells(r, 1).Value) & vbLf
    Next r
    Report = out
    Exit Function
Bad:
    Report = "!err " & Err.Number & " " & Err.Description
End Function
"""


def main() -> int:
    try:
        from pyvbaharness import ExcelSession
    except ImportError:
        print("pyvbaharness is not installed; this script needs it and live Excel")
        return 2
    from pyopenvba import PowerQueryWorkbook
    from pyopenvba.excel import ExcelFile

    SCRATCH.mkdir(parents=True, exist_ok=True)
    seed = SCRATCH / "shared_seed.xlsm"
    path = SCRATCH / "shared.xlsm"
    ExcelFile.create_new(seed)
    book = PowerQueryWorkbook(seed)
    book.add_query("Shared", QUERY)
    book.load_to_sheet("Shared", ["Name"])
    book.save(path)

    with ExcelSession() as excel:
        excel.open_document(path)
        result = excel.run_vba(READER, "Report", timeout=300.0)
        if not result.ok:
            print(f"failed: {result.outcome} {result.message}")
            return 1
        body = str(result.value or "")
    if body.startswith("!err"):
        print(body)
        return 1
    names = sorted({line.strip() for line in body.split("\n") if line.strip()})
    if len(names) < 100:
        print(f"only {len(names)} names came back; something went wrong")
        return 1
    lines = [
        '"""Every name Power Query has, asked of the engine itself.',
        "",
        "Written by scripts/build_m_inventory.py from #shared, which is the",
        "record of everything in scope in an M expression.  Do not edit by",
        "hand.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "from typing import Final",
        "",
        "#: Joined with commas rather than held as a set, because the list",
        "#: costs less to import this way and is split once on first use.",
        "NAMES: Final = (",
    ]
    body_text = ",".join(names)
    for start in range(0, len(body_text), 88):
        lines.append(f"    {body_text[start:start + 88]!r}")
    lines.extend([")", ""])
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"{len(names)} names -> {OUT.relative_to(ROOT)} ({OUT.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
