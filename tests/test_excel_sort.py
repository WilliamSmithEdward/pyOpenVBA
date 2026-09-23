"""Range.Sort and the Sort object, replayed against the in-memory model.

tests/fixtures/sort_order.json is what scripts/measure_sort_order.py saw in
live Excel: one column of values, each beside its place in the list,
sorted four ways -- ascending and descending, ignoring case and matching
it -- and the places read back in their new order. The model sorts the
same list without the text that has characters beyond ASCII, which it
reports it cannot sort, and has to leave the rest in Excel's order.

tests/fixtures/sort.json is what scripts/measure_sort.py saw: layouts that
fill a small table and sort it, dumping A1:E7 afterwards.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
ORDER: dict[str, Any] = json.loads((FIXTURES / "sort_order.json").read_text(encoding="utf-8"))
RECORD: dict[str, Any] = json.loads((FIXTURES / "sort.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _text(expression: str) -> str:
    found = re.fullmatch(r"ChrW\((\d+)\)", expression)
    return chr(int(found.group(1))) if found else expression[1:-1].replace('""', '"')


#: The places of the values the model sorts: every one but text beyond ASCII.
KEPT = [place for place, expression in enumerate([*ORDER["texts"], *ORDER["others"]], start=1)
        if place > len(ORDER["texts"]) or all(ord(char) < 127 for char in _text(expression))]


def _order_module(places: list[int]) -> str:
    """The probe's module over the values at ``places``, each still beside its place in the full list."""
    expressions = [*[f'"\'" & {text}' for text in ORDER["texts"]], *ORDER["others"]]
    total = len(places) + ORDER["blanks"]
    fill = "\n".join(f"    ws.Cells({row}, 1).Value = {expressions[place - 1]}\n    ws.Cells({row}, 2).Value = {place}"
                     for row, place in enumerate(places, start=1))
    blanks = "\n".join(f"    ws.Cells({len(places) + index}, 2).Value = 0" for index in range(1, ORDER["blanks"] + 1))
    return f'''Private Function Placed(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To {total}
        out = out & ws.Cells(r, 2).Value & ","
    Next
    Placed = out
End Function

Private Sub Fill(ws As Object)
    ws.Cells.Clear
{fill}
{blanks}
End Sub

Public Function Probe() As String
    Dim ws As Object, out As String, order As Variant, cased As Variant, i As Long, j As Long
    Set ws = ActiveWorkbook.Worksheets(1)
    order = Array(xlAscending, xlDescending)
    cased = Array(False, True)
    For i = 0 To 1
        For j = 0 To 1
            Fill ws
            ws.Range("A1:B{total}").Sort Key1:=ws.Range("A1"), Order1:=order(i), Header:=xlNo, MatchCase:=cased(j)
            out = out & Placed(ws) & "|"
        Next
    Next
    Probe = out
End Function
'''


def test_text_beyond_ascii_says_it_cannot_be_sorted() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_order_module(list(range(1, len(ORDER["texts"]) + len(ORDER["others"]) + 1))), name="Probe")
    with pytest.raises(VBAUnsupportedError, match="ASCII"):
        app.run("Probe")


@pytest.mark.parametrize("index", range(4), ids=[f"{one['order']}{' match case' if one['match_case'] else ''}"
                                                 for one in ORDER["sorts"]])
def test_values_sort_in_excels_order(index: int) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_order_module(KEPT), name="Probe")
    got = [int(place) for place in str(app.run("Probe")).split("|")[index].split(",") if place]
    # Blank rows carry place 0 in the replay; Excel's record numbers them past the values.
    values = len(ORDER["texts"]) + len(ORDER["others"])
    kept = set(KEPT)
    measured = [0 if place > values else place for place in ORDER["sorts"][index]["places"]
                if place in kept or place > values]
    assert got == measured


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant",
            "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_a_sort_leaves_what_excel_leaves(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]
