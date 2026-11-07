"""What Range.Value, Value2, Formula and FormulaLocal used to report and now answer, replayed from live Excel
(tests/fixtures/value_gaps.json, tests/fixtures/value_gaps/).

Typed strings Excel misreads: a space, a month's name and a time, read one place along, and digits after a point after AM or
PM. Number formats with only text before the exponent. Rows cleared to the last row under columns of different formats, and
what a format reaching column XFD leaves. Built-in styles under a theme a macro changed. And formulas Excel reads that the
model refused or spelled otherwise: a union with no brackets round it, an @ after a reference operator, two points between
cells, a field of a cell, a tab, an array whose rows differ.

Each round's inputs were new, with the model's answers written down first. The scores keep every round, the fixing ones too.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication, _typing

HERE = Path(__file__).parent / "fixtures"
RECORD: dict[str, Any] = json.loads((HERE / "value_gaps.json").read_text(encoding="utf-8"))
#: Cases a batch can ask together, and how many.
BATCHED = ("typed", "format", "formula")
BATCH = 40


def _vb(text: str) -> str:
    return '"' + text.replace('"', '""').replace("\t", '" & vbTab & "') + '"'


def _statements(case: dict[str, Any]) -> list[str]:
    if case["kind"] == "typed":
        return ['Set c = ws.Range("A1"): c.Clear: Err.Clear', f'c.{case["writer"]} = {_vb(case["text"])}',
                's = Err.Number & "|" & TypeName(c.Value) & "|" & CStr(c.Value2) & "|" & c.NumberFormat & "|" & c.Formula']
    if case["kind"] == "format":
        return ['Set c = ws.Range("A1"): c.Clear: Err.Clear', f'c.NumberFormat = {_vb(case["format"])}', 's = Err.Number & "|": Err.Clear',
                f'c.Value2 = {case["value"]}', 's = s & c.NumberFormat & "|" & c.Text']
    cell = case["cell"]
    return [f'Set c = ws.Range("{cell}"): c.ClearContents: Err.Clear', f'c.{case["writer"]} = {_vb(case["text"])}',
            's = Err.Number & "|" & c.Formula & "|" & c.FormulaLocal & "|" & Show(c.Value)']


def _batch(cases: list[dict[str, Any]]) -> str:
    lines = ['Public Function Probe() As String', 'Dim wb, ws, c, s, out, lo: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Sheet1"']
    if cases[0]["kind"] == "formula":
        lines += ['ws.Range("D1").Value = "a": ws.Range("E1").Value = "b": ws.Range("D2").Value = 1: ws.Range("D3").Value = 2: ws.Range("E2").Value = 3: ws.Range("E3").Value = 4',
                  'Set lo = ws.ListObjects.Add(1, ws.Range("D1:E3"), , 1): lo.Name = "Tab"',
                  'ws.Range("A1").Value = 1: ws.Range("A2").Value = 2: ws.Range("A3").Value = 3: ws.Range("B1").Value = 5: ws.Range("B2").Value = 6: ws.Range("B3").Value = 7']
    lines += ["ws.Columns(1).ColumnWidth = 60: ws.Columns(6).ColumnWidth = 60: ws.Columns(8).ColumnWidth = 60", "On Error Resume Next", 'out = ""']
    for case in cases:
        lines += [*_statements(case), 'out = out & "{" & s & "}"']
    lines += ["wb.Close False", "Probe = out", "End Function", ""]
    return RECORD["helper"] + "\n".join(lines)


def _clear(case: dict[str, Any]) -> str:
    lines = ["Public Function Probe() As String", "Dim wb, ws, c, r, s: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)", "On Error Resume Next", *case["setup"],
             f'ws.Rows("{case["rows"]}").{case["action"]}', 's = Err.Number & "|" & ws.UsedRange.Address & "|"', "Err.Clear"]
    for column in case["columns"]:
        for row in (1, 2, 3, 6, 8, 100, 1048576):
            place = f'ws.Range("{column}{row}")'
            lines.append(f's = s & {place}.Interior.Color & "," & {place}.Font.Bold & "," & {place}.Font.Italic & "," & {place}.NumberFormat & ";"')
    lines += ["wb.Close False", "Probe = s", "End Function", ""]
    return "\n".join(lines)


def _theme(case: dict[str, Any]) -> str:
    path = str(HERE / "value_gaps" / case["seed_book"])
    reads = [f's = s & ws.Range("B{i}").Font.Name & "," & ws.Range("B{i}").Font.Size & "," & ws.Range("B{i}").Font.Underline & "," & ws.Range("B{i}").Font.Bold & ";"' for i in range(1, 9)]
    return "\n".join(["Public Function Probe() As String", f"Dim wb, ws, s: Set wb = Workbooks.Open({_vb(path)}): Set ws = wb.Worksheets(1)", "On Error Resume Next", *case["steps"], 's = Err.Number & "|": Err.Clear', *reads, "wb.Close False", "Probe = s", "End Function", ""])


def _run(source: str) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    with pytest.MonkeyPatch.context() as patch:
        # A date typed without a year falls in the year the round ran.
        patch.setattr(_typing, "_this_year", lambda: RECORD["year"])
        return str(app.run("M.Probe"))


@pytest.mark.parametrize("kind", BATCHED)
@pytest.mark.parametrize("name", sorted(RECORD["rounds"]))
def test_each_answer_as_excel_gave_it(name: str, kind: str) -> None:
    cases = [case for case in RECORD["rounds"][name] if case["kind"] == kind]
    wrong: list[tuple[Any, ...]] = []
    for start in range(0, len(cases), BATCH):
        part = cases[start:start + BATCH]
        got = _run(_batch(part))[1:-1].split("}{")
        wrong += [(case.get("text") or case.get("format"), case["excel"], one) for case, one in zip(part, got, strict=True) if case["excel"] != one]
    assert wrong == []


@pytest.mark.parametrize("name", sorted(RECORD["rounds"]))
def test_rows_cleared_as_excel_cleared_them(name: str) -> None:
    wrong = [(case["setup"], case["rows"], case["action"]) for case in RECORD["rounds"][name] if case["kind"] == "clear" and _run(_clear(case)) != case["excel"]]
    assert wrong == []


@pytest.mark.parametrize("name", sorted(RECORD["rounds"]))
def test_built_in_styles_under_a_macros_theme(name: str) -> None:
    wrong = [(case["minor"], case["steps"]) for case in RECORD["rounds"][name] if case["kind"] == "theme" and _run(_theme(case)) != case["excel"]]
    assert wrong == []


def test_the_scores_of_the_last_rounds() -> None:
    assert [RECORD["scores"][name]["differ"] for name in sorted(RECORD["rounds"])] == [0] * len(RECORD["rounds"])
