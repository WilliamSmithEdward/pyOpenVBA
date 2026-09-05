"""Live Excel gate for Power Query (opt-in).

A package that parses and rebuilds is a package agreeing with itself.
This gate hands what pyOpenVBA wrote to Excel and asks the mashup engine
to **evaluate** it: the query has to appear in the workbook's query list
and return its rows.  That is the only check that covers the whole chain
-- the custom XML part, the blob, the section document and the metadata
-- at once.

Opt-in: set ``RUN_LIVE_POWER_QUERY=1`` on a Windows machine with desktop
Excel and ``pyvbaharness`` installed.  ``pyvbaharness`` is a test-time
oracle only; pyOpenVBA never uses COM.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.powerquery import PowerQueryWorkbook

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_POWER_QUERY") != "1" or sys.platform != "win32",
    reason="live Power Query gate: set RUN_LIVE_POWER_QUERY=1 on Windows with Excel installed",
)

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"
_TIMEOUT = 600.0

#: Opens a workbook, lists its queries, then loads one into a sheet and
#: reads the cells back.  A query that does not evaluate shows up here as
#: an error rather than as a plausible empty table.
_PROBES = r'''
Public Function ListQueries(ByVal path As String) As Variant
    Dim wb As Workbook, q As Object, out As String
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    For Each q In wb.Queries
        out = out & q.Name & "|"
    Next q
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ListQueries = "ok:" & out
    Exit Function
Failed:
    ListQueries = "ERR " & Err.Number & " " & Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function

Public Function RefreshTables(ByVal path As String) As Variant
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, out As String, r As Long, c As Long
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            lo.QueryTable.Refresh BackgroundQuery:=False
        Next lo
    Next ws
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            out = out & lo.Name & "@" & lo.Range.Address(False, False) & "="
            For r = 1 To lo.Range.Rows.Count
                For c = 1 To lo.Range.Columns.Count
                    out = out & CStr(lo.Range.Cells(r, c).Value) & ","
                Next c
                out = out & ";"
            Next r
        Next lo
    Next ws
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    RefreshTables = "ok:" & out
    Exit Function
Failed:
    RefreshTables = "ERR " & Err.Number & " " & Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function

Public Function Evaluate(ByVal path As String, ByVal qname As String) As Variant
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, out As String, r As Long, c As Long
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)
    Set lo = ws.ListObjects.Add(SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qname & ";Extended Properties=""""", _
        Destination:=ws.Range("$H$1"))
    lo.QueryTable.CommandType = 2
    lo.QueryTable.CommandText = Array("SELECT * FROM [" & qname & "]")
    lo.QueryTable.Refresh BackgroundQuery:=False
    For r = 1 To lo.Range.Rows.Count
        For c = 1 To lo.Range.Columns.Count
            out = out & CStr(lo.Range.Cells(r, c).Value) & ","
        Next c
        out = out & ";"
    Next r
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Evaluate = "ok:" & out
    Exit Function
Failed:
    Evaluate = "ERR " & Err.Number & " " & Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function
'''


@pytest.fixture(scope="module")
def excel() -> Any:
    harness = pytest.importorskip("pyvbaharness")
    with harness.ExcelSession() as session:
        session.new_workbook()
        yield session


def ask(excel: Any, proc: str, *args: str) -> str:
    """Run one probe and give back what Excel said."""
    call = ", ".join(f'"{arg}"' for arg in args)
    body = f"{_PROBES}\nPublic Function Probe() As Variant\n    Probe = {proc}({call})\nEnd Function\n"
    result = excel.run_vba(body, proc="Probe", timeout=_TIMEOUT)
    shown: list[Any] = list(result.dialogs or [])
    dialogs = [str(getattr(dialog, "message", "")) for dialog in shown]
    assert not dialogs, f"Excel showed a dialog: {dialogs}"
    assert result.outcome == "passed", f"{proc} did not run: {result.outcome} {result.value}"
    answer = str(result.value)
    assert answer.startswith("ok:"), f"{proc} failed in Excel: {answer}"
    return answer[3:]


def edited(tmp_path: Path, source: str, name: str) -> Path:
    out = tmp_path / name
    shutil.copyfile(FIXTURES / source, out)
    return out


def test_excel_sees_the_queries_we_read(excel: Any) -> None:
    """The reader and Excel have to agree about a workbook Excel wrote."""
    listed = ask(excel, "ListQueries", str(FIXTURES / "three_queries.xlsx")).strip("|").split("|")
    assert listed == PowerQueryWorkbook(FIXTURES / "three_queries.xlsx").query_names()


def test_a_workbook_we_rewrote_without_changing_it_still_opens(excel: Any, tmp_path: Path) -> None:
    path = edited(tmp_path, "three_queries.xlsx", "resaved.xlsx")
    PowerQueryWorkbook(path).save()
    assert ask(excel, "ListQueries", str(path)).strip("|").split("|") == [
        "Numbers",
        "Doubled",
        "Count Of Rows",
    ]


def test_a_query_we_added_evaluates_in_excel(excel: Any, tmp_path: Path) -> None:
    path = edited(tmp_path, "three_queries.xlsx", "added.xlsx")
    book = PowerQueryWorkbook(path)
    book.add_query("Extra", "let\r\n    Source = 99\r\nin\r\n    Source")
    book.save()
    assert "Extra" in ask(excel, "ListQueries", str(path))
    assert ask(excel, "Evaluate", str(path), "Extra") == "Extra,;99,;"


def test_an_edited_formula_reaches_the_engine(excel: Any, tmp_path: Path) -> None:
    path = edited(tmp_path, "three_queries.xlsx", "edited.xlsx")
    book = PowerQueryWorkbook(path)
    book.query("Numbers").formula = (
        "let\r\n    Source = {1..3},\r\n"
        '    Table = Table.FromList(Source, Splitter.SplitByNothing(), {"N"})\r\n'
        "in\r\n    Table"
    )
    book.save()
    rows = ask(excel, "Evaluate", str(path), "Doubled")
    assert rows == "N,Twice,;1,2,;2,4,;3,6,;"


def test_a_rename_carries_the_queries_that_use_it(excel: Any, tmp_path: Path) -> None:
    """The dependent query still evaluates, which is what says the
    references were rewritten and not merely the name."""
    path = edited(tmp_path, "three_queries.xlsx", "renamed.xlsx")
    book = PowerQueryWorkbook(path)
    book.rename_query("Numbers", "Renamed Numbers")
    book.save()
    assert "Renamed Numbers" in ask(excel, "ListQueries", str(path))
    assert ask(excel, "Evaluate", str(path), "Doubled").startswith("N,Twice,;1,2,;")


def test_a_removed_query_is_gone_from_excel(excel: Any, tmp_path: Path) -> None:
    path = edited(tmp_path, "three_queries.xlsx", "removed.xlsx")
    book = PowerQueryWorkbook(path)
    book.remove_query("Count Of Rows")
    book.save()
    assert ask(excel, "ListQueries", str(path)).strip("|").split("|") == ["Numbers", "Doubled"]


def test_a_workbook_that_never_held_a_query_gets_a_working_one(excel: Any, tmp_path: Path) -> None:
    """Everything the package needs -- the custom XML part, its
    properties, the relationships and the content type -- has to be
    written from nothing."""
    path = edited(tmp_path, "no_queries.xlsx", "fresh.xlsx")
    book = PowerQueryWorkbook(path)
    book.add_query(
        "FromNothing",
        'let\r\n    Source = Table.FromRecords({[A = 1], [A = 2]})\r\nin\r\n    Source',
    )
    book.save()
    assert ask(excel, "ListQueries", str(path)).strip("|") == "FromNothing"
    assert ask(excel, "Evaluate", str(path), "FromNothing") == "A,;1,;2,;"


def test_a_grouped_workbook_still_opens_and_evaluates(excel: Any, tmp_path: Path) -> None:
    """Excel parses the group blob when it opens the package; a value it
    cannot read stops the queries from loading at all."""
    path = edited(tmp_path, "three_queries.xlsx", "grouped.xlsx")
    book = PowerQueryWorkbook(path)
    staging = book.add_group("Staging")
    book.add_group("Raw", parent=staging)
    book.query("Numbers").move_to_group(staging)
    book.save()
    assert "Numbers" in ask(excel, "ListQueries", str(path))
    assert ask(excel, "Evaluate", str(path), "Numbers").startswith("N,;1,;")


def test_a_query_we_loaded_onto_a_sheet_fills_its_table(excel: Any, tmp_path: Path) -> None:
    """The metadata alone does not load a query -- Excel needs the
    connection, the query table and the table too.  Refreshing has to
    fill the table this wrote."""
    path = edited(tmp_path, "three_queries.xlsx", "loaded.xlsx")
    book = PowerQueryWorkbook(path)
    book.add_query(
        "Loadable",
        'let\r\n    Source = Table.FromRecords({[A = 1, B = "x"], [A = 2, B = "y"]})\r\nin\r\n    Source',
    )
    book.load_to_sheet("Loadable", ["A", "B"], cell="H1")
    book.save()
    assert ask(excel, "RefreshTables", str(path)) == "Loadable@H1:I3=A,B,;1,x,;2,y,;"


def test_a_query_we_unloaded_leaves_no_table_behind(excel: Any, tmp_path: Path) -> None:
    """Excel refuses a connections part that holds no connections, so
    unloading the only loaded query has to take the part away."""
    path = edited(tmp_path, "loaded_to_sheet.xlsx", "unloaded.xlsx")
    book = PowerQueryWorkbook(path)
    assert book.unload("Loaded") is True
    book.save()
    assert ask(excel, "RefreshTables", str(path)) == ""
    assert ask(excel, "ListQueries", str(path)).strip("|") == "Loaded"


def test_a_query_can_be_unloaded_and_loaded_somewhere_else(excel: Any, tmp_path: Path) -> None:
    path = edited(tmp_path, "loaded_to_sheet.xlsx", "moved.xlsx")
    book = PowerQueryWorkbook(path)
    book.unload("Loaded")
    book.load_to_sheet("Loaded", ["A", "B"], cell="D1")
    book.save()
    assert ask(excel, "RefreshTables", str(path)) == "Loaded@D1:E3=A,B,;1,x,;2,y,;"


def test_excel_opens_the_example_workbook(excel: Any, tmp_path: Path) -> None:
    """``examples/power_query_demo.py`` writes twelve queries in four
    groups with six of them loaded onto a sheet.  Excel has to open that
    and list all twelve; refreshing it is left alone here, because the
    queries call public APIs and a test should not need the network."""
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "power_query_demo", Path(__file__).parents[1] / "examples" / "power_query_demo.py"
    )
    assert spec is not None and spec.loader is not None
    demo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(demo)
    out = demo.build(tmp_path / "demo.xlsx")

    listed = ask(excel, "ListQueries", str(out)).strip("|").split("|")
    assert len(listed) == 12
    assert {"Pokedex", "PokemonStats", "Earthquakes", "Rates", "GetFromPokeApi"} <= set(listed)


def test_a_query_with_an_awkward_name_evaluates(excel: Any, tmp_path: Path) -> None:
    path = edited(tmp_path, "three_queries.xlsx", "awkward.xlsx")
    book = PowerQueryWorkbook(path)
    book.add_query("Grand Total €", "let\r\n    Source = 42\r\nin\r\n    Source")
    book.save()
    assert "Grand Total €" in ask(excel, "ListQueries", str(path))
    assert ask(excel, "Evaluate", str(path), "Grand Total €") == "Grand Total €,;42,;"
