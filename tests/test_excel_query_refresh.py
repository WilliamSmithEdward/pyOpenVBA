"""Query tables over a database, refreshed through the host's Refresh function, against what Excel did with an Access file through its
own ACE OLEDB provider (scripts/measure_query_refresh.py, scripts/measure_query_refresh_layout.py).

The function here stands for the Access file: the same rows, chosen by the command's WHERE, as Python values of the types ACE gives.
"""

from __future__ import annotations

import datetime as dt
import importlib.util
import json
import re
import sys
from collections.abc import Callable
from decimal import Decimal
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest
from pyopenvba.apps.excel._query_databases import masked, without_passwords
from pyopenvba.exceptions import BackgroundRefreshError, VBARuntimeError, VBAStepLimitError, VBAUnsupportedError

HERE = Path(__file__).parent / "fixtures"
LAYOUT: dict[str, Any] = json.loads((HERE / "query_refresh_layout.json").read_text(encoding="utf-8"))
TYPES: dict[str, Any] = json.loads((HERE / "query_refresh.json").read_text(encoding="utf-8"))
#: The rows of the Access table: an id and a name.
NUMBERED: list[tuple[int, str]] = [(at, f"n{at}") for at in range(1, 7)]
#: The two rows of the types table, as ACE hands them over.
TYPED: list[list[object]] = [["abc", "1E3", "long text", 7, 1.5, Decimal("12.34"), Decimal("3.25"), dt.datetime(2024, 1, 15), True, None],
         ["00123", "12.50", None, -3, 2.0, Decimal("0"), Decimal("4"), dt.datetime(2024, 2, 20, 13, 30), False, ""]]


def _script(name: str) -> ModuleType:
    folder = Path(__file__).parents[1] / "scripts"
    sys.path.insert(0, str(folder))
    spec = importlib.util.spec_from_file_location(name, folder / f"{name}.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MEASURE = _script("measure_query_refresh_layout")


def numbered(request: HookRequest) -> list[list[object]]:
    assert request.source is not None
    wanted = int(re.search(r"Id <= (\d+)", request.source.command_text).group(1))  # type: ignore[union-attr]
    return [["Id", "Name"], *[[float(at), name] for at, name in NUMBERED if at <= wanted]]


def returning(rows: object) -> Callable[[HookRequest], object]:
    """A function that gives back ``rows``, whatever is asked."""

    def give(request: HookRequest) -> object:
        del request
        return rows

    return give


def failing_with(failure: Exception) -> Callable[[HookRequest], object]:
    def fail(request: HookRequest) -> object:
        del request
        raise failure

    return fail


def _app(function: Callable[[HookRequest], object] = numbered) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook("Refresh", function)
    return app


#: Measured and not built: a ListObject over a database is a table that follows each refresh, with its own parts.
NOT_BUILT = {"list_object", "list_object_shrink"}


@pytest.mark.parametrize("case", sorted(LAYOUT["rows"]))
def test_lays_out_rows_as_excel_did(case: str) -> None:
    if case in NOT_BUILT:
        pytest.xfail("ListObjects.Add over a database is not built")
    app = _app()
    code = MEASURE.CASE.replace("{steps}", "\n".join(f"    {step}" for step in LAYOUT["cases"][case]))
    app.add_module(code, name="M")
    assert str(app.run("Go", "C:/data/rows.accdb")).split("\t") == LAYOUT["rows"][case]


def test_puts_each_type_as_excel_did() -> None:
    names = [column[0] for column in TYPES["columns"]]
    app = _app(returning([names, *TYPED]))
    measure = _script("measure_query_refresh")
    app.add_module(measure.source(), name="M")
    found = str(app.run("Types", "C:/data/plain.accdb")).split("\t")
    assert found == [TYPES["types"]["result"], *TYPES["types"]["cells"]]


BACKGROUND = """Function Go(Optional ByVal how As Long = 0) As String
    Dim q As QueryTable
    Set q = ActiveSheet.QueryTables.Add("OLEDB;Provider=X;Data Source=C:\\data\\rows.accdb", ActiveSheet.Range("A1"), "SELECT * FROM T WHERE Id <= 2")
    If how = 0 Then q.BackgroundQuery = True: q.Refresh
    If how = 1 Then q.Refresh True
    If how = 2 Then q.BackgroundQuery = True: q.Refresh False
    Go = TypeName(ActiveSheet.Range("A2").Value) & " " & q.Refreshing
End Function
Function Read() As String
    Read = ActiveSheet.Range("A2").Value & " " & ActiveSheet.QueryTables(1).Refreshing
End Function
"""


@pytest.mark.parametrize(("how", "during"), [(0, "Empty True"), (1, "Empty True"), (2, "Double False")])
def test_a_background_refresh_waits_for_the_macro(how: int, during: str) -> None:
    """Measured: in the background the sheet stays as it was, Refreshing True, until no macro runs (tests/fixtures/query_refresh.json)."""
    app = _app()
    app.add_module(BACKGROUND, name="M")
    assert app.run("Go", how) == during
    assert app.run("Read") == "1 False"


def test_a_failed_background_refresh_is_raised_after_the_run() -> None:
    calls: list[str] = []

    def failing(request: HookRequest) -> list[list[object]]:
        calls.append(request.member)
        raise OSError("the source is gone")

    app = _app(failing)
    app.add_module(BACKGROUND + """Function Two() As String
    Go
    Go
    Two = "ran"
End Function
""", name="M")
    with pytest.raises(BackgroundRefreshError) as raised:
        app.run("Two")
    assert len(raised.value.failures) == 2 and calls == ["Refresh", "Refresh"]
    assert isinstance(raised.value.__cause__, VBARuntimeError)


def test_a_run_that_fails_carries_its_background_failures() -> None:
    app = _app(failing_with(OSError("gone")))
    app.add_module(BACKGROUND + """Sub Broken()
    Go
    Err.Raise 5
End Sub
""", name="M")
    with pytest.raises(VBARuntimeError) as raised:
        app.run("Broken")
    assert raised.value.number == 5
    assert len(getattr(raised.value, "refresh_failures")) == 1


def test_no_hook_reports_itself() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(BACKGROUND, name="M")
    with pytest.raises(VBAUnsupportedError, match="set_hook"):
        app.run("Go", 2)


@pytest.mark.parametrize("bad", ["rows", [], [[1, 2]], [["a", "b"], [1]], [["a"], [object()]]])
def test_what_the_function_returns_is_checked(bad: object) -> None:
    app = _app(returning(bad))
    app.add_module(BACKGROUND, name="M")
    with pytest.raises(VBARuntimeError) as raised:
        app.run("Go", 2)
    assert raised.value.number == 1004


SECRET = 'Provider=Microsoft.ACE.OLEDB.16.0;Password="";User ID=Admin;Data Source=C:\\x.accdb;Jet OLEDB:Database Password=Secret1;Mode=Share Deny None'


def test_passwords_are_masked_and_dropped_as_excel_drops_them() -> None:
    """The string is the one Excel wrote back with SavePassword True (tests/fixtures/query_refresh.json)."""
    assert "Secret1" not in masked(SECRET) and "Jet OLEDB:Database Password=****" in masked(SECRET)
    assert without_passwords(SECRET) == 'Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=C:\\x.accdb;Mode=Share Deny None'
    measured = TYPES["saved"]["SavePassword=True"]["read"].partition("|")[0]
    assert "Secret1" in measured and "Secret1" not in masked(measured)


def test_a_failing_function_does_not_echo_its_password() -> None:
    def echoing(request: HookRequest) -> list[list[object]]:
        assert request.source is not None
        raise OSError(f"cannot open {request.source.connection}")

    app = _app(echoing)
    app.add_module(BACKGROUND.replace("Data Source=", "Pwd=hunter2;Data Source="), name="M")
    with pytest.raises(VBARuntimeError) as raised:
        app.run("Go", 2)
    assert "hunter2" not in str(raised.value)
    assert "hunter2" in str(raised.value.__cause__)


def test_the_request_hides_its_password() -> None:
    seen: list[HookRequest] = []

    def keeping(request: HookRequest) -> object:
        seen.append(request)
        return [["Id"], [1.0]]

    app = _app(keeping)
    app.add_module(BACKGROUND.replace("Data Source=", "Pwd=hunter2;Data Source="), name="M")
    app.run("Go", 2)
    assert "hunter2" in seen[0].source.connection  # type: ignore[union-attr]
    assert "hunter2" not in repr(seen[0].source)


SPENDING = BACKGROUND + """Function Spend() As Long
    Dim i As Long
    For i = 1 To 100
        Spend = Spend + 1
    Next
End Function
Function SpendTwice() As Long
    SpendTwice = Spend() + Spend()
End Function
Function GoAndSpend() As String
    GoAndSpend = Go() & " " & Spend()
End Function
"""


def test_a_deferred_refresh_is_outside_the_run_step_budget() -> None:
    """William's ruling: the refresh that waits for the macro is not the macro's, and VBA the host's function runs is a new entry with a
    budget of its own (docs/vba_runtime.md, Refresh)."""
    app = ExcelApplication(sandbox=False, step_limit=400)
    app.add_workbook()
    app.add_module(SPENDING, name="M")
    with pytest.raises(VBAStepLimitError):
        app.run("SpendTwice")
    spent: list[object] = []

    def spending(request: HookRequest) -> object:
        spent.append(app.run("Spend"))
        return numbered(request)

    app.set_hook("Refresh", spending)
    assert app.run("GoAndSpend") == "Empty True 100"
    assert spent == [100]
    assert app.run("Read") == "1 False"
