"""ListObjects.Add over a database, refreshed through the host's Refresh function, against what Excel did with an Access file through
its own ACE OLEDB provider (scripts/measure_list_query_add.py, tests/fixtures/list_query_add.json).

The function here stands for the Access file: table T of six numbered rows, the columns a SELECT names, or the whole table for its
name. ACE runs SQL text as SQL even under xlCmdTable, as Excel's refresh showed (table_then_sql), so the function does too.

One difference is named and kept: once a refresh has run, Excel reads back and saves the provider's fuller connection string, where the
model keeps the one the macro set. The expected answers and parts have Excel's string put back to the macro's, and nothing else.
"""

from __future__ import annotations

import html
import importlib.util
import json
import re
import sys
import zipfile
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest
from pyopenvba.exceptions import VBAUnsupportedError

HERE = Path(__file__).parent / "fixtures"
MEASURED: dict[str, Any] = json.loads((HERE / "list_query_add.json").read_text(encoding="utf-8"))
FILE = "C:\\data\\rows.accdb"
#: The connection string as the macro sets it, and as Excel reads it back after a refresh.
GIVEN = f"Provider={MEASURED['provider']};Data Source={FILE}"
_SAVED = re.search(r'<dbPr connection="([^"]*)"', MEASURED["saves"]["save"]["book"]["xl/connections.xml"])
assert _SAVED is not None
#: As Excel writes it in a part, and as it reads.
WRITTEN = _SAVED.group(1)
EXPANDED = html.unescape(WRITTEN)


def _script(name: str) -> ModuleType:
    folder = Path(__file__).parents[1] / "scripts"
    sys.path.insert(0, str(folder))
    spec = importlib.util.spec_from_file_location(name, folder / f"{name}.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MEASURE = _script("measure_list_query_add")
_FIELD = re.compile(r"^(Id \* 2 AS D|Id|Name)$")


def table_t(request: HookRequest) -> list[list[object]]:
    """The rows of T the command asks for."""
    assert request.source is not None
    command = request.source.command_text
    if command == "T":
        # A table's name gives the whole table; SQL text runs as SQL whatever CommandType says, as ACE runs it (table_then_sql).
        assert request.source.command_type == "xlCmdTable"
        fields, last = ["Id", "Name"], 6
    else:
        found = re.fullmatch(r"SELECT (.+) FROM T WHERE Id <= (\d+) ORDER BY Id", command)
        assert found is not None, command
        fields = ["Id", "Name"] if found.group(1) == "*" else [one.strip() for one in found.group(1).split(",")]
        assert all(_FIELD.match(one) for one in fields), fields
        last = int(found.group(2))
    names = ["D" if one.startswith("Id *") else one for one in fields]
    rows: list[list[object]] = [list(names)]
    for at in range(1, last + 1):
        values = {"Id": float(at), "Name": f"n{at}", "D": float(at * 2)}
        rows.append([values[one] for one in names])
    return rows


def _expected(case: str) -> list[str]:
    return [part.replace(EXPANDED, GIVEN) for part in MEASURED["rows"][case]]


def _run(case: str, out: Path) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook("Refresh", table_t)
    code = MEASURE.CASE.replace("{provider}", MEASURED["provider"]).replace(
        "{steps}", "\n".join(f"    {step}" for step in MEASURED["cases"][case]))
    app.add_module(code, name="M")
    return str(app.run("Go", FILE, str(out))).split("\t")


def test_excel_read_its_string_back_fuller_after_a_refresh() -> None:
    """The one named difference is where the measurements say it is: after a refresh, not before."""
    assert EXPANDED.startswith(f"Provider={MEASURED['provider']};User ID=Admin;Data Source={FILE};")
    assert f"Connection=String:OLEDB;{GIVEN} " in MEASURED["rows"]["add"][2]
    assert f"Connection=String:OLEDB;{EXPANDED} " in MEASURED["rows"]["add"][11]


#: Measured and refused: a table added over cells in use, which Excel moves by a rule one case does not settle (the block B2:C4 moved
#: right whole, its third row too, where a lone cell beside the two placeholder cells stays).
NOT_BUILT = {"over_cells"}


@pytest.mark.parametrize("case", sorted(set(MEASURED["rows"]) - NOT_BUILT))
def test_a_list_object_over_a_database_does_what_excel_did(case: str, tmp_path: Path) -> None:
    assert _run(case, tmp_path / f"{case}~") == _expected(case)


@pytest.mark.parametrize("case", sorted(NOT_BUILT))
def test_what_is_not_built_reports_itself(case: str, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match="over cells in use"):
        _run(case, tmp_path / f"{case}~")


def _canonical(text: str) -> str:
    """A part without Excel's random ids, the namespaces only they need, and Excel's fuller connection string."""
    text = re.sub(r' (?:xr\d*|xr16|xr3):uid="\{[0-9A-F-]{36}\}"', "", text)
    text = re.sub(r' mc:Ignorable="[^"]*"| xmlns:(?:mc|xr\d*)="[^"]*"', "", text)
    return text.replace("\r\n", "").replace(WRITTEN, GIVEN)


SAVES = sorted((case, copy) for case, copies in MEASURED["saves"].items() for copy in copies)


@pytest.mark.parametrize(("case", "copy"), SAVES)
def test_a_list_object_over_a_database_saves_as_excel_saved_it(case: str, copy: str, tmp_path: Path) -> None:
    """The connection, the table, its query table and the name of its block, as Excel wrote them."""
    _run(case, tmp_path / f"{case}~")
    path = tmp_path / f"{case}~{'' if copy == 'book' else copy}.xlsx"
    with zipfile.ZipFile(path) as package:
        mine = {name: package.read(name).decode("utf-8") for name in package.namelist()}
    wanted = MEASURED["saves"][case][copy]
    for name in ("xl/connections.xml", "xl/tables/table1.xml", "xl/queryTables/queryTable1.xml", "xl/tables/_rels/table1.xml.rels"):
        assert _canonical(mine[name]) == _canonical(wanted[name]), name
    named = re.search(r"<definedNames>.*</definedNames>", wanted["xl/workbook.xml"])
    assert named is not None and named.group(0) in mine["xl/workbook.xml"]


SAVED = _script("measure_list_query_add_saved")
CHECKED: dict[str, Any] = json.loads((HERE / "list_query_add_saved.json").read_text(encoding="utf-8"))


def _saved_book(book: str, folder: Path) -> Path:
    case, copy, _ = CHECKED["books"][book]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook("Refresh", table_t)
    code = MEASURE.CASE.replace("{provider}", MEASURED["provider"]).replace("{steps}", "\n".join(f"    {step}" for step in MEASURED["cases"][case]))
    app.add_module(code, name="M")
    app.run("Go", FILE, str(folder / f"{book}~"))
    return folder / f"{book}~{copy}.xlsx"


def _checked(book: str, path: Path) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook("Refresh", table_t)
    app.add_module(SAVED.CHECK, name="M")
    return str(app.run("Check", str(path), CHECKED["books"][book][2])).split("\t")


#: A book saved before its first refresh opens in the model, and its refresh, which gives the table its columns, is refused there.
UNREFRESHED = {"save_unrefreshed"}


@pytest.mark.parametrize("book", sorted(set(CHECKED["cases"]) - UNREFRESHED))
def test_excel_opens_and_refreshes_what_the_model_saved(book: str, tmp_path: Path) -> None:
    """Excel opened each book the model saved and refreshed it against the real file (scripts/measure_list_query_add_saved.py); the
    model reads its own save and refreshes it the same way."""
    assert all(" E0 " in line for line in CHECKED["cases"][book] if line.startswith("four"))
    expected = [part.replace(EXPANDED, GIVEN) for part in CHECKED["cases"][book]]
    assert _checked(book, _saved_book(book, tmp_path)) == expected


@pytest.mark.parametrize("case", sorted(CHECKED["dropped"]))
def test_a_dropped_list_object_takes_its_query_table_out_of_the_file(case: str, tmp_path: Path) -> None:
    """Excel's own book with a ListObject over a database, the table deleted or unlisted and saved: the file keeps no connection, no
    query table and no name of its block (scripts/measure_list_query_add_saved.py)."""
    book, how = SAVED.LOADED[case]
    given, out = tmp_path / f"{case}.xlsx", tmp_path / f"{case}_saved.xlsx"
    given.write_bytes((SAVED.LOADED_BOOKS / f"{book}.xlsx").read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SAVED.DROP, name="M")
    measured = CHECKED["dropped"][case]
    assert str(app.run("Drop", str(given), str(out), how)).split("\t") == measured["rows"]
    assert SAVED.kept(out.read_bytes()) == measured["saved"]


@pytest.mark.parametrize("book", sorted(UNREFRESHED))
def test_excel_refreshes_a_saved_placeholder_the_model_refuses(book: str, tmp_path: Path) -> None:
    assert any(line.startswith("four E0 ") for line in CHECKED["cases"][book])
    with pytest.raises(VBAUnsupportedError, match="a file gave"):
        _checked(book, _saved_book(book, tmp_path))
