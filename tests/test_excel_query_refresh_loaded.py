"""Query tables over a database in books Excel saved, opened again: what they read back, how they refresh, and what a second save
writes, against what Excel did (scripts/measure_query_refresh_loaded.py, tests/fixtures/query_refresh_loaded.json).

The books are Excel's own saves, with the run's folder written C:\\data. The function here stands for the Access file and the CSV file
beside it, which hold the same six numbered rows.
"""

from __future__ import annotations

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

HERE = Path(__file__).parent / "fixtures"
MEASURED: dict[str, Any] = json.loads((HERE / "query_refresh_loaded.json").read_text(encoding="utf-8"))["cases"]
BOOKS = HERE / "query_refresh_loaded"


def _script(name: str) -> ModuleType:
    folder = Path(__file__).parents[1] / "scripts"
    sys.path.insert(0, str(folder))
    spec = importlib.util.spec_from_file_location(name, folder / f"{name}.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MEASURE = _script("measure_query_refresh_loaded")


def numbered(request: HookRequest) -> list[list[object]]:
    assert request.source is not None
    wanted = int(re.search(r"Id <= (\d+)", request.source.command_text).group(1))  # type: ignore[union-attr]
    return [["Id", "Name"], *[[float(at), f"n{at}"] for at in range(1, 7) if at <= wanted]]


def _probe(case: str, out: Path) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook("Refresh", numbered)
    app.add_module(MEASURE.PROBE, name="M")
    kind = case.split("_")[1]
    table = "T" if case.startswith("access") else "[rows.csv]"
    return str(app.run("Probe", str((BOOKS / f"{case}.xlsx").resolve()), str(out), kind, table)).split("\t")


def _canonical(text: str) -> str:
    """A part without Excel's random ids and the namespaces only they need."""
    text = re.sub(r' (?:xr\d*|xr16):uid="\{[0-9A-F-]{36}\}"', "", text)
    text = re.sub(r' mc:Ignorable="[^"]*"| xmlns:(?:mc|xr\d*)="[^"]*"', "", text)
    return text.replace("\r\n", "")


QUERIES = sorted(case for case in MEASURED if case.endswith("_query"))
TABLES = sorted(case for case in MEASURED if case.endswith("_table"))


@pytest.mark.parametrize("case", QUERIES)
def test_a_loaded_query_table_reads_and_refreshes_as_excel_did(case: str, tmp_path: Path) -> None:
    assert _probe(case, tmp_path / "again.xlsx") == MEASURED[case]["probe"]


@pytest.mark.parametrize("case", sorted(MEASURED))
def test_a_loaded_query_table_saves_again_as_excel_did(case: str, tmp_path: Path) -> None:
    """The connection, the query table's part and a ListObject's part say what Excel's second save says; the uids of parts the model
    writes again are the named difference."""
    out = tmp_path / "again.xlsx"
    _probe(case, out)
    with zipfile.ZipFile(out) as package:
        mine = {name: package.read(name).decode("utf-8") for name in package.namelist()}
    for name in MEASURED[case]["again"]:
        if name.endswith((".xml",)) and name != "xl/workbook.xml":
            assert _canonical(mine[name]) == _canonical(MEASURED[case]["again"][name]), name
    wanted = re.search(r"<definedName [^>]*>[^<]*</definedName>", MEASURED[case]["again"]["xl/workbook.xml"])
    assert wanted is not None and wanted.group(0) in mine["xl/workbook.xml"]


@pytest.mark.parametrize("case", sorted(MEASURED))
def test_an_untouched_loaded_query_table_saves_excel_parts_as_they_came(case: str, tmp_path: Path) -> None:
    out = tmp_path / "same.xlsx"
    app = ExcelApplication(sandbox=False)
    book = app.open_workbook(BOOKS / f"{case}.xlsx")
    sheet = book.vba_get("Worksheets", [1])
    if case.endswith("_table"):
        assert str(sheet.vba_get("ListObjects", [1]).vba_get("QueryTable", []).vba_get("CommandText", [])).startswith("SELECT")  # type: ignore[attr-defined]
    else:
        assert int(str(sheet.vba_get("QueryTables", []).vba_get("Count", []))) == 1  # type: ignore[attr-defined]
    book.vba_call("SaveAs", [str(out), 51])
    with zipfile.ZipFile(out) as package:
        for name in MEASURED[case]["saved"]:
            if "connections" in name or "queryTable" in name or "tables/" in name:
                assert package.read(name).decode("utf-8") == MEASURED[case]["saved"][name], name


CHECKED: dict[str, list[str]] = json.loads((HERE / "query_refresh_loaded_saved.json").read_text(encoding="utf-8"))["cases"]
SAVED = _script("measure_query_refresh_loaded_saved")


@pytest.mark.parametrize("case", sorted(CHECKED))
def test_excel_opens_and_refreshes_what_the_model_saved(case: str, tmp_path: Path) -> None:
    """Excel opened each book the model refreshed and saved, read it and refreshed it against the real files
    (scripts/measure_query_refresh_loaded_saved.py); the model reads its own save and refreshes it the same way."""
    assert all(" E0 " in line for line in CHECKED[case] if line.startswith("three"))
    kind, source = case.split("_")[1], ("T" if case.startswith("access") else "[rows.csv]")
    out = tmp_path / "model.xlsx"
    app = ExcelApplication(sandbox=False)
    app.open_workbook(BOOKS / f"{case}.xlsx")
    app.set_hook("Refresh", numbered)
    app.add_module(SAVED.MODEL, name="M")
    app.run("Five", kind, source)
    app.workbook.vba_call("SaveAs", [str(out), 51])
    again = ExcelApplication(sandbox=False)
    again.add_workbook()
    again.set_hook("Refresh", numbered)
    again.add_module(SAVED.CHECK, name="M")
    assert str(again.run("Check", str(out), kind, source)).split("\t") == CHECKED[case]


@pytest.mark.parametrize("case", TABLES)
def test_a_loaded_list_query_table_reads_and_refreshes_as_excel_did(case: str, tmp_path: Path) -> None:
    assert _probe(case, tmp_path / "again.xlsx") == MEASURED[case]["probe"]
