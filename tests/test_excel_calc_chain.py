"""A calculation chain kept while Excel can open it and dropped when it cannot (scripts/measure_calc_chain.py)."""

from __future__ import annotations

import importlib.util
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._xml import attributes
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._calc_chain import holds
from pyopenvba.apps.excel._io import WorkbookFileError
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
BOOK = FIXTURES / "calc_chain.xlsx"
PLAIN = FIXTURES / "calc_chain_plain.xlsx"
MEASURED = json.loads((FIXTURES / "calc_chain.json").read_text(encoding="utf-8"))
CHAIN = "xl/calcChain.xml"
RENAMES = ("rename a sheet", "move and rename a sheet with formulas", "rename a sheet and clear a formula")
SAVED = {name: row for name, row in MEASURED["edits"].items() if "model_reads" in row}
#: Where Excel reads the model's save differently, for reasons that are not the chain: the chain kept or dropped, Excel's
#: own save reads the same without one (test_a_missing_chain_changes_nothing).
READ_GAPS = {
    "clear all of formula": "a formula on another sheet keeps its old value in the file",
    "formula over a value": "a formula on another sheet keeps its old value in the file",
    "copy over a formula": "a formula on another sheet keeps its old value in the file",
    "cut a formula": "a formula on another sheet does not follow a Cut",
    "sort formulas": "Sort of relative formulas gives #REF!, and a formula on another sheet keeps its old value",
    "add a sheet": "Excel names the new sheet Sheet1, the model Sheet6",
    "delete then add a sheet": "Excel names the new sheet Sheet1, the model Sheet6",
}


def _book(name: str) -> Path:
    return PLAIN if name in RENAMES else BOOK


def _edited(source: Path, target: Path, edit: str) -> None:
    app = ExcelApplication.open(source, with_vba=False)
    app.add_module(f"Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1)\n    {edit}\nEnd Sub\n", name="E")
    app.run("E.Edit")
    app.save(target)


def _parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def _sheets(parts: dict[str, bytes]) -> dict[int, str]:
    """Each sheetId's xml, through the workbook's relationships."""
    workbook = parts["xl/workbook.xml"].decode("utf-8")
    relationships = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    targets: dict[str, str] = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", relationships):
        found = attributes(tag)
        targets[found["Id"]] = found["Target"]
    out: dict[int, str] = {}
    for tag in re.findall(r"<sheet\b[^>]*/>", workbook):
        found = attributes(tag)
        out[int(found["sheetId"])] = parts["xl/" + targets[found["r:id"]]].decode("utf-8")
    return out


@pytest.mark.parametrize("name", SAVED)
def test_the_chain_stays_exactly_when_excel_opens_it(name: str, tmp_path: Path) -> None:
    row: dict[str, Any] = SAVED[name]
    target = tmp_path / "saved.xlsx"
    _edited(_book(name), target, row["edit"])
    parts = _parts(target)
    original = _parts(_book(name))
    if row["opens_with_old_chain"]:
        assert parts[CHAIN] == original[CHAIN]
        assert b"calcChain" in parts["[Content_Types].xml"]
        assert b"calcChain" in parts["xl/_rels/workbook.xml.rels"]
    else:
        assert CHAIN not in parts
        assert b"calcChain" not in parts["[Content_Types].xml"]
        assert b"calcChain" not in parts["xl/_rels/workbook.xml.rels"]
    assert b"fullCalcOnLoad" not in parts["xl/workbook.xml"]


def test_each_edit_was_decided_by_excel() -> None:
    assert sum(1 for row in SAVED.values() if not row["opens_with_old_chain"]) >= 20
    assert sum(1 for row in SAVED.values() if row["opens_with_old_chain"]) >= 10
    assert all(row["chain_kept"] == row["opens_with_old_chain"] for row in SAVED.values())
    assert not any(row["full_calc_on_load"] for row in SAVED.values())


@pytest.mark.parametrize("name", [name for name, row in MEASURED["edits"].items() if "excel_reads" in row])
def test_a_missing_chain_changes_nothing(name: str) -> None:
    """Excel reads its own save the same with its chain taken out, on opening and after a value changes."""
    row = MEASURED["edits"][name]
    assert row["excel_reads_without_chain"] == row["excel_reads"]


@pytest.mark.parametrize("name", [
    pytest.param(name, marks=pytest.mark.xfail(strict=True, reason=READ_GAPS[name])) if name in READ_GAPS else name
    for name in SAVED if "excel_reads" in SAVED[name]])
def test_excel_read_the_same_from_the_models_save(name: str) -> None:
    assert SAVED[name]["model_reads"] == SAVED[name]["excel_reads"]


@pytest.mark.parametrize("name", list(MEASURED["hand"]))
def test_a_chain_holds_as_excel_opens_it(name: str) -> None:
    """Sheets, inheritance of i, arrays, spills, data tables and flags, on chains written by hand."""
    body = MEASURED["hand"][name]
    chain = f'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{body}</calcChain>'
    assert holds(chain, _sheets(_parts(BOOK))) == MEASURED["hand_opens"][name]


def test_every_chain_excel_saved_holds() -> None:
    """Nothing is dropped from a file only because it was loaded, in any fixture Excel saved with a chain."""
    seen = 0
    for path in FIXTURES.glob("*.xls[xm]"):
        with zipfile.ZipFile(path) as package:
            if CHAIN not in package.namelist():
                continue
            parts = {name: package.read(name) for name in package.namelist() if name.endswith((".xml", ".rels"))}
        seen += 1
        assert holds(parts[CHAIN].decode("utf-8"), _sheets(parts)), path.name
    assert seen >= 2


def test_a_saved_book_keeps_its_chain_gone_on_the_next_save(tmp_path: Path) -> None:
    first, second = tmp_path / "first.xlsx", tmp_path / "second.xlsx"
    _edited(BOOK, first, 'ws.Range("B2").ClearContents')
    _edited(first, second, 'ws.Range("Z100").Value = 1')
    assert CHAIN not in _parts(second)


def test_the_rest_of_a_dropped_save_is_untouched(tmp_path: Path) -> None:
    """Only the part, its Override and its Relationship go."""
    kept, dropped = tmp_path / "kept.xlsx", tmp_path / "dropped.xlsx"
    _edited(BOOK, kept, 'ws.Range("B2").Formula = "=7"')
    _edited(BOOK, dropped, 'ws.Range("B2").Value = 7')
    assert set(_parts(kept)) - set(_parts(dropped)) == {CHAIN}
    assert len(re.findall(rb"<Relationship\b", _parts(kept)["xl/_rels/workbook.xml.rels"])) - 1 == len(
        re.findall(rb"<Relationship\b", _parts(dropped)["xl/_rels/workbook.xml.rels"]))


def test_a_binary_workbook_is_not_opened(tmp_path: Path) -> None:
    """An xlsb holds xl/calcChain.bin, which the model never meets, as it does not open the format's cells."""
    assert MEASURED["formats"][".xlsb"]["chain_parts"] == ["xl/calcChain.bin"]
    path = tmp_path / "book.xlsb"
    path.write_bytes(b"PK")
    with pytest.raises(WorkbookFileError):
        ExcelApplication.open(path, with_vba=False)


def test_each_format_opened_after_the_chain_went() -> None:
    """A template without macros fails even after an edit the chain has nothing to do with: not the chain's doing."""
    formats: dict[str, Any] = MEASURED["formats"]
    for suffix in (".xlsm", ".xltm", ".xlam"):
        assert formats[suffix]["chain_in_excels_file"] and formats[suffix]["opened"] == "0"
    assert formats[".xltx"]["opened"] == formats[".xltx"]["opened_after_unrelated_edit"] == "1004"


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None
OPENER = """Public Function Tryit(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    Tryit = CStr(Err.Number)
    If Not wb Is Nothing Then wb.Close False
    Application.DisplayAlerts = True
End Function
"""


@pytest.mark.skipif(not LIVE, reason="requires explicitly enabled native Excel")
def test_native_excel_opens_the_models_save_after_every_edit(tmp_path: Path) -> None:
    harness: Any = pytest.importorskip("pyvbaharness")
    refused: list[tuple[str, object]] = []
    with harness.ExcelSession() as excel:
        excel.new_document()
        for number, (name, row) in enumerate(MEASURED["edits"].items()):
            target = tmp_path / f"saved{number}.xlsx"
            try:
                _edited(_book(name), target, row["edit"])
            except (VBAUnsupportedError, VBARuntimeError):
                assert "model_refuses" in row, name
                continue
            answer = excel.run_vba(OPENER, "Tryit", (str(target),), timeout=60.0, line_numbers=False)
            if answer.value != "0":
                refused.append((name, answer.value))
    assert not refused
