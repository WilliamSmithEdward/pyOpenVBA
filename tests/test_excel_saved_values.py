"""A formula on a sheet nobody wrote to holds its current value in the file (scripts/measure_saved_values.py)."""

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

FIXTURES = Path(__file__).parent / "fixtures"
BOOK = FIXTURES / "saved_values.xlsx"
MEASURED: dict[str, Any] = json.loads((FIXTURES / "saved_values.json").read_text(encoding="utf-8"))
SAVED: dict[str, dict[str, Any]] = {name: row for name, row in MEASURED["edits"].items() if "model_reads" in row}
_CELL = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.DOTALL)
#: The cells of the sheet Vol holding random and clock values, which no two saves share.
VOLATILE = {"A1", "B1", "E1"}


def _saved(edit: str, target: Path) -> Any:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(f'Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("One")\n    {edit}\nEnd Sub\n', name="E")
    app.run("E.Edit")
    app.save(target)
    return app


def _sheet_parts(package: zipfile.ZipFile) -> dict[str, str]:
    targets: dict[str, str] = {}
    for tag in re.findall(r"<Relationship\b[^>]*/>", package.read("xl/_rels/workbook.xml.rels").decode("utf-8")):
        found = attributes(tag)
        targets[found["Id"]] = found["Target"]
    out: dict[str, str] = {}
    for tag in re.findall(r"<sheet\b[^>]*/>", package.read("xl/workbook.xml").decode("utf-8")):
        found = attributes(tag)
        out[found["name"]] = "xl/" + targets[found["r:id"]]
    return out


def _same(excel: str, model: str) -> bool:
    """Two values as a file holds them: the same t, f attributes and value, a number by what it is worth, since the
    model spells some numbers with fewer digits than Excel does, 0.3333333333333333 for 0.33333333333333331."""
    one, two = excel.split("|"), model.split("|")
    if one[0] != two[0] or one[2] != two[2]:
        return False
    if one[0]:
        return one[1] == two[1]
    return (one[1] == two[1]) or (bool(one[1]) and bool(two[1]) and float(one[1]) == float(two[1]))


@pytest.mark.parametrize("name", [name for name in SAVED if "excel_reads" in SAVED[name]])
def test_excel_reads_the_models_save_as_its_own(name: str) -> None:
    """On opening, and after a value changes."""
    assert SAVED[name]["model_reads"] == SAVED[name]["excel_reads"]


@pytest.mark.parametrize("name", [name for name in SAVED if "excel_cells" in SAVED[name]])
def test_the_file_holds_what_excels_file_holds(name: str) -> None:
    """Every sheet's formulas, but the random and clock values, and the one mark a rewritten sheet does not make yet."""
    excel, model = SAVED[name]["excel_cells"], SAVED[name]["model_cells"]
    assert set(excel) == set(model)
    assert model["#charts"] == excel["#charts"]
    # Attribute by attribute: in manual calculation the model writes calcOnSave and calcCompleted the other way round.
    assert attributes(model["#calcPr"]["tag"]) == attributes(excel["#calcPr"]["tag"])
    assert "fullCalcOnLoad" not in model["#calcPr"]["tag"]
    for sheet, cells in excel.items():
        if sheet.startswith("#"):
            continue
        assert set(cells) == set(model[sheet]), sheet
        for reference, held in cells.items():
            if sheet == "Vol" and reference in VOLATILE:
                continue
            if sheet == "Vol" and reference == "F1" and name == "insert a row":
                # A formula reading a volatile one: Excel marks it ca="1", and a sheet the model rewrites does not.
                continue
            assert _same(held, model[sheet][reference]), (sheet, reference, held, model[sheet][reference])


@pytest.mark.parametrize("name", list(SAVED))
def test_only_the_changed_cells_differ_on_a_sheet_nobody_wrote_to(name: str, tmp_path: Path) -> None:
    """Between the cells, and in every cell whose value did not change, a sheet nobody wrote to keeps its bytes."""
    target = tmp_path / "saved.xlsx"
    _saved(SAVED[name]["edit"], target)
    with zipfile.ZipFile(BOOK) as old, zipfile.ZipFile(target) as new:
        before, after = _sheet_parts(old), _sheet_parts(new)
        for sheet, part in before.items():
            if sheet not in after:
                continue
            first = old.read(part).decode("utf-8")
            second = new.read(after[sheet]).decode("utf-8")
            if sheet not in _written_sheets(name):
                assert _CELL.sub("<c/>", first) == _CELL.sub("<c/>", second), sheet
                for was, now in zip(_CELL.findall(first), _CELL.findall(second)):
                    if was != now:
                        assert _outside_value(was) == _outside_value(now), (sheet, was, now)


def _written_sheets(name: str) -> set[str]:
    """The sheets an edit wrote to, which a save writes whole."""
    edit = SAVED[name]["edit"]
    if name == "insert a row":
        return {"One", "Two", "Three", "Vol", "Link", "Untouched"}
    if name == "delete a sheet nothing reads":
        # The active sheet went, so the one before it is selected now.
        return {"Link"}
    written = {found for found in re.findall(r'Worksheets\("(\w+)"\)', edit)}
    if re.search(r"\bws\.", edit) or not written:
        written.add("One")
    return written


def _outside_value(element: str) -> str:
    """A cell without what a value's change may change: its t, its v, and ca on its f."""
    bare = re.sub(r'(<c\b[^>]*?)\st="[^"]*"', r"\1", element, count=1)
    bare = re.sub(r"<v\b[^>]*?(?:/>|>.*?</v>)", "", bare, flags=re.DOTALL)
    return bare.replace(' ca="1"', "")


def test_a_save_that_changes_no_value_keeps_every_sheet(tmp_path: Path) -> None:
    """All but the one with volatile formulas, which any write works out again, as it does in Excel."""
    target = tmp_path / "saved.xlsx"
    _saved('wb.Worksheets("Untouched").Range("Z9").Value = 1', target)
    with zipfile.ZipFile(BOOK) as old, zipfile.ZipFile(target) as new:
        before, after = _sheet_parts(old), _sheet_parts(new)
        for sheet in ("One", "Two", "Three", "Link"):
            assert old.read(before[sheet]) == new.read(after[sheet]), sheet


def test_a_value_that_changes_kind_changes_its_type(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    _saved('ws.Range("A1").Value = 2', target)
    with zipfile.ZipFile(target) as package:
        two = package.read(_sheet_parts(package)["Two"]).decode("utf-8")
    assert re.search(r'<c r="B3"[^>]*><f>[^<]*</f><v>1</v></c>', two)  # was t="e"
    assert re.search(r'<c r="B4"[^>]* t="str"[^>]*>', two)  # was a number
    assert re.search(r'<c r="B2"[^>]* t="b"[^>]*>', two)


def test_the_saved_value_is_what_the_session_read(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    _saved('ws.Range("A4").Value = 50', target)
    again = ExcelApplication.open(target, with_vba=False)
    again.add_module('Public Function Go() As String\n    Go = ActiveWorkbook.Worksheets("Two").Range("A1").Value & "," & ActiveWorkbook.Worksheets("Three").Range("B1").Value\nEnd Function\n', name="E")
    assert again.run("E.Go") == "56,113"


def test_a_second_save_keeps_the_value(tmp_path: Path) -> None:
    first, second = tmp_path / "first.xlsx", tmp_path / "second.xlsx"
    app = _saved('ws.Range("A4").Value = 50', first)
    app.save(second)
    with zipfile.ZipFile(first) as one, zipfile.ZipFile(second) as two:
        parts = _sheet_parts(one)
        assert one.read(parts["Three"]) == two.read(parts["Three"])


def _cells(path: Path) -> dict[str, dict[str, str]]:
    """Each sheet's formula cells as the probe records them."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, dict[str, str]] = {}
        for sheet, part in _sheet_parts(package).items():
            cells: dict[str, str] = {}
            for element in _CELL.findall(package.read(part).decode("utf-8")):
                head = element[:element.index(">")]
                formula = re.search(r"<f\b([^>]*?)/?>", element)
                if formula is None:
                    continue
                kind = re.search(r'\st="([^"]*)"', head)
                value = re.search(r"<v>(.*?)</v>", element, re.DOTALL)
                cells[attributes(head + ">")["r"]] = f"{kind.group(1) if kind else ''}|{value.group(1) if value else ''}|{formula.group(1).strip()}"
            out[sheet] = cells
    return out


def _ran(name: str, room: Path) -> Path:
    """The file a sequence of saves leaves, and its folders, as the probe made them."""
    (room / "isdir.xlsx").mkdir(parents=True)
    own = room / "book.xlsx"
    own.write_bytes(BOOK.read_bytes())
    app = ExcelApplication.open(own, with_vba=False, sandbox=False)
    edit = MEASURED["sequences"][name]["edit"].replace("\\", os.sep)
    app.add_module(f'Public Sub Run(ByVal out As String)\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("One")\n    {edit}\nEnd Sub\n', name="E")
    app.run("E.Run", str(room))
    return own


@pytest.mark.parametrize("name", list(MEASURED["sequences"]))
def test_saves_one_after_another_write_what_excel_writes(name: str, tmp_path: Path) -> None:
    """A copy, a failed save or an earlier save leaves the book's own Save writing the current values, and takes off a
    mark a save put on once the cell is worked out."""
    row = MEASURED["sequences"][name]
    assert row["model_reads"] == row["excel_reads"]
    excel, model = row["excel_cells"], _cells(_ran(name, tmp_path))
    for sheet, cells in excel.items():
        if sheet.startswith("#"):
            continue
        assert set(cells) == set(model[sheet]), sheet
        for reference, held in cells.items():
            if not (sheet == "Vol" and reference in VOLATILE):
                assert _same(held, model[sheet][reference]), (sheet, reference, held, model[sheet][reference])


def test_a_save_that_fails_midway_loses_no_value(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """After the values are written to the book's package and before the file is: the next save still has them."""
    from pyopenvba.apps.excel import _io

    own = tmp_path / "book.xlsx"
    own.write_bytes(BOOK.read_bytes())
    app = ExcelApplication.open(own, with_vba=False)
    app.add_module('Public Sub Edit()\n    ActiveWorkbook.Worksheets("One").Range("A4").Value = 50\nEnd Sub\n', name="E")
    app.run("E.Edit")

    def fails(*_: object) -> None:
        raise OSError("disk full")

    with monkeypatch.context() as patch:
        patch.setattr(_io, "_write_styles", fails)
        with pytest.raises(OSError):
            _io.save_workbook(app.workbook, tmp_path / "failed.xlsx")
    app.save(tmp_path / "after.xlsx")
    assert _cells(tmp_path / "after.xlsx")["Two"]["A1"].split("|")[1] == "56"


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None
READS = MEASURED["reader"] + """Public Function Reads(ByVal path As String) As String
    Dim wb As Workbook, first As String
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    first = CellsOf(wb)
    wb.Worksheets("One").Range("A3").Value = 100
    Reads = first & "@@" & CellsOf(wb)
    wb.Close False
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Function
"""


@pytest.mark.skipif(not LIVE, reason="requires explicitly enabled native Excel")
def test_native_excel_reads_the_models_save_as_recorded(tmp_path: Path) -> None:
    harness: Any = pytest.importorskip("pyvbaharness")
    wrong: list[str] = []
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number, (name, row) in enumerate(SAVED.items()):
            target = tmp_path / f"saved{number}.xlsx"
            _saved(row["edit"], target)
            answer = excel.run_vba(READS, "Reads", (str(target),), timeout=60.0, line_numbers=False)
            if answer.value != row["model_reads"]:
                wrong.append(name)
        for number, name in enumerate(MEASURED["sequences"]):
            room = tmp_path / f"sequence{number}"
            own = _ran(name, room)
            answer = excel.run_vba(READS, "Reads", (str(own),), timeout=60.0, line_numbers=False)
            if answer.value != MEASURED["sequences"][name]["excel_reads"]:
                wrong.append(name)
    assert not wrong
