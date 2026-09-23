"""Windows and sheet views, replayed against the in-memory model.

tests/fixtures/windows/ is what scripts/measure_windows.py saw in live
Excel: for each workbook, the VBA that set its window up -- cells
selected, panes frozen and split, the window scrolled and zoomed,
gridlines and headings turned off, sheets added, copied and moved --
what the active window answered, and the sheetViews and bookViews Excel
saved; and which window Windows(1), (2) and (3) was as workbooks were
added, activated and closed.

The model runs the same VBA, answers the same and saves the same views.
Two things are the screen's rather than the workbook's, and are left out
of the comparison: the Caption a workbook Excel named SheetN has, where
the model names it BookN, and the window's position and size, which
workbookView keeps in xWindow, yWindow, windowWidth and windowHeight
beside a random xr2:uid. Each workbook Excel saved reads back to the
same answers, and its views are written back byte for byte.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "windows"
RECORD: dict[str, Any] = json.loads((FIXTURES / "windows.json").read_text(encoding="utf-8"))
BOOKS: list[dict[str, Any]] = RECORD["books"]
READS: list[str] = RECORD["reads"]
_SCREEN = re.compile(r' (?:xWindow|yWindow|windowWidth|windowHeight|xr2:uid)="[^"]*"')

#: Where the model's window answers differently, and why.
KNOWN: dict[str, str] = {
    "frozen_zoomed_399": "at 399% the active cell is out of a 700 by 500 window, and Excel froze the middle of it",
}
#: What the model reports as not implemented, and why.
UNSUPPORTED: dict[str, str] = {
    "freeze_out_of_view": "Excel freezes the middle of the window when the active cell is out of view",
    "split_true": "Excel splits the window where the active cell is, and keeps the split in twips",
}
#: What the model answers for and does not save: a split that is not frozen, which Excel keeps in twips.
UNSAVED = ("split_rows", "split_columns", "recorded_unfrozen", "frozen_split_row_moved", "frozen_split_off")
#: The Excel-saved workbooks whose split, in twips, the model does not count in rows and columns.
TWIPS = (*UNSAVED, "split_true")
#: Reads the screen answers rather than the workbook: Caption and WindowState.
_SCREEN_READS = {READS.index("w.Caption"), READS.index("w.WindowState")}


def _probe(book: dict[str, Any]) -> str:
    """The measurement's workbook function, run in the model: set the window up, then read it."""
    return RECORD["reader"] + "\n".join([
        "Public Function Probe() As String",
        "Dim wb As Object, ws As Object, other As Object, failed As Long",
        "Application.DisplayAlerts = False", "Application.ScreenUpdating = False",
        "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", "wb.Activate",
        "ActiveWindow.WindowState = xlNormal", "ActiveWindow.Width = 700", "ActiveWindow.Height = 500",
        "On Error Resume Next", book["making"], "failed = Err.Number", "On Error GoTo 0",
        'Probe = failed & "|" & Seen(ActiveWindow)', "End Function"]) + "\n"


def _answers(seen: str, *, skip: set[int]) -> list[str]:
    """The error the making raised and each read, the screen's reads left out."""
    failed, *reads = seen.split("|")
    return [failed, *(value for index, value in enumerate(reads) if index not in skip)]


def _skipped(book: dict[str, Any]) -> set[int]:
    """Caption is the model's own name for the workbook unless the VBA set it."""
    return set() if book["name"].startswith("caption") else {READS.index("w.Caption")}


def _views(path: Path) -> dict[str, str]:
    """Each sheet's sheetViews element and the workbook's bookViews, as saved."""
    out: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            if name.startswith("xl/worksheets/sheet"):
                found = re.search(r"<sheetViews>.*?</sheetViews>", package.read(name).decode("utf-8"), re.DOTALL)
                out[name] = found.group(0) if found else ""
        found = re.search(r"<bookViews>.*?</bookViews>", package.read("xl/workbook.xml").decode("utf-8"), re.DOTALL)
        out["bookViews"] = found.group(0) if found else ""
    return out


def _param(book: dict[str, Any], known: dict[str, str], unsupported: dict[str, str]) -> Any:
    name = book["name"]
    marks: list[Any] = []
    if name in known:
        marks.append(pytest.mark.xfail(reason=known[name], strict=True))
    if name in unsupported:
        marks.append(pytest.mark.xfail(reason=unsupported[name], raises=VBAUnsupportedError, strict=True))
    return pytest.param(book, marks=marks, id=name)


def _made(book: dict[str, Any]) -> tuple[ExcelApplication, str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_probe(book), name="Probe")
    return app, str(app.run("Probe"))


@pytest.mark.parametrize("book", [_param(book, KNOWN, UNSUPPORTED) for book in BOOKS])
def test_the_window_answers_as_excel_did(book: dict[str, Any]) -> None:
    _, seen = _made(book)
    skip = _skipped(book)
    assert _answers(seen, skip=skip) == _answers(book["seen"], skip=skip)


@pytest.mark.parametrize("book", [_param(book, KNOWN, {**UNSUPPORTED, **{name: "a split that is not frozen is kept in "
                                                                          "twips of the window" for name in UNSAVED}})
                                  for book in BOOKS])
def test_the_model_saves_the_views_excel_saved(book: dict[str, Any], tmp_path: Path) -> None:
    app, _ = _made(book)
    saved = app.save(tmp_path / f"{book['name']}.xlsx", workbook=app.workbook)
    views = {part: _SCREEN.sub("", xml) for part, xml in _views(saved).items()}
    assert views == {part: _SCREEN.sub("", xml) for part, xml in book["views"].items()}


def _reader(skip: set[int]) -> str:
    """The measurement's reads, less the ones in ``skip``, each answering its value or the error it raised."""
    kept = [index for index in range(len(READS)) if index not in skip]
    functions = [f'Private Function R{index}(w As Object) As String\nOn Error GoTo Bad\nR{index} = CStr({READS[index]})\n'
                 f'Exit Function\nBad:\nR{index} = "!" & Err.Number\nEnd Function\n' for index in kept]
    seen = ' & "|" & '.join(f"R{index}(ActiveWindow)" for index in kept)
    return "".join(functions) + f'Public Function Probe() As String\nProbe = "0|" & {seen}\nEnd Function\n'


@pytest.mark.parametrize("book", [_param(book, {}, {name: "a split kept in twips is not counted in rows and columns"
                                                    for name in TWIPS}) for book in BOOKS])
def test_a_workbook_excel_saved_reads_back_the_same(book: dict[str, Any]) -> None:
    app = ExcelApplication.open(FIXTURES / f"{book['name']}.xlsx", with_vba=False)
    skip = _skipped(book) | _SCREEN_READS
    app.add_module(_reader(skip), name="Probe")
    assert _answers(str(app.run("Probe")), skip=set())[1:] == _answers(book["seen"], skip=skip)[1:]


@pytest.mark.parametrize("book", [pytest.param(book, id=book["name"]) for book in BOOKS])
def test_a_view_read_from_a_file_is_written_back_byte_for_byte(book: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(FIXTURES / f"{book['name']}.xlsx", with_vba=False)
    workbook = app.workbook
    for sheet in workbook.sheets_:
        sheet.view.changed = True
    workbook.view.changed = True
    saved = app.save(tmp_path / f"{book['name']}.xlsx")
    assert _views(saved) == _views(FIXTURES / f"{book['name']}.xlsx")


@pytest.mark.parametrize("updating", sorted(RECORD["order"]))
def test_windows_are_listed_front_to_back(updating: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    probe = RECORD["order_probe"] + "\n".join([
        "Public Function Probe() As String",
        f"Probe = Order({updating == 'screen_updating_on'})", "End Function"]) + "\n"
    app.add_module(probe, name="Probe")
    assert str(app.run("Probe")).split("~;~") == RECORD["order"][updating]
