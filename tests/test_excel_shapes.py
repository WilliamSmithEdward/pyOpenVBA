"""Shapes through the workbook: loaded, edited by a macro, saved.

The reader and writer are held to Excel's own file in
``test_shapes_xlsx.py``; this is the other half, that a workbook opened
here shows the same shapes to VBA, that a macro's edit lands in the
file, and that a workbook nobody touched keeps its drawing exactly.
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
WORKBOOK = FIXTURES / "excel_shapes.xlsm"


def opened(tmp_path: Path) -> ExcelApplication:
    copy = tmp_path / "shapes.xlsm"
    shutil.copyfile(WORKBOOK, copy)
    return ExcelApplication.open(copy)


def macro(app: ExcelApplication, body: str, name: str = "S") -> None:
    app.add_module(f"Sub {name}()\n{body}\nEnd Sub\n", name="ShapeModule")
    app.run(name)


def drawing_of(path: Path) -> str:
    return zipfile.ZipFile(path).read("xl/drawings/drawing1.xml").decode("utf-8")


def test_a_workbook_opens_with_its_shapes(tmp_path: Path) -> None:
    app = opened(tmp_path)
    sheet = app.workbook.sheets_[0]
    assert [one.name for one in sheet.shapes_] == [
        "Rect",
        "Rounded",
        "Oval",
        "Box",
        "Line1",
        "Button1",
        "Check1",
        "Drop1",
        "Group1",
    ]


def test_a_macro_sees_what_the_file_holds(tmp_path: Path) -> None:
    app = opened(tmp_path)
    assert app.evaluate("Worksheets(1).Shapes.Count") == 9
    assert app.evaluate('Worksheets(1).Shapes("Rect").OnAction') == "Clicked"
    assert app.evaluate('Worksheets(1).Shapes("Rect").TextFrame.Characters.Text') == "Hello"
    assert app.evaluate('Worksheets(1).Shapes("Button1").Type') == 8


def test_a_form_controls_macro_is_read_without_its_workbook(tmp_path: Path) -> None:
    """The file says [0]!Clicked; a macro asks for the procedure."""
    app = opened(tmp_path)
    assert app.evaluate('Worksheets(1).Shapes("Button1").OnAction') == "Clicked"


def test_a_missing_shape_raises_what_excel_raises(tmp_path: Path) -> None:
    app = opened(tmp_path)
    with pytest.raises(VBARuntimeError) as raised:
        app.evaluate('Worksheets(1).Shapes("NoSuchShape").Name')
    assert raised.value.number == -2147024809


def test_an_untouched_workbook_keeps_its_drawing(tmp_path: Path) -> None:
    """Opened and saved with no edit at all, the drawing is the same bytes."""
    app = opened(tmp_path)
    out = tmp_path / "same.xlsm"
    app.save(out)
    assert drawing_of(out) == drawing_of(WORKBOOK)


def test_a_moved_shape_reaches_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    Worksheets(1).Shapes("Oval").Left = 333')
    out = tmp_path / "moved.xlsm"
    app.save(out)

    again = ExcelApplication.open(out)
    assert again.evaluate('Worksheets(1).Shapes("Oval").Left') == 333
    assert again.evaluate('Worksheets(1).Shapes("Rect").Left') == 10


def test_a_new_shape_reaches_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(
        app,
        "    Dim sh As Object\n"
        "    Set sh = Worksheets(1).Shapes.AddShape(9, 400, 500, 60, 70)\n"
        '    sh.Name = "Fresh"\n'
        '    sh.OnAction = "Clicked"\n'
        '    sh.TextFrame.Characters.Text = "Added"',
    )
    out = tmp_path / "added.xlsm"
    app.save(out)

    again = ExcelApplication.open(out)
    assert again.evaluate('Worksheets(1).Shapes("Fresh").Left') == 400
    assert again.evaluate('Worksheets(1).Shapes("Fresh").Top') == 500
    assert again.evaluate('Worksheets(1).Shapes("Fresh").OnAction') == "Clicked"
    assert again.evaluate('Worksheets(1).Shapes("Fresh").TextFrame.Characters.Text') == "Added"
    assert again.evaluate("Worksheets(1).Shapes.Count") == 10


def test_a_deleted_shape_leaves_the_file(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(app, '    Worksheets(1).Shapes("Rounded").Delete')
    out = tmp_path / "deleted.xlsm"
    app.save(out)

    again = ExcelApplication.open(out)
    assert again.evaluate("Worksheets(1).Shapes.Count") == 8
    assert "Rounded" not in drawing_of(out)


def test_the_macro_a_shape_runs_can_be_changed(tmp_path: Path) -> None:
    app = opened(tmp_path)
    macro(
        app,
        '    Worksheets(1).Shapes("Rect").OnAction = "Other"\n'
        '    Worksheets(1).Shapes("Oval").OnAction = "Clicked"',
    )
    out = tmp_path / "wired.xlsm"
    app.save(out)

    again = ExcelApplication.open(out)
    assert again.evaluate('Worksheets(1).Shapes("Rect").OnAction') == "Other"
    assert again.evaluate('Worksheets(1).Shapes("Oval").OnAction') == "Clicked"


def test_a_workbook_made_from_nothing_can_hold_a_shape(tmp_path: Path) -> None:
    """A sheet with no drawing part gets one, wired up as Excel wires it."""
    app = ExcelApplication()
    app.add_workbook()
    macro(
        app,
        "    Dim sh As Object\n"
        "    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 20, 100, 50)\n"
        '    sh.OnAction = "Go"',
    )
    out = tmp_path / "fresh.xlsx"
    app.save(out)

    package = zipfile.ZipFile(out)
    assert any(name.startswith("xl/drawings/drawing") for name in package.namelist())
    sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert "<drawing r:id=" in sheet

    again = ExcelApplication.open(out)
    assert again.evaluate("Worksheets(1).Shapes.Count") == 1
    assert again.evaluate("Worksheets(1).Shapes(1).OnAction") == "Go"
    assert again.evaluate("Worksheets(1).Shapes(1).Width") == 100
