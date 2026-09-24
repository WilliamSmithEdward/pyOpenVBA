"""Notes, which VBA calls comments, in a saved workbook, replayed against live Excel.

tests/fixtures/notes.json is what scripts/measure_notes.py saw: a workbook
of notes of every shape a macro makes -- text with line feeds, markup,
spaces, nothing and a lot, a note shown, one edited and one deleted, beside
a wide column and a tall row, in a merged cell, by the sheet's last row and
column, on a sheet whose rows and columns moved under them and on one with
a form control -- each read back through Comments, and the parts Excel
saved it in. The model makes the same workbook: each note reads as Excel's
did, and the comments part, the VML part and the sheet are written as
Excel wrote them, the random xr:uid Excel gives each note aside. The
workbook Excel saved opens with its notes as Excel reopened it, and saves
again untouched as it came.

Excel's form control markup is outside this: the model writes a control
of its own, which Excel opens, not Excel's, so the sheet with one compares
only its notes.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE = Path(__file__).parent / "fixtures" / "notes.json"
WORKBOOK = Path(__file__).parent / "fixtures" / "notes.xlsx"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
PARTS: dict[str, str] = RECORD["parts"]
#: Each note's random id, which Excel draws afresh for every note it makes and so no model can repeat.
_UID = re.compile(r'xr:uid="\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}"')
_SHAPE = re.compile(r"<v:shape\b.*?</v:shape>", re.DOTALL)
_ROW = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.DOTALL)
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)
#: The sheets whose parts the model writes as Excel does, by the part numbers Excel gave them.
_NOTES_ONLY = {"Data": 1, "Moved": 2, "Edges": 4}
#: What the model leaves out of a sheet Excel writes, and why.
_MERGED = "Excel writes each cell of a merged block, N2:O3, as a record of its own; the model's Merge writes none"
CELL_GAPS: dict[str, str] = {"sheet1.xml N2": _MERGED, "sheet1.xml O2": _MERGED, "sheet1.xml N3": _MERGED,
                             "sheet1.xml O3": _MERGED}
#: A sheet's notes as the measurement read them: cell, text, shown, shape name, where the shape is, and whether its
#: author is the application's user.
_READ = """Public Function Notes() As String
Dim ws As Object, c As Object, out As String
For Each ws In ActiveWorkbook.Worksheets
    out = out & ws.Name & "`"
    For Each c In ws.Comments
        out = out & c.Parent.Address(False, False) & "|" & Replace(Replace(c.Text, vbCr, "\\r"), vbLf, "\\n") & "|" _
            & c.Visible & "|" & c.Shape.Name & "|" & c.Shape.Left & "|" & c.Shape.Top & "|" & c.Shape.Width & "|" _
            & c.Shape.Height & "|" & (c.Author = Application.UserName) & "~"
    Next
    out = out & "^"
Next
Notes = out
End Function
"""


def _reads(reported: str) -> dict[str, list[dict[str, str]]]:
    found: dict[str, list[dict[str, str]]] = {}
    for sheet in reported.split("^")[:-1]:
        name, _, notes = sheet.partition("`")
        found[name] = []
        for note in notes.split("~")[:-1]:
            cell, text, visible, shape, left, top, width, height, by_user = note.split("|")
            found[name].append({"cell": cell, "text": text.replace("\\r", "\r").replace("\\n", "\n"),
                                "visible": visible, "shape": shape, "left": left, "top": top, "width": width,
                                "height": height, "by_user": by_user})
    return found


@pytest.fixture(scope="module")
def made() -> ExcelApplication:
    """The measurement's workbook, made in the model by the application's user Excel had."""
    app = ExcelApplication()
    app.application.user_name = RECORD["author"]
    app.add_workbook()
    lines = ["Public Sub Writes()", "Dim ws As Object"]
    for index, (name, writes) in enumerate(RECORD["sheets"].items()):
        lines.append("Set ws = ActiveWorkbook.Worksheets(1)" if not index else
                     "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets("
                     "ActiveWorkbook.Worksheets.Count))")
        lines += [f'ws.Name = "{name}"', writes]
    app.add_module("\n".join(lines) + "\nEnd Sub\n" + _READ, name="Writes")
    app.run("Writes")
    return app


@pytest.fixture(scope="module")
def saved(made: ExcelApplication) -> dict[str, str]:
    with tempfile.TemporaryDirectory() as folder:
        path = made.save(Path(folder) / "notes.xlsx")
        with zipfile.ZipFile(path) as package:
            return {name: package.read(name).decode("utf-8") for name in package.namelist() if name in PARTS}


@pytest.mark.parametrize("sheet", list(RECORD["notes"]))
def test_each_note_reads_as_in_excel(made: ExcelApplication, sheet: str) -> None:
    assert _reads(str(made.run("Notes")))[sheet] == RECORD["notes"][sheet]


@pytest.mark.parametrize("part", [name for name in PARTS if name.startswith("xl/comments")])
def test_a_comments_part_is_written_as_excel_writes_it(saved: dict[str, str], part: str) -> None:
    assert _UID.sub("xr:uid", saved[part]) == _UID.sub("xr:uid", PARTS[part])


def test_each_note_has_an_id_of_its_own(saved: dict[str, str]) -> None:
    ids = [found.group(0) for part, text in saved.items() if part.startswith("xl/comments")
           for found in _UID.finditer(text)]
    assert len(ids) == len(set(ids)) == sum(len(notes) for notes in RECORD["notes"].values())


@pytest.mark.parametrize("sheet", list(_NOTES_ONLY))
def test_a_vml_part_draws_the_notes_as_excel_draws_them(saved: dict[str, str], sheet: str) -> None:
    part = f"xl/drawings/vmlDrawing{_NOTES_ONLY[sheet]}.vml"
    assert saved[part] == PARTS[part]


def test_notes_beside_a_form_control_are_drawn_as_excel_draws_them(saved: dict[str, str]) -> None:
    def notes(vml: str) -> list[str]:
        return [shape for shape in _SHAPE.findall(vml) if 'ObjectType="Note"' in shape]

    part = "xl/drawings/vmlDrawing3.vml"
    assert notes(saved[part]) == notes(PARTS[part])
    ids = [_found(r"_x0000_s(\d+)", shape) for shape in _SHAPE.findall(saved[part])]
    # One sequence for the notes and the control, in the order they were made, in the sheet's own block of ids.
    assert ids == ["3073", "3074", "3075"]


def _found(pattern: str, text: str) -> str:
    """What the pattern's first group, or the whole pattern, matches first in ``text``, which it has to."""
    match = re.search(pattern, text)
    assert match is not None, pattern
    return match.group(1) if match.re.groups else match.group(0)


@pytest.mark.parametrize("sheet", list(_NOTES_ONLY))
def test_a_sheet_with_notes_names_its_parts_as_excel_does(saved: dict[str, str], sheet: str) -> None:
    rels = f"xl/worksheets/_rels/sheet{_NOTES_ONLY[sheet]}.xml.rels"
    assert saved[rels] == PARTS[rels]


def _body(xml: str) -> str:
    """A sheet's part from its dimension on, past the root element the model's template writes otherwise."""
    return xml[xml.index("<dimension"):]


@pytest.mark.parametrize("sheet", list(_NOTES_ONLY))
def test_a_sheet_with_notes_writes_its_rows_as_excel_does(saved: dict[str, str], sheet: str) -> None:
    part = f"xl/worksheets/sheet{_NOTES_ONLY[sheet]}.xml"
    model, excel = _body(saved[part]), _body(PARTS[part])
    assert _found(r"<dimension[^>]*>", model) == _found(r"<dimension[^>]*>", excel)

    def starts(xml: str) -> list[str]:
        # Each row's own attributes; whether it holds cells is for the cells to say.
        return [row[:row.index(">")].rstrip("/") for row in _ROW.findall(xml)]

    assert starts(model) == starts(excel)
    assert model[model.index("</sheetData>"):] == excel[excel.index("</sheetData>"):]


def _cells() -> list[Any]:
    out: list[Any] = []
    for sheet, number in _NOTES_ONLY.items():
        for cell in _CELL.findall(PARTS[f"xl/worksheets/sheet{number}.xml"]):
            name = f"sheet{number}.xml {cell}"
            out.append(pytest.param(sheet, cell, id=name, marks=[pytest.mark.xfail(reason=CELL_GAPS[name],
                                                                                   strict=True)]
                                    if name in CELL_GAPS else []))
    return out


@pytest.mark.parametrize(("sheet", "cell"), _cells())
def test_a_sheet_with_notes_writes_its_cells_as_excel_does(saved: dict[str, str], sheet: str, cell: str) -> None:
    part = f"xl/worksheets/sheet{_NOTES_ONLY[sheet]}.xml"

    def element(xml: str) -> str | None:
        found = re.search(rf'<c r="{cell}"[^>]*?(?:/>|>.*?</c>)', xml, re.DOTALL)
        return None if found is None else found.group(0)

    assert element(saved[part]) == element(PARTS[part])


def test_every_gap_is_a_cell() -> None:
    names = {f"sheet{number}.xml {cell}" for number in _NOTES_ONLY.values()
             for cell in _CELL.findall(PARTS[f"xl/worksheets/sheet{number}.xml"])}
    assert set(CELL_GAPS) <= names


@pytest.fixture(scope="module")
def opened() -> ExcelApplication:
    app = ExcelApplication.open(WORKBOOK, with_vba=False)
    app.application.user_name = RECORD["author"]
    app.add_module(_READ, name="Reads")
    return app


@pytest.mark.parametrize("sheet", list(RECORD["opened"]))
def test_the_workbook_excel_saved_opens_with_its_notes(opened: ExcelApplication, sheet: str) -> None:
    assert _reads(str(opened.run("Notes")))[sheet] == RECORD["opened"][sheet]


def test_the_workbook_excel_saved_saves_its_notes_as_they_came() -> None:
    app = ExcelApplication.open(WORKBOOK, with_vba=False)
    app.sheet(1).set_value("B20", 5)
    with tempfile.TemporaryDirectory() as folder, zipfile.ZipFile(app.save(Path(folder) / "again.xlsx")) as package:
        again = {name: package.read(name).decode("utf-8") for name in package.namelist() if name in PARTS}
    for part, text in PARTS.items():
        if part.startswith(("xl/comments", "xl/drawings/vml")):
            assert again[part] == text, part


def test_a_note_added_to_the_workbook_excel_saved_keeps_the_rest_as_they_came() -> None:
    app = ExcelApplication.open(WORKBOOK, with_vba=False)
    app.application.user_name = RECORD["author"]
    app.add_module('Public Sub Add()\nWorksheets("Moved").Range("H8").AddComment "added"\nEnd Sub\n', name="Add")
    app.run("Add")
    with tempfile.TemporaryDirectory() as folder, zipfile.ZipFile(app.save(Path(folder) / "added.xlsx")) as package:
        again = {name: package.read(name).decode("utf-8") for name in package.namelist()}
    comments, vml = again["xl/comments2.xml"], again["xl/drawings/vmlDrawing2.vml"]
    # The notes that were there keep their markup and their ids; the new one comes after them, in cell order.
    for element in re.findall(r"<comment\b.*?</comment>", PARTS["xl/comments2.xml"], re.DOTALL):
        assert element in comments
    for shape in _SHAPE.findall(PARTS["xl/drawings/vmlDrawing2.vml"]):
        assert shape in vml
    assert re.findall(r'<comment ref="([A-Z]+\d+)"', comments) == ["C3", "F5", "H8"]
    assert re.findall(r'id="_x0000_s(\d+)"', vml) == ["2049", "2051", "2052"]
