"""Workbooks other tools write, read and saved as Excel does (scripts/measure_tool_books.py).

tests/fixtures/tool_books/ holds one book made the same way by openpyxl and by XlsxWriter, what Excel read of each, typed
and late bound, and what it read of the file after an edit and a save, its own and the model's.
"""

from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path
from typing import Any
from xml.parsers import expat

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._prefixed_parts import relationships_declared, vml_spelled
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.shapes._xlsx import SheetGrid

HERE = Path(__file__).parent / "fixtures" / "tool_books"
BOOKS: dict[str, Any] = json.loads((HERE / "tool_books.json").read_text(encoding="utf-8"))
READS: list[str] = BOOKS["reads"]
BINDINGS = {"typed": ("Workbook", "Worksheet"), "late": ("Object", "Object")}
SETUP: str = BOOKS["setup"]
SOURCE = """Private Function Shown(ByRef v As Variant) As String
    If IsObject(v) Then
        Shown = "Object"
    ElseIf IsError(v) Then
        Shown = "Error:" & CStr(v)
    ElseIf IsNull(v) Then
        Shown = "Null"
    ElseIf IsEmpty(v) Then
        Shown = "Empty"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function
Function Go(ByVal file As String) As String
    Dim wb As {book}, ws As {sheet}, other As {sheet}, v As Variant, s As String
    Application.DisplayAlerts = False
{setup}    On Error Resume Next
    v = Empty
    v = {read}
    If Err.Number <> 0 Then
        s = "E" & Err.Number
    Else
        s = Shown(v)
    End If
    On Error GoTo 0
    wb.Close False
    Go = s
End Function
"""
ACT = """Sub Act(ByVal file As String)
    Dim wb As Workbook, ws As Worksheet, other As Worksheet
    Application.DisplayAlerts = False
{setup}    {action}
    wb.Save
    wb.Close False
End Sub
"""
def _read(file: Path, read: str, binding: str) -> str:
    book, sheet = BINDINGS[binding]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE.format(book=book, sheet=sheet, setup=SETUP, read=read), name="M")
    return str(app.run("Go", file.as_posix()))


CASES = [(tool, binding, index) for tool in BOOKS["tools"] for binding in BINDINGS for index in range(len(READS))]


@pytest.mark.parametrize(("tool", "binding", "index"), CASES, ids=[f"{tool}-{binding}-{READS[index]}" for tool, binding, index in CASES])
def test_reads_what_excel_reads(tool: str, binding: str, index: int, tmp_path: Path) -> None:
    read = READS[index]
    file = tmp_path / f"{tool}.xlsx"
    shutil.copy(HERE / f"{tool}.xlsx", file)
    assert _read(file, read, binding) == BOOKS["tools"][tool][binding]["excel"][index]


@pytest.mark.parametrize("tool", BOOKS["tools"])
@pytest.mark.parametrize("name", BOOKS["edits"])
def test_a_save_after_an_edit_is_a_file_excel_opens(tool: str, name: str, tmp_path: Path) -> None:
    """Every part parses, as Excel refuses a file with one that does not. A note's VML once lost all it held after its first
    double quote, and a link on an openpyxl sheet was written with a prefix nothing bound."""
    file = tmp_path / f"{tool}.xlsx"
    shutil.copy(HERE / f"{tool}.xlsx", file)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(ACT.format(setup=SETUP, action=BOOKS["edits"][name]), name="M")
    app.run("Act", file.as_posix())
    with zipfile.ZipFile(file) as saved:
        for part in saved.namelist():
            if part.endswith((".xml", ".rels", ".vml")):
                expat.ParserCreate().Parse(saved.read(part), True)


@pytest.mark.parametrize("tool", BOOKS["tools"])
@pytest.mark.parametrize("name", BOOKS["edits"])
def test_excel_read_the_models_save_as_its_own(tool: str, name: str) -> None:
    """Measured live: what Excel read of the file the model saved, against the one it saved itself."""
    entry = BOOKS["tools"][tool]["edits"][name]
    assert [read for read, theirs, ours in zip(READS, entry["excel"], entry["model"], strict=True) if theirs != ours] == []


def test_vml_under_any_prefix_is_spelled_as_excel_spells_it() -> None:
    written = ('<xml><a:shapelayout xmlns:a="urn:schemas-microsoft-com:office:office" xmlns:b="urn:schemas-microsoft-com:vml" b:ext="edit">'
               '<a:idmap b:ext="edit" data="1"/></a:shapelayout>'
               '<a:shape xmlns:a="urn:schemas-microsoft-com:vml" xmlns:b="urn:schemas-microsoft-com:office:office" b:insetmode="auto" id="s">'
               '<c:ClientData xmlns:c="urn:schemas-microsoft-com:office:excel" ObjectType="Note"><c:Row>3</c:Row></c:ClientData></a:shape></xml>')
    spelled = vml_spelled(written)
    expat.ParserCreate(namespace_separator=" ").Parse(spelled, True)
    assert '<o:shapelayout v:ext="edit">' in spelled
    assert '<o:idmap v:ext="edit" data="1"/>' in spelled
    assert '<v:shape o:insetmode="auto" id="s">' in spelled
    assert '<x:ClientData ObjectType="Note"><x:Row>3</x:Row></x:ClientData></v:shape></xml>' in spelled


def test_vml_spelled_as_excel_spells_it_is_left_to_the_byte() -> None:
    written = ('<xml xmlns:v="urn:schemas-microsoft-com:vml"\r\n xmlns:o="urn:schemas-microsoft-com:office:office"\r\n'
               ' xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="s"><br></v:shape></xml>')
    assert vml_spelled(written) == written


def test_the_relationships_prefix_goes_on_the_root() -> None:
    given = b'<?xml version="1.0"?><worksheet xmlns="m"><drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></worksheet>'
    assert relationships_declared(given).startswith(
        b'<?xml version="1.0"?><worksheet xmlns="m" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing')
    held = b'<worksheet xmlns="m" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a/></worksheet>'
    assert relationships_declared(held) is held
    assert relationships_declared(b"<styleSheet/>") == b"<styleSheet/>"
    other = b'<worksheet xmlns="m"><a xmlns:r="urn:other"/></worksheet>'
    assert relationships_declared(other) is other


def test_a_shape_past_a_hidden_line_is_placed_past_it() -> None:
    grid = SheetGrid(column_widths={2: 0.0}, row_heights={1: 0.0}, default_column=48.0, default_row=15.0)
    assert grid.column_at(96.0) == (3, 0)
    assert grid.column_at(150.0)[0] == 4
    assert grid.row_at(15.0) == (2, 0)


def test_the_model_reports_nothing_of_these_books(tmp_path: Path) -> None:
    file = tmp_path / "xlsxwriter.xlsx"
    shutil.copy(HERE / "xlsxwriter.xlsx", file)
    for read in READS:
        try:
            _read(file, read, "typed")
        except VBAUnsupportedError as failure:  # pragma: no cover
            pytest.fail(f"{read}: {failure}")
