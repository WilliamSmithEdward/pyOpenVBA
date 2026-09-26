"""Hyperlinks, replayed against live Excel.

tests/fixtures/hyperlinks/hyperlinks.json is what
scripts/measure_hyperlinks.py saw: links added over every kind of cell,
with every kind of address, their properties set, taken away, moved,
copied, pasted and sorted, on protected sheets and merged cells, the
HYPERLINK function, the Hyperlink style, and the errors -- one module,
each section run in an Excel of its own -- and the workbooks Excel saved
along the way. The model runs the same module and answers as Excel did,
and the workbooks it saves hold the links as Excel's do: each sheet's
part after its cells, its relationships, each cell's style and the
stylesheet. The workbooks Excel saved open with their links as Excel
reopened them, and save again as they came.

Two sections the model refuses, as REFUSED says, and a few parts of the
saved workbooks belong to features outside hyperlinks, as GAPS says.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hyperlinks_file import given, read_address, saved_address, target_of, tidied
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "hyperlinks"
RECORD = json.loads((FIXTURES / "hyperlinks.json").read_text(encoding="utf-8"))
#: The sections the model refuses, and why.
REFUSED = {"Other": "a hyperlink on a shape is not implemented",
           "SubjectInternal": "EmailSubject on a link with no address leaves Excel refusing every change after it"}
#: What of a saved workbook belongs to another feature than hyperlinks, which the model writes otherwise.
GAPS = {
    ("moved.xlsx", "tail"): "Range.Sort's sortState, which the model does not write",
    ("merged.xlsx", "cells"): "Excel gives each cell of a merged block a format of its own; the model's Merge does not",
    ("merged.xlsx", "styles"): "Excel gives each cell of a merged block a format of its own; the model's Merge does not",
    ("rels.xlsx", "styles"): "Excel adds its notes' Tahoma 9 font to the stylesheet; the model does not",
}
#: Each style a macro adds gets a random xr:uid from Excel, which the model leaves out.
_UID = re.compile(r' xr:uid="\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}"')
_CELL = re.compile(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>)')
#: A DOS path on C:, as Excel saved it relative to the folder on C: it saved canon.xlsx in, six deep.
_DOS = "../../../../../../docs/file.xlsx"
#: The Canon section's rows whose link is a DOS path on C:. Saved in a folder on another drive, or in a folder of
#: no drive's off Windows, one takes the rule the rows on Q: take, which every platform checks.
_ON_C = ("A11", "A13", "A34")


def _on_c(folder: Path) -> bool:
    """Whether a folder is on C:, as the one Excel saved canon.xlsx in was."""
    return saved_address("C:\\docs\\file.xlsx", str(folder.resolve())) != "C:\\docs\\file.xlsx"


def _canon(answer: str, folder: Path) -> str:
    """The Canon section's answer as it stands for a workbook saved in ``folder``."""
    if _on_c(folder):
        return answer.replace(_DOS, saved_address("C:\\docs\\file.xlsx", str(folder.resolve())))
    groups = answer.split("~")
    rows = groups[2].split("^")
    groups[2] = "^".join(row for index, row in enumerate(rows, start=1) if f"A{index}" not in _ON_C)
    return "~".join(groups)


def _items(answer: str) -> list[str]:
    return re.split(r"[\^~]", answer)


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], Path]:
    folder = tmp_path_factory.mktemp("links")
    answers: dict[str, str] = {}
    for name, saves in RECORD["sections"].items():
        if name in REFUSED:
            continue
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(RECORD["module"], name="Probe")
        answers[name] = str(app.run(name, *((str(folder) + "/",) if saves else ())))
    return answers, folder


@pytest.mark.parametrize("section", [name for name in RECORD["sections"] if name not in REFUSED])
def test_each_section_answers_as_excel_did(replayed: tuple[dict[str, str], Path], section: str) -> None:
    excel, model = RECORD["runs"][section], replayed[0][section]
    if section == "Canon":
        # Read back, a DOS path is the one the workbook was saved with, relative to the folder it was saved in.
        excel, model = _canon(excel, replayed[1]), (model if _on_c(replayed[1]) else _canon(model, replayed[1]))
    assert _items(model) == _items(excel)


@pytest.mark.parametrize("section", list(REFUSED))
def test_a_section_the_model_refuses_says_so(section: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["module"], name="Probe")
    with pytest.raises(VBAUnsupportedError):
        app.run(section)


# --- the workbooks saved along the way ----------------------------------------------------------------------------


def _parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode("utf-8") for name in package.namelist()
                if name.startswith("xl/worksheets/") or name == "xl/styles.xml"}


def _canon_parts(sheet: str, rels: str, folder: Path) -> tuple[str, str]:
    """A canon.xlsx sheet and its relationships as they stand for a workbook saved in ``folder``: the DOS paths on
    C: made relative to it, or, off C:, the links on C: left out with their relationships."""
    if _on_c(folder):
        return (sheet.replace(_DOS, saved_address("C:\\docs\\file.xlsx", str(folder.resolve()))),
                rels.replace(_DOS, target_of("C:\\docs\\file.xlsx", str(folder.resolve()))))
    for cell in _ON_C:
        found = re.search(rf'<hyperlink ref="{cell}" r:id="(rId\d+)"[^>]*/>', sheet)
        assert found is not None, cell
        sheet = sheet.replace(found.group(0), "")
        rels = re.sub(rf'<Relationship Id="{found.group(1)}"[^>]*/>', "", rels)
    return sheet, rels


def _after_cells(xml: str) -> str:
    """A sheet's part from the end of its cells on, its random ids aside."""
    end = xml.index("</sheetData>") if "</sheetData>" in xml else xml.index("<sheetData/>")
    return _UID.sub("", xml[end:].removeprefix("<sheetData/>").removeprefix("</sheetData>"))


def _saved() -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for path in sorted(FIXTURES.glob("*.xlsx")):
        out += [(path.name, part) for part in _parts(path) if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", part)]
    return out


def _gap(workbook: str, what: str) -> list[pytest.MarkDecorator]:
    reason = GAPS.get((workbook, what))
    return [pytest.mark.xfail(reason=reason, strict=True)] if reason else []


def _pair(workbook: str, part: str, folder: Path) -> tuple[tuple[str, str], tuple[str, str]]:
    """Excel's sheet part and its relationships, and the model's, as they stand for the folder the model saved in."""
    rels = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    theirs, ours = _parts(FIXTURES / workbook), _parts(folder / workbook)
    excel, model = (theirs[part], theirs.get(rels, "")), (ours[part], ours.get(rels, ""))
    if workbook == "canon.xlsx":
        excel = _canon_parts(*excel, folder)
        if not _on_c(folder):
            model = _canon_parts(*model, folder)
    return excel, model


@pytest.mark.parametrize(("workbook", "part"), [pytest.param(*one, marks=_gap(one[0], "tail")) for one in _saved()])
def test_a_sheet_after_its_cells_is_written_as_excel_writes_it(replayed: tuple[dict[str, str], Path], workbook: str,
                                                                part: str) -> None:
    excel, model = _pair(workbook, part, replayed[1])
    assert _after_cells(model[0]) == _after_cells(excel[0])


@pytest.mark.parametrize(("workbook", "part"), _saved())
def test_a_sheets_links_are_written_as_excel_writes_them(replayed: tuple[dict[str, str], Path], workbook: str,
                                                         part: str) -> None:
    def links(xml: str) -> str:
        found = re.search(r"<hyperlinks>.*?</hyperlinks>", xml, re.DOTALL)
        return _UID.sub("", found.group(0)) if found else ""

    excel, model = _pair(workbook, part, replayed[1])
    assert links(model[0]) == links(excel[0])


@pytest.mark.parametrize(("workbook", "part"), _saved())
def test_a_sheets_relationships_are_written_as_excel_writes_them(replayed: tuple[dict[str, str], Path],
                                                                 workbook: str, part: str) -> None:
    excel, model = _pair(workbook, part, replayed[1])
    assert model[1] == excel[1]


@pytest.mark.parametrize(("workbook", "part"), [pytest.param(*one, marks=_gap(one[0], "cells")) for one in _saved()])
def test_each_cell_takes_the_style_excel_gives_it(replayed: tuple[dict[str, str], Path], workbook: str,
                                                  part: str) -> None:
    def styles(xml: str) -> dict[str, str]:
        found = {cell: re.search(r'\bs="\d+"', rest) for cell, rest in _CELL.findall(xml)}
        return {cell: match.group(0) if match else "" for cell, match in found.items()}

    assert styles(_parts(replayed[1] / workbook)[part]) == styles(_parts(FIXTURES / workbook)[part])


@pytest.mark.parametrize("workbook", [pytest.param(path.name, marks=_gap(path.name, "styles"))
                                      for path in sorted(FIXTURES.glob("*.xlsx"))])
def test_the_stylesheet_is_written_as_excel_writes_it(replayed: tuple[dict[str, str], Path], workbook: str) -> None:
    excel = _parts(FIXTURES / workbook)["xl/styles.xml"]
    assert _parts(replayed[1] / workbook)["xl/styles.xml"] == _UID.sub("", excel)


def test_every_gap_is_a_saved_workbook() -> None:
    assert {workbook for workbook, _ in GAPS} <= {path.name for path in FIXTURES.glob("*.xlsx")}


# --- the workbooks Excel saved --------------------------------------------------------------------------------------

#: A workbook Excel saved, the cells read back, and the section and group of its answers that read them in Excel.
_OPENED = {
    "links.xlsx": ([cell for cell, _, _ in RECORD["adds"]], "Adds", 7),
    "reopen.xlsx": (["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A9"], "Reopen", 0),
    "canon.xlsx": ([f"A{row}" for row in range(1, len(RECORD["addresses"]) + len(RECORD["wide_addresses"]) + 1)],
                   "Canon", 2),
}
_READ_BACK = """
Public Function ReadBack(cells As String) As String
    Dim out As String, one As Variant
    For Each one In Split(cells, ",")
        out = out & CellProps(ActiveWorkbook.Worksheets(1).Range(one)) & "^"
    Next
    ReadBack = out
End Function

Public Function Listed() As String
    Listed = Listing(ActiveWorkbook.Worksheets(1))
End Function
"""


def _opened(workbook: str) -> ExcelApplication:
    app = ExcelApplication.open(FIXTURES / workbook, with_vba=False)
    app.add_module(RECORD["module"] + _READ_BACK, name="Probe")
    return app


@pytest.mark.parametrize("workbook", list(_OPENED))
def test_the_workbook_excel_saved_opens_with_its_links(workbook: str) -> None:
    cells, section, group = _OPENED[workbook]
    answer = str(_opened(workbook).run("ReadBack", ",".join(cells)))
    assert _items(answer) == _items(RECORD["runs"][section].split("~")[group])


def test_the_workbook_excel_saved_lists_its_links_in_its_order() -> None:
    assert str(_opened("links.xlsx").run("Listed")) == RECORD["runs"]["Adds"].split("~")[6]


def test_the_workbook_excel_saved_saves_its_links_as_they_came(tmp_path: Path) -> None:
    app = ExcelApplication.open(FIXTURES / "links.xlsx", with_vba=False)
    app.sheet(1).set_value("Z99", 5)
    again = _parts(app.save(tmp_path / "again.xlsx"))
    before = _parts(FIXTURES / "links.xlsx")
    part = "xl/worksheets/sheet1.xml"
    assert again[part][again[part].index("<hyperlinks>"):] == before[part][before[part].index("<hyperlinks>"):]
    assert again["xl/worksheets/_rels/sheet1.xml.rels"] == before["xl/worksheets/_rels/sheet1.xml.rels"]


def test_a_link_added_to_the_workbook_excel_saved_keeps_the_others(tmp_path: Path) -> None:
    app = _opened("links.xlsx")
    app.add_module('Public Sub AddOne()\nWorksheets(1).Hyperlinks.Add Worksheets(1).Range("C1"), '
                   '"http://added.example/"\nEnd Sub\n', name="AddOne")
    app.run("AddOne")
    again = _parts(app.save(tmp_path / "added.xlsx"))["xl/worksheets/sheet1.xml"]
    before = _parts(FIXTURES / "links.xlsx")["xl/worksheets/sheet1.xml"]
    links = re.findall(r"<hyperlink [^>]*/>", again)
    # Each link keeps its markup and its id, numbered as before; the new one comes last, with the next number.
    assert links[:-1] == re.findall(r"<hyperlink [^>]*/>", before)
    assert re.sub(r' xr:uid="[^"]*"', "", links[-1]) == '<hyperlink ref="C1" r:id="rId24"/>'


# --- addresses --------------------------------------------------------------------------------------------------


def test_a_dos_path_is_saved_relative_to_the_workbooks_folder() -> None:
    # Excel saved canon.xlsx six folders deep on the drive of C:\docs.
    assert saved_address("C:\\docs\\file.xlsx", "C:\\a\\b\\c\\d\\e\\f") == _DOS
    assert target_of("file://C:\\docs\\file.xlsx", "C:\\a\\b\\c\\d\\e\\f") == _DOS


def test_a_dos_path_is_saved_relative_to_each_folder_the_workbook_is_saved_in(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub AddOne()\nWorksheets(1).Hyperlinks.Add Worksheets(1).Range("A1"), "C:\\docs\\file.xlsx"'
                   "\nEnd Sub\n", name="AddOne")
    app.run("AddOne")
    for place in ("one", "two/deeper"):
        path = tmp_path / place / "book.xlsx"
        rels = _parts(app.save(path))["xl/worksheets/_rels/sheet1.xml.rels"]
        target = target_of("C:\\docs\\file.xlsx", str(path.parent.resolve()))
        assert f'Target="{target}"' in rels


def test_a_dos_path_on_another_drive_and_a_unc_path_are_saved_whole() -> None:
    assert target_of("D:\\docs\\file.xlsx", "C:\\work") == "file:///D:\\docs\\file.xlsx"
    assert target_of("\\\\server\\share\\file.xlsx", "C:\\work") == "file:///\\\\server\\share\\file.xlsx"
    assert read_address("file:///\\\\server\\share\\file.xlsx") == "\\\\server\\share\\file.xlsx"


def test_an_address_is_tidied_as_excel_tidies_it() -> None:
    record = dict(zip([*RECORD["addresses"]], _items(RECORD["runs"]["Canon"].split("~")[1]), strict=False))
    for given_address, item in record.items():
        address, sub, _ = given(given_address)
        fields = item.split("|")[1].split("`") if "|" in item else ["String:", "String:"]
        assert (f"String:{address}", f"String:{sub}") == (fields[0], fields[1]), given_address
    assert tidied("") == ""
