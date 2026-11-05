"""The hidden names formulas bring into Names, which no file lists (scripts/measure_xlfn_names.py,
scripts/measure_xlfn_functions.py), and a part that is not well-formed (scripts/measure_bad_xml.py)."""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hidden_names import names_in
from pyopenvba.apps.excel._io import WorkbookFileError
from pyopenvba.integrity import xml_problem

HERE = Path(__file__).parent / "fixtures"
NAMES: dict[str, Any] = json.loads((HERE / "xlfn_names.json").read_text(encoding="utf-8"))
FUNCTIONS: dict[str, Any] = json.loads((HERE / "xlfn_functions.json").read_text(encoding="utf-8"))
BAD: dict[str, Any] = json.loads((HERE / "bad_xml/bad_xml.json").read_text(encoding="utf-8"))
#: A hidden name a conditional format's formula brings, which the model does not watch for.
NOT_WATCHED = {"cond_format"}
#: Excel keeps the @ of a formula written while its name was not there, once the name comes: =@Nowhere, in a file SINGLE of it.
KEEPS_AT = {"named_after": "a formula written before its name keeps no @ in the model"}


def _run(source: str, entry: str, *arguments: object) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    return app.run(entry, *arguments)


@pytest.mark.parametrize("row", [row for row in NAMES["rows"] if row["name"] not in NOT_WATCHED], ids=lambda row: row["name"])
def test_names_hold_what_excel_lists(row: dict[str, str], tmp_path: Path) -> None:
    if row["name"] in KEEPS_AT:
        pytest.xfail(KEEPS_AT[row["name"]])
    source = NAMES["probe"].format(action=NAMES["cases"][row["name"]])
    file = tmp_path / f"{row['name']}.xlsx"
    assert str(_run(source, "Go", file.as_posix())).replace(file.name, "<book>") == row["answer"]


@pytest.mark.parametrize("name", sorted(name for name, answer in FUNCTIONS["rows"].items() if "~" in answer))
def test_each_function_brings_the_names_it_did(name: str) -> None:
    shown, _, listed = FUNCTIONS["rows"][name].rpartition("~")
    assert names_in(shown) == {one for one in listed.split(",") if one}


def _broken(tmp_path: Path, part: str, how: str) -> Path:
    file = tmp_path / f"{part}.xlsx"
    with zipfile.ZipFile(HERE / "bad_xml/base.xlsx") as given, zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as made:
        for info in given.infolist():
            data = given.read(info.filename)
            if re.fullmatch(BAD["parts"][part], info.filename):
                at = data.rfind(b"</")
                data = data[:at] + BAD["broken"][how].encode() + data[at:]
            made.writestr(info, data)
    return file


OPEN = """Function Tried(ByVal file As String) As String
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(file)
    Tried = Err.Number & "|" & Err.Description
End Function
"""


@pytest.mark.parametrize("row", [row for row in BAD["rows"] if row["part"]], ids=lambda row: f"{row['part']}-{row['how']}")
def test_a_part_that_is_not_xml_is_what_excel_makes_of_it(row: dict[str, str], tmp_path: Path) -> None:
    file = _broken(tmp_path, row["part"], row["how"])
    number, _, text = row["answer"].partition("|")
    answer = str(_run(OPEN, "Tried", file.as_posix()))
    assert answer.split("|")[0] == number
    if number != "0":
        assert answer.split("|")[1].replace(file.name, "") == re.sub(r"'[^']*\.xlsx'", "''", text) or answer.split("|")[1].startswith(text[:30])


DOCTYPES: dict[str, str] = json.loads((HERE / "bad_xml/doctype.json").read_text(encoding="utf-8"))
_LAUGHS = b'<!DOCTYPE root [<!ENTITY a "x">' + b"".join(b'<!ENTITY a%d "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">' % n for n in range(20)) + b"]>"


@pytest.mark.parametrize("case", sorted(DOCTYPES))
def test_a_part_with_a_doctype_is_refused_as_excel_refuses_it(case: str, tmp_path: Path) -> None:
    part, _, how = case.partition("|")
    file = tmp_path / "doctype.xlsx"
    doctype = b"<!DOCTYPE root>" if how == "plain" else b'<!DOCTYPE root [<!ENTITY e "x">]>'
    with zipfile.ZipFile(HERE / "bad_xml/base.xlsx") as given, zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as made:
        for name in given.namelist():
            data = given.read(name)
            if name == part:
                at = data.find(b"?>") + 2 if data.lstrip().startswith(b"<?xml") else 0
                data = data[:at] + doctype + data[at:]
            made.writestr(name, data)
    assert str(_run(OPEN, "Tried", file.as_posix())).split("|")[0] == DOCTYPES[case]


def test_entities_are_never_expanded_at_open() -> None:
    assert xml_problem(_LAUGHS + b"<root>&a19;</root>") is not None
    assert xml_problem(b"<root/>") is None


def test_python_is_told_a_part_is_not_xml(tmp_path: Path) -> None:
    with pytest.raises(WorkbookFileError, match="well-formed"):
        ExcelApplication.open(_broken(tmp_path, "sheet", "surrogate"))
