"""add_vba_project, held to what Excel, Word and PowerPoint write for a first macro.

tests/fixtures/first_macro/ is what scripts/measure_first_macro.py saw:
each application's file saved with no VBA (before), saved again after a
standard module was added (module), and after a module was added and
removed (emptied). Here the library gives the first file a project, adds
the same module or none, and saves. Its project is the one each
application made, component for component and declaration for
declaration, and every other part is the application's, byte for byte,
but for the differences _new_project.py states: where Office orders the
new relationship and Word's numbering of all of them, Excel's per-build
fileVersion GUID, and Word's vbaData.xml. A part the application changes
on every save alike (first_macro.json's resave_changes) stays as it was.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, NoVBAProjectError, PowerPointFile, UnsupportedFormatError, VBAProjectError, WordFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_vba_project

FIXTURES = Path(__file__).parent / "fixtures" / "first_macro"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "first_macro.json").read_text(encoding="utf-8"))
OPENERS: dict[str, tuple[type[ExcelFile] | type[WordFile] | type[PowerPointFile], str]] = {
    "excel": (ExcelFile, "xl"), "word": (WordFile, "word"), "powerpoint": (PowerPointFile, "ppt")}
MODULE = "Public Sub Hello()\r\nEnd Sub\r\n"
#: Parts whose relationships Office orders, or numbers, its own way: the same relationships, in another order.
REORDERED = {"xl/_rels/workbook.xml.rels", "word/_rels/document.xml.rels", "ppt/_rels/presentation.xml.rels"}
CASES = [(host, label) for host in RECORD for label in ("emptied", "module")]


def _parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def _saved(host: str, label: str, tmp_path: Path) -> Path:
    opener, _ = OPENERS[host]
    suffix = RECORD[host]["suffix"]
    source, out = tmp_path / f"in.{suffix}", tmp_path / f"out.{suffix}"
    shutil.copyfile(FIXTURES / f"{host}_before.{suffix}", source)
    with opener(source) as office_file:
        project = office_file.add_vba_project()
        if label == "module":
            project.add_module("Module1", MODULE)
        office_file.save(out)
    return out


def _project(data: bytes) -> tuple[list[tuple[str, str]], list[str]]:
    cfb = CFB.from_bytes(data)
    modules = [(module.name, module.kind.name) for module in parse_vba_project(cfb).modules]
    declarations = [line for line in cfb.get_stream("PROJECT").decode("latin-1").splitlines()
                    if line.startswith(("Document=", "Module=", "Class="))]
    return modules, declarations


def _relationships(data: bytes) -> list[str]:
    return sorted(re.sub(r' Id="rId\d+"', "", element)
                  for element in re.findall(r"<Relationship\b[^>]*/>", data.decode("utf-8")))


@pytest.mark.parametrize(("host", "label"), CASES)
def test_the_project_is_the_one_the_application_makes(host: str, label: str, tmp_path: Path) -> None:
    _, folder = OPENERS[host]
    ours, theirs = _parts(_saved(host, label, tmp_path)), _parts(FIXTURES / f"{host}_{label}.{RECORD[host]['suffix']}")
    entry = f"{folder}/vbaProject.bin"
    assert (entry in ours) == (entry in theirs)
    if entry in ours:
        assert _project(ours[entry]) == _project(theirs[entry])


@pytest.mark.parametrize(("host", "label"), CASES)
def test_every_other_part_is_the_applications(host: str, label: str, tmp_path: Path) -> None:
    suffix, noise = RECORD[host]["suffix"], set(RECORD[host]["resave_changes"])
    before = _parts(FIXTURES / f"{host}_before.{suffix}")
    ours, theirs = _parts(_saved(host, label, tmp_path)), _parts(FIXTURES / f"{host}_{label}.{suffix}")
    missing = sorted(theirs.keys() - ours.keys())
    if (host, label) == ("word", "module"):
        # Word writes vbaData.xml, related from the project, once there is a macro; the library does not.
        assert missing == ["word/_rels/vbaProject.bin.rels", "word/vbaData.xml"]
    else:
        assert missing == []
    for name in sorted(ours.keys() - {f"{OPENERS[host][1]}/vbaProject.bin"}):
        if name in REORDERED:
            assert _relationships(ours[name]) == _relationships(theirs[name]), name
        elif name == "xl/workbook.xml":
            # Excel changes it on every save; the library changes only the code name it adds, as Excel adds it.
            element = re.compile(r"<workbookPr\b[^>]*>")
            assert element.findall(ours[name].decode("utf-8")) == element.findall(theirs[name].decode("utf-8"))
            assert element.sub("", ours[name].decode("utf-8")) == element.sub("", before[name].decode("utf-8"))
        elif name == "[Content_Types].xml" and (host, label) == ("word", "module"):
            vba_data = rb'<Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/>'
            assert ours[name] == theirs[name].replace(vba_data, b"")
        elif name == "docProps/core.xml":
            # Each save stamps its time here, and Word and PowerPoint count it as a revision: nothing else
            # differs, a save in the same second leaves it out of resave_changes, and the library leaves it alone.
            stamps = re.compile(rb"<dcterms:modified\b[^>]*>[^<]*</dcterms:modified>|<cp:revision>\d+</cp:revision>")
            assert ours[name] == before[name]
            assert stamps.sub(b"", theirs[name]) == stamps.sub(b"", before[name])
        elif name in noise:
            assert ours[name] == before[name], name
        else:
            assert ours[name] == theirs[name], name


def test_excel_names_are_excels_own_and_the_saved_file_reads_back(tmp_path: Path) -> None:
    out = _saved("excel", "module", tmp_path)
    with ExcelFile(out) as workbook:
        assert workbook.module_names() == [name for name, _ in RECORD["excel"]["module"]["components"]]
        assert workbook.get_module("Module1").endswith(MODULE)


def test_a_file_with_a_project_already_is_refused(tmp_path: Path) -> None:
    out = _saved("word", "module", tmp_path)
    with WordFile(out) as document, pytest.raises(VBAProjectError, match="already has a VBA project"):
        document.add_vba_project()


@pytest.mark.parametrize("name", ["workbook.xls", "document.doc", "presentation.ppt"])
def test_a_binary_file_is_not_given_one(name: str) -> None:
    opener = {"xls": ExcelFile, "doc": WordFile, "ppt": PowerPointFile}[name.rsplit(".", 1)[1]]
    with opener(Path(__file__).parent / "fixtures" / "no_vba" / name) as office_file:
        with pytest.raises(UnsupportedFormatError, match="not supported"):
            office_file.add_vba_project()
        assert not office_file.has_vba_project()


def test_the_refusal_for_a_file_with_none_says_it_can_be_given_one() -> None:
    with ExcelFile(FIXTURES / "excel_before.xlsm") as workbook, pytest.raises(NoVBAProjectError) as raised:
        workbook.get_module("Module1")
    assert "add_vba_project()" in str(raised.value)
    with ExcelFile(Path(__file__).parent / "fixtures" / "no_vba" / "workbook.xls") as workbook:
        with pytest.raises(NoVBAProjectError) as raised:
            workbook.get_module("Module1")
    assert "add_vba_project()" not in str(raised.value)


def test_a_project_added_but_not_saved_lists_as_one(tmp_path: Path) -> None:
    shutil.copyfile(FIXTURES / "excel_before.xlsm", tmp_path / "book.xlsm")
    with ExcelFile(tmp_path / "book.xlsm") as workbook:
        workbook.add_vba_project()
        assert workbook.has_vba_project()
        assert workbook.module_names() == ["ThisWorkbook", "Sheet1", "Sheet2", "Chart3", "Sheet4"]
        assert CFB.from_bytes(workbook.vba_project_bytes()).get_stream("PROJECT")
