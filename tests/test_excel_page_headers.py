"""PageSetup headers and footers, replayed from live Excel (scripts/measure_page_headers.py,
tests/fixtures/page_headers.json).

Excel's macros run in the model as Excel ran them: part one sets, reads and saves, and part two opens files whose
headers were written into a workbook the model saved, as the measurement wrote them into one Excel saved, then reads
and saves them twice. Each reading and each saved headerFooter element is Excel's.
"""

from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._page_headers import section, with_section
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / "fixtures" / "page_headers.json").read_text(encoding="utf-8"))
CHILDREN = ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")
_ELEMENT = re.compile(r"<headerFooter\b[^>]*?(?:/>|>.*?</headerFooter>)", re.S)


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


def records(text: str) -> list[list[str]]:
    """The macro's answer as the measurement splits Excel's: name, notes, and each reading."""
    out: list[list[str]] = []
    for entry in text.split("~"):
        if entry:
            name, notes, dump = entry.split("^")
            out.append([name, notes, *dump.split("¦")])
    return out


def elements(path: Path) -> list[str | None]:
    """Each sheet's headerFooter element as the file has it."""
    found: list[str | None] = []
    with zipfile.ZipFile(path) as package:
        for index in range(1, 10):
            part = f"xl/worksheets/sheet{index}.xml"
            if part not in package.namelist():
                break
            element = _ELEMENT.search(package.read(part).decode("utf-8"))
            found.append(element.group(0) if element else None)
    return found


def written(base: Path, target: Path, attributes: str, children: dict[str, str]) -> None:
    """``base`` with a headerFooter element of ``attributes`` and ``children``, spelled as Excel spells one, where the
    schema puts it."""
    inner = "".join(f"<{name}>{children[name].replace('&', '&amp;').replace('<', '&lt;')}</{name}>"
                    for name in CHILDREN if name in children)
    element = f"<headerFooter{attributes}>{inner}</headerFooter>"
    with zipfile.ZipFile(base) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                found = re.search(r"<pageSetup\b[^>]*/>", text) or re.search(r"<pageMargins\b[^>]*/>", text)
                assert found is not None and "<headerFooter" not in text
                data = (text[:found.end()] + element + text[found.end():]).encode("utf-8")
            out.writestr(info, data)


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[list[str]], Path]:
    """Both parts run in the model: their readings and the folder their files are in."""
    folder = tmp_path_factory.mktemp("page_headers").resolve()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(portable(RECORD["build"]), name="Build")
    made = str(app.run("Build.Build", str(folder)))
    for key, (attributes, children, _step) in RECORD["crafted"].items():
        written(folder / "default.xlsx", folder / f"crafted_{key}.xlsx", attributes, children)
    app.add_module(portable(RECORD["opened"]), name="Opened")
    opened = str(app.run("Opened.Opened", str(folder)))
    return records(made) + records(opened), folder


def test_each_property_reads_as_excels_does(replayed: tuple[list[list[str]], Path]) -> None:
    found, _folder = replayed
    assert found == RECORD["cases"]


def test_each_save_writes_excels_header_element(replayed: tuple[list[list[str]], Path]) -> None:
    _found, folder = replayed
    for name, measured in RECORD["files"].items():
        if not name.startswith("crafted_"):
            assert elements(folder / f"{name}.xlsx") == measured, name


@pytest.mark.parametrize("name", list(RECORD["too_long"]))
def test_a_header_longer_than_excel_opens_is_refused(name: str, tmp_path: Path) -> None:
    length, excel = RECORD["too_long"][name]
    assert excel == "E1004:Method 'Open' of object 'Workbooks' failed"
    app = ExcelApplication()
    app.add_workbook()
    app.save(tmp_path / "base.xlsx")
    written(tmp_path / "base.xlsx", tmp_path / "long.xlsx", "", {"oddHeader": "&C" + "a" * (length - 2)})
    opened = ExcelApplication.open(tmp_path / "long.xlsx")
    with pytest.raises(VBAUnsupportedError, match="longer than 255"):
        opened.evaluate("Worksheets(1).PageSetup.CenterHeader")


def test_a_header_with_pictures_is_read_and_not_set(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.save(tmp_path / "base.xlsx")
    written(tmp_path / "base.xlsx", tmp_path / "pictures.xlsx", "", {"oddHeader": "&C&G"})
    with zipfile.ZipFile(tmp_path / "pictures.xlsx") as package:
        parts = {info.filename: package.read(info.filename) for info in package.infolist()}
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    parts["xl/worksheets/sheet1.xml"] = sheet.replace("</headerFooter>", '</headerFooter><legacyDrawingHF r:id="rId9"/>'
                                                     ).encode("utf-8")
    with zipfile.ZipFile(tmp_path / "pictures.xlsx", "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)
    opened = ExcelApplication.open(tmp_path / "pictures.xlsx")
    assert opened.evaluate("Worksheets(1).PageSetup.CenterHeader") == "&G"
    opened.add_module('Public Sub Retitle()\n    Worksheets(1).PageSetup.CenterHeader = "x"\nEnd Sub\n', name="Edit")
    with pytest.raises(VBAUnsupportedError, match="header pictures"):
        opened.run("Retitle")


def test_sections_read_and_set_as_excel_reads_and_sets_them() -> None:
    # Strings the fixture's cases arrive at, and what Excel read and wrote of them.
    assert [section("&L&C&R", code) for code in "LCR"] == ["&C", "&R", ""]
    assert section("&C&&Rb", "C") == "&" and section("&C&&Rb", "R") == ""
    assert section("&&&Lx", "C") == "&&"
    assert with_section("&L&C&R", "C", "&R") == "&L&R&C&R"
    assert with_section("&La&Cb&Lc", "L", "z") == "&Cb&Lz"
    assert with_section("abc&Rr", "C", "abc") == "&Rr&Cabc"
    assert with_section("&Ca&", "R", "b") == "&Ca&&Rb"
