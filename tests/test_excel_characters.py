"""Range.Characters and the rich text of a cell, replayed from live Excel (scripts/measure_characters.py,
tests/fixtures/characters.json): part of a cell's text read and changed, its font among it, what edits, copies,
sorts and a save do to the runs, and the runs a file read back. Then how tall a row and how wide AutoFit makes a
column for text in runs of fonts (scripts/measure_rich_text_fit.py, tests/fixtures/rich_text_fit.json)."""

from __future__ import annotations

import json
import os
import re
import zipfile
from collections.abc import Collection
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / "fixtures" / "characters.json").read_text(encoding="utf-8"))
FIT: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "rich_text_fit.json").read_text(
    encoding="utf-8"))
#: The fit cases the model does not measure, which report themselves unsupported, as a cell in such a font does.
UNMEASURED_FITS = ("superscript_last", "subscript_last")
#: One case of the fit probe: its sheet, its runs and its answer.
_CASE = re.compile(r'Set ws = wb\.Worksheets\.Add\(After:=wb\.Worksheets\(wb\.Worksheets\.Count\)\)\n'
                   r'ws\.Name = "(\w+)"\n.*?s = s & "#\1=" & Shape\(ws\)\n', re.DOTALL)


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp("characters").resolve()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(portable(RECORD["probe"]), name="Probe")
    return str(app.run("Probe.Probe", str(folder))).split("~"), folder


@pytest.mark.parametrize("index", range(len(RECORD["answers"])),
                         ids=lambda index: RECORD["answers"][index].split(" ")[0][:40])
def test_each_answer_is_excels(replayed: tuple[list[str], Path], index: int) -> None:
    assert replayed[0][index] == RECORD["answers"][index]


def _runs(item: str) -> list[tuple[str, str]]:
    """A string item's runs, as the text and the rPr each one has, a CR LF read as an XML reader reads it in text,
    as LF: Excel writes a line feed so, the model as it writes its other text."""
    return [(text.replace("\r\n", "\n"), font)
            for font, text in re.findall(r"<r>(?:<rPr>(.*?)</rPr>)?<t[^>]*>(.*?)</t></r>", item, re.DOTALL)]


def _rich_cells(sheet: str, shared: str = "") -> dict[str, list[tuple[str, str]]]:
    """Each cell of rich text by its reference, from its shared string or the inline one it holds."""
    items = re.findall(r"<si>(.*?)</si>", shared, re.DOTALL)
    out: dict[str, list[tuple[str, str]]] = {}
    for reference, attributes, body in re.findall(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', sheet, re.DOTALL):
        if 't="s"' in attributes:
            index = re.search(r"<v>(\d+)</v>", body)
            runs = _runs(items[int(index.group(1))]) if index else []
        else:
            runs = _runs(body)
        if runs:
            out[reference] = runs
    return out


def test_the_saved_runs_are_excels(replayed: tuple[list[str], Path]) -> None:
    """The model writes rich text inline where Excel shares it, each cell's runs with the rPr Excel gives them."""
    with zipfile.ZipFile(replayed[1] / "rich.xlsx") as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert _rich_cells(sheet) == _rich_cells(RECORD["sheet"], RECORD["shared_strings"])


def _without(macro: str, names: Collection[str]) -> str:
    """The fit probe without the cases named."""
    return _CASE.sub(lambda found: "" if found.group(1) in names else found.group(0), macro)


def _kept(path: Path) -> dict[str, dict[str, str]]:
    """Each sheet's first row and columns as a saved file keeps them, by the sheet's name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        relationships = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets: dict[str, str] = {}
        for element in re.findall(r"<Relationship\b[^>]*>", relationships):
            identity, target = re.search(r'Id="([^"]+)"', element), re.search(r'Target="([^"]+)"', element)
            if identity and target:
                targets[identity.group(1)] = target.group(1).removeprefix("/xl/")
        out: dict[str, dict[str, str]] = {}
        for element in re.findall(r"<sheet\b[^>]*>", workbook):
            name, identity = re.search(r'name="([^"]+)"', element), re.search(r'r:id="([^"]+)"', element)
            assert name is not None and identity is not None
            sheet = package.read(f"xl/{targets[identity.group(1)]}").decode("utf-8")
            row, columns = re.search(r'<row r="1"[^>]*>', sheet), re.search(r"<cols>.*?</cols>", sheet)
            out[name.group(1)] = {"row": row.group(0) if row else "", "cols": columns.group(0) if columns else ""}
    return out


@pytest.mark.parametrize("probe", FIT["probes"], ids=lambda probe: probe["normal"])
def test_rich_text_fits_as_excels(probe: dict[str, Any], tmp_path: Path) -> None:
    """Each run's font counts in its row's height as a cell in that font would, and AutoFit measures each run in
    its own font, an italic run's reach after it; the file keeps the rows and columns Excel's does."""
    path = tmp_path / "fit.xlsx"
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_without(probe["macro"], UNMEASURED_FITS), name="Probe")
    got = str(app.run("Probe.Probe", str(path)))
    assert got.split("#") == [piece for piece in probe["result"].split("#")
                              if piece.split("=")[0] not in UNMEASURED_FITS]
    kept = _kept(path)
    assert {name: kept[name] for name in probe["file"] if name not in UNMEASURED_FITS} == {
        name: found for name, found in probe["file"].items() if name not in UNMEASURED_FITS}


@pytest.mark.parametrize("name", UNMEASURED_FITS)
def test_a_superscript_or_subscript_run_reports_itself_unsupported(name: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_without(FIT["probes"][0]["macro"], [case for case in FIT["cases"] if case != name]), name="Probe")
    with pytest.raises(VBAUnsupportedError, match="superscript|subscript"):
        app.run("Probe.Probe", str(tmp_path / "fit.xlsx"))
