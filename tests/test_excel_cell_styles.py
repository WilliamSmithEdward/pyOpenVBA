"""Named cell styles: scripts/measure_cell_styles.py and measure_shared_edges.py, replayed against the model.

tests/fixtures/cell_styles/cell_styles.json keeps the probe module live
Excel ran and every answer it gave: a new workbook's Styles collection
and each style's properties, every built-in style given to a cell in
both orders, styles over formats and formats over styles, styles a macro
adds, whole rows and columns, a macro's fill over a style's, a protected
sheet and the errors. The model runs the same module and answers the
same, read for read.

Each workbook Excel saved on the way is compared with the model's: every
section of the stylesheet as Excel wrote it, the cell styles' xfs and
names included, and each cell's format by what its xf says. Excel numbers
cell xfs in the order it made them where the model numbers them in the
order a save meets them, so those are compared as a set. Excel also
rewrites the xfs of a file it opens, flags recomputed and xfs that then
agree merged, where the model keeps a file's xfs as it read them; for the
workbooks saved after an open the flags are left out.

shared_edges.xlsx holds cells whose styles both draw the edge between
them; the model reads each edge as Excel did.
"""

from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "cell_styles"
DIMENSIONS = Path(__file__).parent / "fixtures" / "dimensions"
RECORD: dict[str, Any] = json.loads((FIXTURES / "cell_styles.json").read_text(encoding="utf-8"))
EDGES: dict[str, Any] = json.loads((FIXTURES / "shared_edges.json").read_text(encoding="utf-8"))
#: The stylesheet sections compared as Excel wrote them.
SECTIONS = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellStyles", "dxfs")
#: The workbooks Excel saved, by the section that saved them; the rest come from Files.
SAVED = {"every_style.xlsx": "EveryCell", "every_style_reverse.xlsx": "EveryCellBack", "mixed.xlsx": "Mixed",
         "custom.xlsx": "Custom", "whole.xlsx": "Whole", "fills.xlsx": "Fills", "themed.xlsx": "Themed",
         "calibri.xlsx": "Calibri", **{f"{name}.xlsx": "Files" for name, _, _ in RECORD["files"]}}
REOPENED = {f"{name}.xlsx" for name, reopen, _ in RECORD["files"] if reopen}


@pytest.fixture(scope="module")
def model(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], Path]:
    """Every section run in the model, and the folder its workbooks went to."""
    folder = tmp_path_factory.mktemp("cell_styles")
    target, source = str(folder) + os.sep, str(FIXTURES) + os.sep
    plan = {"EveryStyle": (), "EveryCell": (target,), "EveryCellBack": (target,), "Mixed": (target,),
            "Custom": (target,), "Whole": (target,), "Fills": (target,), "Protected": (), "Errors": (),
            "Files": (source, target), "Themed": (source, target),
            "Calibri": (str(DIMENSIONS) + os.sep, target)}
    answers: dict[str, str] = {}
    for name, args in plan.items():
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(RECORD["module"], name="Probe")
        answers[name] = str(app.run(name, *args))
    return answers, folder


@pytest.mark.parametrize("section", list(RECORD["runs"]))
def test_a_section_answers_as_excel_did(model: tuple[dict[str, str], Path], section: str) -> None:
    ours, theirs = re.split(r"[\^~]", model[0][section]), re.split(r"[\^~]", RECORD["runs"][section])
    differing = [(index, [(field, mine, want) for field, (mine, want) in
                          enumerate(zip(one.split("`"), other.split("`"), strict=False)) if mine != want])
                 for index, (one, other) in enumerate(zip(ours, theirs, strict=True)) if one != other]
    assert not differing


def _sections(path: Path, *, flags: bool) -> tuple[dict[str, str], list[str], dict[str, str]]:
    """A saved workbook's stylesheet sections, its cell xfs as a set, and each cell's xf, flags left out if asked."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found: dict[str, str] = {}
    for section in SECTIONS:
        match = re.search(rf"<{section}\b[^>]*/>|<{section}\b[^>]*>.*?</{section}>", styles, re.DOTALL)
        # Excel gives each style a macro adds a random xr:uid, which the model leaves out.
        found[section] = re.sub(r' xr:uid="[^"]*"', "", match.group(0)) if match else ""
    block = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", styles, re.DOTALL)
    assert block is not None
    xfs = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", block.group(1), re.DOTALL)
    if not flags:
        xfs = [re.sub(r' apply\w+="1"', "", xf) for xf in xfs]
    cells = {reference: xfs[int(index)] for reference, index in re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet)}
    return found, sorted(set(xfs)), cells


@pytest.mark.parametrize("name", sorted(SAVED))
def test_a_saved_workbook_is_written_as_excel_wrote_it(model: tuple[dict[str, str], Path], name: str) -> None:
    flags = name not in REOPENED
    ours, our_xfs, our_cells = _sections(model[1] / name, flags=flags)
    theirs, their_xfs, their_cells = _sections(FIXTURES / name, flags=flags)
    assert ours == theirs
    assert our_cells == their_cells
    if flags:
        assert our_xfs == their_xfs


def test_an_edge_two_cells_draw_reads_as_excel_reads_it() -> None:
    app = ExcelApplication.open(FIXTURES / "shared_edges.xlsx", with_vba=False)
    app.add_module(EDGES["reader"] + "Public Function Run() As String\nRun = Edges(ActiveWorkbook.Worksheets(1))\n"
                   "End Function\n", name="Reader")
    ours = str(app.run("Run")).split("|")[: len(EDGES["pairs"])]
    assert [pair for pair, mine, want in zip(EDGES["pairs"], ours, EDGES["answers"], strict=True) if mine != want] == []


@pytest.mark.parametrize("statement", [
    'ActiveWorkbook.Styles("Good").Font.Bold = True', 'ActiveWorkbook.Styles("Good").NumberFormat = "0.00"',
    'ActiveWorkbook.Styles("Good").IncludeNumber = True', 'ActiveWorkbook.Styles("Good").Interior.Color = 255',
    'ActiveWorkbook.Styles("Heading 1").Borders(-4107).Weight = 2', 'ActiveWorkbook.Styles("Good").Delete',
    "ActiveWorkbook.Styles.Merge ActiveWorkbook", 'ActiveWorkbook.Styles.Add ""'])
def test_what_is_not_modelled_reports_itself(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Public Sub Run()\n{statement}\nEnd Sub\n", name="Probe")
    with pytest.raises(VBAUnsupportedError):
        app.run("Run")
