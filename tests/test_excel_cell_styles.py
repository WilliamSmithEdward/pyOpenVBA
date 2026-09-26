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
names and the cell xfs with their flags included, and the xf each cell
points at.

shared_edges.xlsx holds cells whose styles both draw the edge between
them; the model reads each edge as Excel did.

tests/fixtures/cell_styles/style_changes.json (scripts/
measure_style_changes.py) does the same for styles a macro changes,
deletes and merges, and for the order Excel's tables write what a session
makes, in new workbooks and in files opened again: one holding entries
nothing uses (planted.xlsx), one holding a font and xfs twice
(duplicates.xlsx), and one listing its custom formats out of order
(renumbered.xlsx).
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
SECTIONS = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs")
#: The workbooks Excel saved, by the section that saved them; the rest come from Files.
SAVED = {"every_style.xlsx": "EveryCell", "every_style_reverse.xlsx": "EveryCellBack", "mixed.xlsx": "Mixed",
         "custom.xlsx": "Custom", "whole.xlsx": "Whole", "fills.xlsx": "Fills", "themed.xlsx": "Themed",
         "calibri.xlsx": "Calibri", **{f"{name}.xlsx": "Files" for name, _, _ in RECORD["files"]}}
CHANGES: dict[str, Any] = json.loads((FIXTURES / "style_changes.json").read_text(encoding="utf-8"))
#: The sections the model refuses where Excel does what it should not guess at: Normal's font size and face,
#: which the sheets' geometry rests on, and a style of a workbook that is not the active one, which Excel
#: changes in the active workbook instead.
REFUSED = {"NormalFont", "Guards2", "Guards3", "Elsewhere"}
#: The workbooks the style changes saved, new ones and files opened again.
CHANGED = {f"{name}.xlsx" for name in ("follow", "normal", "setting", "delete", "merge", "whole_changed", "order",
                                       "quirks", "after_delete", "reach", "delete_kept", "merge_order", "spelled",
                                       "reach_font", "reopen_changed", "planted_saved", "duplicates_saved",
                                       "renumbered_saved")}


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


def _sections(path: Path) -> tuple[dict[str, str], list[tuple[str, str]]]:
    """A saved workbook's stylesheet sections, and the xf each of its first sheet's cells points at."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found: dict[str, str] = {}
    for section in SECTIONS:
        match = re.search(rf"<{section}\b[^>]*/>|<{section}\b[^>]*>.*?</{section}>", styles, re.DOTALL)
        # Excel gives each style a macro adds a random xr:uid, which the model leaves out.
        found[section] = re.sub(r' xr:uid="[^"]*"', "", match.group(0)) if match else ""
    return found, re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet)


@pytest.mark.parametrize("name", sorted(SAVED))
def test_a_saved_workbook_is_written_as_excel_wrote_it(model: tuple[dict[str, str], Path], name: str) -> None:
    assert _sections(model[1] / name) == _sections(FIXTURES / name)


def test_an_edge_two_cells_draw_reads_as_excel_reads_it() -> None:
    app = ExcelApplication.open(FIXTURES / "shared_edges.xlsx", with_vba=False)
    app.add_module(EDGES["reader"] + "Public Function Run() As String\nRun = Edges(ActiveWorkbook.Worksheets(1))\n"
                   "End Function\n", name="Reader")
    ours = str(app.run("Run")).split("|")[: len(EDGES["pairs"])]
    assert [pair for pair, mine, want in zip(EDGES["pairs"], ours, EDGES["answers"], strict=True) if mine != want] == []


@pytest.mark.parametrize("statement", [
    'ActiveWorkbook.Styles.Add ""', 'ActiveWorkbook.Styles("Normal").Font.Size = 12',
    'ActiveWorkbook.Styles("Normal").Font.Name = "Arial"', 'ActiveWorkbook.Styles("Normal").Font.Bold = True',
    'Dim s As Object\nSet s = ActiveWorkbook.Styles("Good")\nWorkbooks.Add\ns.Font.Bold = True',
    'Dim v As Variant\nv = ActiveWorkbook.Styles("Input").Borders.LineStyle',
    'ActiveWorkbook.Styles("Good").Font.ThemeFont = 0'])
def test_what_is_not_modelled_reports_itself(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Public Sub Run()\n{statement}\nEnd Sub\n", name="Probe")
    with pytest.raises(VBAUnsupportedError):
        app.run("Run")


@pytest.fixture(scope="module")
def changes(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str | None], Path]:
    """Every section of the style changes run in the model, None where it reports itself unsupported, and the
    folder its workbooks went to."""
    folder = tmp_path_factory.mktemp("style_changes")
    target, source = str(folder) + os.sep, str(FIXTURES) + os.sep
    answers: dict[str, str | None] = {}
    for name, folders in CHANGES["sections"].items():
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(CHANGES["module"], name="Probe")
        try:
            answers[name] = str(app.run(name, *{"target": (target,), "": (), "both": (source, target)}[folders]))
        except VBAUnsupportedError:
            answers[name] = None
    return answers, folder


@pytest.mark.parametrize("section", [name for name in CHANGES["sections"] if name not in REFUSED])
def test_a_style_change_answers_as_excel_did(changes: tuple[dict[str, str | None], Path], section: str) -> None:
    answer = changes[0][section]
    assert answer is not None
    ours, theirs = re.split(r"[\^~]", answer), re.split(r"[\^~]", CHANGES["runs"][section])
    differing = [(index, [(field, mine, want) for field, (mine, want) in
                          enumerate(zip(one.split("`"), other.split("`"), strict=False)) if mine != want])
                 for index, (one, other) in enumerate(zip(ours, theirs, strict=True)) if one != other]
    assert not differing


@pytest.mark.parametrize("section", sorted(REFUSED))
def test_what_the_model_leaves_to_excel_reports_itself(changes: tuple[dict[str, str | None], Path],
                                                       section: str) -> None:
    assert changes[0][section] is None


@pytest.mark.parametrize("name", sorted(CHANGED))
def test_a_changed_workbook_is_written_as_excel_wrote_it(changes: tuple[dict[str, str | None], Path],
                                                         name: str) -> None:
    assert _sections(changes[1] / name) == _sections(FIXTURES / name)
