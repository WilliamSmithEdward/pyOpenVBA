"""The shape-type tables against what Office answered.

``scripts/measure_shape_types.py`` adds one shape of every type in live
Excel, Word and PowerPoint, records the name each host gave it, saves
the file and reads the preset geometry back out of the markup.  The
three tables in the library are held to that file here.

The mapping is not guessable and it was got wrong once: several preset
names belong to a different shape than their name suggests, so 11 is
`plus` rather than `cross`, 12 is `pentagon` rather than `star5`, and
`star5` is 92.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.powerpoint import PowerPointApplication
from pyopenvba.apps.word import WordApplication
from pyopenvba.shapes._values import PRESET_GEOMETRY, SHAPE_NAMES, WORD_SHAPE_NAMES

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"


def measured() -> dict[str, dict[str, str]]:
    """What each host answered, flattened to one table per question."""
    body: dict[str, object] = json.loads(
        (FIXTURES / "shape_types.json").read_text(encoding="utf-8")
    )
    excel: dict[str, dict[str, str]] = body["excel"]  # type: ignore[assignment]
    return {
        "preset": {number: row["preset"] for number, row in excel.items()},
        "excel": {number: row["name"] for number, row in excel.items()},
        "word": dict(body["word_names"]),  # type: ignore[arg-type]
        "powerpoint": dict(body["powerpoint_names"]),  # type: ignore[arg-type]
    }


def stem_of(name: str) -> str:
    """A default name without the number Office counts with."""
    return re.sub(r"\s+\d+$", "", name)


MEASURED = measured()
TYPES = sorted(MEASURED["preset"], key=int)


@pytest.mark.parametrize("number", TYPES)
def test_the_preset_is_the_one_office_wrote(number: str) -> None:
    assert PRESET_GEOMETRY[int(number)] == MEASURED["preset"][number]


@pytest.mark.parametrize("number", TYPES)
def test_excel_and_powerpoint_name_it_the_same_way(number: str) -> None:
    excel = stem_of(MEASURED["excel"][number])
    assert excel == stem_of(MEASURED["powerpoint"][number]), "the two were measured to agree"
    assert SHAPE_NAMES[int(number)] == excel


@pytest.mark.parametrize("number", TYPES)
def test_word_names_it_its_own_way(number: str) -> None:
    assert WORD_SHAPE_NAMES[int(number)] == stem_of(MEASURED["word"][number])


def test_every_type_this_offers_was_measured() -> None:
    assert set(PRESET_GEOMETRY) == {int(one) for one in TYPES}
    assert set(SHAPE_NAMES) == set(PRESET_GEOMETRY)
    assert set(WORD_SHAPE_NAMES) == set(PRESET_GEOMETRY)


# --- what a new shape is called -----------------------------------------------------


def test_excel_counts_shapes_per_sheet_and_never_back_down() -> None:
    """Rectangle 1, Oval 2, delete both, Rectangle 3, as Excel does."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(
        "Function Names() As String\n"
        "    Dim a As String, b As String, c As String, sh As Object\n"
        "    a = ActiveSheet.Shapes.AddShape(1, 10, 10, 20, 20).Name\n"
        "    b = ActiveSheet.Shapes.AddShape(9, 40, 10, 20, 20).Name\n"
        "    For Each sh In ActiveSheet.Shapes\n"
        "        sh.Delete\n"
        "    Next sh\n"
        "    c = ActiveSheet.Shapes.AddShape(1, 10, 10, 20, 20).Name\n"
        '    Names = a & "|" & b & "|" & c\n'
        "End Function\n",
        name="Naming",
    )
    assert app.run("Names") == "Rectangle 1|Oval 2|Rectangle 3"


def test_a_second_sheet_starts_again() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(
        "Function Names() As String\n"
        "    Dim s As Object\n"
        "    ActiveSheet.Shapes.AddShape 1, 10, 10, 20, 20\n"
        "    Set s = ActiveWorkbook.Worksheets.Add\n"
        '    Names = s.Shapes.AddShape(1, 10, 10, 20, 20).Name\n'
        "End Function\n",
        name="Naming",
    )
    assert app.run("Names") == "Rectangle 1"


def test_powerpoint_counts_the_same_way_per_slide() -> None:
    app = PowerPointApplication()
    app.add_presentation()
    app.add_module(
        "Function Names() As String\n"
        "    Dim a As String, b As String, c As String, sh As Object, sl As Object\n"
        "    Set sl = ActivePresentation.Slides(1)\n"
        "    a = sl.Shapes.AddShape(1, 10, 10, 20, 20).Name\n"
        "    b = sl.Shapes.AddShape(9, 40, 10, 20, 20).Name\n"
        "    For Each sh In sl.Shapes\n"
        "        sh.Delete\n"
        "    Next sh\n"
        "    c = sl.Shapes.AddShape(1, 10, 10, 20, 20).Name\n"
        '    Names = a & "|" & b & "|" & c\n'
        "End Function\n",
        name="Naming",
    )
    assert app.run("Names") == "Rectangle 1|Oval 2|Rectangle 3"


def test_word_counts_the_same_way_per_document() -> None:
    app = WordApplication()
    app.add_document()
    app.add_module(
        "Function Names() As String\n"
        "    Dim a As String, b As String, c As String, sh As Object\n"
        "    a = ActiveDocument.Shapes.AddShape(1, 100, 100, 20, 20).Name\n"
        "    b = ActiveDocument.Shapes.AddShape(9, 140, 100, 20, 20).Name\n"
        "    For Each sh In ActiveDocument.Shapes\n"
        "        sh.Delete\n"
        "    Next sh\n"
        "    c = ActiveDocument.Shapes.AddShape(1, 100, 100, 20, 20).Name\n"
        '    Names = a & "|" & b & "|" & c\n'
        "End Function\n",
        name="Naming",
    )
    assert app.run("Names") == "Rectangle 1|Oval 2|Rectangle 3"


def test_the_geometry_written_is_the_measured_one() -> None:
    """A shape a macro adds carries the preset Office would have used."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(
        "Sub Draw()\n"
        "    ActiveSheet.Shapes.AddShape 11, 10, 10, 20, 20\n"
        "    ActiveSheet.Shapes.AddShape 12, 40, 10, 20, 20\n"
        "    ActiveSheet.Shapes.AddShape 92, 70, 10, 20, 20\n"
        "End Sub\n",
        name="Draw",
    )
    app.run("Draw")
    presets = [one.geometry for one in app.workbook.sheets_[0].shapes_]
    assert presets == ["plus", "pentagon", "star5"]
