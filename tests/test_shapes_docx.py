"""Reading a document's shapes, against what Word reported.

``scripts/measure_shapes.py word`` had Word build the shapes, asked
what it then said about each, and kept both the document and the
answers.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from pyopenvba.shapes import Shape
from pyopenvba.shapes._docx import read_document, without, written

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
DOCUMENT = FIXTURES / "word_shapes.docm"


def body() -> str:
    return zipfile.ZipFile(DOCUMENT).read("word/document.xml").decode("utf-8")


def measured() -> list[dict[str, str]]:
    answers = json.loads((FIXTURES / "word_shapes.json").read_text(encoding="utf-8"))
    return list(answers["shapes"])


def floating() -> list[Shape]:
    return [one for one in read_document(body()) if one.placement == "anchored"]


def inline() -> list[Shape]:
    return [one for one in read_document(body()) if one.placement == "inline"]


def test_the_floating_shapes_come_back_in_words_order() -> None:
    """Word lists them by z-order, not by where the runs sit."""
    want = [one["name"] for one in measured() if one["placement"] == "anchored"]
    assert [one.name for one in floating()] == want


def test_one_inline_shape_is_found() -> None:
    assert len(inline()) == 1
    assert inline()[0].kind == "horizontalLine"


@pytest.mark.parametrize(
    "want",
    [one for one in measured() if one["placement"] == "anchored"],
    ids=lambda one: str(one["name"]),
)
def test_the_shape_is_where_word_said(want: dict[str, str]) -> None:
    shape = next(one for one in floating() if one.name == want["name"])
    where = (shape.left, shape.top, shape.width, shape.height)
    assert where == pytest.approx(
        (float(want["left"]), float(want["top"]), float(want["width"]), float(want["height"])),
        abs=0.75,
    )


def test_the_inline_shape_is_the_size_word_said() -> None:
    want = next(one for one in measured() if one["placement"] == "inline")
    shape = inline()[0]
    assert (shape.width, shape.height) == pytest.approx(
        (float(want["width"]), float(want["height"])), abs=0.01
    )


def test_the_kinds_are_what_word_calls_them() -> None:
    kinds = {one.name: one.mso_type for one in floating()}
    want = {one["name"]: int(one["type"]) for one in measured() if one["placement"] == "anchored"}
    assert kinds == want


def test_a_shapes_text_comes_back() -> None:
    shapes = {one.name: one for one in floating()}
    assert shapes["Rect"].text == "Hello"
    assert shapes["Box"].text == "Two\nlines"
    assert shapes["Oval"].text == ""


def test_writing_back_an_untouched_document_changes_nothing() -> None:
    original = body()
    assert written(read_document(original), original) == original


def test_a_moved_shape_reaches_the_markup() -> None:
    original = body()
    shapes = read_document(original)
    moved = next(one for one in shapes if one.name == "Oval")
    moved.left, moved.top = 111.0, 222.0
    out = written(shapes, original)
    again = {one.name: one for one in read_document(out)}
    assert (again["Oval"].left, again["Oval"].top) == (111.0, 222.0)
    assert (again["Rect"].left, again["Rect"].top) == (-32.0, -12.0)


def test_a_resized_shape_changes_both_places_the_size_is_written() -> None:
    original = body()
    shapes = read_document(original)
    resized = next(one for one in shapes if one.name == "Rect")
    resized.width, resized.height = 200.0, 100.0
    out = written(shapes, original)
    again = {one.name: one for one in read_document(out)}
    assert (again["Rect"].width, again["Rect"].height) == (200.0, 100.0)
    # The shape's own transform follows the extent, or Word draws it
    # at the old size inside the new box.
    assert '<a:ext cx="2540000" cy="1270000"/>' in out


def test_a_renamed_shape_keeps_everything_else() -> None:
    original = body()
    shapes = read_document(original)
    next(one for one in shapes if one.name == "Oval").name = "Circle"
    again = {one.name: one for one in read_document(written(shapes, original))}
    assert "Circle" in again
    assert again["Circle"].width == 70.0


def test_a_deleted_shape_takes_its_run_with_it() -> None:
    original = body()
    shapes = read_document(original)
    gone = next(one for one in shapes if one.name == "Oval")
    out = without(gone, original)
    assert "Oval" not in {one.name for one in read_document(out)}
    assert len(read_document(out)) == len(shapes) - 1
