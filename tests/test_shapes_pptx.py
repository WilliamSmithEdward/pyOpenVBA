"""Reading a slide's shapes, against what PowerPoint reported.

``scripts/measure_shapes.py powerpoint`` had PowerPoint build the
shapes, asked what it then said about each, and kept both the
presentation and the answers.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from pyopenvba.shapes import Shape
from pyopenvba.shapes._pptx import read_slide, slide_order, with_macro, written

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
DECK = FIXTURES / "powerpoint_shapes.pptm"


def slide() -> str:
    return zipfile.ZipFile(DECK).read("ppt/slides/slide1.xml").decode("utf-8")


def layout() -> str:
    """The layout behind the slide, which places its placeholders."""
    return zipfile.ZipFile(DECK).read("ppt/slideLayouts/slideLayout1.xml").decode("utf-8")


def measured() -> list[dict[str, str]]:
    body = json.loads((FIXTURES / "powerpoint_shapes.json").read_text(encoding="utf-8"))
    return list(body["shapes"])


def by_name() -> dict[str, Shape]:
    return {one.name: one for one in read_slide(slide(), layout())}


def test_every_shape_powerpoint_reported_is_read() -> None:
    assert [one.name for one in read_slide(slide(), layout())] == [one["name"] for one in measured()]


@pytest.mark.parametrize("want", measured(), ids=lambda one: str(one["name"]))
def test_the_shape_is_where_powerpoint_said(want: dict[str, str]) -> None:
    shape = by_name()[want["name"]]
    where = (shape.left, shape.top, shape.width, shape.height)
    assert where == pytest.approx(
        (float(want["left"]), float(want["top"]), float(want["width"]), float(want["height"])),
        abs=0.75,
    )


def test_the_kinds_are_what_powerpoint_calls_them() -> None:
    kinds = {name: shape.mso_type for name, shape in by_name().items()}
    assert kinds == {one["name"]: int(one["type"]) for one in measured()}


def test_the_macro_a_click_runs_is_read() -> None:
    assert by_name()["Rect"].macro == "Clicked"
    assert by_name()["Oval"].macro == ""


def test_a_shapes_text_comes_back() -> None:
    assert by_name()["Rect"].text == "Hello"
    assert by_name()["Box"].text == "Two\nlines"


def test_a_group_holds_its_members() -> None:
    group = by_name()["Group1"]
    assert group.kind == "group"
    assert [one.name for one in group.children] == ["Grouped1", "Grouped2"]


def test_the_slide_order_is_the_list_not_the_part_numbers() -> None:
    presentation = zipfile.ZipFile(DECK).read("ppt/presentation.xml").decode("utf-8")
    assert slide_order(presentation) == ["rId2"]


def test_writing_back_an_untouched_slide_changes_nothing() -> None:
    original = slide()
    assert written(read_slide(original), original) == original


def test_a_moved_shape_leaves_the_others_alone() -> None:
    original = slide()
    shapes = read_slide(original)
    moved = next(one for one in shapes if one.name == "Oval")
    moved.left, moved.top = 300.0, 400.0
    again = {one.name: one for one in read_slide(written(shapes, original))}
    assert (again["Oval"].left, again["Oval"].top) == (300.0, 400.0)
    assert again["Rect"].macro == "Clicked"
    assert again["Rect"].left == 20.0


def test_the_macro_can_be_set_and_cleared() -> None:
    original = slide()
    shapes = read_slide(original)
    oval = next(one for one in shapes if one.name == "Oval")
    rect = next(one for one in shapes if one.name == "Rect")
    oval.macro = "Other"
    rect.macro = ""
    again = {one.name: one for one in read_slide(written(shapes, original))}
    assert again["Oval"].macro == "Other"
    assert again["Rect"].macro == ""


def test_setting_a_macro_puts_the_link_inside_the_name() -> None:
    """The action is a child of p:cNvPr, which is where PowerPoint keeps it."""
    markup = '<p:sp><p:nvSpPr><p:cNvPr id="9" name="X"/><p:cNvSpPr/></p:nvSpPr></p:sp>'
    out = with_macro(markup, "Go")
    assert '<p:cNvPr id="9" name="X"><a:hlinkClick r:id="" action="ppaction://macro?name=Go"/>' in out


def test_a_new_shape_is_written_and_read_back() -> None:
    original = slide()
    shapes = read_slide(original)
    shapes.append(
        Shape(
            name="Added",
            kind="shape",
            geometry="ellipse",
            left=11.0,
            top=22.0,
            width=33.0,
            height=44.0,
            text="New",
            macro="Clicked",
            shape_id=99,
        )
    )
    again = {one.name: one for one in read_slide(written(shapes, original))}
    added = again["Added"]
    assert (added.left, added.top, added.width, added.height) == (11.0, 22.0, 33.0, 44.0)
    assert (added.text, added.macro, added.geometry) == ("New", "Clicked", "ellipse")


def test_a_deleted_shape_is_gone() -> None:
    original = slide()
    shapes = [one for one in read_slide(original) if one.name != "Oval"]
    assert "Oval" not in {one.name for one in read_slide(written(shapes, original))}
