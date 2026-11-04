"""A shape selected on a sheet, and what Selection answers then.

Selection is no Range while a shape is selected: one shape answers as the old drawing object of its kind, several as
DrawingObjects (tests/fixtures/excel_everyday_10.json). A sheet keeps its shapes selected while another sheet is active,
editing cells leaves them selected, and selecting or activating a cell ends it. The cells selected before are still
there, which is where the selection goes when the shape is deleted, and the active cell never moves.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._shapes import Shape, ShapeObject, ShapeRange
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, to_bool

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

_DRAWN = {"rect": "Rectangle", "roundRect": "Rectangle", "ellipse": "Oval", "line": "Line", "arc": "Arc"}


def _type_name(shape: Shape) -> str | None:
    """The drawing object a shape answers as, for those measured; another kind of shape does not join the selection."""
    if shape.kind == "textBox":
        return "TextBox"
    if shape.kind == "picture":
        return "Picture"
    if shape.kind == "formControl" and shape.control is not None:
        return {"Button": "Button", "CheckBox": "CheckBox", "Radio": "OptionButton", "Drop": "DropDown", "List": "ListBox"}.get(shape.control.kind)
    return _DRAWN.get(shape.geometry) if shape.kind in {"autoShape", "shape", "line", "connector"} else None


def select(sheet: Worksheet, shape: Shape, replace: object) -> None:
    """Shape.Select: on the active sheet, instead of what was selected unless Replace is False."""
    if sheet is not sheet.book.active_sheet or _type_name(shape) is None:
        return
    keeping = replace is not MISSING and not to_bool(replace)
    held = [one for one in selected(sheet) if one is not shape] if keeping else []
    sheet.selected_shapes = [*held, shape]
    sheet.active_chart = None


def selected(sheet: Worksheet) -> list[Shape]:
    """The sheet's selected shapes that are still on it."""
    return [shape for shape in sheet.selected_shapes if any(one is shape for one in sheet.shapes_)]


def answer(sheet: Worksheet) -> object | None:
    """What Selection answers for the sheet's selected shapes, if it has any."""
    shapes = selected(sheet)
    if not shapes:
        return None
    if len(shapes) == 1 and shapes[0].kind == "formControl":
        from pyopenvba.apps.excel._old_controls import item_of

        # What the macro recorder writes next goes to this: Selection.OnAction, Selection.Characters.Text.
        return item_of(sheet, shapes[0])
    return SelectedDrawing(sheet, shapes[0]) if len(shapes) == 1 else SelectedDrawings(sheet, shapes)


class SelectedDrawing(ExcelObject):
    """One selected shape, as Selection answers it."""

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet = sheet
        self.shape = ShapeObject(sheet, shape)
        self.vba_type_name = _type_name(shape) or "Shape"

    @member
    def Name(self) -> object:
        return self.shape.vba_get("Name")

    @member
    def Left(self) -> object:
        return self.shape.vba_get("Left")

    @setter("Left")
    def _set_left(self, value: object) -> None:
        self.shape.vba_set("Left", value)

    @member
    def Top(self) -> object:
        return self.shape.vba_get("Top")

    @setter("Top")
    def _set_top(self, value: object) -> None:
        self.shape.vba_set("Top", value)

    @member
    def ShapeRange(self) -> object:
        return ShapeRange([self.shape])

    @method
    def Delete(self) -> object:
        self.shape.Delete()
        return EMPTY


class SelectedDrawings(ExcelObject):
    """Several selected shapes, as Selection answers them."""

    vba_type_name = "DrawingObjects"

    def __init__(self, sheet: Worksheet, shapes: list[Shape]) -> None:
        self.sheet = sheet
        self.shapes = [ShapeObject(sheet, shape) for shape in shapes]

    @member
    def Count(self) -> object:
        return VBAInt(len(self.shapes), "Long")

    @member
    def ShapeRange(self) -> object:
        return ShapeRange(list(self.shapes))
