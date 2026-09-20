"""A worksheet's shapes, as a macro reaches them.

``ActiveSheet.Shapes`` is the collection, ``Shapes(1)`` or
``Shapes("Rectangle 1")`` is one shape, and what the object model
answers about each is what live Excel answered in
``tests/fixtures/excel_model/probes.txt``: the default names it gives a
new shape, the Type numbers, the Single it reports a position in, and
the error a shape that is not there raises.

The state is the same :class:`pyopenvba.shapes.Shape` the file layer
reads and writes, so a macro's edit and a file's contents are never two
models of the same thing.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    VBAInt,
    VBASingle,
    error,
    to_number,
    to_text,
)
from pyopenvba.shapes._values import PRESET_GEOMETRY, Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: What Excel calls a new shape, by the msoShapeType it was asked for.
#: Measured: AddShape(1) is "Rectangle 1", AddShape(5) is "Rounded
#: Rectangle 1", AddShape(9) is "Oval 1".
AUTO_SHAPE_NAMES: dict[int, str] = {
    1: "Rectangle",
    2: "Parallelogram",
    3: "Trapezoid",
    4: "Diamond",
    5: "Rounded Rectangle",
    6: "Octagon",
    7: "Isosceles Triangle",
    8: "Right Triangle",
    9: "Oval",
    10: "Hexagon",
    11: "Cross",
    12: "5-Point Star",
    16: "Can",
    17: "Cube",
}

#: What Excel calls a new one of each other kind.
KIND_NAMES: dict[str, str] = {
    "textBox": "TextBox",
    "line": "Straight Connector",
    "formControl": "Button",
    "picture": "Picture",
    "group": "Group",
}

#: The form controls AddFormControl makes, by xlFormControl.
FORM_CONTROLS: dict[int, str] = {
    0: "Button",
    1: "Check Box",
    2: "Drop Down",
    4: "Group Box",
    5: "Label",
    6: "List Box",
    7: "Option Button",
    8: "Scroll Bar",
    9: "Spinner",
}

#: Excel raises this for a shape that is not there, by name or number.
#: Measured rather than assumed: it is not 1004 but the HRESULT behind
#: "The item with the specified name wasn't found".
ERR_NO_SUCH_SHAPE = -2147024809


class Shapes(VBACollection):
    """The shapes on one worksheet."""

    vba_type_name = "Shapes"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [ShapeObject(self.sheet, one) for one in self.sheet.shapes_]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for shape in self.sheet.shapes_:
                if shape.name.lower() == index.lower():
                    return ShapeObject(self.sheet, shape)
            raise error(ERR_NO_SUCH_SHAPE, f"there is no shape called {index}")
        position = int(to_number(index))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(ERR_NO_SUCH_SHAPE, f"there is no shape number {position}")

    @method
    def AddShape(
        self,
        Type: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
    ) -> object:
        wanted = int(to_number(Type)) if Type is not MISSING else 1
        shape = self._made(
            kind="shape",
            geometry=PRESET_GEOMETRY.get(wanted, "rect"),
            stem=AUTO_SHAPE_NAMES.get(wanted, "Rectangle"),
            box=(Left, Top, Width, Height),
        )
        shape.auto_shape_type = wanted
        return ShapeObject(self.sheet, shape)

    @method
    def AddTextbox(
        self,
        Orientation: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
    ) -> object:
        shape = self._made(
            kind="textBox", geometry="rect", stem="TextBox", box=(Left, Top, Width, Height)
        )
        return ShapeObject(self.sheet, shape)

    @method
    def AddLine(
        self,
        BeginX: object = MISSING,
        BeginY: object = MISSING,
        EndX: object = MISSING,
        EndY: object = MISSING,
    ) -> object:
        """A line, which Excel places by its two ends rather than a box."""
        start_x, start_y = _number(BeginX), _number(BeginY)
        end_x, end_y = _number(EndX), _number(EndY)
        shape = self._made(
            kind="line",
            geometry="line",
            stem="Straight Connector",
            box=(
                min(start_x, end_x),
                min(start_y, end_y),
                abs(end_x - start_x),
                abs(end_y - start_y),
            ),
        )
        return ShapeObject(self.sheet, shape)

    @method
    def AddFormControl(
        self,
        Type: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
    ) -> object:
        from pyopenvba.shapes._values import ControlInfo

        wanted = int(to_number(Type)) if Type is not MISSING else 0
        shape = self._made(
            kind="formControl",
            geometry="rect",
            stem=FORM_CONTROLS.get(wanted, "Button"),
            box=(Left, Top, Width, Height),
        )
        shape.control = ControlInfo(kind=FORM_CONTROLS.get(wanted, "Button").replace(" ", ""))
        return ShapeObject(self.sheet, shape)

    @method
    def Range(self, Index: object = MISSING) -> object:
        """The shapes an argument names, as a range of them."""
        if Index is MISSING:
            raise error(449)
        wanted: list[object] = [*Index] if isinstance(Index, list) else [Index]  # type: ignore[misc]
        items = self.vba_items()
        chosen = [self.vba_lookup(one, items) for one in wanted]
        return ShapeRange([one for one in chosen if isinstance(one, ShapeObject)])

    @method
    def SelectAll(self) -> object:
        return EMPTY

    def _made(
        self,
        *,
        kind: str,
        geometry: str,
        stem: str,
        box: tuple[object, object, object, object] | tuple[float, float, float, float],
    ) -> Shape:
        left, top, width, height = (_number(one) for one in box)
        shape = Shape(
            name=self._free_name(stem),
            kind=kind,
            geometry=geometry,
            left=left,
            top=top,
            width=width,
            height=height,
            shape_id=self._free_id(),
        )
        self.sheet.shapes_.append(shape)
        self.sheet.drawing_changed()
        return shape

    def _free_name(self, stem: str) -> str:
        """The name Excel would give the next shape of this kind.

        Excel counts within the kind and within the sheet, so a second
        rectangle is "Rectangle 2" even when an oval was added between.
        """
        taken = {one.name for one in self.sheet.shapes_}
        number = 1
        while f"{stem} {number}" in taken:
            number += 1
        return f"{stem} {number}"

    def _free_id(self) -> int:
        return max((one.shape_id for one in self.sheet.shapes_), default=1) + 1


class ShapeRange(VBACollection):
    """Several shapes answered as one, which is what Range gives back."""

    vba_type_name = "ShapeRange"

    def __init__(self, shapes: list[ShapeObject]) -> None:
        self.shapes = shapes

    def vba_items(self) -> list[object]:
        return list(self.shapes)

    @method
    def Group(self) -> object:
        raise _unsupported("grouping shapes")

    @method
    def Delete(self) -> object:
        for shape in self.shapes:
            shape.Delete()
        return EMPTY


class ShapeObject(VBAObject):
    """One shape on a sheet."""

    vba_type_name = "Shape"

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet = sheet
        self.shape = shape

    # -- what it is

    @member
    def Name(self) -> object:
        return self.shape.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        self.shape.name = to_text(value)
        self.sheet.drawing_changed()

    @member
    def Type(self) -> object:
        return VBAInt(self.shape.mso_type, "Long")

    @member
    def AutoShapeType(self) -> object:
        """Which AutoShape it is, which a text box answers as a rectangle."""
        return VBAInt(getattr(self.shape, "auto_shape_type", 1), "Long")

    @member
    def ZOrderPosition(self) -> object:
        return VBAInt(self._position(), "Long")

    @member
    def Visible(self) -> object:
        return VBAInt(-1, "Long")

    @member
    def Placement(self) -> object:
        """xlMoveAndSize, which is what a new shape gets."""
        return VBAInt(1, "Long")

    # -- where it is

    @member
    def Left(self) -> object:
        return VBASingle(self.shape.left)

    @setter("Left")
    def _set_left(self, value: object) -> None:
        self.shape.left = _number(value)
        self.sheet.drawing_changed()

    @member
    def Top(self) -> object:
        return VBASingle(self.shape.top)

    @setter("Top")
    def _set_top(self, value: object) -> None:
        self.shape.top = _number(value)
        self.sheet.drawing_changed()

    @member
    def Width(self) -> object:
        return VBASingle(self.shape.width)

    @setter("Width")
    def _set_width(self, value: object) -> None:
        self.shape.width = _number(value)
        self.sheet.drawing_changed()

    @member
    def Height(self) -> object:
        return VBASingle(self.shape.height)

    @setter("Height")
    def _set_height(self, value: object) -> None:
        self.shape.height = _number(value)
        self.sheet.drawing_changed()

    @member
    def TopLeftCell(self) -> object:
        return self._cell_at(self.shape.left, self.shape.top)

    @member
    def BottomRightCell(self) -> object:
        return self._cell_at(
            self.shape.left + self.shape.width, self.shape.top + self.shape.height
        )

    # -- what it does

    @member
    def OnAction(self) -> object:
        return self.shape.macro

    @setter("OnAction")
    def _set_onaction(self, value: object) -> None:
        self.shape.macro = to_text(value)
        self.sheet.drawing_changed()

    @member
    def TextFrame(self) -> object:
        return TextFrame(self)

    @member
    def TextFrame2(self) -> object:
        return TextFrame2(self)

    @member
    def GroupItems(self) -> object:
        if self.shape.kind != "group":
            raise error(ERR_NO_SUCH_SHAPE, "this shape is not a group")
        return GroupItems(self.sheet, self.shape)

    @method
    def Delete(self) -> object:
        self.sheet.shapes_ = [one for one in self.sheet.shapes_ if one is not self.shape]
        self.sheet.drawing_changed()
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:
        return EMPTY

    @method
    def Copy(self) -> object:
        raise _unsupported("copying a shape")

    def _position(self) -> int:
        for index, one in enumerate(self.sheet.shapes_, start=1):
            if one is self.shape:
                return index
        return 1

    def _cell_at(self, left: float, top: float) -> object:
        """The cell a point falls in, which is how Excel answers these."""
        from pyopenvba._a1 import Area
        from pyopenvba.apps.excel._model import Range

        column, across = 1, 0.0
        while across + self.sheet.column_width_points(column) <= left + 1e-9:
            across += self.sheet.column_width_points(column)
            column += 1
        row, down = 1, 0.0
        while down + self.sheet.row_height_points(row) <= top + 1e-9:
            down += self.sheet.row_height_points(row)
            row += 1
        return Range(self.sheet, [Area(row, column, row, column, self.sheet.name)])


class GroupItems(VBACollection):
    """The shapes inside a group."""

    vba_type_name = "GroupShapes"

    def __init__(self, sheet: Worksheet, group: Shape) -> None:
        self.sheet = sheet
        self.group = group

    def vba_items(self) -> list[object]:
        return [ShapeObject(self.sheet, one) for one in self.group.children]


class TextFrame(VBAObject):
    """A shape's text, reached the way a macro reaches it."""

    vba_type_name = "TextFrame"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    @member
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:
        return Characters(self.owner)

    @member
    def AutoSize(self) -> object:
        return VBAInt(0, "Long")


class TextFrame2(VBAObject):
    """The newer way to the same text."""

    vba_type_name = "TextFrame2"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    @member
    def TextRange(self) -> object:
        return Characters(self.owner)

    @member
    def HasText(self) -> object:
        return VBAInt(-1 if self.owner.shape.text else 0, "Long")


class Characters(VBAObject):
    """The characters of a shape's text."""

    vba_type_name = "Characters"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    @member(default=True)
    def Text(self) -> object:
        return self.owner.shape.text

    @setter("Text")
    def _set_text(self, value: object) -> None:
        self.owner.shape.text = to_text(value)
        self.owner.sheet.drawing_changed()

    @member
    def Count(self) -> object:
        return VBAInt(len(self.owner.shape.text), "Long")


def _number(value: object) -> float:
    """One of the four measurements a shape is placed by."""
    if value is MISSING or value is None:
        return 0.0
    return float(to_number(value))


def _unsupported(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is real Excel that pyOpenVBA does not implement")
