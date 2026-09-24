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

from collections.abc import Sequence
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
from pyopenvba.shapes._values import PRESET_GEOMETRY, SHAPE_NAMES, Shape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: What Excel calls a new shape, by the msoShapeType it was asked for.
#: Measured in live Excel by scripts/measure_shape_types.py, which is
#: also where PowerPoint's identical list comes from.
AUTO_SHAPE_NAMES = SHAPE_NAMES

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


class _Drawn(VBAObject):
    """A shape or a part of one, which a sheet protecting its drawing objects keeps a macro from changing."""

    sheet: Worksheet

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, f"setting {self.vba_type_name}.{name}")
        super().vba_set(name, value, args, named, by_ref=by_ref)


class Shapes(VBACollection):
    """The shapes on one worksheet."""

    vba_type_name = "Shapes"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        from pyopenvba.apps.excel._notes import NoteShape

        # A note's box is a shape too, "Comment 1" (tests/fixtures/excel_model/); where it comes among the drawing's
        # shapes is not measured, and the notes follow them here.
        return [*(ShapeObject(self.sheet, one) for one in self.sheet.shapes_),
                *(NoteShape(self.sheet, *key) for key in self.sheet.notes)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for shape in self.sheet.shapes_:
                if shape.name.lower() == index.lower():
                    return ShapeObject(self.sheet, shape)
            for key, note in self.sheet.notes.items():
                if note.name.lower() == index.lower():
                    from pyopenvba.apps.excel._notes import NoteShape

                    return NoteShape(self.sheet, *key)
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
        from pyopenvba.apps.excel._protection import drawing_enforced

        if drawing_enforced(self.sheet):
            # Measured: what AddShape says on a sheet protecting its drawing objects.
            raise error(1004, "The specified value is out of range.")
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
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, "Shapes.AddTextbox")
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
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, "Shapes.AddLine")
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
        from pyopenvba.apps.excel._protection import guard_drawing
        from pyopenvba.shapes._values import ControlInfo

        guard_drawing(self.sheet, "Shapes.AddFormControl")
        wanted = int(to_number(Type)) if Type is not MISSING else 0
        from pyopenvba.apps.excel._radios import creation_group

        first_button = creation_group(self.sheet, wanted, _number(Left), _number(Top), _number(Width), _number(Height))
        shape = self._made(
            kind="formControl",
            geometry="rect",
            stem=FORM_CONTROLS.get(wanted, "Button"),
            box=(Left, Top, Width, Height),
        )
        shape.control = ControlInfo(kind=FORM_CONTROLS.get(wanted, "Button").replace(" ", ""))
        if wanted in (2, 6):
            shape.control.kind = "Drop" if wanted == 2 else "List"
        if wanted in (4, 7, 8, 9):
            shape.control.kind = {4: "GBox", 7: "Radio", 8: "Scroll", 9: "Spin"}[wanted]
        if wanted in (1, 7):
            shape.control.value = -4146
        if wanted == 7:
            shape.control.first_button = first_button
        if wanted in (8, 9):
            shape.control.maximum = 100 if wanted == 8 else 30000
        # A control is numbered in the sheet's block of VML ids, 1025 on in the first sheet to have one, apart from
        # the drawing shapes, which start at 2, and in one sequence with the sheet's notes (tests/fixtures/notes.json).
        # The VML names it by that number too, and Excel will not open a file where the two disagree.
        from pyopenvba.apps.excel._notes import next_shape_id

        shape.shape_id = next_shape_id(self.sheet)
        if wanted == 7:
            from pyopenvba.apps.excel._radios import groups

            groups(self.sheet)
        if wanted == 4:
            from pyopenvba.apps.excel._radios import added_box

            added_box(self.sheet, shape)
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
        """The name Excel would give the next shape.

        One counter per sheet, over every kind: a rectangle and then an
        oval are "Rectangle 1" and "Oval 2".  It does not go back down
        when shapes are deleted -- delete both of those and the next
        rectangle is "Rectangle 3" -- and a new sheet starts again at
        one.  Measured in live Excel, and the same in all three hosts.
        """
        self.sheet.shape_count += 1
        return f"{stem} {self.sheet.shape_count}"

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


class ShapeObject(_Drawn):
    """One shape on a sheet."""

    vba_type_name = "Shape"

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet = sheet
        self.shape = shape

    # -- what it is

    @member
    def ControlFormat(self) -> object:
        if self.shape.control is None:
            raise error(1004, "this shape is not a form control")
        return ControlFormat(self.sheet, self.shape)

    @member
    def DrawingObject(self) -> object:
        from pyopenvba.apps.excel._controls import list_control

        list_control(self.shape)
        return ListDrawingObject(self.sheet, self.shape)

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
        from pyopenvba.apps.excel._protection import guard_drawing
        from pyopenvba.apps.excel._radios import deleting

        guard_drawing(self.sheet, "Shape.Delete")
        deleting(self.sheet, self.shape)
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


class Characters(_Drawn):
    """The characters of a shape's text."""

    vba_type_name = "Characters"

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape
        self.sheet = shape.sheet

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


class ControlFormat(_Drawn):
    """The implemented checkbox, list-item and selection properties."""

    vba_type_name = "ControlFormat"

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet = sheet
        self.shape = shape

    def _numeric_property(self, name: str) -> object:
        from pyopenvba.apps.excel._controls import numeric_control, numeric_maximum

        control = numeric_control(self.shape)
        if name == "page_change" and control.kind == "Spin":
            raise error(438, "spinners do not expose LargeChange")
        value = numeric_maximum(control) if name == "maximum" else getattr(control, name)
        return VBAInt(value, "Long")

    def _set_numeric_property(self, name: str, value: object) -> None:
        from pyopenvba.apps.excel._controls import set_numeric_property

        self._numeric_property(name)
        try:
            set_numeric_property(self.sheet, self.shape, name, int(to_number(value)))
        except ValueError as exc:
            raise error(1004, str(exc)) from exc

    @member
    def Min(self) -> object:
        return self._numeric_property("minimum")

    @setter("Min")
    def _set_min(self, value: object) -> None:
        self._set_numeric_property("minimum", value)

    @member
    def Max(self) -> object:
        return self._numeric_property("maximum")

    @setter("Max")
    def _set_max(self, value: object) -> None:
        self._set_numeric_property("maximum", value)

    @member
    def SmallChange(self) -> object:
        return self._numeric_property("increment")

    @setter("SmallChange")
    def _set_small_change(self, value: object) -> None:
        self._set_numeric_property("increment", value)

    @member
    def LargeChange(self) -> object:
        return self._numeric_property("page_change")

    @setter("LargeChange")
    def _set_large_change(self, value: object) -> None:
        self._set_numeric_property("page_change", value)

    @member
    def LinkedCell(self) -> object:
        assert self.shape.control is not None
        if self.shape.control.kind == "Radio":
            from pyopenvba.apps.excel._radios import link

            return link(self.sheet, self.shape)
        return self.shape.control.linked_cell

    @member
    def ListCount(self) -> object:
        from pyopenvba.apps.excel._controls import list_count

        return VBAInt(list_count(self.sheet, self.shape), "Long")

    @member
    def MultiSelect(self) -> object:
        from pyopenvba.apps.excel._controls import list_control

        control = list_control(self.shape)
        if control.kind == "Drop":
            raise error(1004, "MultiSelect cannot be read on a dropdown")
        return VBAInt({"single": -4142, "multi": -4154, "extended": 3}[control.selection_mode], "Long")

    @setter("MultiSelect")
    def _set_multi_select(self, value: object) -> None:
        from pyopenvba.apps.excel._controls import set_mode

        try:
            set_mode(self.sheet, self.shape, int(to_number(value)))
        except ValueError as exc:
            raise error(1004, str(exc)) from exc

    @member
    def List(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._controls import items
        from pyopenvba.interpreter._values import NULL, VBAArray

        entries = items(self.sheet, self.shape)
        if Index is MISSING:
            return VBAArray([(1, len(entries))], items=list(entries)) if entries else NULL
        index = int(to_number(Index))
        if not 1 <= index <= len(entries):
            raise error(1004, "list index is outside the list")
        return entries[index - 1]

    @setter("List")
    def _set_list(self, Index: object, value: object) -> None:
        if Index is MISSING:
            self._replace_list(value)
            return
        self._edit_items("set", int(to_number(Index)), to_text(value))

    def _replace_list(self, value: object) -> None:
        from itertools import product
        from pyopenvba.interpreter._values import NULL, VBAArray

        if value is NULL:
            raise error(1004, "Null cannot be assigned to a list")
        if isinstance(value, VBAArray):
            if value.size == 0:
                raise error(13, "an empty array cannot be assigned to a list")
            # Excel traverses dimensions in row-major order here, unlike
            # VBA's internal column-major storage.
            entries = [value.get(list(index)) for index in product(
                *(range(low, high + 1) for low, high in value.bounds))]
        else:
            entries = [value]
        self._edit_items("clear")
        for entry in entries:
            if not isinstance(entry, str):
                # Excel retains the successfully written prefix on error.
                raise error(1004, "list array entries must be strings")
            self._edit_items("add", text=entry)

    def _edit_items(self, operation: str, index: int = 0, text: str = "", count: int = 1) -> None:
        from pyopenvba.apps.excel._controls import edit_items

        try:
            edit_items(self.sheet, self.shape, operation, index, text, count)
        except ValueError as exc:
            raise error(1004, str(exc)) from exc

    @method
    def AddItem(self, Text: object = MISSING, Index: object = MISSING) -> object:
        self._edit_items("add", 0 if Index is MISSING else int(to_number(Index)), to_text(Text))
        return EMPTY

    @method
    def RemoveItem(self, Index: object = MISSING, Count: object = MISSING) -> object:
        self._edit_items("remove", int(to_number(Index)), count=1 if Count is MISSING else int(to_number(Count)))
        return EMPTY

    @method
    def RemoveAllItems(self) -> object:
        self._edit_items("clear")
        return EMPTY

    @member
    def ListFillRange(self) -> object:
        from pyopenvba.apps.excel._controls import list_count

        list_count(self.sheet, self.shape)
        assert self.shape.control is not None
        return self.shape.control.list_range

    @setter("ListFillRange")
    def _set_list_fill_range(self, value: object) -> None:
        from pyopenvba.apps.excel._controls import list_control
        from pyopenvba.apps.excel._shape_api import update_control_shape

        list_control(self.shape)
        try:
            update_control_shape(self.sheet, self.shape, linked_cell=None, list_range=to_text(value))
        except ValueError as exc:
            raise error(1004, str(exc)) from exc

    @setter("LinkedCell")
    def _set_linked_cell(self, value: object) -> None:
        from pyopenvba.apps.excel._controls import set_link

        set_link(self.sheet, self.shape, to_text(value))

    @member
    def Value(self) -> object:
        from pyopenvba.apps.excel._controls import value_control, refresh

        if self.shape.control is not None and self.shape.control.selection_mode != "single":
            raise error(1004, "a multi-selection list has no scalar Value")
        value_control(self.shape)
        refresh(self.sheet, self.shape)
        assert self.shape.control is not None
        return VBAInt(self.shape.control.value, "Long")

    @setter("Value")
    def _set_value(self, value: object) -> None:
        from pyopenvba.apps.excel._controls import set_value

        try:
            # Excel treats Boolean True as on and truncates fractional
            # values here, unlike VBA's usual CLng conversion.
            wanted = int(value) if isinstance(value, bool) else int(to_number(value))
            set_value(self.sheet, self.shape, wanted)
        except ValueError as exc:
            raise error(1004, str(exc)) from exc


class ListDrawingObject(_Drawn):
    """The Selected property of a form list's legacy drawing object."""

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet, self.shape = sheet, shape
        self.vba_type_name = "ListBox" if shape.control and shape.control.kind == "List" else "DropDown"

    @member
    def Selected(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._controls import list_count, list_control

        index = int(to_number(Index))
        if not 1 <= index <= list_count(self.sheet, self.shape):
            raise error(1004, "selection index is outside the list")
        control = list_control(self.shape)
        return index in control.selected_indices if control.selection_mode != "single" else index == control.value

    @setter("Selected")
    def _set_selected(self, Index: object, value: object) -> None:
        from pyopenvba.apps.excel._controls import list_control, set_selection, set_value
        from pyopenvba.interpreter._values import to_bool

        index = int(to_number(Index))
        self.Selected(Index)  # validate before changing state
        control = list_control(self.shape)
        if control.selection_mode == "single":
            if to_bool(value):
                set_value(self.sheet, self.shape, index)
            elif control.value == index:
                set_value(self.sheet, self.shape, 0)
        else:
            selected = set(control.selected_indices)
            if to_bool(value):
                selected.add(index)
            else:
                selected.discard(index)
            set_selection(self.sheet, self.shape, sorted(selected))


def _number(value: object) -> float:
    """One of the four measurements a shape is placed by."""
    if value is MISSING or value is None:
        return 0.0
    return float(to_number(value))


def _unsupported(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is real Excel that pyOpenVBA does not implement")
