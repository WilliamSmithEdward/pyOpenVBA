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

import re
from collections.abc import Sequence
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.apps.excel._creator import Created
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    VBAArray,
    VBAInt,
    VBASingle,
    error,
    to_bool,
    to_integer,
    to_number,
    to_text,
)
from pyopenvba.shapes._values import (
    PRESET_GEOMETRY,
    SHAPE_NAMES,
    Box,
    ControlInfo,
    Shape,
    TextLook,
    child_space,
    emu,
    to_child,
    to_sheet,
    union,
)

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
#: Placement's xlMoveAndSize, xlMove and xlFreeFloating as the anchor's editAs (tests/fixtures/chart_object_more.json).
PLACEMENTS = ("", "oneCell", "absolute")


class _Drawn(Created):
    """A shape or a part of one, which a sheet protecting its drawing objects keeps a macro from changing."""

    sheet: Worksheet

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, f"setting {self.vba_type_name}.{name}")
        super().vba_set(name, value, args, named, by_ref=by_ref)


class Shapes(VBACollection, Created):
    """The shapes on one worksheet."""

    vba_type_name = "Shapes"

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Parent(self) -> object:
        return self.sheet

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
            # A shape inside a group answers to its name too (tests/fixtures/shape_group_items.json).
            for group in self.sheet.shapes_:
                for shape in group.children:
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
        return _chosen(self, Index)

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

    def new_chart_shape(self, left: float, top: float, width: float, height: float) -> Shape:
        return self._made(kind='chart', geometry='', stem='Chart', box=(left, top, width, height))

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
        # A group's members keep their ids, which no new shape may take, and an ungrouped group keeps its own
        # while its shapes could be regrouped (tests/fixtures/shape_group_items.json).
        drawn = [*self.sheet.shapes_]
        for one in drawn:
            drawn.extend(one.children)
        drawn += [one.former_group for one in drawn if one.former_group is not None]
        return max((one.shape_id for one in drawn), default=1) + 1


class ShapeRange(VBACollection, Created):
    """Several shapes answered as one, which is what Range gives back."""

    vba_type_name = "ShapeRange"

    def __init__(self, shapes: list[ShapeObject]) -> None:
        self.shapes = shapes

    def vba_items(self) -> list[object]:
        return list(self.shapes)

    @member
    def Type(self) -> object:
        kinds = {shape.shape.mso_type for shape in self.shapes}
        return VBAInt(next(iter(kinds)) if len(kinds) == 1 else -2, 'Long')

    def shared(self, name: str) -> object:
        """What every shape in the range reads for ``name``: the one value they share, or for a measurement they
        do not share Excel's mixed value, -2147483648 (tests/fixtures/shape_range_properties.json)."""
        values = [shape.vba_get(name) for shape in self.shapes]
        if all(value == values[0] for value in values[1:]):
            return values[0]
        if isinstance(values[0], VBASingle):
            return VBASingle(-2147483648.0)
        raise VBAUnsupportedError(f"A ShapeRange whose shapes differ in {name} is not implemented")

    def spread(self, name: str, value: object) -> None:
        """A setting given to every shape in the range."""
        for shape in self.shapes:
            shape.vba_set(name, value)

    @member
    def Left(self) -> object:
        return self.shared("Left")

    @setter("Left")
    def _left(self, value: object) -> None:
        self.spread("Left", value)

    @member
    def Top(self) -> object:
        return self.shared("Top")

    @setter("Top")
    def _top(self, value: object) -> None:
        self.spread("Top", value)

    @member
    def Width(self) -> object:
        return self.shared("Width")

    @setter("Width")
    def _width(self, value: object) -> None:
        self.spread("Width", value)

    @member
    def Height(self) -> object:
        return self.shared("Height")

    @setter("Height")
    def _height(self, value: object) -> None:
        self.spread("Height", value)

    @member
    def Rotation(self) -> object:
        return self.shared("Rotation")

    @setter("Rotation")
    def _rotation(self, value: object) -> None:
        self.spread("Rotation", value)

    @member
    def Visible(self) -> object:
        values = {int(to_integer(shape.vba_get("Visible"), "Long")) for shape in self.shapes}
        # Shapes that differ read msoTriStateMixed (tests/fixtures/shape_visible.json).
        return VBAInt(values.pop() if len(values) == 1 else -2, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        self.spread("Visible", value)

    @member
    def AlternativeText(self) -> object:
        return self.shared("AlternativeText")

    @setter("AlternativeText")
    def _alternative_text(self, value: object) -> None:
        self.spread("AlternativeText", value)

    @member
    def Title(self) -> object:
        return self.shared("Title")

    @setter("Title")
    def _title(self, value: object) -> None:
        self.spread("Title", value)

    @member
    def LockAspectRatio(self) -> object:
        return self.shared("LockAspectRatio")

    @setter("LockAspectRatio")
    def _lock_aspect(self, value: object) -> None:
        self.spread("LockAspectRatio", value)

    @member
    def HorizontalFlip(self) -> object:
        return self.shared("HorizontalFlip")

    @member
    def VerticalFlip(self) -> object:
        return self.shared("VerticalFlip")

    # -- tests/fixtures/shape_range_more.json: each shape on its own, -2 where they differ

    def _either(self, name: str) -> object:
        values = {int(to_number(shape.vba_get(name))) for shape in self.shapes}
        return VBAInt(values.pop() if len(values) == 1 else -2, "Long")

    @member
    def Decorative(self) -> object:
        return self._either("Decorative")

    @setter("Decorative")
    def _decorative(self, value: object) -> None:
        self.spread("Decorative", value)

    @member
    def AutoShapeType(self) -> object:
        return self._either("AutoShapeType")

    @setter("AutoShapeType")
    def _auto_shape_type(self, value: object) -> None:
        self.spread("AutoShapeType", value)

    @method
    def IncrementLeft(self, Increment: object = MISSING) -> object:  # noqa: N803 - native named argument
        for shape in self.shapes:
            shape.IncrementLeft(Increment)
        return EMPTY

    @method
    def IncrementTop(self, Increment: object = MISSING) -> object:  # noqa: N803 - native named argument
        for shape in self.shapes:
            shape.IncrementTop(Increment)
        return EMPTY

    @method
    def IncrementRotation(self, Increment: object = MISSING) -> object:  # noqa: N803 - native named argument
        for shape in self.shapes:
            shape.IncrementRotation(Increment)
        return EMPTY

    @method
    def Flip(self, FlipCmd: object = MISSING) -> object:  # noqa: N803 - native named argument
        """Each shape turned over where it stands, its rotation mirrored."""
        for shape in self.shapes:
            shape.Flip(FlipCmd)
        return EMPTY

    @method
    def ZOrder(self, ZOrderCmd: object = MISSING) -> object:  # noqa: N803 - native named argument
        """The shapes moved as one block, in their own order: to the front or back, or one step forward or back,
        where none steps past another shape of the range."""
        command = int(to_integer(ZOrderCmd, "Long"))
        if command not in (0, 1, 2, 3):
            raise error(-2147024809, "The parameter is incorrect.")
        if not self.shapes:
            return EMPTY
        sheet = self.shapes[0].sheet
        for shape in self.shapes:
            shape.guard("ZOrder")
        order = sheet.shapes_
        chosen = [any(one is shape.shape for shape in self.shapes) for one in order]
        if command in (0, 1):
            moved = [one for one, pick in zip(order, chosen, strict=True) if pick]
            rest = [one for one, pick in zip(order, chosen, strict=True) if not pick]
            sheet.shapes_ = rest + moved if command == 0 else moved + rest
        else:
            step = 1 if command == 2 else -1
            places = range(len(order) - 2, -1, -1) if step == 1 else range(1, len(order))
            for place in places:
                if chosen[place] and not chosen[place + step]:
                    order[place], order[place + step] = order[place + step], order[place]
                    chosen[place], chosen[place + step] = chosen[place + step], chosen[place]
        sheet.drawing_changed()
        return EMPTY

    @member
    def Name(self) -> object:
        if len(self.shapes) != 1:
            raise error(-2147024891, "Access denied")
        return self.shapes[0].vba_get("Name")

    @setter("Name")
    def _name(self, value: object) -> None:
        # Excel lets no ShapeRange rename its shapes: one shape is a bad parameter, several access denied.
        raise error(-2147024809 if len(self.shapes) == 1 else -2147024891, "The parameter is incorrect.")

    @method
    def Duplicate(self) -> object:
        """A copy of each shape, after all the sheet's shapes, as one range (tests/fixtures/shape_duplicate.json)."""
        for shape in self.shapes:
            shape.guard("Duplicate")
        return ShapeRange([ShapeObject(shape.sheet, duplicated(shape.sheet, shape.shape)) for shape in self.shapes])

    @method
    def Group(self) -> object:
        """One group of these shapes, as Excel makes it (tests/fixtures/shape_group.json).

        It takes the place of its topmost member and holds them in the order the sheet does, whatever order the
        range named them in. Its box is theirs together, its name "Group" and the sheet's next number, its id the
        next one. Fewer than two shapes raise 1004."""
        if len(self.shapes) < 2:
            raise error(1004, "Grouping needs two shapes or more")
        for shape in self.shapes:
            shape.guard("Group")
            if (shape.shape.source or shape.shape.kind not in ("shape", "textBox", "line") or shape.shape.rotation
                    or shape.shape.flip_h or shape.shape.flip_v):
                raise _unsupported("grouping this shape")
        return ShapeObject(self.shapes[0].sheet, grouped(self.shapes[0].sheet, [one.shape for one in self.shapes]))

    @method
    def Align(self, AlignCmd: object = MISSING, RelativeTo: object = MISSING) -> object:  # noqa: N803 - native named arguments
        """The shapes lined up on the edge or the middle of their box together (tests/fixtures/shape_align.json).

        Excel aligns to the shapes whatever RelativeTo says, takes msoTrue, msoFalse and msoCTrue there and raises
        -2147024809 for anything else, a command it lacks, or a single shape."""
        command = int(to_number(AlignCmd))
        if command not in range(6) or int(to_number(RelativeTo)) not in (-1, 0, 1) or len(self.shapes) < 2:
            raise error(-2147024809, "The parameter is incorrect.")
        for shape in self.shapes:
            shape.guard("Align")
        across = command < 3
        boxes = [_box(one.shape) for one in self.shapes]
        start = min(box[0 if across else 1] for box in boxes)
        end = max(box[0 if across else 1] + box[2 if across else 3] for box in boxes)
        for shape, box in zip(self.shapes, boxes, strict=True):
            size = box[2 if across else 3]
            place = (start, (start + end) / 2 - size / 2, end - size)[command % 3]
            shape.vba_set("Left" if across else "Top", place)
        return EMPTY

    @method
    def Distribute(self, DistributeCmd: object = MISSING, RelativeTo: object = MISSING) -> object:  # noqa: N803 - native named arguments
        """The shapes spaced evenly, in order of their left or top edges, between the first one's start and the
        last one's end, which stay put (tests/fixtures/shape_align.json). Fewer than three shapes, a command Excel
        lacks or a RelativeTo other than msoTrue, msoFalse and msoCTrue raise -2147024809."""
        command = int(to_number(DistributeCmd))
        if command not in (0, 1) or int(to_number(RelativeTo)) not in (-1, 0, 1) or len(self.shapes) < 3:
            raise error(-2147024809, "The parameter is incorrect.")
        for shape in self.shapes:
            shape.guard("Distribute")
        edge, size = (0, 2) if command == 0 else (1, 3)
        ordered = sorted(((_box(one.shape), one) for one in self.shapes), key=lambda pair: pair[0][edge])
        start = ordered[0][0][edge]
        gap = (ordered[-1][0][edge] + ordered[-1][0][size] - start - sum(box[size] for box, _ in ordered)) / (
            len(ordered) - 1)
        place = start
        for box, shape in ordered:
            shape.vba_set("Left" if command == 0 else "Top", place)
            place += box[size] + gap
        return EMPTY

    @method
    def Regroup(self) -> object:
        """The group Ungroup took these shapes out of, made again with the next name and id; shapes that came out
        of no group raise 1004 (tests/fixtures/shape_group_items.json)."""
        former = self.shapes[0].shape.former_group if self.shapes else None
        if former is None or any(one.shape.former_group is not former for one in self.shapes):
            raise error(1004, "These shapes were not ungrouped together")
        sheet = self.shapes[0].sheet
        members = [one for one in sheet.shapes_ if one.former_group is former]
        # After a member moves Excel writes the members inside the old group's child space, where the saved file
        # draws them elsewhere than Excel answers; that is left unmodelled.
        moved = tuple(map(emu, union([_box(one) for one in members]))) != tuple(map(emu, _box(former)))
        if len(members) != len(self.shapes) or former.child_box is not None or moved:
            raise _unsupported("regrouping shapes moved since they were ungrouped")
        return ShapeObject(sheet, grouped(sheet, members))

    @method
    def Ungroup(self) -> object:
        """Every group in the range ungrouped, their shapes answered as one range; shapes that are not groups are
        passed over, and a range with no group raises 1004 (tests/fixtures/shape_group_items.json)."""
        groups = [one for one in self.shapes if one.shape.kind == "group"]
        if not groups:
            raise error(1004, "There is no group to ungroup")
        members: list[ShapeObject] = []
        for group in groups:
            ungrouped = group.Ungroup()
            assert isinstance(ungrouped, ShapeRange)
            members += ungrouped.shapes
        return ShapeRange(members)

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

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def HasChart(self) -> object:
        return VBAInt(-1 if self.shape.kind == 'chart' else 0, 'Long')

    @member
    def Chart(self) -> object:
        from pyopenvba.apps.excel._charts import Chart

        if self.shape.kind != 'chart':
            raise error(1004, 'This shape is not a chart')
        return Chart(self.sheet, self.shape)

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

    @setter("AutoShapeType")
    def _auto_shape_type(self, value: object) -> None:
        """Another AutoShape in the same place, which only swaps the saved preset geometry
        (tests/fixtures/shape_range_more.json)."""
        wanted = int(to_number(value))
        if self.shape.source or self.shape.kind != "shape" or wanted not in PRESET_GEOMETRY:
            raise _unsupported("this AutoShapeType")
        self.shape.auto_shape_type = wanted
        self.restyled().geometry = PRESET_GEOMETRY[wanted]

    @member
    def ZOrderPosition(self) -> object:
        return VBAInt(self._position(), "Long")

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.shape.hidden else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        # Any number but 0 shows the shape, as Excel takes it (tests/fixtures/shape_visible.json).
        wanted = to_bool(value)
        self.restyled().hidden = not wanted

    @member
    def Placement(self) -> object:
        """xlMoveAndSize, xlMove or xlFreeFloating, as the anchor's editAs holds it; a new shape gets
        xlMoveAndSize."""
        return VBAInt(PLACEMENTS.index(self.shape.edit_as) + 1, "Long")

    # -- where it is

    @member
    def Left(self) -> object:
        return VBASingle(self.shape.left)

    @setter("Left")
    def _set_left(self, value: object) -> None:
        self._placed(left=_number(value))

    @member
    def Top(self) -> object:
        return VBASingle(self.shape.top)

    @setter("Top")
    def _set_top(self, value: object) -> None:
        self._placed(top=_number(value))

    def _placed(self, **changes: float) -> None:
        """The shape given a new left, top, width or height (tests/fixtures/shape_group_items.json).

        A new group carries its members with it, scaled with it, and keeps its child space as it was. A member of
        one moves inside that child space, which then becomes the members' boxes together, and the group's box
        follows at the scale it had."""
        shape = self.shape
        group = parent_group(self.sheet, shape)
        if shape.kind == "group" and not shape.source:
            inside = [to_child(shape, _box(one)) for one in shape.children]
            shape.child_box = child_space(shape)
            _set_box(shape, (changes.get("left", shape.left), changes.get("top", shape.top),
                             changes.get("width", shape.width), changes.get("height", shape.height)))
            for one, held in zip(shape.children, inside, strict=True):
                _set_box(one, to_sheet(shape, held))
        elif group is not None and not group.source:
            for name, value in changes.items():
                setattr(shape, name, value)
            space = union([to_child(group, _box(one)) for one in group.children])
            _set_box(group, to_sheet(group, space))
            group.child_box = space
        else:
            for name, value in changes.items():
                setattr(shape, name, value)
        self.sheet.drawing_changed()

    def restyled(self) -> Shape:
        """The shape, for a change a save writes, into the markup a loaded shape keeps as well
        (tests/fixtures/loaded_shape/cases.json)."""
        self.sheet.drawing_changed()
        return self.shape

    @member
    def AlternativeText(self) -> object:
        return self.shape.description

    @setter("AlternativeText")
    def _alternative_text(self, value: object) -> None:
        wanted = to_text(value)
        self.restyled().description = wanted

    @member
    def Title(self) -> object:
        return self.shape.title

    @setter("Title")
    def _title(self, value: object) -> None:
        wanted = to_text(value)
        self.restyled().title = wanted

    @member
    def Decorative(self) -> object:
        return VBAInt(-1 if self.shape.decorative else 0, "Long")

    @setter("Decorative")
    def _decorative(self, value: object) -> None:
        wanted = to_bool(value)
        shape = self.restyled()
        shape.decorative = wanted
        if wanted:
            # A decorative shape needs no alternative text, and Excel takes it away, title and all.
            shape.description = shape.title = ""

    @member
    def Rotation(self) -> object:
        return VBASingle(self.shape.rotation % 360)

    @setter("Rotation")
    def _rotation(self, value: object) -> None:
        wanted = float(VBASingle(_number(value)))
        self.restyled().rotation = wanted

    @method
    def IncrementRotation(self, Increment: object = MISSING) -> object:  # noqa: N803 - native named argument
        self.guard("IncrementRotation")
        wanted = float(VBASingle(_number(Increment)))
        shape = self.restyled()
        # Unlike Rotation, which the file keeps as given, an increment lands inside 0 to 360
        # (tests/fixtures/loaded_shape/cases.json).
        shape.rotation = (shape.rotation + wanted) % 360
        return EMPTY

    @method
    def IncrementLeft(self, Increment: object = MISSING) -> object:  # noqa: N803 - native named argument
        self.guard("IncrementLeft")
        # Excel stops a shape at the sheet's edge.
        self._placed(left=max(0.0, self.shape.left + _number(Increment)))
        return EMPTY

    @method
    def IncrementTop(self, Increment: object = MISSING) -> object:  # noqa: N803 - native named argument
        self.guard("IncrementTop")
        self._placed(top=max(0.0, self.shape.top + _number(Increment)))
        return EMPTY

    @member
    def ID(self) -> object:
        """The id the drawing gives the shape, as Excel answers it (tests/fixtures/shape_duplicate.json)."""
        return VBAInt(self.shape.shape_id, "Long")

    @member
    def HorizontalFlip(self) -> object:
        return VBAInt(-1 if self.shape.flip_h else 0, "Long")

    @member
    def VerticalFlip(self) -> object:
        return VBAInt(-1 if self.shape.flip_v else 0, "Long")

    @method
    def Flip(self, FlipCmd: object = MISSING) -> object:  # noqa: N803 - native named argument
        """Turns the shape over, msoFlipHorizontal (0) or msoFlipVertical (1); doing it twice turns it back
        (tests/fixtures/shape_flip_order.json)."""
        self.guard("Flip")
        command = int(to_integer(FlipCmd, "Long"))
        if command not in (0, 1):
            raise error(-2147024809, "The parameter is incorrect.")
        shape = self.restyled()
        if command == 0:
            shape.flip_h = not shape.flip_h
        else:
            shape.flip_v = not shape.flip_v
        if shape.rotation:
            # Turning a shape over mirrors its turn: 30 degrees becomes 330 (tests/fixtures/shape_duplicate.json).
            shape.rotation = -shape.rotation % 360
        return EMPTY

    @method
    def Duplicate(self) -> object:
        """A copy of the shape, 12 points right and down from it, on top of the others
        (tests/fixtures/shape_duplicate.json)."""
        self.guard("Duplicate")
        return ShapeObject(self.sheet, duplicated(self.sheet, self.shape))

    @method
    def ZOrder(self, ZOrderCmd: object = MISSING) -> object:  # noqa: N803 - native named argument
        """Moves the shape to the front (0) or back (1), or one place forward (2) or back (3), among the sheet's
        shapes, whose order is the drawing's and the Shapes collection's (tests/fixtures/shape_flip_order.json)."""
        self.guard("ZOrder")
        command = int(to_integer(ZOrderCmd, "Long"))
        if command not in (0, 1, 2, 3):
            raise error(-2147024809, "The parameter is incorrect.")
        shapes = self.sheet.shapes_
        place = next(index for index, one in enumerate(shapes) if one is self.shape)
        target = {0: len(shapes) - 1, 1: 0, 2: min(place + 1, len(shapes) - 1), 3: max(place - 1, 0)}[command]
        if target != place:
            shapes.insert(target, shapes.pop(place))
            self.sheet.drawing_changed()
        return EMPTY

    def guard(self, name: str) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, f"Shape.{name}")

    @member
    def LockAspectRatio(self) -> object:
        return VBAInt(-1 if self.shape.lock_aspect else 0, "Long")

    @setter("LockAspectRatio")
    def _lock_aspect(self, value: object) -> None:
        wanted = to_bool(value)
        self.restyled().lock_aspect = wanted

    @member
    def Locked(self) -> object:
        return self.shape.locked

    @setter("Locked")
    def _locked(self, value: object) -> None:
        wanted = to_bool(value)
        self.restyled().locked = wanted

    @member
    def Width(self) -> object:
        return VBASingle(self.shape.width)

    @setter("Width")
    def _set_width(self, value: object) -> None:
        self._placed(width=_number(value))

    @member
    def Height(self) -> object:
        return VBASingle(self.shape.height)

    @setter("Height")
    def _set_height(self, value: object) -> None:
        self._placed(height=_number(value))

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
            raise error(1004, "this shape is not a group")
        return GroupItems(self.sheet, self.shape)

    @method
    def Ungroup(self) -> object:
        """The group's shapes put back in its place, in its order, answered as one range; a shape that is no
        group raises -2147024891 (tests/fixtures/shape_group.json)."""
        if self.shape.kind != "group":
            raise error(-2147024891, "Access denied")
        self.guard("Ungroup")
        if self.shape.source:
            raise _unsupported("ungrouping a group loaded from a file")
        members = self.shape.children
        for one in members:
            one.former_group = self.shape
        place = self._position() - 1
        self.sheet.shapes_[place:place + 1] = members
        self.sheet.drawing_changed()
        return ShapeRange([ShapeObject(self.sheet, one) for one in members])

    @member
    def ParentGroup(self) -> object:
        """The group a shape is in; a shape in none raises 1004 (tests/fixtures/shape_group_items.json)."""
        group = parent_group(self.sheet, self.shape)
        if group is None:
            raise error(1004, "This shape is in no group")
        return ShapeObject(self.sheet, group)

    @method
    def Delete(self) -> object:
        from pyopenvba.apps.excel._protection import guard_drawing
        from pyopenvba.apps.excel._radios import deleting

        guard_drawing(self.sheet, "Shape.Delete")
        group = parent_group(self.sheet, self.shape)
        if group is not None:
            self._left_group(group)
            return EMPTY
        deleting(self.sheet, self.shape)
        if self.shape.kind == 'chart':
            chart = self.sheet.chart_states.pop(self.shape.shape_id, None)
            if chart is not None and chart.part:
                self.sheet.deleted_chart_parts.add(chart.part)
        self.sheet.shapes_ = [one for one in self.sheet.shapes_ if one is not self.shape]
        self.sheet.drawing_changed()
        return EMPTY

    def _left_group(self, group: Shape) -> None:
        """A member deleted from its group (tests/fixtures/shape_group_items.json): the last one left takes the
        group's place, or the group's child space becomes the rest's boxes together, its box following."""
        if group.source:
            raise _unsupported("deleting a shape inside a group loaded from a file")
        group.children = [one for one in group.children if one is not self.shape]
        if len(group.children) == 1:
            place = [index for index, one in enumerate(self.sheet.shapes_) if one is group][0]
            self.sheet.shapes_[place] = group.children[0]
        else:
            space = union([to_child(group, _box(one)) for one in group.children])
            _set_box(group, to_sheet(group, space))
            group.child_box = space
        self.sheet.drawing_changed()

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


class GroupItems(VBACollection, Created):
    """The shapes inside a group."""

    vba_type_name = "GroupShapes"

    def __init__(self, sheet: Worksheet, group: Shape) -> None:
        self.sheet = sheet
        self.group = group

    def vba_items(self) -> list[object]:
        return [ShapeObject(self.sheet, one) for one in self.group.children]

    @member
    def Parent(self) -> object:
        """The group itself, as Excel answers it (tests/fixtures/shape_group_items.json)."""
        return ShapeObject(self.sheet, self.group)

    @method
    def Range(self, Index: object = MISSING) -> object:
        return _chosen(self, Index)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for shape in self.group.children:
                if shape.name.lower() == index.lower():
                    return ShapeObject(self.sheet, shape)
            raise error(ERR_NO_SUCH_SHAPE, f"the group holds no shape called {index}")
        return super().vba_lookup(index, items)


class _TextLayout(Created):
    """What TextFrame and TextFrame2 share of a shape's text layout: its margins and orientation, as
    tests/fixtures/text_frame.json and tests/fixtures/text_frame2.json measured them on a new rectangle and text
    box."""

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    def _look(self) -> TextLook:
        """The shape's text layout, which a loaded shape's markup holds and the model does not read yet."""
        shape = self.owner.shape
        if shape.source or shape.kind not in ("shape", "textBox"):
            raise _unsupported("the text layout of this shape")
        if shape.text_look is None:
            shape.text_look = TextLook(align="l" if shape.kind == "shape" else None,
                                       vert="horz" if shape.kind == "textBox" else None)
        return shape.text_look

    def _changed(self) -> TextLook:
        self.owner.sheet.drawing_changed()
        return self._look()

    def _margin(self, side: int) -> object:
        value = self._look().margins[side]
        return VBASingle((7.2, 3.6, 7.2, 3.6)[side] if value is None else value)

    def _set_margin(self, side: int, value: object) -> None:
        wanted = float(VBASingle(_number(value)))
        if wanted < 0:
            raise error(-2147024809, "The parameter is incorrect.")
        self._changed().margins[side] = wanted

    @member
    def MarginLeft(self) -> object:
        return self._margin(0)

    @setter("MarginLeft")
    def _margin_left(self, value: object) -> None:
        self._set_margin(0, value)

    @member
    def MarginTop(self) -> object:
        return self._margin(1)

    @setter("MarginTop")
    def _margin_top(self, value: object) -> None:
        self._set_margin(1, value)

    @member
    def MarginRight(self) -> object:
        return self._margin(2)

    @setter("MarginRight")
    def _margin_right(self, value: object) -> None:
        self._set_margin(2, value)

    @member
    def MarginBottom(self) -> object:
        return self._margin(3)

    @setter("MarginBottom")
    def _margin_bottom(self, value: object) -> None:
        self._set_margin(3, value)

    @member
    def Orientation(self) -> object:
        return VBAInt(_ORIENTATIONS.index(self._look().vert or "horz") + 1, "Long")

    @setter("Orientation")
    def _orientation(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in range(1, 7):
            raise error(-2147024809, "The parameter is incorrect.")
        look = self._look()
        if wanted == 1 and look.vert is None:
            return
        self._changed().vert = _ORIENTATIONS[wanted - 1]


class TextFrame(_TextLayout):
    """A shape's text, reached the way a macro reaches it."""

    vba_type_name = "TextFrame"

    @member
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:
        return Characters(self.owner)

    @member
    def AutoSize(self) -> object:
        return VBAInt(0, "Long")

    @member
    def AutoMargins(self) -> object:
        # Excel raises 1004 for AutoMargins either way, on a rectangle and a text box alike.
        raise error(1004, "Unable to get the AutoMargins property of the TextFrame class")

    @setter("AutoMargins")
    def _auto_margins(self, value: object) -> None:
        raise error(1004, "Unable to set the AutoMargins property of the TextFrame class")

    @member
    def HorizontalAlignment(self) -> object:
        look = self._look()
        if look.align is None:
            # A text box that names no alignment reads right to left as right-aligned.
            return VBAInt(-4152 if look.rtl else -4131, "Long")
        return VBAInt(_ALIGN_READ[look.align], "Long")

    @setter("HorizontalAlignment")
    def _horizontal_alignment(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in _ALIGN_SET:
            raise _unsupported(f"HorizontalAlignment {wanted}")
        self._changed().align = _ALIGN_SET[wanted]

    @member
    def VerticalAlignment(self) -> object:
        return VBAInt({"t": -4160, "ctr": -4108, "b": -4107}[self._look().anchor], "Long")

    @setter("VerticalAlignment")
    def _vertical_alignment(self, value: object) -> None:
        wanted = int(to_number(value))
        # Justify, distributed and 5 all centre the text.
        anchors = {-4160: "t", -4108: "ctr", -4107: "b", -4130: "ctr", -4117: "ctr", 5: "ctr"}
        if wanted not in anchors:
            raise _unsupported(f"VerticalAlignment {wanted}")
        self._changed().anchor = anchors[wanted]

    @member
    def HorizontalOverflow(self) -> object:
        return VBAInt(("overflow", "clip").index(self._look().horizontal_overflow), "Long")

    @setter("HorizontalOverflow")
    def _horizontal_overflow(self, value: object) -> None:
        self._overflow("horizontal_overflow", ("overflow", "clip"), value)

    @member
    def VerticalOverflow(self) -> object:
        return VBAInt(("overflow", "clip", "ellipsis").index(self._look().vertical_overflow), "Long")

    @setter("VerticalOverflow")
    def _vertical_overflow(self, value: object) -> None:
        self._overflow("vertical_overflow", ("overflow", "clip", "ellipsis"), value)

    def _overflow(self, field_name: str, choices: tuple[str, ...], value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in range(len(choices)):
            raise error(-2147467259, "Unspecified error")
        look = self._look()
        if getattr(look, field_name) != choices[wanted]:
            setattr(self._changed(), field_name, choices[wanted])
            # An AutoShape whose overflow changes is written in full from then on; a text box is not.
            look.full = look.full or self.owner.shape.kind == "shape"

    @member
    def ReadingOrder(self) -> object:
        rtl = self._look().rtl
        return VBAInt(-2 if rtl is None else -5004 if rtl else -5003, "Long")

    @setter("ReadingOrder")
    def _reading_order(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in (-5004, -5003):
            raise error(1004, "Unable to set the ReadingOrder property of the TextFrame class")
        look = self._look()
        if wanted == -5004 and look.rtl is False:
            raise _unsupported("turning text from left to right back to right to left")
        if wanted == -5003 and look.rtl and look.align in ("l", "r"):
            # Back from right to left, Excel mirrors a left or right alignment.
            look.align = "r" if look.align == "l" else "l"
        self._changed().rtl = wanted == -5004


#: HorizontalAlignment as each algn reads, and what each value Excel takes writes: general is left, fill (5)
#: distributed and center across selection (7) centred.
_ALIGN_READ = {"l": -4131, "ctr": -4108, "r": -4152, "just": -4130, "justLow": -4117}
_ALIGN_SET = {-4131: "l", -4108: "ctr", -4152: "r", -4130: "just", -4117: "justLow", 1: "l", 5: "justLow", 7: "ctr"}
#: msoTextOrientation 1 to 6 as bodyPr's vert.
_ORIENTATIONS = ("horz", "vert270", "vert", "eaVert", "mongolianVert", "wordArtVert")


class TextFrame2(_TextLayout):
    """The newer way to the same text (tests/fixtures/text_frame2.json).

    A text box writes each of these settings as given; a rectangle leaves out one set back to its default."""

    vba_type_name = "TextFrame2"

    @member
    def TextRange(self) -> object:
        return Characters(self.owner)

    @member
    def HasText(self) -> object:
        return VBAInt(-1 if self.owner.shape.text else 0, "Long")

    def _given(self, value: str | bool, default: str | bool) -> str | bool | None:
        """What the file holds for a setting: a rectangle drops one set back to its default."""
        return None if value == default and self.owner.shape.kind == "shape" else value

    @member
    def AutoSize(self) -> object:
        return VBAInt(2 if self._look().autofit == "norm" else 0, "Long")

    @setter("AutoSize")
    def _auto_size(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted == 1:
            raise _unsupported("fitting a shape to its text")
        if wanted not in (0, 2):
            raise error(-2147024809, "The parameter is incorrect.")
        given = self._given("norm" if wanted == 2 else "no", "no")
        self._changed().autofit = given if isinstance(given, str) else None

    @member
    def HorizontalAnchor(self) -> object:
        return VBAInt(2 if self._look().anchor_center else 1, "Long")

    @setter("HorizontalAnchor")
    def _horizontal_anchor(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in (1, 2):
            raise error(-2147024809, "The parameter is incorrect.")
        given = self._given(wanted == 2, False)
        self._changed().anchor_center = given if isinstance(given, bool) else None

    @member
    def VerticalAnchor(self) -> object:
        return VBAInt({"t": 1, "ctr": 3, "b": 4}[self._look().anchor], "Long")

    @setter("VerticalAnchor")
    def _vertical_anchor(self, value: object) -> None:
        # The baseline anchors, 2 and 5, raise like any value Excel lacks.
        anchors = {1: "t", 3: "ctr", 4: "b"}
        wanted = int(to_number(value))
        if wanted not in anchors:
            raise error(-2147024809, "The parameter is incorrect.")
        self._changed().anchor = anchors[wanted]

    @member
    def NoTextRotation(self) -> object:
        return VBAInt(-1 if self._look().upright else 0, "Long")

    @setter("NoTextRotation")
    def _no_text_rotation(self, value: object) -> None:
        self._changed().upright = to_bool(value)

    @member
    def WordWrap(self) -> object:
        return VBAInt(0 if self._look().wrap == "none" else -1, "Long")

    @setter("WordWrap")
    def _word_wrap(self, value: object) -> None:
        given = self._given("square" if to_bool(value) else "none", "square")
        self._changed().wrap = given if isinstance(given, str) else None

    @method
    def DeleteText(self) -> object:
        """The text gone, its empty paragraph left without a size, as Excel leaves it."""
        look = self._changed()
        self.owner.shape.text = ""
        look.deleted = True
        return EMPTY


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

    # -- tests/fixtures/control_format_more.json, on new drop-downs, list boxes, check boxes and buttons

    def _settings(self, *kinds: str) -> ControlInfo:
        """The control's settings, for the kinds that have this one; the rest raise 438."""
        from pyopenvba.shapes._xlsx import CONTROL_KINDS

        control = self.shape.control
        kind = CONTROL_KINDS.get(control.kind if control else "Button", "Button")
        if control is None or kind not in kinds:
            if kind in ("Drop", "List", "Checkbox", "Button"):
                raise error(438, "Object doesn't support this property or method")
            raise VBAUnsupportedError(f"This setting of a {kind} control is not implemented")
        if self.shape.source:
            raise VBAUnsupportedError("This setting of a control loaded from a file is not implemented")
        return control

    def _changed(self, *kinds: str) -> ControlInfo:
        from pyopenvba.apps.excel._protection import guard_drawing

        control = self._settings(*kinds)
        guard_drawing(self.sheet, "ControlFormat")
        self.sheet.drawing_changed()
        return control

    @member
    def DropDownLines(self) -> object:
        lines = self._settings("Drop").drop_lines
        return VBAInt(8 if lines is None else lines, "Long")

    @setter("DropDownLines")
    def _drop_down_lines(self, value: object) -> None:
        wanted = int(to_number(value))
        control = self._settings("Drop")
        if not 0 <= wanted <= 32767:
            raise error(1004, "Unable to set the DropDownLines property of the ControlFormat class")
        self._changed("Drop")
        control.drop_lines = wanted

    @member
    def Enabled(self) -> object:
        return not self._settings("Drop", "List", "Checkbox", "Button").disabled

    @setter("Enabled")
    def _enabled(self, value: object) -> None:
        wanted = to_bool(value)
        self._changed("Drop", "List", "Checkbox", "Button").disabled = not wanted

    @member
    def PrintObject(self) -> object:
        from pyopenvba.shapes._xlsx import CONTROL_KINDS

        control = self._settings("Drop", "List", "Checkbox", "Button")
        if control.print_object is None:
            # A button starts out not printing; the other controls print.
            return CONTROL_KINDS.get(control.kind, "Button") != "Button"
        return control.print_object

    @setter("PrintObject")
    def _print_object(self, value: object) -> None:
        wanted = to_bool(value)
        self._changed("Drop", "List", "Checkbox", "Button").print_object = wanted

    @member
    def LockedText(self) -> object:
        return self._settings("Checkbox", "Button").lock_text

    @setter("LockedText")
    def _locked_text(self, value: object) -> None:
        wanted = to_bool(value)
        self._changed("Checkbox", "Button").lock_text = wanted

    @member
    def ListIndex(self) -> object:
        self._settings("Drop", "List")
        return self.Value()

    @setter("ListIndex")
    def _list_index(self, value: object) -> None:
        self._settings("Drop", "List")
        self._set_value(value)

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


#: Every stem a shape the model makes is named from, so a copy can tell a default name from one given.
_STEMS = frozenset({*SHAPE_NAMES.values(), "TextBox", "Straight Connector"})


def duplicated(sheet: Worksheet, shape: Shape) -> Shape:
    """Shape.Duplicate: the copy Excel makes, placed 12 points right and down, after every other shape, with the
    next id. A name of Excel's own making, stem and number, takes the sheet's next one; a name a macro gave is kept
    as it is (tests/fixtures/shape_duplicate.json)."""
    if shape.source or shape.kind not in ("shape", "textBox"):
        raise VBAUnsupportedError("Duplicating this shape is not implemented")
    found = re.fullmatch(r"(.+) \d+", shape.name)
    shapes = Shapes(sheet)
    if found is not None and found.group(1) in _STEMS:
        name = shapes._free_name(found.group(1))  # pyright: ignore[reportPrivateUsage]
    else:
        sheet.shape_count += 1
        name = shape.name
    copy = replace(shape, name=name, shape_id=shapes._free_id(),  # pyright: ignore[reportPrivateUsage]
                   left=shape.left + 12, top=shape.top + 12, children=[], source="", source_at=0, former_group=None)
    sheet.shapes_.append(copy)
    sheet.drawing_changed()
    return copy


def grouped(sheet: Worksheet, shapes: list[Shape]) -> Shape:
    """A new group of these shapes, as Excel makes one (tests/fixtures/shape_group.json).

    It takes the place of its topmost member and holds them in the order the sheet does. Its box is theirs
    together, its name "Group" and the sheet's next number, its id the next one."""
    members = [one for one in sheet.shapes_ if any(one is shape for shape in shapes)]
    made = Shapes(sheet)
    group = Shape(name=made._free_name("Group"), kind="group", shape_id=made._free_id(),  # pyright: ignore[reportPrivateUsage]
                  children=members)
    _set_box(group, union([_box(one) for one in members]))
    place = max(index for index, one in enumerate(sheet.shapes_) if any(one is member for member in members))
    place -= len(members) - 1
    sheet.shapes_ = [one for one in sheet.shapes_ if not any(one is member for member in members)]
    sheet.shapes_.insert(place, group)
    for one in members:
        one.former_group = None
    sheet.drawing_changed()
    return group


def parent_group(sheet: Worksheet, shape: Shape) -> Shape | None:
    """The group on the sheet that holds this shape, if one does."""
    for group in sheet.shapes_:
        if any(one is shape for one in group.children):
            return group
    return None


def _box(shape: Shape) -> Box:
    return (shape.left, shape.top, shape.width, shape.height)


def _set_box(shape: Shape, box: Box) -> None:
    shape.left, shape.top, shape.width, shape.height = box


def _chosen(shapes: VBACollection, index: object) -> ShapeRange:
    """Shapes.Range and GroupShapes.Range: the shapes an argument names, as a range of them."""
    if index is MISSING:
        raise error(449)
    if isinstance(index, VBAArray):
        wanted: list[object] = list(index.elements())
    else:
        wanted = [*index] if isinstance(index, list) else [index]  # type: ignore[misc]
    items = shapes.vba_items()
    chosen = [shapes.vba_lookup(one, items) for one in wanted]
    return ShapeRange([one for one in chosen if isinstance(one, ShapeObject)])


def _number(value: object) -> float:
    """One of the four measurements a shape is placed by."""
    if value is MISSING or value is None:
        return 0.0
    return float(to_number(value))


def _unsupported(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is real Excel that pyOpenVBA does not implement")
