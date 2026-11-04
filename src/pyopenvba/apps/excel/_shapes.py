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

import math
import re
from pathlib import Path
from collections.abc import Callable, Iterator, Sequence
from copy import deepcopy
from dataclasses import fields, replace
from typing import TYPE_CHECKING, ClassVar, cast

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.apps.excel._creator import Created, ItemDefault
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
from pyopenvba.shapes._adjustments import ADJUSTMENTS, DARK_TEXT, EXCEL_ALIASES, EXCEL_GEOMETRY, EXCEL_SHAPE_NAMES
from pyopenvba.shapes._xlsx import text_run_spans
from pyopenvba.shapes._values import (
    Box,
    EMU_PER_POINT,
    ControlInfo,
    Shape,
    TextLook,
    child_space,
    emu,
    points,
    to_child,
    to_sheet,
    union,
)

if TYPE_CHECKING:
    from pyopenvba.shapes._values import FillLine
    from pyopenvba.apps.excel._charts import ChartState
    from pyopenvba.apps.excel._model import Worksheet

#: What Excel calls a new shape, by the msoShapeType it was asked for.
#: Measured in live Excel by scripts/measure_shape_types.py, which is
#: also where PowerPoint's identical list comes from.
AUTO_SHAPE_NAMES = EXCEL_SHAPE_NAMES

#: What Excel calls a new one of each other kind.
KIND_NAMES: dict[str, str] = {
    "textBox": "TextBox",
    "line": "Straight Connector",
    "formControl": "Button",
    "picture": "Picture",
    "group": "Group",
}

#: The form controls AddFormControl makes, by xlFormControl.
#: ConnectionSiteCount for each AutoShape the model makes, and for the other kinds; ShapeStyle for each kind; a
#: button is the one control measured (tests/fixtures/shape_scale.json).
_CONNECTION_SITES = {1: 4, 2: 6, 3: 4, 4: 4, 5: 4, 6: 8, 7: 6, 8: 6, 9: 8, 10: 6, 11: 4, 12: 6, 13: 5, 14: 6, 15: 8,
                     16: 4, 17: 8, 18: 8, 19: 8, 20: 3, 92: 5}
_KIND_SITES = {"textBox": 4, "line": 2, "chart": 4, "button": 4, "group": 0}
_SHAPE_STYLES = {"shape": 9, "textBox": 0, "line": 10009, "chart": 0, "button": 0, "group": 9}
#: Shape.FormControlType for each kind of control the model makes: the XlFormControl constants.
_CONTROL_TYPES = {"Button": 0, "CheckBox": 1, "Drop": 2, "GBox": 4, "Label": 5, "List": 6, "Radio": 7, "Scroll": 8,
                  "Spin": 9}
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


class Shapes(VBACollection, ItemDefault, Created):
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
    def AddPolyline(self, SafeArrayOfPoints: object = MISSING) -> object:  # noqa: N803 - native named argument
        """A freeform through the points, closed when the last is the first; fewer than two points raise 1004
        (tests/fixtures/shape_polyline.json)."""
        found = _points(SafeArrayOfPoints)
        if len(found) < 2:
            raise error(1004, "AddPolyline method of Shapes class failed")
        closed = len(found) > 2 and found[-1] == found[0]
        steps: list[tuple[str, tuple[tuple[int, int], ...]]] = [("moveTo", (found[0],))]
        steps += [("lnTo", (one,)) for one in found[1:len(found) - closed]]
        if closed:
            steps.append(("close", ()))
        return ShapeObject(self.sheet, self._freeform(found, steps))

    @method
    def AddCurve(self, SafeArrayOfPoints: object = MISSING) -> object:  # noqa: N803 - native named argument
        """A freeform of Bezier curves: a start and three points for each curve, or 1004
        (tests/fixtures/shape_polyline.json)."""
        found = _points(SafeArrayOfPoints)
        if len(found) < 4 or (len(found) - 1) % 3:
            raise error(1004, "AddCurve method of Shapes class failed")
        steps: list[tuple[str, tuple[tuple[int, int], ...]]] = [("moveTo", (found[0],))]
        steps += [("cubicBezTo", tuple(found[index:index + 3])) for index in range(1, len(found), 3)]
        return ShapeObject(self.sheet, self._freeform(found, steps))

    def _freeform(self, found: list[tuple[int, int]],
                  steps: list[tuple[str, tuple[tuple[int, int], ...]]]) -> Shape:
        """A new freeform over these points, in EMU: its box their bounds and one EMU more each way, kept on the sheet,
        and its path measured from the bounds' corner."""
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, "Shapes.AddPolyline")
        across, down = min(x for x, _ in found), min(y for _, y in found)
        wide, tall = max(x for x, _ in found) - across + 1, max(y for _, y in found) - down + 1
        shape = Shape(name=self._free_name("Freeform"), kind="freeform", auto_shape_type=138,
                      left=max(across, 0) / EMU_PER_POINT, top=max(down, 0) / EMU_PER_POINT,
                      width=wide / EMU_PER_POINT, height=tall / EMU_PER_POINT, shape_id=self._free_id(),
                      path=[(command, tuple((x - across, y - down) for x, y in points)) for command, points in steps])
        self.sheet.shapes_.append(shape)
        self.sheet.drawing_changed()
        return shape

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
        if wanted not in EXCEL_GEOMETRY:
            raise error(1004, "Application-defined or object-defined error")
        shape = self._made(
            kind="shape",
            geometry=EXCEL_GEOMETRY[wanted],
            stem=AUTO_SHAPE_NAMES[wanted],
            box=(Left, Top, Width, Height),
        )
        shape.auto_shape_type = EXCEL_ALIASES.get(wanted, wanted)
        book = self.sheet.book
        if book.shape_default_markup:
            # After SetShapesDefaultProperties a new AutoShape takes its look, and writes its text layout out.
            shape.fill_line = deepcopy(cast("FillLine | None", book.shape_default_look))
            shape.text_look = TextLook(align="l", full=True)
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
    def AddPicture(
        self,
        Filename: object = MISSING,
        LinkToFile: object = MISSING,
        SaveWithDocument: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
    ) -> object:
        """A picture from a PNG, JPEG or GIF file, kept in the workbook as it is. A linked picture, a BMP, which Excel
        converts, and other kinds of image report themselves (tests/fixtures/shape_add_picture.json,
        tests/fixtures/shape_picture_formats.json)."""
        from pyopenvba.apps.excel._protection import guard_drawing

        self.sheet.book.application.outside("Shapes.AddPicture")
        guard_drawing(self.sheet, "Shapes.AddPicture")
        linked, saved = int(to_number(LinkToFile)), int(to_number(SaveWithDocument))
        if not linked and not saved:
            raise error(1004, "AddPicture method of Shapes class failed")
        if linked:
            raise _unsupported("a linked picture")
        try:
            source = Path(to_text(Filename))
            data = source.read_bytes()
        except OSError:
            raise error(1004, "AddPicture method of Shapes class failed") from None
        wide, tall, kind = _image_points(data)
        # The part keeps the file's own extension where it is an image's, whatever the file holds.
        extension = source.suffix[1:] if source.suffix[1:] in ("png", "jpg", "jpeg", "gif") else kind
        width, height = _number(Width), _number(Height)
        if (width < 0 and width != -1) or (height < 0 and height != -1):
            raise _unsupported("a picture with a negative size")
        natural = width == -1 and height == -1
        # A size of 0 comes out as the smallest there is (1588 EMU), and -1 as the image's own.
        across = wide if width == -1 else max(width, 1588 / EMU_PER_POINT)
        down = tall if height == -1 else max(height, 1588 / EMU_PER_POINT)
        # A picture takes two of the sheet's numbers and two of its ids.
        self.sheet.shape_count += 1
        shape = Shape(name=self._free_name("Picture"), kind="picture", left=_number(Left), top=_number(Top), width=across,
                      height=down, shape_id=self._free_id() + 1, edit_as="oneCell", lock_aspect=natural, image=data,
                      image_extension=extension)
        self.sheet.shapes_.append(shape)
        self.sheet.drawing_changed()
        return ShapeObject(self.sheet, shape)

    @method
    def AddPicture2(
        self,
        Filename: object = MISSING,
        LinkToFile: object = MISSING,
        SaveWithDocument: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
        Compress: object = MISSING,
    ) -> object:
        """AddPicture; Compress changed nothing in a small image (tests/fixtures/shape_add_picture.json)."""
        return self.AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)

    @method
    def AddCallout(
        self,
        Type: object = MISSING,
        Left: object = MISSING,
        Top: object = MISSING,
        Width: object = MISSING,
        Height: object = MISSING,
    ) -> object:
        """A callout, which is the AutoShape its type stands for: 1 and 2 make the same one
        (tests/fixtures/shape_builders.json)."""
        chosen = int(to_number(Type))
        if chosen not in CALLOUTS:
            raise error(1004, "Application-defined or object-defined error")
        return self.vba_call("AddShape", (CALLOUTS[chosen], Left, Top, Width, Height))

    @method
    def BuildFreeform(self, EditingType: object = MISSING, X1: object = MISSING, Y1: object = MISSING) -> object:
        """A FreeformBuilder starting at this point; its editing type changes nothing measured
        (tests/fixtures/shape_builders.json)."""
        if int(to_number(EditingType)) not in (0, 1):
            raise error(1004, "BuildFreeform method of Shapes class failed")
        return FreeformBuilder(self, (emu(_number(X1)), emu(_number(Y1))))

    @method
    def AddConnector(
        self,
        Type: object = MISSING,
        BeginX: object = MISSING,
        BeginY: object = MISSING,
        EndX: object = MISSING,
        EndY: object = MISSING,
    ) -> object:
        """A connector, placed by its two ends, flipped where the end comes before the beginning
        (tests/fixtures/shape_adjustments.json)."""
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, "Shapes.AddConnector")
        chosen = int(to_number(Type))
        if chosen not in CONNECTORS:
            raise error(1004, "Application-defined or object-defined error")
        start_x, start_y = _number(BeginX), _number(BeginY)
        end_x, end_y = _number(EndX), _number(EndY)
        geometry, stem = CONNECTORS[chosen]
        shape = self._made(
            kind="line",
            geometry=geometry,
            stem=stem,
            box=(min(start_x, end_x), min(start_y, end_y), abs(end_x - start_x), abs(end_y - start_y)),
        )
        shape.flip_h, shape.flip_v = end_x < start_x, end_y < start_y
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
        # A control is numbered in the sheet's block of VML ids, 1025 on in the first sheet to have one, apart from
        # the drawing shapes, which start at 2, and in one sequence with the sheet's notes (tests/fixtures/notes.json).
        # The VML names it by that number too, and Excel will not open a file where the two disagree.
        from pyopenvba.apps.excel._notes import next_named

        stem = FORM_CONTROLS.get(wanted, "Button")
        name, identifier = next_named(self.sheet, stem)
        shape = self._made(kind="formControl", geometry="rect", stem=stem, box=(Left, Top, Width, Height), name=name)
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
        shape.shape_id = identifier
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
        name: str | None = None,
    ) -> Shape:
        left, top, width, height = (_number(one) for one in box)
        shape = Shape(
            name=self._free_name(stem) if name is None else name,
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

    @method
    def AddChart2(self, Style: object = MISSING, XlChartType: object = MISSING, Left: object = MISSING,  # noqa: N803
                  Top: object = MISSING, Width: object = MISSING, Height: object = MISSING,  # noqa: N803
                  NewLayout: object = MISSING) -> object:  # noqa: N803
        from pyopenvba.apps.excel._charts import add_chart

        return ShapeObject(self.sheet, add_chart(self.sheet, Style, XlChartType, (Left, Top, Width, Height), NewLayout))

    @method
    def AddChart(self, XlChartType: object = MISSING, Left: object = MISSING, Top: object = MISSING,  # noqa: N803
                 Width: object = MISSING, Height: object = MISSING) -> object:  # noqa: N803
        from pyopenvba.apps.excel._charts import add_chart

        return ShapeObject(self.sheet, add_chart(self.sheet, 2, XlChartType, (Left, Top, Width, Height), False))

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
        # A form control's id is in the VML's block, and drawing shapes go on counting from 2 past it
        # (tests/fixtures/comment_shape_names.json).
        return max((one.shape_id for one in drawn if one.kind != "formControl"), default=1) + 1


class ShapeRange(VBACollection, Created):
    """Several shapes answered as one, which is what Range gives back."""

    vba_type_name = "ShapeRange"

    def __init__(self, shapes: list[ShapeObject]) -> None:
        self.shapes = shapes

    def vba_items(self) -> list[object]:
        return list(self.shapes)

    @member
    def Adjustments(self, Index: object = MISSING) -> object:
        found = RangeAdjustments(self)
        return found if Index is MISSING else found.vba_call("Item", (Index,))

    @setter("Adjustments")
    def _adjustments(self, Index: object, value: object) -> None:
        RangeAdjustments(self).vba_set("Item", value, (Index,))

    @member
    def ConnectorFormat(self) -> object:
        return RangeConnector(self)

    @member
    def Shadow(self) -> object:
        from pyopenvba.apps.excel._shape_effects import RangeShadow

        return RangeShadow(self)

    @member
    def Glow(self) -> object:
        from pyopenvba.apps.excel._shape_effects import RangeGlow

        return RangeGlow(self)

    @member
    def SoftEdge(self) -> object:
        from pyopenvba.apps.excel._shape_effects import RangeSoftEdge

        return RangeSoftEdge(self)

    @member
    def Reflection(self) -> object:
        from pyopenvba.apps.excel._shape_effects import RangeReflection

        return RangeReflection(self)

    @method
    def PickUp(self) -> object:
        """One shape's look; a range of more is out of range (tests/fixtures/shape_painter.json)."""
        if len(self.shapes) != 1:
            raise error(-2147024809, "The specified value is out of range.")
        return self.shapes[0].PickUp()

    @method
    def Apply(self) -> object:
        for one in self.shapes:
            one.Apply()
        return EMPTY

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

    # -- what one shape answers for itself; several answer ID and ZOrderPosition with 70 and what they share of the
    # rest (tests/fixtures/missing_reads.json)

    @member
    def Application(self) -> object:
        return self.shapes[0].sheet.book.application

    @member
    def Parent(self) -> object:
        return self.shapes[0].sheet

    def single(self, name: str) -> object:
        if len(self.shapes) != 1:
            raise error(70, "Permission denied")
        return self.shapes[0].vba_get(name)

    @member
    def ID(self) -> object:
        return self.single("ID")

    @member
    def ZOrderPosition(self) -> object:
        return self.single("ZOrderPosition")

    @member
    def HasChart(self) -> object:
        return self.shared("HasChart")

    @member
    def Child(self) -> object:
        return self.shared("Child")

    @member
    def Connector(self) -> object:
        return self.shared("Connector")

    @member
    def Fill(self) -> object:
        from pyopenvba.apps.excel._shape_format import RangeFill

        return RangeFill(self)

    @member
    def Line(self) -> object:
        from pyopenvba.apps.excel._shape_format import RangeLine

        return RangeLine(self)

    def mixed(self, name: str) -> object:
        """What every shape reads for ``name``, or -2, Excel's mixed value, where they differ
        (tests/fixtures/shape_scale.json)."""
        values = [shape.vba_get(name) for shape in self.shapes]
        return values[0] if all(value == values[0] for value in values[1:]) else VBAInt(-2, "Long")

    @member
    def BlackWhiteMode(self) -> object:
        return self.mixed("BlackWhiteMode")

    @member
    def ConnectionSiteCount(self) -> object:
        return self.mixed("ConnectionSiteCount")

    @member
    def ShapeStyle(self) -> object:
        return self.mixed("ShapeStyle")

    @member
    def BackgroundStyle(self) -> object:
        return self.mixed("BackgroundStyle")

    @member
    def GraphicStyle(self) -> object:
        return self.mixed("GraphicStyle")

    @member
    def Chart(self) -> object:
        # A range holding a shape with no chart raises as that shape does.
        for shape in self.shapes:
            if shape.shape.kind != "chart":
                return shape.vba_get("Chart")
        if len(self.shapes) != 1:
            raise _unsupported("Chart of a ShapeRange of several charts")
        return self.shapes[0].vba_get("Chart")

    @member
    def TextFrame(self) -> object:
        if len(self.shapes) != 1:
            raise _unsupported("TextFrame of a ShapeRange of several shapes")
        return self.shapes[0].vba_get("TextFrame")

    @member
    def TextFrame2(self) -> object:
        if len(self.shapes) != 1:
            raise _unsupported("TextFrame2 of a ShapeRange of several shapes")
        return self.shapes[0].vba_get("TextFrame2")

    @method
    def ScaleWidth(self, Factor: object = MISSING, RelativeToOriginalSize: object = MISSING,  # noqa: N803
                   Scale: object = MISSING) -> object:  # noqa: N803 - native named arguments
        for shape in self.shapes:
            shape.vba_call("ScaleWidth", [Factor, RelativeToOriginalSize, Scale])
        return EMPTY

    @method
    def ScaleHeight(self, Factor: object = MISSING, RelativeToOriginalSize: object = MISSING,  # noqa: N803
                    Scale: object = MISSING) -> object:  # noqa: N803 - native named arguments
        for shape in self.shapes:
            shape.vba_call("ScaleHeight", [Factor, RelativeToOriginalSize, Scale])
        return EMPTY

    @member
    def GroupItems(self) -> object:
        if len(self.shapes) != 1:
            raise _unsupported("GroupItems of a ShapeRange of several shapes")
        return self.shapes[0].vba_get("GroupItems")

    @member
    def ParentGroup(self) -> object:
        if len(self.shapes) != 1:
            raise _unsupported("ParentGroup of a ShapeRange of several shapes")
        return self.shapes[0].vba_get("ParentGroup")

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

    @method
    def Select(self, Replace: object = MISSING) -> object:
        from pyopenvba.apps.excel._selected_shapes import select

        for index, shape in enumerate(self.shapes):
            select(shape.sheet, shape.shape, Replace if index == 0 else False)
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
            # tests/fixtures/core_reads.json
            raise error(-2147024809, 'The specified value is out of range.')
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

        if self.shape.kind == "chart":
            from pyopenvba.apps.excel._charts import ChartObject

            return ChartObject(self.sheet, self.shape)
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
        """Which AutoShape it is, which a text box answers as a rectangle, and a line or connector as mixed."""
        if self.shape.kind == "line":
            return VBAInt(-2, "Long")
        return VBAInt(getattr(self.shape, "auto_shape_type", 1), "Long")

    @setter("AutoShapeType")
    def _auto_shape_type(self, value: object) -> None:
        """Another AutoShape in the same place, which only swaps the saved preset geometry
        (tests/fixtures/shape_range_more.json)."""
        wanted = int(to_number(value))
        if self.shape.source or self.shape.kind != "shape" or wanted not in EXCEL_GEOMETRY:
            raise _unsupported("this AutoShapeType")
        before = AUTO_SHAPE_NAMES.get(getattr(self.shape, "auto_shape_type", 1))
        self.shape.auto_shape_type = wanted
        self.restyled().geometry = EXCEL_GEOMETRY[wanted]
        self.shape.adjustments = None
        # A default name follows the new type, keeping its number (tests/fixtures/shape_adjustments.json).
        found = re.fullmatch(rf"{re.escape(before or '')} (\d+)", self.shape.name)
        if found:
            self.shape.name = f"{AUTO_SHAPE_NAMES[wanted]} {found.group(1)}"

    @method
    def PickUp(self) -> object:
        """Keep this shape's look for Apply, in the application, so another workbook can take it too
        (tests/fixtures/shape_painter.json)."""
        if self.shape.source or self.shape.kind != "shape" or self.shape.text_look is not None:
            raise _unsupported("PickUp from this shape")
        self.sheet.book.application.picked_up_shape = deepcopy(self.shape.fill_line)
        self.sheet.book.application.shape_picked_up = True
        return EMPTY

    @method
    def Apply(self) -> object:
        """Give this shape the look PickUp kept, in place of its own. Its empty text's format is written out."""
        application = self.sheet.book.application
        shape = self.shape
        if (not application.shape_picked_up or shape.source or shape.kind != "shape"
                or shape.text or shape.text_look is not None or shape.geometry in DARK_TEXT):
            raise _unsupported("Apply with nothing picked up, or to this shape")
        shape.fill_line, shape.painted = cast("FillLine | None", deepcopy(application.picked_up_shape)), True
        self.restyled()
        return EMPTY

    @method
    def SetShapesDefaultProperties(self) -> object:
        """This AutoShape's fill, line and effects for later AutoShapes in the workbook, in the theme when it is saved.
        Text boxes and lines keep their own. Text settings carry over too, which is not built, and report themselves
        (tests/fixtures/shape_columns.json)."""
        from pyopenvba.shapes._xlsx import shape_default_markup

        shape = self.shape
        if shape.source or shape.kind != "shape" or shape.text_look is not None or shape.text:
            raise _unsupported("SetShapesDefaultProperties from this shape")
        book = self.sheet.book
        book.shape_default_look = deepcopy(shape.fill_line)
        book.shape_default_markup = shape_default_markup(shape)
        return EMPTY

    @member
    def PictureFormat(self) -> object:
        return PictureFormat(self)

    @member
    def Callout(self) -> object:
        return CalloutFormat(self)

    @member
    def Nodes(self) -> object:
        return ShapeNodes(self)

    @member
    def Vertices(self) -> object:
        return _vba_points(_freeform_points(self))

    @member
    def Adjustments(self, Index: object = MISSING) -> object:
        """The shape's Adjustments, or with an index the one value, as its default Item."""
        found = Adjustments(self)
        return found if Index is MISSING else found.vba_call("Item", (Index,))

    @setter("Adjustments")
    def _adjustments(self, Index: object, value: object) -> None:
        Adjustments(self).vba_set("Item", value, (Index,))

    @member
    def ConnectorFormat(self) -> object:
        return ConnectorFormat(self)

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

    @setter("Placement")
    def _set_placement(self, value: object) -> None:
        # tests/fixtures/shape_cell_moves.json
        wanted = int(to_integer(value, "Long"))
        if wanted not in (1, 2, 3):
            raise VBAUnsupportedError(f"Shape.Placement {wanted} is not implemented")
        self.restyled().edit_as = PLACEMENTS[wanted - 1]

    # -- where it is

    @member
    def Left(self) -> object:
        self.sheet.dims.settle_shapes()
        return VBASingle(self.shape.left)

    @setter("Left")
    def _set_left(self, value: object) -> None:
        self._placed(left=_number(value))

    @member
    def Top(self) -> object:
        self.sheet.dims.settle_shapes()
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

    def placed_by_cells(self, **changes: float) -> None:
        """The shape moved, and sized, with the cells under it (see _shape_anchors)."""
        self._placed(**changes)

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
        self.sheet.dims.settle_shapes()
        return VBASingle(self.shape.width)

    @setter("Width")
    def _set_width(self, value: object) -> None:
        self._placed(width=_number(value))

    @member
    def Height(self) -> object:
        self.sheet.dims.settle_shapes()
        return VBASingle(self.shape.height)

    @setter("Height")
    def _set_height(self, value: object) -> None:
        self._placed(height=_number(value))

    @member
    def TopLeftCell(self) -> object:
        self.sheet.dims.settle_shapes()
        return self._anchored("top left") or self._cell_at(self.shape.left, self.shape.top)

    @member
    def BottomRightCell(self) -> object:
        self.sheet.dims.settle_shapes()
        return self._anchored("bottom right") or self._cell_at(
            self.shape.left + self.shape.width, self.shape.top + self.shape.height
        )

    def _anchored(self, corner: str) -> object:
        """The cell a corner was anchored to when the shape last moved with its cells, hidden or not
        (tests/fixtures/shape_cell_moves.json)."""
        from pyopenvba._a1 import Area
        from pyopenvba.apps.excel._model import Range
        from pyopenvba.apps.excel._shape_anchors import anchored_cell

        found = anchored_cell(self.sheet, self.shape, corner)
        if found is None:
            return None
        column, row = found
        return Range(self.sheet, [Area(row, column, row, column, self.sheet.name)])

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
    def Child(self) -> object:
        """Whether the shape is inside a group (tests/fixtures/missing_reads.json)."""
        return VBAInt(0 if parent_group(self.sheet, self.shape) is None else -1, "Long")

    @member
    def Connector(self) -> object:
        """Whether the shape is a connector, which no shape the model makes but a line is
        (tests/fixtures/missing_reads.json)."""
        if self.shape.source:
            raise _unsupported("Connector of a loaded shape")
        return VBAInt(-1 if self.shape.kind == "line" else 0, "Long")

    @member
    def FormControlType(self) -> object:
        """Which form control the shape is; any other shape raises 1004 (tests/fixtures/missing_reads.json)."""
        if self.shape.control is None:
            raise error(1004, "Unable to get the FormControlType property of the Shape class")
        if self.shape.control.kind not in _CONTROL_TYPES:
            raise _unsupported(f"FormControlType of a {self.shape.control.kind} control")
        return VBAInt(_CONTROL_TYPES[self.shape.control.kind], "Long")

    @member
    def HasSmartArt(self) -> object:
        return VBAInt(0, "Long")

    @member
    def BlackWhiteMode(self) -> object:
        """How the shape prints in black and white: automatically, as every shape the model makes does."""
        if self.shape.source:
            raise _unsupported("BlackWhiteMode of a loaded shape")
        return VBAInt(1, "Long")

    @member
    def Hyperlink(self) -> object:
        """A shape's hyperlink, which no shape the model makes has (tests/fixtures/missing_reads.json)."""
        if self.shape.source:
            raise _unsupported("Hyperlink of a loaded shape")
        raise error(1004, "Unable to get the Hyperlink property of the Shape class")

    def on_sheet(self) -> None:
        # A deleted shape has nothing to answer with (tests/fixtures/shape_range_fill_line.json).
        if not any(shape is self.shape for shape in self.sheet.shapes_):
            raise error(424, "Object required")

    @member
    def Fill(self) -> object:
        from pyopenvba.apps.excel._shape_format import FillFormat

        self.on_sheet()
        return FillFormat(self)

    @member
    def Line(self) -> object:
        from pyopenvba.apps.excel._shape_format import LineFormat

        self.on_sheet()
        return LineFormat(self)

    @member
    def Shadow(self) -> object:
        from pyopenvba.apps.excel._shape_effects import ShadowFormat

        self.on_sheet()
        return ShadowFormat(self)

    @member
    def Glow(self) -> object:
        from pyopenvba.apps.excel._shape_effects import GlowFormat

        self.on_sheet()
        return GlowFormat(self)

    @member
    def SoftEdge(self) -> object:
        from pyopenvba.apps.excel._shape_effects import SoftEdgeFormat

        self.on_sheet()
        return SoftEdgeFormat(self)

    @member
    def Reflection(self) -> object:
        from pyopenvba.apps.excel._shape_effects import ReflectionFormat

        self.on_sheet()
        return ReflectionFormat(self)

    def _measured_kind(self, what: str) -> str:
        """The kind the style reads were measured on: a button stands apart from the other controls; loaded shapes,
        freeforms and other controls are refused (tests/fixtures/shape_scale.json)."""
        shape = self.shape
        kind = "button" if shape.control is not None and shape.control.kind == "Button" else shape.kind
        if shape.source or kind not in _SHAPE_STYLES or (
                kind == "shape" and getattr(shape, "auto_shape_type", 1) not in _CONNECTION_SITES):
            raise _unsupported(f"{what} of this shape")
        return kind

    @member
    def ConnectionSiteCount(self) -> object:
        kind = self._measured_kind("ConnectionSiteCount")
        count = _CONNECTION_SITES[getattr(self.shape, "auto_shape_type", 1)] if kind == "shape" else _KIND_SITES[kind]
        return VBAInt(count, "Long")

    @member
    def ShapeStyle(self) -> object:
        return VBAInt(_SHAPE_STYLES[self._measured_kind("ShapeStyle")], "Long")

    @member
    def BackgroundStyle(self) -> object:
        self._measured_kind("BackgroundStyle")
        return VBAInt(0, "Long")

    @member
    def GraphicStyle(self) -> object:
        self._measured_kind("GraphicStyle")
        return VBAInt(0, "Long")

    @method
    def ScaleWidth(self, Factor: object = MISSING, RelativeToOriginalSize: object = MISSING,  # noqa: N803
                   Scale: object = MISSING) -> object:  # noqa: N803 - native named arguments
        self._scaled("width", Factor, RelativeToOriginalSize, Scale)
        return EMPTY

    @method
    def ScaleHeight(self, Factor: object = MISSING, RelativeToOriginalSize: object = MISSING,  # noqa: N803
                    Scale: object = MISSING) -> object:  # noqa: N803 - native named arguments
        self._scaled("height", Factor, RelativeToOriginalSize, Scale)
        return EMPTY

    def _scaled(self, size: str, factor: object, relative: object, scale: object) -> None:
        """The shape's width or height times ``factor``, a Single, never below 0, kept where Scale anchors it: its top
        left, its middle or its bottom right, never past the sheet's edge, each in whole EMU. A shape that keeps its
        aspect ratio scales both ways. Only a picture scales to its original size (tests/fixtures/shape_scale.json)."""
        if factor is MISSING or relative is MISSING:
            raise error(449, "Argument not optional")
        self.guard("ScaleWidth" if size == "width" else "ScaleHeight")
        wanted = float(VBASingle(_number(factor)))
        anchor = 0 if scale is MISSING else int(to_integer(scale, "Long"))
        if to_bool(relative) or anchor not in (0, 1, 2):
            raise error(-2147024809, "The specified value is out of range.")
        shape = self.shape
        if shape.rotation % 360 or shape.kind == "group" or parent_group(self.sheet, shape) is not None:
            raise _unsupported("scaling a rotated shape, a group or a shape in one")
        self.sheet.dims.settle_shapes()
        sizes = ["width", "height"] if shape.lock_aspect else [size]
        changes: dict[str, float] = {}
        for dimension in sizes:
            start = "left" if dimension == "width" else "top"
            old = getattr(shape, dimension)
            new = points(emu(max(0.0, old * wanted)))
            changes[dimension] = new
            changes[start] = points(emu(max(0.0, getattr(shape, start) - (new - old) * (0.0, 0.5, 1.0)[anchor])))
        self._placed(**changes)

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
        if self.shape.kind == "chart" and Replace is MISSING:
            # Selecting a chart's shape makes it the active chart with its chart area selected
            # (tests/fixtures/active_chart.json).
            from pyopenvba.apps.excel._charts import ChartObject

            ChartObject(self.sheet, self.shape).made_active("area")
            return EMPTY
        from pyopenvba.apps.excel._selected_shapes import select

        select(self.sheet, self.shape, Replace)
        return EMPTY

    @method
    def Copy(self) -> object:
        if self.shape.kind != "chart":
            raise _unsupported("copying a shape")
        chart_copied(self.sheet, self.shape)
        return EMPTY

    @method
    def Cut(self) -> object:
        """A chart goes at once, and can be pasted after (tests/fixtures/chart_copies.json)."""
        if self.shape.kind != "chart":
            raise _unsupported("cutting a shape")
        chart_copied(self.sheet, self.shape)
        self.Delete()
        return EMPTY

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


#: Shapes.AddCallout's types and the AutoShape type each makes.
CALLOUTS = {1: 113, 2: 113, 3: 119, 4: 120}


class FreeformBuilder(Created):
    """Shapes.BuildFreeform (tests/fixtures/shape_builders.json): nodes kept until ConvertToShape, which can be called
    again and makes another shape each time. The result is what AddPolyline and AddCurve make of the same points: a
    line back to the start closes the path and is not written, and a curve back to it closes it after the curve. A
    curve with msoEditingAuto is smoothed by a rule not measured, and reports itself."""

    vba_type_name = "FreeformBuilder"

    def __init__(self, shapes: Shapes, start: tuple[int, int]) -> None:
        self.shapes, self.start = shapes, start
        self.nodes: list[tuple[str, tuple[tuple[int, int], ...]]] = []

    @member
    def Parent(self) -> object:
        return self.shapes

    @member
    def Application(self) -> object:
        return self.shapes.sheet.book.application

    @method
    def AddNodes(self, SegmentType: object = MISSING, EditingType: object = MISSING, X1: object = MISSING,
                 Y1: object = MISSING, X2: object = MISSING, Y2: object = MISSING, X3: object = MISSING,
                 Y3: object = MISSING) -> object:
        segment, editing = int(to_number(SegmentType)), int(to_number(EditingType))
        if segment not in (0, 1) or editing not in (0, 1):
            raise error(-2147024809, "The specified value is out of range.")
        first = (emu(_number(X1)), emu(_number(Y1)))
        if segment == 0:
            # A line takes only its first point.
            self.nodes.append(("lnTo", (first,)))
            return EMPTY
        if editing == 0:
            raise _unsupported("a smoothed curve in BuildFreeform")
        second = first if X2 is MISSING else (emu(_number(X2)), emu(_number(Y2)))
        third = second if X3 is MISSING else (emu(_number(X3)), emu(_number(Y3)))
        self.nodes.append(("cubicBezTo", (first, second, third)))
        return EMPTY

    @method
    def ConvertToShape(self) -> object:
        if not self.nodes:
            raise error(1004, "ConvertToShape method of FreeformBuilder class failed")
        steps = [("moveTo", (self.start,)), *self.nodes]
        found = [point for _, points in steps for point in points]
        last_command, last_points = steps[-1]
        if last_points[-1] == self.start and (last_command == "cubicBezTo" or len(steps) > 2):
            steps = steps + [("close", ())] if last_command == "cubicBezTo" else [*steps[:-1], ("close", ())]
        return ShapeObject(self.shapes.sheet, self.shapes._freeform(found, steps))  # pyright: ignore[reportPrivateUsage]


#: Shapes.AddConnector's types: the geometry each saves and what a new one is called.
CONNECTORS = {1: ("straightConnector1", "Straight Arrow Connector"), 2: ("bentConnector3", "Elbow Connector"),
              3: ("curvedConnector3", "Curved Connector")}


def _adjustable(owner: ShapeObject) -> tuple[tuple[str, int, bool], ...]:
    """The adjustments of a shape the model made."""
    shape = owner.shape
    if shape.source or shape.kind not in ("shape", "textBox", "line", "freeform"):
        raise _unsupported("the adjustments of this shape")
    return ADJUSTMENTS.get(shape.geometry, ())


class Adjustments(Created):
    """A shape's adjustment values (tests/fixtures/shape_adjustments.json): each read as a fraction, or as an angle
    from -180 to 180. Past the end is error 5, and an index below 1 is out of range. Setting one writes them all, unless
    every one is back at its default."""

    vba_type_name = "Adjustments"

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application

    @member
    def Count(self) -> object:
        return VBAInt(len(_adjustable(self.owner)), "Long")

    def values(self) -> tuple[int, ...]:
        kept = self.owner.shape.adjustments
        return kept if kept is not None else tuple(default for _, default, _ in _adjustable(self.owner))

    def place(self, index: object) -> int:
        table = _adjustable(self.owner)
        chosen = int(to_number(index))
        if chosen < 1:
            raise error(-2147024809, "The specified value is out of range.")
        if chosen > len(table):
            raise error(5, "Invalid procedure call or argument")
        return chosen - 1

    @member(default=True, kind="method")
    def Item(self, Index: object = MISSING) -> object:
        at = self.place(Index)
        value = self.values()[at]
        if _adjustable(self.owner)[at][2]:
            degrees = value / 60000
            return VBASingle(degrees - 360 if degrees > 180 else degrees)
        return VBASingle(value / 100000)

    @setter("Item")
    def _item(self, Index: object, value: object) -> None:
        at = self.place(Index)
        table = _adjustable(self.owner)
        number = float(VBASingle(float(to_number(value))))
        written = round(number * 60000) % 21600000 if table[at][2] else round(number * 100000)
        values = list(self.values())
        values[at] = written
        self.owner.restyled()
        self.owner.shape.adjustments = None if values == [default for _, default, _ in table] else tuple(values)


def _jpeg_points(data: bytes) -> tuple[float, float]:
    """A JPEG's own size in points: its frame's pixels at its JFIF density in inches, or at 96 an inch with none."""
    at, across, down = 2, 0.0, 0.0
    wide, tall = 0, 0
    while at + 9 < len(data) and data[at] == 0xFF:
        marker, length = data[at + 1], int.from_bytes(data[at + 2:at + 4], "big")
        if marker == 0xE0 and data[at + 4:at + 9] == b"JFIF\0":
            units = data[at + 11]
            if units == 2:
                raise _unsupported("a JPEG with its density in centimetres")
            if units == 1:
                across, down = int.from_bytes(data[at + 12:at + 14], "big"), int.from_bytes(data[at + 14:at + 16], "big")
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            tall, wide = int.from_bytes(data[at + 5:at + 7], "big"), int.from_bytes(data[at + 7:at + 9], "big")
            break
        at += 2 + length
    if not wide or not tall:
        raise _unsupported("a JPEG this cannot read the size of")
    return wide * 72 / (across or 96), tall * 72 / (down or 96)


def _image_points(data: bytes) -> tuple[float, float, str]:
    """An image's own size in points and the extension its kind takes, from what the file holds."""
    if data[:2] == b"\xff\xd8":
        return (*_jpeg_points(data), "jpg")
    if data[:4] == b"GIF8":
        return (int.from_bytes(data[6:8], "little") * 0.75, int.from_bytes(data[8:10], "little") * 0.75, "gif")
    return (*_png_points(data), "png")


def _png_points(data: bytes) -> tuple[float, float]:
    """A PNG's own size in points: its pixels at its pHYs density, or at 96 a inch with none. Other images are not read."""
    if data[:8] != b"\x89PNG\r\n\x1a\n" or data[12:16] != b"IHDR":
        raise _unsupported("a picture that is not a PNG, JPEG or GIF")
    wide, tall = int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    at = data.find(b"pHYs", 33, 200)
    if at >= 0 and data[at + 12] == 1:
        across, down = int.from_bytes(data[at + 4:at + 8], "big"), int.from_bytes(data[at + 8:at + 12], "big")
        # In whole EMU, rounded down.
        return (int(wide / across * 36000000) / EMU_PER_POINT, int(tall / down * 36000000) / EMU_PER_POINT)
    return wide * 0.75, tall * 0.75


#: Each callout AutoShape as (lines, border, accent).
_CALLOUT_TYPES = {109: (1, True, False), 111: (2, True, False), 112: (3, True, False), 113: (1, False, False),
                  114: (1, False, True), 115: (2, False, True), 116: (3, False, True), 119: (2, False, False),
                  120: (3, False, False), 122: (1, True, True), 123: (2, True, True), 124: (3, True, True)}
_NOT_A_CALLOUT = (-2147024809, "The specified value is out of range.")
#: CalloutFormat.Angle's named angles from the vertical, in degrees.
_CALLOUT_ANGLES = {2: 30.0, 3: 45.0, 4: 60.0, 5: 90.0}


class CalloutFormat(Created):
    """A callout's look (tests/fixtures/shape_columns.json), read from and written to its adjustments: pairs of y and x
    as fractions of the shape, one more pair than it has lines. Gap and Drop move the whole leader; Length and Angle
    its second point, a one-line callout's only where measured. Other shapes have this object, and using it is out of
    range, but for AutoAttach and AutoLength, which read false."""

    vba_type_name = "CalloutFormat"

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def lines(self) -> tuple[int, bool, bool]:
        found = _CALLOUT_TYPES.get(getattr(self.owner.shape, "auto_shape_type", 1)) if self.owner.shape.kind == "shape" else None
        if found is None:
            raise error(*_NOT_A_CALLOUT)
        return found

    def values(self) -> list[int]:
        return list(Adjustments(self.owner).values())

    def put(self, values: list[int]) -> None:
        self.owner.restyled()
        self.owner.shape.adjustments = tuple(values)

    @member
    def Parent(self) -> object:
        return self.owner

    def retype(self, lines: int, border: bool, accent: bool) -> None:
        wanted = next(key for key, one in _CALLOUT_TYPES.items() if one == (lines, border, accent))
        shape = self.owner.shape
        if lines != self.lines()[0] and shape.adjustments is not None:
            raise _unsupported("a callout's Type with its adjustments changed")
        self.owner.restyled()
        shape.auto_shape_type, shape.geometry = wanted, EXCEL_GEOMETRY[wanted]

    @member
    def Type(self) -> object:
        return VBAInt({1: 1, 2: 3, 3: 4}[self.lines()[0]], "Long")

    @setter("Type")
    def _type(self, value: object) -> None:
        lines = {1: 1, 2: 1, 3: 2, 4: 3}.get(int(to_number(value)))
        _, border, accent = self.lines()
        if lines is None:
            raise error(*_NOT_A_CALLOUT)
        self.retype(lines, border, accent)

    @member
    def Border(self) -> object:
        return VBAInt(-1 if self.lines()[1] else 0, "Long")

    @setter("Border")
    def _border(self, value: object) -> None:
        lines, _, accent = self.lines()
        self.retype(lines, int(to_number(value)) != 0, accent)

    @member
    def Accent(self) -> object:
        return VBAInt(-1 if self.lines()[2] else 0, "Long")

    @setter("Accent")
    def _accent(self, value: object) -> None:
        lines, border, _ = self.lines()
        self.retype(lines, border, int(to_number(value)) != 0)

    @member
    def AutoAttach(self) -> object:
        return VBAInt(0, "Long")

    @setter("AutoAttach")
    def _auto_attach(self, value: object) -> None:
        raise error(445, "Object doesn't support this action")

    @member
    def AutoLength(self) -> object:
        return VBAInt(0, "Long")

    @setter("AutoLength")
    def _auto_length(self, value: object) -> None:
        raise error(450, "Wrong number of arguments or invalid property assignment")

    @method
    def AutomaticLength(self) -> object:
        raise error(445, "Object doesn't support this action")

    @member
    def Gap(self) -> object:
        self.lines()
        return VBASingle(max(0.0, -self.values()[1] / 100000 * self.owner.shape.width))

    @setter("Gap")
    def _gap(self, value: object) -> None:
        self.lines()
        values = self.values()
        shift = round(-_number(value) / self.owner.shape.width * 100000) - values[1]
        self.put([one + shift if index % 2 else one for index, one in enumerate(values)])

    @member
    def Drop(self) -> object:
        self.lines()
        return VBASingle(self.values()[0] / 100000 * self.owner.shape.height)

    @setter("Drop")
    def _drop(self, value: object) -> None:
        raise error(450, "Wrong number of arguments or invalid property assignment")

    @member
    def DropType(self) -> object:
        self.lines()
        values = self.values()
        if values[1] >= 0:
            return VBAInt(-2, "Long")
        at = values[0]
        return VBAInt({0: 2, 50000: 3, 100000: 4}.get(at, 1 if 0 < at < 100000 else -2), "Long")

    def dropped(self, wanted: float) -> None:
        values = self.values()
        shift = round(wanted / self.owner.shape.height * 100000) - values[0]
        self.put([one if index % 2 else one + shift for index, one in enumerate(values)])

    @method
    def CustomDrop(self, Drop: object = MISSING) -> object:
        self.lines()
        self.dropped(_number(Drop))
        return EMPTY

    @method
    def PresetDrop(self, DropType: object = MISSING) -> object:
        self.lines()
        height = self.owner.shape.height
        wanted = {1: 3.6, 2: 0.0, 3: height / 2, 4: height}.get(int(to_number(DropType)))
        if wanted is not None:
            self.dropped(wanted)
        return EMPTY

    def span(self, segment: int) -> tuple[float, float]:
        """A leader's segment across and down in points."""
        values = self.values()
        shape = self.owner.shape
        return ((values[segment * 2 + 3] - values[segment * 2 + 1]) / 100000 * shape.width,
                (values[segment * 2 + 2] - values[segment * 2]) / 100000 * shape.height)

    @member
    def Length(self) -> object:
        self.lines()
        return VBASingle(math.hypot(*self.span(0)))

    @method
    def CustomLength(self, Length: object = MISSING) -> object:
        """The new length along the old leader's fractions taken as points, so its direction changes, and a negative one
        turns it back through its start."""
        lines = self.lines()[0]
        values = self.values()
        shape = self.owner.shape
        wanted = _number(Length)
        if lines == 1:
            across, down = values[3] - values[1], values[2] - values[0]
            size = math.hypot(across, down)
            values[3] = values[1] + round(wanted * across / size / shape.width * 100000)
            values[2] = values[0] + round(wanted * down / size / shape.height * 100000)
        elif lines == 2:
            # The first segment's end moves; the leader's far end stays.
            values[3] = values[1] + round(-wanted / shape.width * 100000)
        else:
            raise _unsupported("CustomLength on a three-line callout")
        self.put(values)
        return EMPTY

    @member
    def Angle(self) -> object:
        lines = self.lines()[0]
        across, down = self.span(min(lines, 2) - 1)
        if not across and not down:
            return VBAInt(5, "Long")
        degrees = math.degrees(math.atan2(abs(across), abs(down)))
        return VBAInt(min(_CALLOUT_ANGLES, key=lambda key: abs(_CALLOUT_ANGLES[key] - degrees)), "Long")

    @setter("Angle")
    def _angle(self, value: object) -> None:
        lines = self.lines()[0]
        degrees = _CALLOUT_ANGLES.get(int(to_number(value)))
        if degrees is None:
            raise error(5, "Invalid procedure call or argument")
        if lines != 1:
            raise _unsupported("the Angle of a callout of several lines")
        values = self.values()
        shape = self.owner.shape
        across, down = self.span(0)
        length = math.hypot(across, down)
        turned = math.radians(degrees)
        values[3] = values[1] + round(math.copysign(length * math.sin(turned), across) / shape.width * 100000)
        values[2] = values[0] + round(length * math.cos(turned) / shape.height * 100000)
        self.put(values)


class PictureCrop(Created):
    """A new picture's Crop, uncropped (tests/fixtures/shape_picture_formats.json). Setting a crop moves and sizes the shape
    by rules not worked out, including a bottom crop of the shape's fifth on the first, and reports itself."""

    vba_type_name = "Crop"

    def __init__(self, format_: PictureFormat) -> None:
        self.format = format_

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        raise _unsupported(f"setting Crop.{name}")

    def shape(self) -> Shape:
        self.format.picture()
        return self.format.owner.shape

    @member
    def ShapeWidth(self) -> object:
        return VBASingle(self.shape().width)

    @member
    def ShapeHeight(self) -> object:
        return VBASingle(self.shape().height)

    @member
    def ShapeLeft(self) -> object:
        return VBASingle(self.shape().left)

    @member
    def ShapeTop(self) -> object:
        return VBASingle(self.shape().top)

    @member
    def PictureWidth(self) -> object:
        return VBASingle(self.shape().width)

    @member
    def PictureHeight(self) -> object:
        return VBASingle(self.shape().height)

    @member
    def PictureOffsetX(self) -> object:
        self.shape()
        return VBASingle(0.0)

    @member
    def PictureOffsetY(self) -> object:
        self.shape()
        return VBASingle(0.0)


class PictureFormat(Created):
    """What a new picture reads and sets (tests/fixtures/shape_add_picture.json, tests/fixtures/shape_picture_formats.json);
    other shapes give this object, and using it is out of range. Brightness and contrast run 0 to 1, with 0.5 as the
    file's 0. ColorType's 4, a watermark, sets them both, and any ColorType clears them first."""

    vba_type_name = "PictureFormat"

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def picture(self) -> None:
        if self.owner.shape.kind != "picture":
            raise error(-2147024809, "The specified value is out of range.")
        if self.owner.shape.source:
            raise _unsupported("the PictureFormat of a loaded picture")

    @member
    def Parent(self) -> object:
        return self.owner

    def lum(self, at: int) -> object:
        self.picture()
        found = self.owner.shape.picture_lum
        return VBASingle(0.5 if found is None else (found[at] / 100000 + 1) / 2)

    def set_lum(self, at: int, wanted: float) -> None:
        self.picture()
        if not 0 <= wanted <= 1:
            raise error(-2147024809, "The specified value is out of range.")
        self.put_lum(at, wanted)

    def put_lum(self, at: int, wanted: float) -> None:
        shape = self.owner.shape
        if shape.picture_clear is not None:
            raise _unsupported("brightness or contrast on a picture with a transparent colour")
        values = list(shape.picture_lum or (0, 0))
        values[at] = round((wanted - 0.5) * 200000)
        shape.picture_lum = (values[0], values[1])
        self.owner.restyled()

    def stepped(self, at: int, value: object) -> None:
        self.picture()
        now = float(self.lum(at))  # pyright: ignore[reportArgumentType]
        self.put_lum(at, min(1.0, max(0.0, now + float(VBASingle(float(to_number(value)))))))

    @member
    def Brightness(self) -> object:
        return self.lum(0)

    @setter("Brightness")
    def _brightness(self, value: object) -> None:
        self.set_lum(0, float(VBASingle(float(to_number(value)))))

    @member
    def Contrast(self) -> object:
        return self.lum(1)

    @setter("Contrast")
    def _contrast(self, value: object) -> None:
        self.set_lum(1, float(VBASingle(float(to_number(value)))))

    @method
    def IncrementBrightness(self, Increment: object = MISSING) -> object:
        self.stepped(0, Increment)
        return EMPTY

    @method
    def IncrementContrast(self, Increment: object = MISSING) -> object:
        self.stepped(1, Increment)
        return EMPTY

    @member
    def ColorType(self) -> object:
        self.picture()
        return VBAInt(self.owner.shape.picture_colour_type, "Long")

    @setter("ColorType")
    def _color_type(self, value: object) -> None:
        self.picture()
        wanted = int(to_number(value))
        if not 1 <= wanted <= 4:
            raise error(-2147024809, "The specified value is out of range.")
        shape = self.owner.shape
        shape.picture_colour_type = wanted
        shape.picture_lum = (70000, -70000) if wanted == 4 else None
        self.owner.restyled()

    @member
    def TransparentBackground(self) -> object:
        shape = self.owner.shape
        if shape.kind != "picture":
            # Measured: a line reads true and an AutoShape false, though neither has a picture.
            return VBAInt(-1 if shape.kind == "line" else 0, "Long")
        return VBAInt(-1 if shape.picture_transparent else 0, "Long")

    @setter("TransparentBackground")
    def _transparent_background(self, value: object) -> None:
        self.picture()
        wanted = bool(to_number(value))
        if not wanted and self.owner.shape.picture_clear is not None:
            raise _unsupported("turning a picture's transparent colour off")
        shape = self.owner.shape
        shape.picture_transparent, shape.picture_fill = wanted, "clear" if wanted else "solid"
        self.owner.restyled()

    @member
    def TransparencyColor(self) -> object:
        self.picture()
        chosen = self.owner.shape.picture_clear
        if chosen is None:
            return VBAInt(-2147483648, "Long")
        red, green, blue = (int(chosen[at:at + 2], 16) for at in (0, 2, 4))
        return VBAInt(red | green << 8 | blue << 16, "Long")

    @setter("TransparencyColor")
    def _transparency_color(self, value: object) -> None:
        self.picture()
        shape = self.owner.shape
        if shape.picture_lum is not None or shape.picture_colour_type != 1:
            raise _unsupported("a transparent colour on a picture with its brightness, contrast or colour changed")
        colour = int(to_number(value)) & 0xFFFFFF
        shape.picture_clear = f"{colour & 255:02X}{colour >> 8 & 255:02X}{colour >> 16 & 255:02X}"
        self.owner.restyled()

    @member
    def Crop(self) -> object:
        return PictureCrop(self)

    @setter("CropLeft")
    def _crop_left(self, value: object) -> None:
        raise _unsupported("setting a picture's crop")

    @setter("CropRight")
    def _crop_right(self, value: object) -> None:
        raise _unsupported("setting a picture's crop")

    @setter("CropTop")
    def _crop_top(self, value: object) -> None:
        raise _unsupported("setting a picture's crop")

    @setter("CropBottom")
    def _crop_bottom(self, value: object) -> None:
        raise _unsupported("setting a picture's crop")

    @member
    def CropLeft(self) -> object:
        self.picture()
        return VBASingle(0.0)

    @member
    def CropRight(self) -> object:
        self.picture()
        return VBASingle(0.0)

    @member
    def CropTop(self) -> object:
        self.picture()
        return VBASingle(0.0)

    @member
    def CropBottom(self) -> object:
        self.picture()
        return VBASingle(0.0)


def _freeform_points(owner: ShapeObject) -> list[tuple[int, tuple[float, float]]]:
    """A freeform's points in order, each with the segment type of the step that holds it, on the sheet in points; a
    close adds none. A first point takes the type of the step after it."""
    shape = owner.shape
    if shape.source or shape.kind != "freeform":
        raise _unsupported("the Nodes and Vertices of this shape")
    found: list[tuple[int, tuple[float, float]]] = []
    for command, points in shape.path:
        for x, y in points:
            found.append((1 if command == "cubicBezTo" else 0,
                          (shape.left + x / EMU_PER_POINT, shape.top + y / EMU_PER_POINT)))
    if len(found) > 1:
        found[0] = (found[1][0], found[0][1])
    return found


def _vba_points(points: list[tuple[int, tuple[float, float]]]) -> VBAArray:
    """Points as Excel gives them: Singles, n rows of x and y from 1."""
    made = VBAArray([(1, len(points)), (1, 2)], element_type="Single")
    for row, (_, (x, y)) in enumerate(points, start=1):
        made.set([row, 1], VBASingle(x))
        made.set([row, 2], VBASingle(y))
    return made


class ShapeNodes(Created):
    """A freeform's nodes (tests/fixtures/shape_painter.json). Other shapes have nodes of their preset's that are not
    modelled, and report themselves."""

    vba_type_name = "ShapeNodes"

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def vba_iterate(self) -> Iterator[object]:
        return iter([ShapeNode(self.owner, index) for index in range(len(_freeform_points(self.owner)))])

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Count(self) -> object:
        return VBAInt(len(_freeform_points(self.owner)), "Long")

    @member(default=True, kind="method")
    def Item(self, Index: object = MISSING) -> object:
        chosen = int(to_number(Index))
        if not 1 <= chosen <= len(_freeform_points(self.owner)):
            raise error(-2147024809, "The specified value is out of range.")
        return ShapeNode(self.owner, chosen - 1)


class ShapeNode(Created):
    vba_type_name = "ShapeNode"

    def __init__(self, owner: ShapeObject, index: int) -> None:
        self.owner, self.index = owner, index

    def point(self) -> tuple[int, tuple[float, float]]:
        return _freeform_points(self.owner)[self.index]

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def SegmentType(self) -> object:
        return VBAInt(self.point()[0], "Long")

    @member
    def EditingType(self) -> object:
        # A line's later nodes have none to read, which Excel reports as a failure.
        if self.point()[0] == 0 and self.index > 0:
            raise error(-2147467259, "Method 'EditingType' of object 'ShapeNode' failed")
        return VBAInt(0, "Long")

    @member
    def Points(self) -> object:
        return _vba_points([self.point()])


class RangeAdjustments(Created):
    """A ShapeRange's Adjustments (tests/fixtures/shape_range_more_format.json): what its shapes share, or the mixed
    value. Past any shape's end is error 5, and a set stops there, the shapes before it having taken it."""

    vba_type_name = "Adjustments"

    def __init__(self, shapes: ShapeRange) -> None:
        self.range = shapes

    def parts(self) -> list[Adjustments]:
        return [Adjustments(one) for one in self.range.shapes]

    @member
    def Parent(self) -> object:
        return self.range

    @member
    def Count(self) -> object:
        counts = [part.vba_get("Count") for part in self.parts()]
        return counts[0] if all(one == counts[0] for one in counts) else VBAInt(-2147483648, "Long")

    @member(default=True, kind="method")
    def Item(self, Index: object = MISSING) -> object:
        found = [part.vba_call("Item", (Index,)) for part in self.parts()]
        return found[0] if all(one == found[0] for one in found) else VBASingle(-2147483648.0)

    @setter("Item")
    def _item(self, Index: object, value: object) -> None:
        for part in self.parts():
            part.vba_set("Item", value, (Index,))


class RangeConnector(Created):
    """A ShapeRange's ConnectorFormat (tests/fixtures/shape_range_more_format.json): reads its first line's, and is
    out of range with no line among them. Setting its Type is always out of range."""

    vba_type_name = "ConnectorFormat"

    def __init__(self, shapes: ShapeRange) -> None:
        self.range = shapes

    def first(self) -> ConnectorFormat:
        for one in self.range.shapes:
            if one.shape.kind == "line":
                return ConnectorFormat(one)
        raise error(-2147024809, "The specified value is out of range.")

    @member
    def Parent(self) -> object:
        return self.range

    @member
    def Type(self) -> object:
        return self.first().vba_get("Type")

    @setter("Type")
    def _type(self, value: object) -> None:
        raise error(-2147024809, "The specified value is out of range.")

    @member
    def BeginConnected(self) -> object:
        return self.first().vba_get("BeginConnected")

    @member
    def EndConnected(self) -> object:
        return self.first().vba_get("EndConnected")


class ConnectorFormat(Created):
    """A line or connector's connections (tests/fixtures/shape_adjustments.json). Every other shape gives this object,
    but reading or setting anything on it is out of range. Connecting to shapes is not implemented."""

    vba_type_name = "ConnectorFormat"

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    def connector(self) -> Shape:
        shape = self.owner.shape
        if shape.source:
            raise _unsupported("the ConnectorFormat of a loaded shape")
        if shape.kind != "line":
            raise error(-2147024809, "The specified value is out of range.")
        return shape

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application

    @member
    def Type(self) -> object:
        geometry = self.connector().geometry
        return VBAInt(next((key for key, (name, _) in CONNECTORS.items() if name == geometry), 1), "Long")

    @setter("Type")
    def _type(self, value: object) -> None:
        shape = self.connector()
        wanted = int(to_number(value))
        if wanted not in CONNECTORS:
            raise error(-2147024809, "The specified value is out of range.")
        current = int(to_number(self.vba_get("Type")))
        if (current, wanted) not in ((1, 2), (2, 3)) or shape.adjustments is not None:
            raise _unsupported("this change of connector type")
        self.owner.restyled()
        if current == 2:
            # An elbow's default name changes as a curve is made of it, keeping its number.
            found = re.fullmatch(r"Elbow Connector (\d+)", shape.name)
            if found:
                shape.name = f"Connector: Elbow {found.group(1)}"
        else:
            # A straight line has no adjustment to leave at its default, so this one is written.
            shape.adjustments = (50000,)
        shape.geometry = CONNECTORS[wanted][0]

    @member
    def BeginConnected(self) -> object:
        self.connector()
        return VBAInt(0, "Long")

    @member
    def EndConnected(self) -> object:
        self.connector()
        return VBAInt(0, "Long")


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

    @member
    def Application(self) -> object:
        return self.sheet.book.application

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

    #: A base the two types share, not a type of its own.
    vba_type_name = ""

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape

    def _look(self) -> TextLook:
        """The shape's text layout, which a loaded shape's markup holds and the model does not read yet."""
        shape = self.owner.shape
        if shape.source or shape.kind not in ("shape", "textBox"):
            raise _unsupported("the text layout of this shape")
        if shape.text_look is None:
            shape.text_look = TextLook(align="l" if shape.kind == "shape" else None,
                                       vert="horz" if shape.kind == "textBox" else None,
                                       font_scheme="dk1" if shape.kind == "textBox" else "lt1")
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
        return Characters(self.owner, Start, Length)

    @setter("Characters")
    def _set_characters(self, value: object, Start: object = MISSING, Length: object = MISSING) -> None:
        # Characters has no default member to take a value.
        raise error(438, "Object doesn't support this property or method")

    @member
    def AutoSize(self) -> object:
        # A Boolean, False for a shape that does not size to its text (tests/fixtures/core_reads.json).
        return False

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


class TextColumn2(Created):
    """TextFrame2.Column: 1 to 16 columns and the gap between them. A shape with no text has this object, and using it is
    out of range. It has no Parent (tests/fixtures/shape_columns.json)."""

    vba_type_name = "TextColumn2"

    def __init__(self, frame: TextFrame2) -> None:
        self.frame = frame

    def look(self) -> TextLook:
        if self.frame.owner.shape.kind == "line":
            raise error(-2147024809, "The specified value is out of range.")
        return self.frame._look()  # pyright: ignore[reportPrivateUsage]

    @member
    def Parent(self) -> object:
        raise error(438, "Object doesn't support this property or method")

    @member
    def Number(self) -> object:
        return VBAInt(self.look().columns, "Long")

    @setter("Number")
    def _number(self, value: object) -> None:
        wanted = int(to_number(value))
        self.look()
        if not 1 <= wanted <= 16:
            raise error(-2147024809, "The specified value is out of range.")
        self.frame._changed().columns = wanted  # pyright: ignore[reportPrivateUsage]

    @member
    def Spacing(self) -> object:
        return VBASingle(self.look().column_spacing)

    @setter("Spacing")
    def _spacing(self, value: object) -> None:
        wanted = float(VBASingle(_number(value)))
        self.look()
        if wanted < 0:
            raise error(-2147024809, "The specified value is out of range.")
        self.frame._changed().column_spacing = wanted  # pyright: ignore[reportPrivateUsage]


class TextFrame2(_TextLayout):
    """The newer way to the same text (tests/fixtures/text_frame2.json).

    A text box writes each of these settings as given; a rectangle leaves out one set back to its default."""

    vba_type_name = "TextFrame2"

    @member
    def Column(self) -> object:
        return TextColumn2(self)

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application

    @member
    def TextRange(self) -> object:
        return TextRange2(self.owner)

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
        shape = self.owner.shape
        if shape.text_marks and shape.text_marks[-1] or shape.text_end:
            raise _unsupported("DeleteText over a text whose last run has settings of its own")
        shape.text, shape.text_marks, shape.text_end = "", [], None
        shape.text_edits.append((0, 0, 0, True))
        look.deleted = True
        return EMPTY


def _shape_text(value: object) -> str:
    """Text as the model holds it: \\n ends a paragraph, as vbCr and vbCrLf do; \\x0b breaks a line within one, as vbLf does.
    Both read back as vbLf (tests/fixtures/shape_runs.json)."""
    return to_text(value).replace("\r\n", "\r").replace("\n", "\x0b").replace("\r", "\n")


def _edit_text(owner: ShapeObject, at: int, removed: int, new: str) -> None:
    """Replace part of a shape's text. What comes takes the settings of the first character that went; inserted text, those
    of the character before it or, at the start, after it. Replacing it all can leave its last settings on the paragraph's
    end (tests/fixtures/shape_runs.json)."""
    shape = owner.shape
    whole = removed == len(shape.text)
    marks = shape.text_marks
    if whole and not marks:
        if shape.text_end is not None and new and not shape.text:
            shape.text_marks = [dict(shape.text_end) for _ in new]
            shape.text_end = None
    elif whole:
        first, last = marks[0], marks[-1]
        if shape.text_end is None and (not new or _run_differs(shape, first, last)):
            shape.text_end = dict(last)
        shape.text_marks = [dict(first) for _ in new]
    elif marks:
        like = marks[at] if removed else marks[max(at - 1, 0)]
        shape.text_marks = [*marks[:at], *(dict(like) for _ in new), *marks[at + removed:]]
    if not any(shape.text_marks):
        shape.text_marks = []
    shape.text = shape.text[:at] + new + shape.text[at + removed:]
    shape.text_edits.append((at, removed, len(new), whole))
    owner.sheet.drawing_changed()


def _run_differs(shape: Shape, first: dict[str, object], last: dict[str, object]) -> bool:
    return len(text_run_spans(replace(shape, text="ab", text_marks=[first, last]))) == 2


class _ShapeText(_Drawn):
    """What Characters and TextRange2 share: some of a shape's text, its Text and its Count."""

    vba_type_name = ""

    def __init__(self, shape: ShapeObject) -> None:
        self.owner = shape
        self.sheet = shape.sheet

    def span(self) -> tuple[int, int]:
        """Where this text starts, from 0, and how long it is."""
        raise NotImplementedError

    def shown(self) -> str:
        at, count = self.span()
        return self.owner.shape.text[at:at + count].replace("\x0b", "\n")

    def replaced(self, value: object) -> None:
        at, count = self.span()
        _edit_text(self.owner, at, count, _shape_text(value))

    @member
    def Text(self) -> object:
        return self.shown()

    @setter("Text")
    def _set_text(self, value: object) -> None:
        self.replaced(value)

    @member
    def Count(self) -> object:
        return VBAInt(self.span()[1], "Long")


class Characters(_ShapeText):
    """The characters of a shape's text, or some of them (tests/fixtures/shape_runs.json)."""

    vba_type_name = "Characters"

    def __init__(self, shape: ShapeObject, start: object = MISSING, length: object = MISSING) -> None:
        super().__init__(shape)
        self.start = start
        self.length = length

    def span(self) -> tuple[int, int]:
        """Where these characters start in the text, from 0, and how many they are. A start before the text is its first
        character, one past it raises; a length that is not positive, or too long, runs to the end."""
        size = len(self.owner.shape.text)
        if self.start is MISSING and self.length is MISSING:
            return 0, size
        start = 1 if self.start is MISSING else max(int(to_number(self.start)), 1)
        if start > size:
            raise error(1004, "Unable to get the Characters property of the TextFrame class")
        length = 0 if self.length is MISSING else int(to_number(self.length))
        rest = size - start + 1
        return start - 1, rest if length <= 0 else min(length, rest)

    @member
    def Font(self) -> object:
        return ShapeFont(self)

    @member
    def Caption(self) -> object:
        return self.shown()

    @setter("Caption")
    def _set_caption(self, value: object) -> None:
        self.replaced(value)

    @method
    def Insert(self, String: object = MISSING) -> object:
        self.replaced("" if String is MISSING else String)
        return True

    @method
    def Delete(self) -> object:
        self.replaced("")
        return True


#: What Characters.Font writes for each Underline it takes.
_FONT_UNDERLINES = {2: "words", -4119: "dbl", -4142: "none"}
#: A font name's attributes in the file, for those measured. Others would need the font's own data, and report themselves.
_FONT_LATIN = {"Arial": '<a:latin typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>',
               "": '<a:latin typeface=""/>'}


class _RunLook(TextLook):
    """The look of part of a shape's text, as its first character has it. What is set on it goes to these marks, or, for all
    of the text, to the whole text's look in place of any character's own."""

    __slots__ = ("targets",)

    def __init__(self, base: TextLook, first: dict[str, object], marks: list[dict[str, object]], whole: bool) -> None:
        object.__setattr__(self, "targets", None)
        TextLook.__init__(self, **{one.name: getattr(base, one.name) for one in fields(TextLook)})
        for name, value in first.items():
            object.__setattr__(self, name, value)
        object.__setattr__(self, "targets", (base, marks, whole))

    def __setattr__(self, name: str, value: object) -> None:
        object.__setattr__(self, name, value)
        targets: tuple[TextLook, list[dict[str, object]], bool] | None = getattr(self, "targets")
        if targets is None:
            return
        base, marks, whole = targets
        if whole:
            setattr(base, name, value)
        for mark in marks:
            if whole:
                mark.pop(name, None)
            else:
                mark[name] = value


class _ShapeFontBase(Created):
    """What Characters.Font and TextRange2.Font share (tests/fixtures/shape_font.json), over all of a shape's text or part
    of it (tests/fixtures/shape_runs.json): a range with more than one setting reads as its first character does, but for
    the members in `mixed`."""

    vba_type_name = ""
    #: What a member answers when these settings differ across the range, by the member's name.
    mixed: ClassVar[dict[str, tuple[tuple[str, ...], Callable[[_ShapeFontBase], object]]]] = {}

    def __init__(self, characters: _ShapeText) -> None:
        self.characters = characters
        self.frame = TextFrame(characters.owner)

    def look(self) -> TextLook:
        base = self.frame._look()  # pyright: ignore[reportPrivateUsage]
        shape = self.characters.owner.shape
        at, count = self.characters.span()
        if not shape.text and shape.text_end:
            # No text reads as what its paragraph's end holds.
            return _RunLook(base, shape.text_end, [], False)
        if not shape.text_marks and count == len(shape.text):
            return base
        return _RunLook(base, shape.text_marks[at] if shape.text_marks and count else {}, [], False)

    def changed(self) -> TextLook:
        base = self.frame._changed()  # pyright: ignore[reportPrivateUsage]
        shape = self.characters.owner.shape
        at, count = self.characters.span()
        if count == len(shape.text):
            if not shape.text_marks and shape.text_end is None:
                return base
            marks = [*shape.text_marks, *([] if shape.text_end is None else [shape.text_end])]
            return _RunLook(base, shape.text_marks[0] if shape.text_marks else {}, marks, True)
        if not shape.text_marks:
            shape.text_marks = [{} for _ in shape.text]
        return _RunLook(base, shape.text_marks[at], shape.text_marks[at:at + count], False)

    def values(self, names: tuple[str, ...]) -> list[tuple[object, ...]]:
        """These settings of each character in the range."""
        shape = self.characters.owner.shape
        base = self.frame._look()  # pyright: ignore[reportPrivateUsage]
        at, count = self.characters.span()
        if not shape.text_marks or not count:
            return [tuple(getattr(self.look(), name) for name in names)]
        return [tuple(mark.get(name, getattr(base, name)) for name in names) for mark in shape.text_marks[at:at + count]]

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        spec = self.vba_member(name)
        rule = self.mixed.get(spec.name) if spec is not None else None
        if rule is not None and len(set(self.values(rule[0]))) > 1:
            return rule[1](self)
        return super().vba_get(name, args, named)

    def text_box(self) -> bool:
        return self.characters.owner.shape.kind == "textBox"

    @member
    def Parent(self) -> object:
        return self.characters

    @member
    def Size(self) -> object:
        size = self.look().font_size
        return VBAInt(int(size), "Long") if size == int(size) else size

    @member
    def Name(self) -> object:
        latin = self.look().font_latin
        return "Aptos Narrow" if latin is None else next(name for name, markup in _FONT_LATIN.items() if markup == latin)

    @setter("Name")
    def _name(self, value: object) -> None:
        wanted = to_text(value)
        if wanted not in _FONT_LATIN:
            raise _unsupported(f"the font {wanted!r} in a shape")
        self.changed().font_latin = _FONT_LATIN[wanted]


class ShapeFont(_ShapeFontBase):
    """All of a shape's text's font (tests/fixtures/shape_font.json)."""

    vba_type_name = "Font"

    @member
    def Bold(self) -> object:
        return bool(self.look().font_bold)

    @setter("Bold")
    def _bold(self, value: object) -> None:
        self.changed().font_bold = bool(to_number(value))

    @member
    def Italic(self) -> object:
        return bool(self.look().font_italic)

    @setter("Italic")
    def _italic(self, value: object) -> None:
        self.changed().font_italic = bool(to_number(value))

    @setter("Size")
    def _size(self, value: object) -> None:
        wanted = float(to_number(value))
        # A size of 0 is left alone; others below the measured 10.5 go through as given.
        if wanted > 0:
            self.changed().font_size = wanted

    @member
    def Color(self) -> object:
        from pyopenvba.apps.excel._text_paint import font_ink

        return VBAInt(font_ink(self), "Long")

    @setter("Color")
    def _color(self, value: object) -> None:
        colour = int(to_number(value)) & 0xFFFFFF
        look = self.changed()
        look.font_rgb, look.font_index, look.font_tint = f"{colour & 255:02X}{colour >> 8 & 255:02X}{colour >> 16 & 255:02X}", None, None
        look.font_no_fill, look.font_theme, look.font_mods = False, None, ()

    @member
    def ColorIndex(self) -> object:
        look = self.look()
        if look.font_index is not None:
            return VBAInt(look.font_index - 7, "Long")
        if look.font_rgb is None:
            return VBAInt(-4105 if self.text_box() else 2, "Long")
        if look.font_rgb == "000000":
            return VBAInt(-4105, "Long")
        palette = self.characters.owner.sheet.book.stylesheet.colors.indexed[8:64]
        if look.font_rgb not in palette:
            raise _unsupported("the ColorIndex of a font colour not in the palette")
        return VBAInt(palette.index(look.font_rgb) + 1, "Long")

    @setter("ColorIndex")
    def _color_index(self, value: object) -> None:
        wanted = int(to_number(value))
        look = self.changed()
        if wanted == -4105:
            look.font_rgb, look.font_index, look.font_tint = "000000", None, None
            return
        palette = self.characters.owner.sheet.book.stylesheet.colors.indexed[8:64]
        if not 1 <= wanted <= len(palette):
            raise _unsupported("this ColorIndex of a shape's font")
        look.font_rgb, look.font_index, look.font_tint = palette[wanted - 1], wanted + 7, None

    @member
    def Underline(self) -> object:
        chosen = self.look().font_underline
        if chosen is None:
            return VBAInt(-4142, "Long")
        found = next((key for key, name in _FONT_UNDERLINES.items() if name == chosen), None)
        if found is None:
            raise _unsupported("this Underline of a shape's font")
        return VBAInt(found, "Long")

    @setter("Underline")
    def _underline(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in _FONT_UNDERLINES:
            raise _unsupported("this Underline of a shape's font")
        self.changed().font_underline = _FONT_UNDERLINES[wanted]

    @member
    def Strikethrough(self) -> object:
        return self.look().font_strike == "sngStrike"

    @setter("Strikethrough")
    def _strikethrough(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Strikethrough off a shape's font")
        self.changed().font_strike = "sngStrike"

    @member
    def Superscript(self) -> object:
        return self.look().font_baseline == 30000

    @setter("Superscript")
    def _superscript(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Superscript off a shape's font")
        self.changed().font_baseline = 30000

    @member
    def Subscript(self) -> object:
        return self.look().font_baseline == -25000

    @setter("Subscript")
    def _subscript(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Subscript off a shape's font")
        self.changed().font_baseline = -25000

    @member
    def FontStyle(self) -> object:
        look = self.look()
        return {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic", (True, True): "Bold Italic"}[
            (bool(look.font_bold), bool(look.font_italic))]

    @setter("FontStyle")
    def _font_style(self, value: object) -> None:
        styles = {"Regular": (False, False), "Bold": (True, False), "Italic": (False, True), "Bold Italic": (True, True)}
        wanted = to_text(value)
        if wanted not in styles:
            raise _unsupported(f"the FontStyle {wanted!r} of a shape's font")
        look = self.changed()
        look.font_bold, look.font_italic = styles[wanted]

    @member
    def ThemeColor(self) -> object:
        raise error(1004, "Unable to get the ThemeColor property of the Font class")

    @setter("ThemeColor")
    def _theme_color(self, value: object) -> None:
        raise error(1004, "Unable to set the ThemeColor property of the Font class")

    @member
    def TintAndShade(self) -> object:
        tint = self.look().font_tint
        return VBASingle(0.0 if tint is None else 1 - tint / 100000)

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if wanted <= 0:
            raise _unsupported("this TintAndShade of a shape's font")
        look = self.changed()
        look.font_tint, look.font_rgb, look.font_index = round((1 - wanted) * 100000), None, None

    @member
    def ThemeFont(self) -> object:
        return VBAInt(0, "Long")

    @setter("ThemeFont")
    def _theme_font(self, value: object) -> None:
        raise error(1004, "Unable to set the ThemeFont property of the Font class")

    @member
    def OutlineFont(self) -> object:
        return VBAInt(19, "Long")

    @setter("OutlineFont")
    def _outline_font(self, value: object) -> None:
        # It changes nothing.
        self.look()

    @member
    def Shadow(self) -> object:
        return self.look().font_shadow

    @setter("Shadow")
    def _shadow(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Shadow off a shape's font")
        self.changed().font_shadow = True

    @member
    def Background(self) -> object:
        raise error(1004, "Unable to get the Background property of the Font class")


def _tri(value: object) -> object:
    return VBAInt(-1 if value else 0, "Long")


class ShapeFont2(_ShapeFontBase):
    """TextRange2.Font (tests/fixtures/shape_font2.json), the same settings as Characters.Font, read as MsoTriState, with a
    few more. Its Fill, Line and effects, far-east and other font names are not built."""

    vba_type_name = "Font2"

    @member
    def Fill(self) -> object:
        from pyopenvba.apps.excel._text_paint import TextFill

        return TextFill(self)

    @member
    def Line(self) -> object:
        from pyopenvba.apps.excel._text_paint import TextLineFormat

        return TextLineFormat(self)

    @member
    def Bold(self) -> object:
        return _tri(self.look().font_bold)

    @setter("Bold")
    def _bold(self, value: object) -> None:
        self.changed().font_bold = bool(to_number(value))

    @member
    def Italic(self) -> object:
        return _tri(self.look().font_italic)

    @setter("Italic")
    def _italic(self, value: object) -> None:
        self.changed().font_italic = bool(to_number(value))

    @setter("Size")
    def _size(self, value: object) -> None:
        wanted = float(to_number(value))
        if not 0 < wanted <= 4000:
            raise error(-2147024809, "The specified value is out of range.")
        self.changed().font_size = wanted

    @member
    def NameAscii(self) -> object:
        return self.vba_get("Name")

    @setter("NameAscii")
    def _name_ascii(self, value: object) -> None:
        self.vba_set("Name", value)

    @member
    def UnderlineStyle(self) -> object:
        chosen = self.look().font_underline
        if chosen in (None, "none"):
            return VBAInt(0, "Long")
        found = {"words": 1, "sng": 2, "dbl": 3}.get(chosen or "")
        if found is None:
            raise _unsupported("this UnderlineStyle of a shape's font")
        return VBAInt(found, "Long")

    @setter("UnderlineStyle")
    def _underline_style(self, value: object) -> None:
        wanted = {0: "none", 2: "sng", 3: "dbl"}.get(int(to_number(value)))
        if wanted is None:
            raise _unsupported("this UnderlineStyle of a shape's font")
        self.changed().font_underline = wanted

    @member
    def Strike(self) -> object:
        return VBAInt({"sngStrike": 1, "dblStrike": 2}.get(self.look().font_strike or "", 0), "Long")

    @setter("Strike")
    def _strike(self, value: object) -> None:
        wanted = {0: "noStrike", 1: "sngStrike", 2: "dblStrike"}.get(int(to_number(value)))
        if wanted is None:
            raise _unsupported("this Strike of a shape's font")
        self.changed().font_strike = wanted

    @member
    def StrikeThrough(self) -> object:
        return _tri(self.look().font_strike == "sngStrike")

    @setter("StrikeThrough")
    def _strike_through(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking StrikeThrough off a shape's font")
        self.changed().font_strike = "sngStrike"

    @member
    def DoubleStrikeThrough(self) -> object:
        return _tri(self.look().font_strike == "dblStrike")

    @setter("DoubleStrikeThrough")
    def _double_strike_through(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking DoubleStrikeThrough off a shape's font")
        self.changed().font_strike = "dblStrike"

    @member
    def Superscript(self) -> object:
        return _tri((self.look().font_baseline or 0) > 0)

    @member
    def Subscript(self) -> object:
        return _tri((self.look().font_baseline or 0) < 0)

    @setter("Superscript")
    def _superscript(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Superscript off a shape's font")
        self.changed().font_baseline = 30000

    @setter("Subscript")
    def _subscript(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Subscript off a shape's font")
        self.changed().font_baseline = -25000

    @member
    def BaselineOffset(self) -> object:
        return VBASingle((self.look().font_baseline or 0) / 100000)

    @setter("BaselineOffset")
    def _baseline_offset(self, value: object) -> None:
        self.changed().font_baseline = round(float(VBASingle(float(to_number(value)))) * 100000)

    @member
    def Caps(self) -> object:
        return VBAInt({"small": 1, "all": 2}.get(self.look().font_cap or "", 0), "Long")

    @setter("Caps")
    def _caps(self, value: object) -> None:
        wanted = {1: "small", 2: "all"}.get(int(to_number(value)))
        if wanted is None:
            raise _unsupported("this Caps of a shape's font")
        self.changed().font_cap = wanted

    @member
    def Allcaps(self) -> object:
        return _tri(self.look().font_cap == "all")

    @setter("Allcaps")
    def _allcaps(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Allcaps off a shape's font")
        self.changed().font_cap = "all"

    @member
    def Smallcaps(self) -> object:
        return _tri(self.look().font_cap == "small")

    @setter("Smallcaps")
    def _smallcaps(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Smallcaps off a shape's font")
        self.changed().font_cap = "small"

    @member
    def Spacing(self) -> object:
        return VBASingle((self.look().font_spacing or 0) / 100)

    @setter("Spacing")
    def _spacing(self, value: object) -> None:
        self.changed().font_spacing = round(float(to_number(value)) * 100)

    @member
    def Kerning(self) -> object:
        return VBASingle((self.look().font_kern or 0) / 100)

    @setter("Kerning")
    def _kerning(self, value: object) -> None:
        self.changed().font_kern = round(float(to_number(value)) * 100)

    @member
    def Equalize(self) -> object:
        return _tri(self.look().font_normalize)

    @setter("Equalize")
    def _equalize(self, value: object) -> None:
        if not to_number(value):
            raise _unsupported("taking Equalize off a shape's font")
        self.changed().font_normalize = True


_COLOUR_FIELDS = ("font_rgb", "font_index", "font_tint", "font_theme", "font_mods", "font_no_fill")
_MIXED_NUMBER = VBAInt(-2147483648, "Long")
_MIXED_SINGLE = VBASingle(-2147483648.0)


def _colour_index_of_mixed(_: _ShapeFontBase) -> object:
    raise error(-2147467259, "Unspecified error")


def _mixed_style(font: _ShapeFontBase) -> object:
    bold, italic = (all(one[0] for one in font.values((name,))) for name in ("font_bold", "font_italic"))
    return {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic", (True, True): "Bold Italic"}[(bold, italic)]


ShapeFont.mixed = {
    "Bold": (("font_bold",), lambda font: all(one[0] for one in font.values(("font_bold",)))),
    "Italic": (("font_italic",), lambda font: all(one[0] for one in font.values(("font_italic",)))),
    "FontStyle": (("font_bold", "font_italic"), _mixed_style),
    "Name": (("font_latin",), lambda _: ""),
    "Color": (_COLOUR_FIELDS, lambda _: _MIXED_NUMBER),
    "ColorIndex": (_COLOUR_FIELDS, _colour_index_of_mixed),
    "TintAndShade": (_COLOUR_FIELDS, lambda _: _MIXED_SINGLE),
    "Shadow": (("font_shadow",), lambda _: True),
}
ShapeFont2.mixed = {
    "Bold": (("font_bold",), lambda _: VBAInt(-2, "Long")),
    "Italic": (("font_italic",), lambda _: VBAInt(-2, "Long")),
    "Name": (("font_latin",), lambda _: ""),
    "Kerning": (("font_kern",), lambda _: _MIXED_SINGLE),
    "Equalize": (("font_normalize",), lambda _: VBAInt(-2, "Long")),
}
#: What Words can read here: plain text, with no breaks, and no space at its start. Rules for the rest are not measured.
_WORD = re.compile(r"\w+ *|[^\w\s]+ *", re.ASCII)


class TextRange2(_ShapeText):
    """TextFrame2.TextRange, the same text as Characters, and the ranges inside it (tests/fixtures/shape_runs.json).

    `kind` is what the range counts and indexes: nothing for TextRange itself, which is one item, or Characters, Runs, Words
    or Paragraphs, each cut to the range. A range kept while the text changes follows its characters."""

    vba_type_name = "TextRange2"

    def __init__(self, shape: ShapeObject, at: int | None = None, length: int = 0, kind: str = "") -> None:
        super().__init__(shape)
        self.at = at
        self.size = length
        self.kind = kind
        self.seen = len(shape.shape.text_edits)

    def span(self) -> tuple[int, int]:
        edits = self.owner.shape.text_edits
        for at, removed, added, whole in edits[self.seen:]:
            if self.at is None:
                break
            if whole:
                self.at = None
            elif at + removed <= self.at and removed:
                self.at += added - removed
            elif at > self.at + self.size:
                continue
            else:
                raise _unsupported("a TextRange2 kept while the text changes at or inside it")
        self.seen = len(edits)
        return (0, len(self.owner.shape.text)) if self.at is None else (self.at, self.size)

    def parts(self, kind: str) -> list[tuple[int, int]]:
        """What this range holds of a kind: where each starts, from 0, and its length."""
        start, count = self.span()
        shape = self.owner.shape
        if kind == "Characters":
            return [(at, 1) for at in range(start, start + count)]
        if kind == "Runs":
            found = text_run_spans(shape)
        elif kind == "Paragraphs":
            found = [(match.start(), match.end() - match.start()) for match in re.finditer(r"[^\n]*\n|[^\n]+$", shape.text)]
        elif kind == "Words":
            if _WORD.sub("", shape.text):
                raise _unsupported(f"the Words of {shape.text!r}")
            found = [(match.start(), match.end() - match.start()) for match in _WORD.finditer(shape.text)]
        else:
            return [(start, count)]
        cut = [(max(at, start), min(at + size, start + count) - max(at, start)) for at, size in found]
        return [one for one in cut if one[1] > 0]

    def counted(self, kind: str, start: object, length: object) -> TextRange2:
        """Runs, Words or Paragraphs: the whole range counting them, or one of them and any after it."""
        if start is MISSING and length is MISSING:
            self.span()
            return TextRange2(self.owner, self.at, self.size, kind)
        parts = self.parts(kind)
        first = 1 if start is MISSING else int(to_number(start))
        many = 1 if length is MISSING else int(to_number(length))
        if many == 0:
            raise _unsupported(f"{kind} of no length")
        if first < 1 or many < 1 or first + many - 1 > len(parts):
            raise error(-2147024809, "The parameter is incorrect.")
        begin = parts[first - 1][0]
        return TextRange2(self.owner, begin, parts[first + many - 2][0] + parts[first + many - 2][1] - begin, kind)

    def _inserted(self, place: int, value: object) -> object:
        new = _shape_text("" if value is MISSING else value)
        _edit_text(self.owner, place, 0, new)
        return TextRange2(self.owner, place, len(new), "Characters")

    @member
    def Font(self) -> object:
        if self.span()[1] == 0 and self.owner.shape.text:  # Excel sets the word around the place, by rules not worked out.
            raise _unsupported("the Font of a TextRange2 with no characters")
        return ShapeFont2(self)

    @member
    def Parent(self) -> object:
        return TextFrame2(self.owner)

    @member
    def Start(self) -> object:
        return VBAInt(self.span()[0] + 1, "Long")

    @member
    def Length(self) -> object:
        return VBAInt(self.span()[1], "Long")

    @member
    def Count(self) -> object:
        return VBAInt(len(self.parts(self.kind)) if self.kind else 1, "Long")

    @member
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:
        """Either part left out, or -1, is its default: the first character, and one of them. With both, all of them."""
        at, count = self.span()
        start = -1 if Start is MISSING else int(to_number(Start))
        length = -1 if Length is MISSING else int(to_number(Length))
        if start == -1 and length == -1:
            return TextRange2(self.owner, self.at, self.size, "Characters")
        start, length = 1 if start == -1 else start, 1 if length == -1 else length
        if start < 1 or length < 0 or start + length - 1 > count:
            raise error(-2147024809, "The parameter is incorrect.")
        return TextRange2(self.owner, at + start - 1, length, "Characters")

    @member
    def Runs(self, Start: object = MISSING, Length: object = MISSING) -> object:
        return self.counted("Runs", Start, Length)

    @member
    def Words(self, Start: object = MISSING, Length: object = MISSING) -> object:
        return self.counted("Words", Start, Length)

    @member
    def Paragraphs(self, Start: object = MISSING, Length: object = MISSING) -> object:
        return self.counted("Paragraphs", Start, Length)

    @method
    def Item(self, Index: object = MISSING) -> object:
        number = int(to_number(Index))
        if not self.kind and number == 1:
            return self
        parts = self.parts(self.kind) if self.kind else []
        if not 1 <= number <= len(parts):
            raise error(-2147024809, "The parameter is incorrect.")
        return TextRange2(self.owner, *parts[number - 1], self.kind)

    @method
    def InsertAfter(self, NewText: object = MISSING) -> object:
        at, count = self.span()
        return self._inserted(at + count, NewText)

    @method
    def InsertBefore(self, NewText: object = MISSING) -> object:
        return self._inserted(self.span()[0], NewText)

    @method
    def Delete(self) -> object:
        at, count = self.span()
        if self.at is not None and not count:
            raise error(-2147467259, "Unspecified error")
        _edit_text(self.owner, at, count, "")
        return EMPTY

    def vba_default_call(self, args: Sequence[object], named: dict[str, object]) -> object:
        raise error(451, "Property let procedure not defined and property get procedure did not return an object")

    def vba_iterate(self) -> Iterator[object]:
        if not self.kind:
            return iter([self])
        return iter([TextRange2(self.owner, at, size, self.kind) for at, size in self.parts(self.kind)])


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

    @member(name="_Default")
    def hidden_default(self) -> object:
        """The hidden _Default: Value, read and set (tests/fixtures/hidden_defaults.json)."""
        return self.Value()

    @setter("_Default")
    def _set_hidden_default(self, value: object) -> None:
        self._set_value(value)


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
_STEMS = frozenset({*EXCEL_SHAPE_NAMES.values(), "Picture", "TextBox", "Straight Connector", "Straight Arrow Connector",
                    "Elbow Connector", "Curved Connector"})


def duplicated(sheet: Worksheet, shape: Shape) -> Shape:
    """Shape.Duplicate: the copy Excel makes, placed 12 points right and down, after every other shape, with the
    next id. A name of Excel's own making, stem and number, takes the sheet's next one; a name a macro gave is kept
    as it is (tests/fixtures/shape_duplicate.json)."""
    chart = sheet.chart_states.get(shape.shape_id) if shape.kind == "chart" else None
    if chart is not None and chart.original:
        raise VBAUnsupportedError("Duplicating a chart read from a file is not implemented")
    if chart is None and (shape.source or shape.kind not in ("shape", "textBox", "picture")):
        raise VBAUnsupportedError("Duplicating this shape is not implemented")
    found = re.fullmatch(r"(.+) \d+", shape.name)
    shapes = Shapes(sheet)
    if found is not None and (found.group(1) in _STEMS or chart is not None and found.group(1) == "Chart"):
        name = shapes._free_name(found.group(1))  # pyright: ignore[reportPrivateUsage]
    else:
        sheet.shape_count += 1
        name = shape.name
    copy = replace(shape, name=name, shape_id=shapes._free_id(),  # pyright: ignore[reportPrivateUsage]
                   left=shape.left + 12, top=shape.top + 12, children=[], source="", source_at=0, former_group=None,
                   # The copy's looks are its own, so changing one leaves the original alone.
                   text_look=deepcopy(shape.text_look), text_marks=deepcopy(shape.text_marks), text_end=deepcopy(shape.text_end), text_edits=[],
                   fill_line=deepcopy(shape.fill_line))
    if chart is not None:
        # A chart of its own, with a part of its own when it is saved (tests/fixtures/chart_copies.json).
        own = deepcopy(chart)
        own.part, own.dirty = "", True
        sheet.chart_states[copy.shape_id] = own
    sheet.shapes_.append(copy)
    sheet.drawing_changed()
    if chart is not None:
        from pyopenvba.apps.excel import _events
        from pyopenvba.apps.excel._charts import ChartObject

        _events.new_chart(sheet.book, ChartObject(sheet, copy).Chart())
    return copy


def chart_copied(sheet: Worksheet, shape: Shape) -> None:
    """A chart on the clipboard, which is not Range's: CutCopyMode stays 0. It is what the chart was then, so it
    outlives the chart, and lasts until cells are copied (tests/fixtures/chart_copies.json)."""
    chart = sheet.chart_states.get(shape.shape_id)
    if chart is None or chart.original:
        raise VBAUnsupportedError("Copying a chart read from a file is not implemented")
    application = sheet.book.application
    application.copied_chart = (deepcopy(shape), deepcopy(chart), application.clipboard)


def chart_pasted(sheet: Worksheet, destination: object) -> bool:
    """The copied chart at ``destination`` or the selection, if there is one to paste. Only the active sheet takes it;
    another raises 1004, and Excel's own Paste also fails at random (tests/fixtures/chart_copies.json)."""
    from pyopenvba.apps.excel._charts import ChartObject
    from pyopenvba.apps.excel._model import Range

    application = sheet.book.application
    copied = application.copied_chart
    if copied is None or copied[2] is not application.clipboard:
        return False
    if sheet.book.ActiveSheet() is not sheet:
        raise error(1004, "Paste method of Worksheet class failed")
    place = destination if isinstance(destination, Range) else sheet.selection
    if not isinstance(place, Range) or place.sheet is not sheet:
        raise VBAUnsupportedError("Pasting a chart anywhere but a cell on its sheet is not implemented")
    found = re.fullmatch(r"(.+) \d+", copied[0].name)
    shapes = Shapes(sheet)
    if found is not None and found.group(1) == "Chart":
        name = shapes._free_name("Chart")  # pyright: ignore[reportPrivateUsage]
    else:
        sheet.shape_count += 1
        name = copied[0].name
    made = replace(deepcopy(copied[0]), name=name, shape_id=shapes._free_id(),  # pyright: ignore[reportPrivateUsage]
                   left=float(to_number(place.Left())), top=float(to_number(place.Top())), source="", source_at=0)
    own = deepcopy(copied[1])
    own.part, own.dirty = "", True
    numbered_afresh(own)
    sheet.chart_states[made.shape_id] = own
    sheet.shapes_.append(made)
    sheet.drawing_changed()
    ChartObject(sheet, made).made_active("area")
    return True


def numbered_afresh(own: ChartState) -> None:
    """A chart pasted or copied with its sheet numbers its unique ids afresh, series by series: a series with a dPt
    passes one over, then each point takes the next, its dPt and its dLbl the same, and then the series
    (tests/fixtures/chart_copies.json)."""
    number = 0
    for entry in own.full_series or own.series:
        if any(point.unique >= 0 for point in entry.points.values()):
            number += 1
        for _place, point in sorted(entry.points.items()):
            if point.unique >= 0:
                point.unique = number
            if point.label_unique >= 0:
                point.label_unique = number
            number += point.unique >= 0 or point.label_unique >= 0
        entry.unique, number = number, number + 1
    own.next_unique = own.unique_base = number


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


def _points(value: object) -> list[tuple[int, int]]:
    """A SafeArrayOfPoints, n rows of x and y in points, as EMU."""
    if not isinstance(value, VBAArray) or value.dimensions != 2 or value.bounds[1][1] - value.bounds[1][0] != 1:
        raise _unsupported("this array of points")
    first, last = value.bounds[0]
    column = value.bounds[1][0]
    return [(emu(float(to_number(value.get([row, column])))), emu(float(to_number(value.get([row, column + 1])))))
            for row in range(first, last + 1)]


def _number(value: object) -> float:
    """One of the four measurements a shape is placed by."""
    if value is MISSING or value is None:
        return 0.0
    return float(to_number(value))


def _unsupported(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is real Excel that pyOpenVBA does not implement")
