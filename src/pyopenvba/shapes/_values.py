"""What a shape is, in the words all three Office hosts share.

Excel, Word and PowerPoint keep shapes in different parts and reach
them through different elements, but a shape is the same object in all
three: a name, a geometry, a place, a size, some text, and -- in Excel
and PowerPoint -- a macro a click runs.  This holds that vocabulary;
each host's own module owns its markup.

Where a host has nothing to put in a field the field stays empty
rather than being given a default, so "no macro" and "this host cannot
hold a macro" stay apart.
"""

from __future__ import annotations

from dataclasses import dataclass, field

#: English Metric Units in one point.  Office measures a shape in EMU
#: in the file and in points through the object model, and 12700 is the
#: whole of the conversion: 914400 to the inch, 72 points to the inch.
EMU_PER_POINT = 12700

#: What a shape is, told apart the way the object model tells them
#: apart rather than by which element holds it.
ShapeKind = str

KINDS: frozenset[str] = frozenset(
    {
        "shape",  # an AutoShape: msoAutoShape
        "textBox",  # msoTextBox
        "line",  # a line or connector: msoLine
        "picture",  # msoPicture
        "chart",  # a chart in a graphic frame: msoChart
        "group",  # msoGroup, holding others
        "formControl",  # Excel's own button, check box, drop-down
        "activeX",  # an OLE control, listed but not made here
        "placeholder",  # PowerPoint's title and body boxes
        "table",  # a table in a graphic frame
        "other",  # something real that none of the above describes
    }
)

#: What the object model answers for Shape.Type, so a reader can be
#: held to the number Office itself reported.
MSO_TYPE: dict[str, int] = {
    "shape": 1,
    "group": 6,
    "formControl": 8,
    "line": 9,
    "picture": 13,
    "textBox": 17,
    "chart": 3,
    "activeX": 12,
    "placeholder": 14,
    "table": 19,
}

#: The preset geometry behind the msoShapeType a macro asks for.  Only
#: the ones this can write are here; reading carries whatever the file
#: holds, named or not.
PRESET_GEOMETRY: dict[int, str] = {
    1: "rect",  # msoShapeRectangle
    5: "roundRect",  # msoShapeRoundedRectangle
    9: "ellipse",  # msoShapeOval
    2: "parallelogram",
    3: "trapezoid",
    4: "diamond",
    6: "octagon",
    7: "triangle",
    8: "rtTriangle",
    10: "hexagon",
    11: "cross",
    12: "star5",
    16: "can",
    17: "cube",
}


@dataclass(slots=True)
class ControlInfo:
    """The parts of an Excel form control that are not the drawing.

    A control's state lives in its own part, and what it is wired to --
    the cell it writes and the range it lists -- lives there with it.
    """

    kind: str = "button"
    linked_cell: str = ""
    list_range: str = ""
    value: int = 0
    #: The macro as the sheet holds it, which is ``[0]!Name`` for one
    #: in the workbook the control is in.
    macro_text: str = ""
    #: The control part this came from, so an edit lands back in it.
    part_name: str = ""
    #: The relationship the sheet names it by.
    relationship: str = ""


@dataclass(slots=True)
class Shape:
    """One shape on a sheet, a slide or a page.

    ``left``, ``top``, ``width`` and ``height`` are in points, the unit
    the object model uses, converted from the file's EMU on the way in
    and back on the way out.
    """

    name: str = ""
    kind: ShapeKind = "shape"
    #: The preset geometry, such as rect, roundRect or ellipse.
    geometry: str = ""
    left: float = 0.0
    top: float = 0.0
    width: float = 0.0
    height: float = 0.0
    text: str = ""
    #: The procedure a click runs, where the host has one.
    macro: str = ""
    #: The id the file gives it, unique within its part.
    shape_id: int = 0
    #: A form control's own settings, where it is one.
    control: ControlInfo | None = None
    #: A group's members, in the order the file holds them.
    children: list[Shape] = field(default_factory=lambda: [])
    #: Which AutoShape it is, as Shape.AutoShapeType answers.  A text
    #: box answers 1 as well, which is what Excel does.
    auto_shape_type: int = 1
    #: Word only: whether the shape sits in the text or floats beside it.
    placement: str = ""
    #: Where the shape's markup was, so an untouched one is written
    #: back exactly as it came in.
    source: str = ""

    @property
    def mso_type(self) -> int:
        """What Shape.Type answers for this kind."""
        return MSO_TYPE.get(self.kind, 1)

    def __repr__(self) -> str:
        where = f"{self.left:g},{self.top:g} {self.width:g}x{self.height:g}"
        macro = f" macro={self.macro}" if self.macro else ""
        return f"<Shape {self.name!r} {self.kind} {where}{macro}>"


def points(emu: str | int | float) -> float:
    """EMU as the points the object model reports."""
    try:
        return int(emu) / EMU_PER_POINT
    except (TypeError, ValueError):
        return 0.0


def emu(value: float) -> int:
    """Points as the EMU the file holds, rounded as Office rounds."""
    return int(round(value * EMU_PER_POINT))
