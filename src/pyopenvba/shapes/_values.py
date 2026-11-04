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
        "freeform",  # AddPolyline or AddCurve: msoFreeform
        "textBox",  # msoTextBox
        "line",  # a line or connector: msoLine
        "picture",  # msoPicture
        "chart",  # a chart in a graphic frame: msoChart
        "smartArt",  # a diagram in a graphic frame: msoSmartArt
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
    "freeform": 5,
    "group": 6,
    "formControl": 8,
    "line": 9,
    "picture": 13,
    "textBox": 17,
    "chart": 3,
    "smartArt": 24,
    "activeX": 12,
    "placeholder": 14,
    "table": 19,
}

#: The preset geometry behind the msoShapeType a macro asks for.
#:
#: Measured rather than reasoned about, by `scripts/measure_shape_types.py`:
#: it adds one of each in live Office and reads the `a:prstGeom` back out
#: of the file.  The two names are not related, and several of the
#: presets belong to a different shape than their name suggests -- 11 is
#: `plus` and not `cross`, 12 is `pentagon` and not `star5`, and `star5`
#: is 92.  Only the ones this can write are here; reading carries
#: whatever the file holds, named or not.
PRESET_GEOMETRY: dict[int, str] = {
    1: "rect",  # msoShapeRectangle
    2: "parallelogram",
    3: "trapezoid",
    4: "diamond",
    5: "roundRect",  # msoShapeRoundedRectangle
    6: "octagon",
    7: "triangle",  # msoShapeIsoscelesTriangle
    8: "rtTriangle",
    9: "ellipse",  # msoShapeOval
    10: "hexagon",
    11: "plus",  # msoShapeCross
    12: "pentagon",  # msoShapeRegularPentagon
    13: "can",
    14: "cube",
    15: "bevel",
    16: "foldedCorner",
    17: "smileyFace",
    18: "donut",
    19: "noSmoking",
    20: "blockArc",
    92: "star5",  # msoShape5pointStar
}

#: What Excel and PowerPoint call a new shape of each type.  The two
#: agree on every one of them; Word does not, so it has its own table.
SHAPE_NAMES: dict[int, str] = {
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
    12: "Regular Pentagon",
    13: "Can",
    14: "Cube",
    15: "Bevel",
    16: "Folded Corner",
    17: "Smiley Face",
    18: "Donut",
    19: '"No" Symbol',
    20: "Block Arc",
    92: "5-Point Star",
}

#: Word's own names for the same shapes.  It renamed several of them
#: and leaves the rest alone.
WORD_SHAPE_NAMES: dict[int, str] = {
    **SHAPE_NAMES,
    5: "Rectangle: Rounded Corners",
    12: "Pentagon",
    13: "Cylinder",
    15: "Rectangle: Beveled",
    16: "Rectangle: Folded Corner",
    18: "Circle: Hollow",
    19: '"Not Allowed" Symbol',
    92: "Star: 5 Points",
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
    #: Excel's single, multi, or extended mode; multi modes use selected_indices.
    selection_mode: str = "single"
    items: list[str] = field(default_factory=lambda: [])
    selected_indices: list[int] = field(default_factory=lambda: [])
    minimum: int = 0
    #: None requests the measured creation default for a spinner/scroll bar.
    maximum: int | None = None
    increment: int = 1
    page_change: int = 10
    drop_width: int = 31
    #: Starts a radio group in control order; only its leader stores the link.
    first_button: bool = False
    #: Runtime group-box shape id (0 for the sheet); reconstructed on open.
    radio_group: int | None = None
    #: DropDownLines once set, Enabled turned round, PrintObject once set (a button starts False, the rest True), and
    #: LockedText (tests/fixtures/control_format_more.json).
    drop_lines: int | None = None
    disabled: bool = False
    print_object: bool | None = None
    lock_text: bool = True


@dataclass(slots=True)
class TextLook:
    """Excel: how a new shape lays out its text, as TextFrame sets it and bodyPr and pPr keep it
    (tests/fixtures/text_frame.json)."""

    #: lIns, tIns, rIns and bIns in points, None until set.
    margins: list[float | None] = field(default_factory=lambda: [None, None, None, None])
    #: The paragraphs' algn, None while the file leaves it out.
    align: str | None = None
    #: bodyPr's anchor.
    anchor: str = "t"
    #: bodyPr's vert, None while the file leaves it out.
    vert: str | None = None
    horizontal_overflow: str = "clip"
    vertical_overflow: str = "clip"
    #: The paragraphs' rtl, None while the file leaves it out.
    rtl: bool | None = None
    #: Once an AutoShape's overflow has changed, Excel writes bodyPr with every attribute and a text warp.
    full: bool = False
    #: TextFrame2 (tests/fixtures/text_frame2.json): bodyPr's autofit child, "norm" or "no"; its anchorCtr; its
    #: upright; its wrap, "square" or "none"; each None while the file leaves it out.
    autofit: str | None = None
    anchor_center: bool | None = None
    upright: bool = False
    wrap: str | None = None
    #: Characters.Font on all the text, as the runs write it (tests/fixtures/shape_font.json): the attributes in the
    #: file's order, each None until set; the colour as RGB, with SchemeColor-like index or a tint of the style's own; a
    #: shadow; and the latin font's own attributes.
    font_size: float = 11.0
    font_bold: bool | None = None
    font_italic: bool | None = None
    font_underline: str | None = None
    font_strike: str | None = None
    font_kern: int | None = None
    font_cap: str | None = None
    font_spacing: int | None = None
    font_normalize: bool = False
    font_baseline: int | None = None
    font_rgb: str | None = None
    font_index: int | None = None
    font_tint: int | None = None
    font_shadow: bool = False
    font_latin: str | None = None
    #: TextFrame2.Column: bodyPr's numCol, and its spcCol in points (tests/fixtures/shape_columns.json).
    columns: int = 1
    column_spacing: float = 0.0
    #: DeleteText leaves the empty paragraph's end without a size.
    deleted: bool = False


@dataclass(slots=True)
class ShapeEffects:
    """A shape's Shadow, Glow, SoftEdge and Reflection, as the file spells them (tests/fixtures/shape_effects.json)."""

    #: The file has an effectLst though nothing is in it.
    list_only: bool = False
    #: Glow, SoftEdge or Reflection has been set, even back to nothing.
    touched: bool = False
    shadow_on: bool = False
    shadow_visible: bool = False
    shadow_moved: bool = False
    shadow_rgb: str | None = None
    shadow_distance: int = 37357
    shadow_direction: int = 2700000
    shadow_blur: int = 63500
    shadow_alpha: int | None = None
    shadow_size: int | None = None
    shadow_obscured: bool = False
    glow: bool = False
    glow_radius: int | None = None
    glow_rgb: str | None = None
    glow_scheme: str | None = None
    glow_alpha: int | None = None
    soft_edge: int = 0
    reflection: bool = False
    reflection_type: int = 0
    reflection_start: int | None = None
    reflection_end: int | None = None
    reflection_place: int | None = None
    reflection_distance: int | None = None
    reflection_blur: int | None = None


@dataclass(frozen=True, slots=True)
class GradientPoint:
    """One stop of a gradient fill: its position and alpha in thousandths of a percent, and its colour."""

    position: int
    rgb: str | None = None
    scheme: str | None = None
    mods: tuple[tuple[str, int], ...] = ()
    alpha: int | None = None

    def colour(self) -> tuple[str | None, str | None, tuple[tuple[str, int], ...]]:
        return self.rgb, self.scheme, self.mods


#: FillFormat.GradientStyle by angle in sixty-thousandths of a degree, with its opposite.
_STYLE_ANGLES = {5400000: 1, 16200000: 1, 0: 2, 10800000: 2, 2700000: 3, 13500000: 3, 8100000: 4, 18900000: 4}


@dataclass(frozen=True, slots=True)
class FillGradient:
    """Excel: a gradient fill (tests/fixtures/shape_fill_gradients.json). What FillFormat reads of it follows from its
    stops and angle, whatever made it."""

    stops: tuple[GradientPoint, ...]
    #: In sixty-thousandths of a degree, or None for a path, whose style 5 or 7 is then kept, and a corner's variant.
    angle: int | None = None
    path: int | None = None
    corner: int | None = None
    rotate: bool = True
    #: How it was made, until its stops change: OneColorGradient's degree, PresetGradient's number, and variant 3 or 4
    #: when those folded more stops than the variants can be read from.
    degree: float | None = None
    preset: int | None = None
    mirror: int | None = None

    def colour_type(self) -> int:
        """msoGradientOneColor, TwoColors, PresetColors or MultiColor."""
        if self.preset is not None:
            return 3
        if self.degree is not None:
            return 1
        stops = self.stops
        if len(stops) == 2 or (len(stops) == 3 and stops[2].colour() in (stops[0].colour(), stops[1].colour())):
            return 2
        return 4

    def shown_style(self) -> int:
        if self.path is not None:
            return self.path
        return -2 if self.angle is None else _STYLE_ANGLES.get(self.angle, -2)

    def shown_variant(self) -> int:
        """What GradientVariant reads: a corner's own, else a folding if the stops show one, else which way it runs."""
        if self.corner is not None:
            return self.corner
        if self.shown_style() == -2:
            return 0
        folded = self.mirror
        if folded is None and len(self.stops) == 3 and self.colour_type() == 2:
            folded = 3 if self.stops[2].colour() == self.stops[0].colour() else 4
        if self.path is not None:
            return 1 if folded is None else 3
        if folded is not None:
            return folded
        return 2 if self.angle is not None and self.angle >= 10800000 else 1


@dataclass(slots=True)
class FillLine:
    """Excel: what Shape.Fill and Shape.Line set on a new shape, each None while the shape's style gives it
    (tests/fixtures/shape_fill_line.json). Excel writes only what was set."""

    #: The fill's colour as RRGGBB, its alpha in thousandths of a percent, and whether Solid wrote it out.
    fill_rgb: str | None = None
    fill_alpha: int | None = None
    fill_solid: bool = False
    #: The line's colour as RRGGBB, its width in EMU, and its prstDash and cap.
    line_rgb: str | None = None
    line_width: int | None = None
    line_dash: str | None = None
    line_cap: str | None = None
    #: Excel: Border.LineStyle = xlNone on a chart element, which writes the line as a width and no fill.
    line_off: bool = False
    #: Excel: LineFormat.Style on a trendline or error bars, the compound line the file writes.
    line_compound: str | None = None
    #: Excel: Border.Weight on a trendline, whose black line still reads as automatic.
    line_by_weight: bool = False
    #: Whether the fill or line is hidden, each kept in the spPr's extLst meanwhile.
    fill_hidden: bool = False
    line_hidden: bool = False
    #: Hiding a fill the style gave, or any line, writes the line out in full with an effectLst, for good.
    line_full: bool = False
    effects: bool = False
    #: Its Shadow, Glow, SoftEdge and Reflection, once any was set.
    extra: ShapeEffects | None = None
    #: On a chart element, the one effect set: its member and value (tests/fixtures/chart_protection.json).
    chart_effect: tuple[str, object] | None = None
    #: A line that has been hidden writes a line shape's absent fill as noFill.
    no_fill: bool = False
    #: Line.Transparency writes the line out in full with its ends, and this alpha, None at 0.
    line_ends: bool = False
    line_alpha: int | None = None
    #: The width and dash a hidden line had when it was hidden, which its ln goes on holding.
    frozen_width: int | None = None
    frozen_dash: str | None = None
    #: A fill colour set while the line is hidden loses the hidden line.
    line_lost: bool = False
    #: Hidden through a range, after its first shape: a noFill and no copy of the colour, which is gone.
    fill_bare: bool = False
    #: A colour's theme name, its lumMod, lumOff, tint and shade in the order set, and a SchemeColor's palette index,
    #: fill then line (tests/fixtures/shape_line_ends.json).
    fill_scheme: str | None = None
    fill_mods: tuple[tuple[str, int], ...] = ()
    fill_index: int | None = None
    line_scheme: str | None = None
    line_mods: tuple[tuple[str, int], ...] = ()
    line_index: int | None = None
    #: TintAndShade takes the place of the shade the style gives a line.
    line_plain: bool = False
    #: A line's back colour with no pattern to take it, which leaves an empty solidFill.
    line_open: bool = False
    #: A patterned line's prst and back colour; its fore colour is the line's own.
    line_pattern: str | None = None
    line_back_rgb: str | None = None
    line_back_scheme: str | None = None
    #: A patterned fill's prst and back colour, white until set; its fore colour is the fill's own
    #: (tests/fixtures/shape_fill_patterns.json).
    fill_pattern: str | None = None
    fill_back_rgb: str | None = None
    fill_back_scheme: str | None = None
    fill_back_mods: tuple[tuple[str, int], ...] = ()
    fill_gradient: FillGradient | None = None
    #: InsetPen's algn.
    line_align: str | None = None
    #: Each end's type, w and len as set, and those a hidden line had when it was hidden.
    line_head: tuple[str | None, str | None, str | None] = (None, None, None)
    line_tail: tuple[str | None, str | None, str | None] = (None, None, None)
    frozen_head: tuple[str | None, str | None, str | None] = (None, None, None)
    frozen_tail: tuple[str | None, str | None, str | None] = (None, None, None)


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
    #: PowerPoint only: what a click does, as ActionSetting.Action
    #: answers it.  Held apart from the macro's name because
    #: PowerPoint remembers ppActionRunMacro with no macro named yet.
    action: int = 0
    #: Word only: whether the shape sits in the text or floats beside it.
    placement: str = ""
    #: Where the shape's markup was, so an untouched one is written
    #: back exactly as it came in.
    source: str = ""
    #: Where a load placed it, to tell whether it has moved since.
    loaded_box: tuple[float, float, float, float] | None = None
    #: The hlinkClick its name carries in a file: its relationship, tip and place; a load makes the first the address.
    link: tuple[str, str, str] | None = None
    #: Where in the part that markup started.
    source_at: int = 0
    #: Word only: the anchor's relativeHeight, which is the order Word
    #: reports its floating shapes in.  It rises as shapes are added.
    z_order: int = 0
    #: Excel: AlternativeText and Title, kept as the file's descr and title.
    description: str = ""
    title: str = ""
    decorative: bool = False
    #: Excel: Rotation as given, which the file keeps unwrapped; reads wrap it into 0 to 360.
    rotation: float = 0.0
    #: Excel: LockAspectRatio once set, None until then; Locked.
    lock_aspect: bool | None = None
    locked: bool = True
    #: Excel: a new picture's image file and, once saved, its relationship in the drawing part.
    image: bytes | None = None
    image_relationship: str = ""
    #: Excel: Shape.Apply has written the empty text's format out.
    painted: bool = False
    #: Excel: Adjustments, as the file's values in the order of ADJUSTMENTS; None while the geometry's defaults stand.
    adjustments: tuple[int, ...] | None = None
    #: Excel: HorizontalFlip and VerticalFlip, the xfrm's flipH and flipV.
    flip_h: bool = False
    flip_v: bool = False
    #: Excel: Visible msoFalse, the cNvPr's hidden.
    hidden: bool = False
    #: Excel: PrintObject, the clientData's fPrintsWithSheet.
    prints: bool = True
    #: Excel: Placement as the anchor's editAs: "" for xlMoveAndSize, oneCell for xlMove, absolute for
    #: xlFreeFloating.
    edit_as: str = ""
    #: Excel: a new group's child space (chOff and chExt) in points, None while it is the group's own box. Its
    #: members are held where the sheet shows them and mapped into it when written.
    child_box: tuple[float, float, float, float] | None = None
    #: Excel: the group Ungroup took this shape out of, which Regroup puts back together.
    former_group: Shape | None = field(default=None, compare=False)
    #: Excel: the text layout TextFrame set on a new shape, None while it is the one Excel starts with.
    text_look: TextLook | None = None
    #: Excel: a new freeform's path, each step its command (moveTo, lnTo, cubicBezTo or close) and its points in EMU
    #: from the path's corner (tests/fixtures/shape_polyline.json).
    path: list[tuple[str, tuple[tuple[int, int], ...]]] = field(default_factory=lambda: [])
    #: Excel: the fill and line Shape.Fill and Shape.Line set, None while it is the style's
    #: (tests/fixtures/shape_fill_line.json).
    fill_line: FillLine | None = None

    @property
    def mso_type(self) -> int:
        """What Shape.Type answers for this kind: a connector is an AutoShape where a plain line is not, and types 109 to
        124 are callouts (tests/fixtures/shape_adjustment_table.json)."""
        if self.kind == "line" and self.geometry != "line":
            return 1
        if self.kind == "shape" and 109 <= getattr(self, "auto_shape_type", 1) <= 124:
            return 2
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


Box = tuple[float, float, float, float]


def child_space(group: Shape) -> Box:
    """A group's child space: where its members sit in the file, against its own box on the sheet."""
    return group.child_box or (group.left, group.top, group.width, group.height)


def to_child(group: Shape, box: Box) -> Box:
    """A box on the sheet, as the group's child space holds it."""
    across, down, wide, tall = child_space(group)
    x_scale = wide / group.width if group.width else 1.0
    y_scale = tall / group.height if group.height else 1.0
    return (across + (box[0] - group.left) * x_scale, down + (box[1] - group.top) * y_scale, box[2] * x_scale,
            box[3] * y_scale)


def to_sheet(group: Shape, box: Box) -> Box:
    """A box in the group's child space, as the sheet shows it."""
    across, down, wide, tall = child_space(group)
    x_scale = group.width / wide if wide else 1.0
    y_scale = group.height / tall if tall else 1.0
    return (group.left + (box[0] - across) * x_scale, group.top + (box[1] - down) * y_scale, box[2] * x_scale,
            box[3] * y_scale)


def union(boxes: list[Box]) -> Box:
    """The smallest box holding all of these."""
    left = min(one[0] for one in boxes)
    top = min(one[1] for one in boxes)
    return (left, top, max(one[0] + one[2] for one in boxes) - left, max(one[1] + one[3] for one in boxes) - top)
