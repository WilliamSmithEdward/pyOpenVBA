"""Shape.Fill and Shape.Line, as tests/fixtures/shape_fill_line.json measured them on a new AutoShape, text box and
line: what each reads from the shape's style while nothing is set, and what a macro sets, which a save writes and
nothing more, until something makes Excel write out the style:

- Hiding a fill the style gave, or hiding any line, writes the line out in full with an effectLst, and those stay after
  it is shown again. A fill already set is hidden with no more than a noFill and a hidden copy of its colour.
- While a line is hidden its ln keeps the width and dash it had, and changes reach only the hidden copy.
- A line's Transparency, 0 included, writes the line out in full with its ends, and no fill or effectLst.
- A fill colour set while the line is hidden loses the hidden line, which then reads mixed values.

The style's colours are the default Office theme's; a workbook with another theme, a loaded shape, and combinations
that were not measured report themselves unsupported."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._styles import DEFAULT_THEME, bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._shapes import ShapeObject, ShapeRange
    from pyopenvba.shapes._values import Shape

#: What each kind's style gives, as BGR Longs: the fill colour (None where it has no fill) and its theme colour, the
#: line colour and its theme colour, and the line's width in points.
STYLE = {
    "shape": (8544277, 5, 3351556, 5, 1.5),
    "textBox": (16777215, 2, 12369084, 2, 0.75),
    "line": (None, -2, 8544277, 5, 1.5),
}
#: LineFormat.DashStyle and how the file writes it: its prstDash and the cap it takes.
DASHES = {4: ("dash", None), 3: ("dot", "rnd")}
_INVALID = (-2147024809, "The specified value is out of range.")
#: What Excel's mixed value reads as in a Single.
_MIXED = -2147483648.0
#: The colour a fill hidden bare comes back as (tests/fixtures/shape_range_fill_line.json).
BARE_COLOUR = "BBE0E3"


def _rgb(colour: int) -> str:
    """A BGR Long as the file's RRGGBB."""
    return f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


def _shown(value: object) -> bool:
    """A Visible given as msoTrue or msoFalse, the two measured."""
    number = int(to_integer(value, "Long"))
    if number not in (0, -1):
        raise VBAUnsupportedError("Visible set to anything but msoTrue or msoFalse is not implemented")
    return number == -1


def _unmeasured(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is not implemented")


class _Format(Created):
    """What Fill and Line share: the shape, its kind's style, and the look a change writes."""

    vba_type_name = ""

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    @property
    def shape(self) -> Shape:
        return self.owner.shape

    def style(self) -> tuple[int | None, int, int, int, float]:
        shape = self.shape
        if shape.source or shape.kind not in STYLE:
            raise VBAUnsupportedError("The fill and line of this shape are not implemented")
        if self.owner.sheet.book.stylesheet.colors.theme != DEFAULT_THEME:
            raise VBAUnsupportedError("The fill and line of a shape under another theme are not implemented")
        return STYLE[shape.kind]

    def look(self) -> FillLine:
        return self.shape.fill_line or FillLine()

    def change(self) -> FillLine:
        self.style()
        self.owner.guard("Fill" if isinstance(self, FillFormat) else "Line")
        if self.shape.fill_line is None:
            self.shape.fill_line = FillLine()
        self.owner.restyled()
        return self.shape.fill_line

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.sheet.book.application


class FillFormat(_Format):
    vba_type_name = "FillFormat"

    def filled(self) -> bool:
        return self.style()[0] is not None

    @member
    def Visible(self) -> object:
        look = self.look()
        return VBAInt(-1 if self.filled() and not look.fill_hidden and not look.fill_bare else 0, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = _shown(value)
        if not self.filled():
            # A line has no fill: hiding it changes nothing, and showing it is refused.
            if wanted:
                raise error(*_INVALID)
            return
        before = self.look()
        if before.fill_bare:
            # A fill hidden bare has lost its colour: it comes back as this one.
            if wanted:
                look = self.change()
                look.fill_bare, look.fill_rgb = False, BARE_COLOUR
            return
        if wanted != before.fill_hidden:
            return
        if before.line_hidden:
            raise _unmeasured("Showing or hiding a fill while the line is hidden")
        unset = self.shape.kind == "shape" and before.fill_rgb is None and before.fill_alpha is None
        if unset and not before.fill_solid and before.line_cap is not None:
            raise _unmeasured("Hiding a fill under a line with a round cap")
        look = self.change()
        look.fill_hidden = not wanted
        if not wanted and unset and not look.fill_solid:
            look.fill_solid = look.line_full = look.effects = True

    @member
    def ForeColor(self) -> object:
        return ColorFormat(self, fore=True)

    @member
    def BackColor(self) -> object:
        return ColorFormat(self, fore=False)

    @member
    def Type(self) -> object:
        # msoFillSolid, or the mixed value for a line, which has no fill, or one hidden bare.
        return VBAInt(1 if self.filled() and not self.look().fill_bare else -2, "Long")

    @member
    def Transparency(self) -> object:
        if not self.filled() or self.look().fill_bare:
            # No fill to see through: it reads fully transparent.
            return VBASingle(1.0)
        alpha = self.look().fill_alpha
        return VBASingle(0.0 if alpha is None else (100000 - alpha) / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        self.style()
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        if self.shape.kind == "line" or self.look().line_hidden or self.look().fill_bare:
            raise _unmeasured("The fill transparency of a line, under a hidden line, or of a fill hidden bare,")
        # Setting it shows a hidden fill.
        look = self.change()
        look.fill_alpha, look.fill_hidden = round((1 - wanted) * 100000), False

    @method
    def Solid(self) -> object:
        self.style()
        before = self.look()
        if self.shape.kind != "shape" or before.fill_hidden or before.line_hidden or before.fill_bare:
            raise _unmeasured("Solid on this kind of shape, or with the fill or line hidden,")
        self.change().fill_solid = True
        return EMPTY


class LineFormat(_Format):
    vba_type_name = "LineFormat"

    def kept(self) -> FillLine:
        """The look before a change to the line, which a line lost to a fill colour takes no more of."""
        self.style()
        before = self.look()
        if before.line_lost:
            raise _unmeasured("Changing a line lost to a fill colour")
        return before

    @member
    def Visible(self) -> object:
        self.style()
        return VBAInt(0 if self.look().line_hidden else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = _shown(value)
        before = self.kept()
        if wanted != before.line_hidden:
            return
        if before.line_ends or before.line_cap is not None:
            raise _unmeasured("Hiding or showing a line with its ends or a round cap written")
        look = self.change()
        look.line_hidden = not wanted
        if wanted:
            look.frozen_width = look.frozen_dash = None
            return
        look.line_full = look.effects = True
        look.frozen_width, look.frozen_dash = look.line_width, look.line_dash
        if self.shape.kind == "line":
            look.no_fill = True
        elif self.shape.kind == "shape":
            look.fill_solid = True

    @member
    def ForeColor(self) -> object:
        return ColorFormat(self, fore=True)

    @member
    def Weight(self) -> object:
        style = self.style()
        look = self.look()
        if look.line_lost:
            return VBASingle(_MIXED)
        return VBASingle(style[4] if look.line_width is None else points(look.line_width))

    @setter("Weight")
    def _weight(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        self.kept()
        if wanted < 0:
            raise error(*_INVALID)
        self.change().line_width = emu(wanted)

    @member
    def DashStyle(self) -> object:
        self.style()
        look = self.look()
        if look.line_lost:
            return VBAInt(-2, "Long")
        return VBAInt(next((key for key, (name, _) in DASHES.items() if name == look.line_dash), 1), "Long")

    @setter("DashStyle")
    def _dash_style(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        before = self.kept()
        if wanted not in DASHES or self.shape.kind == "textBox" or before.line_ends:
            raise VBAUnsupportedError("This line dash style is not implemented")
        if DASHES[wanted][1] is not None and before.line_full:
            raise _unmeasured("A round cap on a line written out in full")
        look = self.change()
        look.line_dash, look.line_cap = DASHES[wanted]

    @member
    def Style(self) -> object:
        # msoLineSingle.
        self.style()
        return VBAInt(-2 if self.look().line_lost else 1, "Long")

    @member
    def Transparency(self) -> object:
        self.style()
        look = self.look()
        if look.line_lost:
            return VBASingle(_MIXED)
        return VBASingle(0.0 if look.line_alpha is None else (100000 - look.line_alpha) / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        before = self.kept()
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        if before.line_hidden or before.line_dash is not None:
            raise _unmeasured("The transparency of a hidden or dashed line")
        look = self.change()
        look.line_ends = True
        look.line_alpha = None if wanted == 0 else round((1 - wanted) * 100000)


class ColorFormat(Created):
    """A fill's or line's colour: its own once RGB is set, else its style's theme colour."""

    vba_type_name = "ColorFormat"

    def __init__(self, owner: _Format, *, fore: bool) -> None:
        self.owner, self.fore = owner, fore

    def styled(self) -> tuple[int | None, int]:
        """The style's colour and its theme colour."""
        fill, fill_theme, line, line_theme, _ = self.owner.style()
        return (fill, fill_theme) if isinstance(self.owner, FillFormat) else (line, line_theme)

    def own(self) -> str | None:
        look = self.owner.look()
        return look.fill_rgb if isinstance(self.owner, FillFormat) else look.line_rgb

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.owner.sheet.book.application

    def lost(self) -> bool:
        return isinstance(self.owner, LineFormat) and self.owner.look().line_lost

    def bare(self) -> bool:
        return isinstance(self.owner, FillFormat) and self.owner.look().fill_bare

    @member
    def RGB(self) -> object:
        colour, _ = self.styled()
        if not self.fore:
            # A new shape's back colour is white; a line, with no fill, reads 0.
            return VBAInt(0 if colour is None else 16777215, "Long")
        if self.lost() or self.bare():
            return VBAInt(16777215, "Long")
        own = self.own()
        if own is not None:
            return VBAInt(bgr(own), "Long")
        return VBAInt(0 if colour is None else colour, "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        wanted = int(to_integer(value, "Long")) & 0xFFFFFF
        colour, _ = self.styled()
        if not self.fore:
            raise VBAUnsupportedError("Setting a back colour is not implemented")
        if colour is None:
            # A line has no fill to colour.
            raise error(*_INVALID)
        before = self.owner.look()
        if before.line_lost:
            raise _unmeasured("Changing a colour once a hidden line is lost")
        if isinstance(self.owner, LineFormat):
            if self.owner.shape.kind == "textBox":
                raise VBAUnsupportedError("Setting a text box's line colour is not implemented")
            self.owner.change().line_rgb = _rgb(wanted)
            return
        if before.line_hidden and self.owner.shape.kind != "shape":
            raise _unmeasured("A fill colour for this kind of shape while its line is hidden")
        # Setting it shows a hidden fill, and loses a hidden line.
        look = self.owner.change()
        look.fill_rgb, look.fill_hidden, look.fill_bare, look.line_lost = _rgb(wanted), False, False, look.line_hidden

    @member
    def ObjectThemeColor(self) -> object:
        if not self.fore:
            raise VBAUnsupportedError("The theme colour of a back colour is not implemented")
        _, theme = self.styled()
        return VBAInt(0 if self.own() is not None or self.lost() or self.bare() else theme, "Long")

    @member
    def Type(self) -> object:
        # msoColorTypeRGB once set, msoColorTypeScheme from the style, mixed where there is no fill.
        if not self.fore:
            raise VBAUnsupportedError("The type of a back colour is not implemented")
        colour, _ = self.styled()
        if colour is None:
            return VBAInt(-2, "Long")
        return VBAInt(1 if self.own() is not None or self.lost() or self.bare() else 2, "Long")


def _alike(values: list[object], mixed: object) -> object:
    return values[0] if all(value == values[0] for value in values[1:]) else mixed


class RangeFill(FillFormat):
    """A ShapeRange's fill (tests/fixtures/shape_range_fill_line.json): what its shapes with a fill share, a line having
    none to count. A fill hidden beside one that is not, or hidden differently, reads mixed throughout. Setting
    Visible does nothing while the range reads that already; hiding hides the first shape as it would alone and leaves
    the rest a bare noFill, which loses their colours."""

    vba_type_name = "FillFormat"

    def __init__(self, shapes: ShapeRange) -> None:
        super().__init__(shapes.shapes[0])
        self.range = shapes

    def parts(self) -> list[FillFormat]:
        on_sheet = self.owner.sheet.shapes_
        return [FillFormat(one) for one in self.range.shapes if any(shape is one.shape for shape in on_sheet)]

    def bearers(self) -> list[FillFormat]:
        parts = self.parts()
        return [part for part in parts if part.filled()] or parts

    def agreed(self) -> bool:
        return len({(part.look().fill_hidden, part.look().fill_bare) for part in self.bearers()}) == 1

    def merged(self, name: str, mixed: object) -> object:
        if not self.agreed():
            return mixed
        return _alike([part.vba_get(name) for part in self.bearers()], mixed)

    def still(self, what: str) -> None:
        """Refuses a change to fills already hidden, which Excel treats in ways one case does not settle."""
        if any(part.look().fill_hidden or part.look().fill_bare for part in self.bearers()):
            raise _unmeasured(f"{what} of a range with a hidden fill")

    @member
    def Parent(self) -> object:
        return self.range

    @member
    def Visible(self) -> object:
        bearers = self.bearers()
        if len(bearers) == 1:
            return bearers[0].vba_get("Visible")
        looks = [part.look() for part in bearers]
        hidden = any(look.fill_hidden for look in looks) or all(look.fill_bare for look in looks)
        return VBAInt(0 if hidden else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = _shown(value)
        if (self.vba_get("Visible") == -1) == wanted:
            return
        parts = self.parts()
        if wanted:
            for part in parts:
                if part.filled():
                    part.vba_set("Visible", -1)
            return
        if not parts[0].filled() or any(part.look().line_hidden or part.look().fill_bare for part in parts):
            raise _unmeasured("Hiding the fills of a range that starts with a line, or with a line hidden,")
        parts[0].vba_set("Visible", 0)
        for part in parts[1:]:
            look = part.change()
            if not part.filled():
                look.no_fill = True
                continue
            look.fill_bare, look.fill_rgb, look.fill_alpha, look.fill_solid = True, None, None, False

    @member
    def ForeColor(self) -> object:
        return RangeColor(self, fore=True)

    @member
    def BackColor(self) -> object:
        return RangeColor(self, fore=False)

    @member
    def Type(self) -> object:
        return self.merged("Type", VBAInt(-2, "Long"))

    @member
    def Transparency(self) -> object:
        return self.merged("Transparency", VBASingle(_MIXED))

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        self.still("The transparency")
        for part in self.bearers():
            part.vba_set("Transparency", value)

    @method
    def Solid(self) -> object:
        for part in self.bearers():
            part.vba_get("Solid")
        return EMPTY


class RangeLine(LineFormat):
    """A ShapeRange's line (tests/fixtures/shape_range_fill_line.json): what all its shapes share. A line hidden beside
    one that is not reads mixed in its colour and transparency, and goes on sharing its weight and dash."""

    vba_type_name = "LineFormat"

    def __init__(self, shapes: ShapeRange) -> None:
        super().__init__(shapes.shapes[0])
        self.range = shapes

    def parts(self) -> list[LineFormat]:
        on_sheet = self.owner.sheet.shapes_
        return [LineFormat(one) for one in self.range.shapes if any(shape is one.shape for shape in on_sheet)]

    def bearers(self) -> list[LineFormat]:
        return self.parts()

    def agreed(self) -> bool:
        return len({part.look().line_hidden for part in self.parts()}) == 1

    def spread(self, name: str, value: object) -> None:
        for part in self.parts():
            part.vba_set(name, value)

    def alike(self, name: str, mixed: object) -> object:
        return _alike([part.vba_get(name) for part in self.parts()], mixed)

    @member
    def Parent(self) -> object:
        return self.range

    @member
    def Visible(self) -> object:
        return VBAInt(0 if any(part.look().line_hidden for part in self.parts()) else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = _shown(value)
        if (self.vba_get("Visible") == -1) != wanted:
            self.spread("Visible", value)

    @member
    def ForeColor(self) -> object:
        return RangeColor(self, fore=True)

    @member
    def Weight(self) -> object:
        return self.alike("Weight", VBASingle(_MIXED))

    @setter("Weight")
    def _weight(self, value: object) -> None:
        self.spread("Weight", value)

    @member
    def DashStyle(self) -> object:
        return self.alike("DashStyle", VBAInt(-2, "Long"))

    @setter("DashStyle")
    def _dash_style(self, value: object) -> None:
        self.spread("DashStyle", value)

    @member
    def Style(self) -> object:
        return self.alike("Style", VBAInt(-2, "Long"))

    @member
    def Transparency(self) -> object:
        return self.alike("Transparency", VBASingle(_MIXED)) if self.agreed() else VBASingle(_MIXED)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        self.spread("Transparency", value)


class RangeColor(ColorFormat):
    """A ShapeRange's fill or line colour: one unit, its RGB, theme colour and type all read mixed if any differs."""

    vba_type_name = "ColorFormat"

    def __init__(self, owner: RangeFill | RangeLine, *, fore: bool) -> None:
        super().__init__(owner, fore=fore)
        self.whole = owner

    def parts(self) -> list[ColorFormat]:
        return [ColorFormat(part, fore=self.fore) for part in self.whole.bearers()]

    def unit(self) -> tuple[object, object, object] | None:
        """The colour its shapes share, or None."""
        if not self.fore:
            raise VBAUnsupportedError("The theme colour and type of a range's back colour are not implemented")
        if not self.whole.agreed():
            return None
        found = [(part.vba_get("RGB"), part.vba_get("ObjectThemeColor"), part.vba_get("Type")) for part in self.parts()]
        return found[0] if all(one == found[0] for one in found[1:]) else None

    @member
    def Parent(self) -> object:
        return self.whole

    @member
    def RGB(self) -> object:
        mixed = VBAInt(-2147483648, "Long")
        if self.fore:
            unit = self.unit()
            return mixed if unit is None else unit[0]
        if not self.whole.agreed():
            return mixed
        return _alike([part.vba_get("RGB") for part in self.parts()], mixed)

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        if isinstance(self.whole, RangeFill):
            self.whole.still("The colour")
        if self.fore:
            for part in self.parts():
                part.vba_set("RGB", value)
            return
        # A range's back colour changes no colour: it writes the fill out, as Solid does.
        if not isinstance(self.whole, RangeFill) or any(part.shape.kind != "shape" for part in self.whole.bearers()):
            raise _unmeasured("Setting the back colour of this range")
        for part in self.whole.bearers():
            part.vba_get("Solid")

    @member
    def ObjectThemeColor(self) -> object:
        unit = self.unit()
        return VBAInt(-2, "Long") if unit is None else unit[1]

    @member
    def Type(self) -> object:
        unit = self.unit()
        return VBAInt(-2, "Long") if unit is None else unit[2]
