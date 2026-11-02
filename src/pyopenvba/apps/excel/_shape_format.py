"""Shape.Fill and Shape.Line, as tests/fixtures/shape_fill_line.json measured them on a new AutoShape, text box and
line: what each reads from the shape's style while nothing is set, and the colour, transparency, weight and dash a
macro sets, which a save writes and nothing more.

The style's colours are the default Office theme's; a workbook with another theme, a loaded shape, and the settings
that make Excel write out the whole style (Visible, a line's Transparency) report themselves unsupported."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._styles import DEFAULT_THEME, bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._shapes import ShapeObject
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


def _rgb(colour: int) -> str:
    """A BGR Long as the file's RRGGBB."""
    return f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


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
        return VBAInt(-1 if self.filled() else 0, "Long")

    @member
    def ForeColor(self) -> object:
        return ColorFormat(self, fore=True)

    @member
    def BackColor(self) -> object:
        return ColorFormat(self, fore=False)

    @member
    def Type(self) -> object:
        # msoFillSolid, or the mixed value for a line, which has no fill.
        return VBAInt(1 if self.filled() else -2, "Long")

    @member
    def Transparency(self) -> object:
        if not self.filled():
            # A line has no fill to see through: it reads fully transparent.
            return VBASingle(1.0)
        alpha = self.look().fill_alpha
        return VBASingle(0.0 if alpha is None else (100000 - alpha) / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        self.style()
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        if self.shape.kind != "shape":
            raise VBAUnsupportedError("The fill transparency of this kind of shape is not implemented")
        self.change().fill_alpha = round((1 - wanted) * 100000)

    @method
    def Solid(self) -> object:
        if self.shape.kind != "shape":
            self.style()
            raise VBAUnsupportedError("Solid on this kind of shape is not implemented")
        self.change().fill_solid = True
        return EMPTY


class LineFormat(_Format):
    vba_type_name = "LineFormat"

    @member
    def Visible(self) -> object:
        self.style()
        return VBAInt(-1, "Long")

    @member
    def ForeColor(self) -> object:
        return ColorFormat(self, fore=True)

    @member
    def Weight(self) -> object:
        width = self.look().line_width
        return VBASingle(self.style()[4] if width is None else points(width))

    @setter("Weight")
    def _weight(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        self.style()
        if wanted < 0:
            raise error(*_INVALID)
        self.change().line_width = emu(wanted)

    @member
    def DashStyle(self) -> object:
        self.style()
        dash = self.look().line_dash
        return VBAInt(next((key for key, (name, _) in DASHES.items() if name == dash), 1), "Long")

    @setter("DashStyle")
    def _dash_style(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        self.style()
        if wanted not in DASHES or self.shape.kind == "textBox":
            raise VBAUnsupportedError("This line dash style is not implemented")
        look = self.change()
        look.line_dash, look.line_cap = DASHES[wanted]

    @member
    def Style(self) -> object:
        # msoLineSingle.
        self.style()
        return VBAInt(1, "Long")

    @member
    def Transparency(self) -> object:
        self.style()
        return VBASingle(0.0)


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

    @member
    def RGB(self) -> object:
        colour, _ = self.styled()
        if not self.fore:
            # A new shape's back colour is white; a line, with no fill, reads 0.
            return VBAInt(0 if colour is None else 16777215, "Long")
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
        if isinstance(self.owner, LineFormat) and self.owner.shape.kind == "textBox":
            raise VBAUnsupportedError("Setting a text box's line colour is not implemented")
        look = self.owner.change()
        if isinstance(self.owner, FillFormat):
            look.fill_rgb = _rgb(wanted)
        else:
            look.line_rgb = _rgb(wanted)

    @member
    def ObjectThemeColor(self) -> object:
        if not self.fore:
            raise VBAUnsupportedError("The theme colour of a back colour is not implemented")
        _, theme = self.styled()
        return VBAInt(0 if self.own() is not None else theme, "Long")

    @member
    def Type(self) -> object:
        # msoColorTypeRGB once set, msoColorTypeScheme from the style, mixed where there is no fill.
        if not self.fore:
            raise VBAUnsupportedError("The type of a back colour is not implemented")
        colour, _ = self.styled()
        if colour is None:
            return VBAInt(-2, "Long")
        return VBAInt(1 if self.own() is not None else 2, "Long")
