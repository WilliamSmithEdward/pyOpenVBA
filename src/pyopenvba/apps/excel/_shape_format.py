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

from collections.abc import Callable, Iterable, Iterator
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._held_shapes import ShapePart
from pyopenvba.apps.excel._styles import DEFAULT_THEME
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import MISSING, member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillGradient, FillLine, GradientPoint, emu, points

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
#: LineFormat's arrowhead values and what the file writes for each (tests/fixtures/shape_line_ends.json).
ARROW_STYLES = {1: "none", 2: "triangle", 3: "arrow", 4: "stealth", 5: "diamond", 6: "oval"}
ARROW_SIZES = {1: "sm", 2: "med", 3: "lg"}
#: MsoPatternType 1 to 54 as the file names them.
PATTERNS = ("pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70", "pct75", "pct80", "pct90",
            "dkHorz", "dkVert", "dkDnDiag", "dkUpDiag", "smCheck", "trellis", "ltHorz", "ltVert", "ltDnDiag",
            "ltUpDiag", "smGrid", "dotDmnd", "wdDnDiag", "wdUpDiag", "dashUpDiag", "dashDnDiag", "narVert", "narHorz",
            "dashVert", "dashHorz", "lgConfetti", "lgGrid", "horzBrick", "lgCheck", "smConfetti", "zigZag",
            "solidDmnd", "diagBrick", "openDmnd", "plaid", "sphere", "weave", "dotGrid", "divot", "shingle", "wave",
            "horz", "vert", "cross", "dnDiag", "upDiag", "diagCross")
#: MsoThemeColorIndex 1 to 16 as the file names them, and where DEFAULT_THEME holds each.
SCHEMES = {1: "dk1", 2: "lt1", 3: "dk2", 4: "lt2", 5: "accent1", 6: "accent2", 7: "accent3", 8: "accent4", 9: "accent5",
           10: "accent6", 11: "hlink", 12: "folHlink", 13: "tx1", 14: "bg1", 15: "tx2", 16: "bg2"}
_THEME_SLOT = {"lt1": 0, "dk1": 1, "lt2": 2, "dk2": 3, "accent1": 4, "accent2": 5, "accent3": 6, "accent4": 7,
               "accent5": 8, "accent6": 9, "hlink": 10, "folHlink": 11, "tx1": 1, "bg1": 0, "tx2": 3, "bg2": 2}
#: The shade the style gives each kind's line.
_LINE_SHADES = {"shape": 15000, "textBox": 50000, "line": 0}
#: SchemeColor past the palette: Windows' own colours, of which these three do not vary (66 to 79 are refused).
_SCHEME_EXTRAS = {64: "000000", 65: "FFFFFF", 80: "FFFFE1"}

#: PresetGradient 1 to 24 as stops: position and colour (tests/fixtures/shape_fill_gradients.json).
GRADIENT_PRESETS: tuple[tuple[tuple[int, str], ...], ...] = (
    ((0, "000082"), (30000, "66008F"), (64999, "BA0066"), (89999, "FF0000"), (100000, "FF8200"),),
    ((0, "000000"), (20000, "000040"), (50000, "400040"), (75000, "8F0040"), (89999, "F27300"), (100000, "FFBF00"),),
    ((0, "000000"), (39999, "0A128C"), (70000, "181CC7"), (88000, "7005D4"), (100000, "8C3D91"),),
    ((0, "5E9EFF"), (39999, "85C2FF"), (70000, "C4D6EB"), (100000, "FFEBFA"),),
    ((0, "DCEBF5"), (8000, "83A7C3"), (13000, "768FB9"), (21001, "83A7C3"), (52000, "FFFFFF"), (56000, "9C6563"),
     (58000, "80302D"), (71001, "C0524E"), (94000, "EBDAD4"), (100000, "55261C"),),
    ((0, "FC9FCB"), (13000, "F8B049"), (21001, "F8B049"), (63000, "FEE7F2"), (67000, "F952A0"), (69000, "C50849"),
     (82001, "B43E85"), (100000, "F8B049"),),
    ((0, "03D4A8"), (25000, "21D6E0"), (75000, "0087E6"), (100000, "005CBF"),),
    ((0, "CCCCFF"), (17999, "99CCFF"), (36000, "9966FF"), (61000, "CC99FF"), (82001, "99CCFF"), (100000, "CCCCFF"),),
    ((0, "FFF200"), (45000, "FF7A00"), (70000, "FF0300"), (100000, "4D0808"),),
    ((0, "8488C4"), (53000, "D4DEFF"), (83000, "D4DEFF"), (100000, "96AB94"),),
    ((0, "DDEBCF"), (50000, "9CB86E"), (100000, "156B13"),),
    ((0, "3399FF"), (16000, "00CCCC"), (47000, "9999FF"), (60001, "2E6792"), (71001, "3333CC"), (81000, "1170FF"),
     (100000, "006699"),),
    ((0, "FBEAC7"), (17999, "FEE7F2"), (36000, "FAC77D"), (61000, "FBA97D"), (82001, "FBD49C"), (100000, "FEE7F2"),),
    ((0, "FFEFD1"), (64999, "F0EBD5"), (100000, "D1C39F"),),
    ((0, "D6B19C"), (30000, "D49E6C"), (70000, "A65528"), (100000, "663012"),),
    ((0, "A603AB"), (21001, "0819FB"), (35001, "1A8D48"), (52000, "FFFF00"), (73000, "EE3F17"), (88000, "E81766"),
     (100000, "A603AB"),),
    ((0, "FF3399"), (25000, "FF6633"), (50000, "FFFF00"), (75000, "01A78F"), (100000, "3366FF"),),
    ((0, "E6DCAC"), (12000, "E6D78A"), (30000, "C7AC4C"), (45000, "E6D78A"), (77000, "C7AC4C"), (100000, "E6DCAC"),),
    ((0, "FBE4AE"), (13000, "BD922A"), (21001, "BD922A"), (63000, "FBE4AE"), (67000, "BD922A"), (69000, "835E17"),
     (82001, "A28949"), (100000, "FAE3B7"),),
    ((0, "825600"), (13000, "FFA800"), (28000, "825600"), (42999, "FFA800"), (58000, "825600"), (72000, "FFA800"),
     (87000, "825600"), (100000, "FFA800"),),
    ((0, "FFFFFF"), (16000, "1F1F1F"), (17999, "FFFFFF"), (42000, "636363"), (53000, "CFCFCF"), (66000, "CFCFCF"),
     (75999, "1F1F1F"), (78999, "FFFFFF"), (100000, "7F7F7F"),),
    ((0, "CBCBCB"), (13000, "5F5F5F"), (21001, "5F5F5F"), (63000, "FFFFFF"), (67000, "B2B2B2"), (69000, "292929"),
     (82001, "777777"), (100000, "EAEAEA"),),
    ((0, "FFFFFF"), (7001, "E6E6E6"), (32001, "7D8496"), (47000, "E6E6E6"), (85001, "7D8496"), (100000, "E6E6E6"),),
    ((0, "000082"), (13000, "0047FF"), (28000, "000082"), (42999, "0047FF"), (58000, "000082"), (72000, "0047FF"),
     (87000, "000082"), (100000, "0047FF"),),
)
#: The lengthways styles' angles at variant 1.
_GRADIENT_ANGLES = {1: 90, 2: 0, 3: 45, 4: 135}


def _linear(channel: float) -> float:
    return channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4


def _gamma(channel: float) -> float:
    return 12.92 * channel if channel <= 0.0031308 else 1.055 * channel ** (1 / 2.4) - 0.055


def _lightness(red: float, green: float, blue: float, change: Callable[[float], float]) -> tuple[float, float, float]:
    """An RGB colour with its HSL lightness changed. Written out because colorsys changed in 3.11."""
    high, low = max(red, green, blue), min(red, green, blue)
    light = (high + low) / 2
    new = max(0.0, min(1.0, change(light)))
    if high == low:
        return new, new, new
    chroma = high - low
    saturation = chroma / (high + low) if light <= 0.5 else chroma / (2 - high - low)
    if high == red:
        hue = (green - blue) / chroma % 6
    elif high == green:
        hue = (blue - red) / chroma + 2
    else:
        hue = (red - green) / chroma + 4
    span = (1 - abs(2 * new - 1)) * saturation
    rise = span * (1 - abs(hue % 2 - 1))
    base = new - span / 2
    spread = ((span, rise, 0.0), (rise, span, 0.0), (0.0, span, rise), (0.0, rise, span), (rise, 0.0, span),
              (span, 0.0, rise))[int(hue) % 6]
    return spread[0] + base, spread[1] + base, spread[2] + base


def painted(rrggbb: str, mods: Iterable[tuple[str, int]]) -> int:
    """A colour under its shade, tint, lumMod and lumOff, in order, as a BGR Long: shade and tint work on linear light,
    lumMod and lumOff on HSL lightness, nothing rounded until the end (tests/fixtures/shape_line_ends.json)."""
    red, green, blue = (int(rrggbb[index:index + 2], 16) / 255 for index in (0, 2, 4))
    for name, value in mods:
        amount = value / 100000
        if name == "lumMod":
            red, green, blue = _lightness(red, green, blue, lambda light: light * amount)
        elif name == "lumOff":
            red, green, blue = _lightness(red, green, blue, lambda light: light + amount)
        elif name == "shade":
            red, green, blue = (_gamma(_linear(one) * amount) for one in (red, green, blue))
        else:
            red, green, blue = (_gamma(_linear(one) * amount + 1 - amount) for one in (red, green, blue))
    return round(red * 255) | round(green * 255) << 8 | round(blue * 255) << 16


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


class _Format(ShapePart):
    """What Fill and Line share: the shape, its kind's style, and the look a change writes."""

    vba_type_name = ""

    def __init__(self, owner: ShapeObject) -> None:
        self.owner = owner

    @property
    def shape(self) -> Shape:
        return self.owner.shape

    def style(self) -> tuple[int | None, int, int, int, float]:
        shape = self.shape
        if not shape.source and shape.kind == "freeform":
            # An open path has a line's style and a closed one an AutoShape's (tests/fixtures/shape_builders.json).
            closed = any(command == "close" for command, _ in shape.path)
            return STYLE["shape" if closed else "line"]
        if shape.source or shape.kind not in STYLE:
            raise VBAUnsupportedError("The fill and line of this shape are not implemented")
        if self.owner.sheet.book.stylesheet.colors.theme != DEFAULT_THEME:
            raise VBAUnsupportedError("The fill and line of a shape under another theme are not implemented")
        return STYLE[shape.kind]

    def look(self) -> FillLine:
        return self.shape.fill_line or FillLine()

    def change(self) -> FillLine:
        self.style()
        if self.shape.kind == "freeform":
            raise _unmeasured("Changing the fill or line of a freeform")
        self.owner.guard("Fill" if isinstance(self, FillFormat) else "Line")
        if self.shape.fill_line is None:
            self.shape.fill_line = FillLine()
        self.owner.restyled()
        return self.shape.fill_line

    def alone(self) -> None:
        """What a ShapeRange does not take yet, which it refuses here."""

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
        unset = (self.shape.kind == "shape" and before.fill_rgb is None and before.fill_alpha is None
                 and before.fill_scheme is None and not before.fill_mods and before.fill_index is None
                 and before.fill_pattern is None)
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
        # msoFillSolid, msoFillPatterned or msoFillGradient, or the mixed value for a line, which has no fill, or one
        # hidden bare.
        if self.filled() and self.look().fill_gradient is not None:
            return VBAInt(3, "Long")
        if self.filled() and self.look().fill_pattern is not None:
            return VBAInt(2, "Long")
        return VBAInt(1 if self.filled() and not self.look().fill_bare else -2, "Long")

    @member
    def Transparency(self) -> object:
        if not self.filled() or self.look().fill_bare:
            # No fill to see through: it reads fully transparent.
            return VBASingle(1.0)
        alpha = self.look().fill_alpha
        found = self.look().fill_gradient
        if found is not None:
            # Each stop has its own, and stops that differ read as mixed.
            alphas = {stop.alpha for stop in found.stops}
            if len(alphas) > 1:
                return VBASingle(_MIXED)
            alpha = alphas.pop()
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
        if look.fill_gradient is not None:
            alpha = None if wanted == 0 else look.fill_alpha
            look.fill_gradient = replace(
                look.fill_gradient, stops=tuple(replace(stop, alpha=alpha) for stop in look.fill_gradient.stops))

    @method
    def Solid(self) -> object:
        self.style()
        before = self.look()
        if self.shape.kind != "shape" or before.fill_hidden or before.line_hidden or before.fill_bare:
            raise _unmeasured("Solid on this kind of shape, or with the fill or line hidden,")
        if before.fill_gradient is not None and any(stop.alpha is not None for stop in before.fill_gradient.stops):
            raise _unmeasured("Solid on a gradient with a transparency")
        look = self.change()
        look.fill_solid = True
        look.fill_pattern = look.fill_back_rgb = look.fill_back_scheme = look.fill_gradient = None
        look.fill_back_mods = ()
        return EMPTY

    @method
    def Patterned(self, pattern: object) -> object:
        self.alone()
        wanted = int(to_integer(pattern, "Long"))
        before = self.look()
        self.style()
        if not self.filled() or not 1 <= wanted <= len(PATTERNS):
            raise error(*_INVALID)
        if before.line_hidden or before.fill_bare:
            raise _unmeasured("A pattern under a hidden line, or on a fill hidden bare,")
        if before.fill_gradient is not None and any(stop.alpha is not None for stop in before.fill_gradient.stops):
            raise _unmeasured("A pattern on a gradient with a transparency")
        look = self.change()
        # It keeps the back colour a pattern or gradient had; a solid fill has none and starts on white.
        look.fill_pattern, look.fill_hidden, look.fill_gradient = PATTERNS[wanted - 1], False, None
        return EMPTY

    def gradient(self) -> FillGradient:
        """The gradient a stop or angle belongs to; other fills have no stops."""
        self.style()
        found = self.look().fill_gradient
        if found is None:
            raise error(*_INVALID)
        return found

    def start_gradient(self, style: object, variant: object, stops: tuple[GradientPoint, ...], *, degree: float | None = None,
                       preset: int | None = None) -> None:
        """Any of the three gradients: the fore and back colours become its first and last stops. A hidden fill stays
        hidden, and a transparency goes."""
        chosen_style, chosen_variant = int(to_integer(style, "Long")), int(to_integer(variant, "Long"))
        if not self.filled() or chosen_style not in (1, 2, 3, 4, 5, 7) or not 1 <= chosen_variant <= 4:
            raise error(*_INVALID)
        if self.look().fill_bare or self.look().fill_index is not None:
            raise _unmeasured("A gradient on a fill hidden bare, or from a SchemeColor,")
        look = self.change()
        made = laid_out(stops, chosen_style, chosen_variant)
        turned = chosen_style in (5, 7) or chosen_variant in (1, 3)
        whole = chosen_style == 5 or (chosen_variant < 3 and (chosen_style != 7 or chosen_variant == 1))
        look.fill_gradient = FillGradient(
            made, None if chosen_style in (5, 7) else (_GRADIENT_ANGLES[chosen_style] + (0 if turned else 180)) * 60000,
            chosen_style if chosen_style in (5, 7) else None, chosen_variant if chosen_style == 5 else None,
            True, degree, preset if whole else None,
            chosen_variant if chosen_style != 5 and chosen_variant > 2 else None)
        look.fill_pattern, look.fill_alpha = None, None
        synced(look)

    def first_colour(self) -> GradientPoint:
        """The fill's fore colour as a stop."""
        look, style = self.look(), self.style()
        scheme = look.fill_scheme if look.fill_rgb is not None or look.fill_scheme else SCHEMES[style[1]]
        return GradientPoint(0, look.fill_rgb, None if look.fill_rgb is not None else scheme, look.fill_mods)

    def back_colour(self) -> GradientPoint:
        look = self.look()
        if look.fill_back_rgb is None and look.fill_back_scheme is None:
            return GradientPoint(100000, "FFFFFF")
        return GradientPoint(100000, look.fill_back_rgb, look.fill_back_scheme, look.fill_back_mods)

    @method
    def TwoColorGradient(self, style: object, variant: object) -> object:
        self.alone()
        self.style()
        if self.filled():
            self.start_gradient(style, variant, (self.first_colour(), self.back_colour()))
        else:
            raise error(*_INVALID)
        return EMPTY

    @method
    def OneColorGradient(self, style: object, variant: object, degree: object) -> object:
        self.alone()
        self.style()
        wanted = float(VBASingle(float(to_number(degree))))
        if not self.filled() or not 0 <= wanted <= 1:
            raise error(*_INVALID)
        first = self.first_colour()
        # Past the middle it tints towards white, before it shades towards black; each reaches the end at 0 and 1.
        change = (("shade", round(wanted * 200000)),) if wanted < 0.5 else (
            (("tint", round((1 - wanted) * 200000)),) if wanted > 0.5 else ())
        second = GradientPoint(100000, first.rgb, first.scheme, first.mods + change)
        self.start_gradient(style, variant, (first, second), degree=_degree_read(wanted))
        return EMPTY

    @method
    def PresetGradient(self, style: object, variant: object, preset: object) -> object:
        self.alone()
        self.style()
        chosen = int(to_integer(preset, "Long"))
        if not self.filled() or not 1 <= chosen <= len(GRADIENT_PRESETS):
            raise error(*_INVALID)
        stops = tuple(GradientPoint(position, colour) for position, colour in GRADIENT_PRESETS[chosen - 1])
        self.start_gradient(style, variant, stops, preset=chosen)
        return EMPTY

    def read_only(self) -> None:
        raise error(450, "Wrong number of arguments or invalid property assignment")

    @member
    def GradientColorType(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        return VBAInt(-2 if found is None else found.colour_type(), "Long")

    @setter("GradientColorType")
    def _gradient_colour_type(self, value: object) -> None:
        self.read_only()

    @member
    def GradientStyle(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        return VBAInt(-2 if found is None else found.shown_style(), "Long")

    @setter("GradientStyle")
    def _gradient_style(self, value: object) -> None:
        self.read_only()

    @member
    def GradientVariant(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        return VBAInt(0 if found is None else found.shown_variant(), "Long")

    @setter("GradientVariant")
    def _gradient_variant(self, value: object) -> None:
        self.read_only()

    @member
    def GradientDegree(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        if found is not None and found.degree is not None:
            return VBASingle(found.degree)
        if found is not None and found.stops[0].colour() == found.stops[-1].colour():
            return VBASingle(0.5)
        raise error(70, "Permission denied")

    @setter("GradientDegree")
    def _gradient_degree(self, value: object) -> None:
        self.read_only()

    @member
    def PresetGradientType(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        return VBAInt(-2 if found is None or found.preset is None else found.preset, "Long")

    @setter("PresetGradientType")
    def _preset_gradient_type(self, value: object) -> None:
        self.read_only()

    @member
    def GradientAngle(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        if found is None or found.angle is None:
            raise error(-2147467259, "Method 'GradientAngle' of object 'FillFormat' failed")
        return VBASingle(found.angle / 60000)

    @setter("GradientAngle")
    def _gradient_angle(self, value: object) -> None:
        self.alone()
        self.style()
        wanted = float(VBASingle(float(to_number(value))))
        found = self.look().fill_gradient
        if found is None or found.angle is None or wanted < 0:
            raise error(-2147467259, "Method 'GradientAngle' of object 'FillFormat' failed")
        look = self.change()
        look.fill_gradient = replace(found, angle=round(wanted * 60000) % 21600000, mirror=None)

    @member
    def RotateWithObject(self) -> object:
        self.alone()
        self.style()
        found = self.look().fill_gradient
        if not self.filled():
            return VBAInt(-2, "Long")
        return VBAInt(-1 if found is not None and found.rotate else 0, "Long")

    @setter("RotateWithObject")
    def _rotate_with_object(self, value: object) -> None:
        self.alone()
        wanted = _shown(value)
        found = self.gradient()
        self.change().fill_gradient = replace(found, rotate=wanted)

    @member
    def GradientStops(self) -> object:
        self.alone()
        return GradientStops(self)

    @member
    def Pattern(self) -> object:
        self.alone()
        self.style()
        chosen = self.look().fill_pattern
        return VBAInt(-2 if chosen is None else PATTERNS.index(chosen) + 1, "Long")

    @setter("Pattern")
    def _pattern(self, value: object) -> None:
        # Patterned sets it; this has no Let.
        raise error(450, "Wrong number of arguments or invalid property assignment")


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
        if before.line_ends or before.line_cap is not None or before.line_align is not None or before.line_open:
            raise _unmeasured("Hiding or showing a line with its ends, a round cap, an InsetPen or no colour written")
        look = self.change()
        look.line_hidden = not wanted
        if wanted:
            look.frozen_width = look.frozen_dash = None
            look.frozen_head = look.frozen_tail = (None, None, None)
            return
        look.line_full = look.effects = True
        look.frozen_width, look.frozen_dash = look.line_width, look.line_dash
        look.frozen_head, look.frozen_tail = look.line_head, look.line_tail
        if self.shape.kind == "line":
            look.no_fill = True
        elif self.shape.kind == "shape":
            look.fill_solid = True

    @member
    def ForeColor(self) -> object:
        return ColorFormat(self, fore=True)

    @member
    def BackColor(self) -> object:
        self.alone()
        return ColorFormat(self, fore=False)

    def arrow(self, end: str, part: int) -> object:
        self.alone()
        self.style()
        chosen = getattr(self.look(), f"line_{end}")[part]
        table = ARROW_STYLES if part == 0 else ARROW_SIZES
        return VBAInt(next((key for key, name in table.items() if name == chosen), 1 if part == 0 else 2), "Long")

    def set_arrow(self, end: str, part: int, value: object) -> None:
        self.alone()
        table = ARROW_STYLES if part == 0 else ARROW_SIZES
        wanted = int(to_integer(value, "Long"))
        self.kept()
        # Only an open line has ends to put arrowheads on.
        if self.shape.kind != "line" or wanted not in table:
            raise error(*_INVALID)
        self.put_arrow(end, part, table[wanted])

    def put_arrow(self, end: str, part: int, written: str) -> None:
        before = self.kept()
        if before.line_ends and part:
            raise _unmeasured("The size of an arrowhead on a line written out with its ends")
        look = self.change()
        ends = list(getattr(look, f"line_{end}"))
        ends[part] = written
        setattr(look, f"line_{end}", tuple(ends))

    @member
    def BeginArrowheadStyle(self) -> object:
        return self.arrow("head", 0)

    @setter("BeginArrowheadStyle")
    def _begin_style(self, value: object) -> None:
        self.set_arrow("head", 0, value)

    @member
    def BeginArrowheadWidth(self) -> object:
        return self.arrow("head", 1)

    @setter("BeginArrowheadWidth")
    def _begin_width(self, value: object) -> None:
        self.set_arrow("head", 1, value)

    @member
    def BeginArrowheadLength(self) -> object:
        return self.arrow("head", 2)

    @setter("BeginArrowheadLength")
    def _begin_length(self, value: object) -> None:
        self.set_arrow("head", 2, value)

    @member
    def EndArrowheadStyle(self) -> object:
        return self.arrow("tail", 0)

    @setter("EndArrowheadStyle")
    def _end_style(self, value: object) -> None:
        self.set_arrow("tail", 0, value)

    @member
    def EndArrowheadWidth(self) -> object:
        return self.arrow("tail", 1)

    @setter("EndArrowheadWidth")
    def _end_width(self, value: object) -> None:
        self.set_arrow("tail", 1, value)

    @member
    def EndArrowheadLength(self) -> object:
        return self.arrow("tail", 2)

    @setter("EndArrowheadLength")
    def _end_length(self, value: object) -> None:
        self.set_arrow("tail", 2, value)

    @member
    def Pattern(self) -> object:
        self.alone()
        self.style()
        chosen = self.look().line_pattern
        return VBAInt(-2 if chosen is None else PATTERNS.index(chosen) + 1, "Long")

    @setter("Pattern")
    def _pattern(self, value: object) -> None:
        self.alone()
        wanted = int(to_integer(value, "Long"))
        before = self.kept()
        if not 1 <= wanted <= len(PATTERNS):
            raise error(*_INVALID)
        if (before.line_hidden or before.line_full or before.line_ends or before.line_dash is not None
                or any(before.line_head) or any(before.line_tail)):
            raise _unmeasured("A pattern on a hidden or dashed line, or one with its ends set,")
        look = self.change()
        if before.line_pattern is None:
            # The solid colour goes: a pattern starts from the default black on white.
            look.line_rgb = look.line_scheme = look.line_index = None
            look.line_mods, look.line_plain, look.line_open = (), False, False
        look.line_pattern = PATTERNS[wanted - 1]

    @member
    def InsetPen(self) -> object:
        self.alone()
        self.style()
        return VBAInt(-1 if self.look().line_align == "in" else 0, "Long")

    @setter("InsetPen")
    def _inset_pen(self, value: object) -> None:
        self.alone()
        wanted = _shown(value)
        before = self.kept()
        if before.line_hidden or before.line_full or before.line_ends:
            raise _unmeasured("InsetPen on a hidden line, or one written out in full,")
        self.change().line_align = "in" if wanted else "ctr"

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
        if before.line_pattern is not None and (
                before.line_rgb is not None or before.line_scheme is not None or before.line_back_rgb is not None
                or before.line_back_scheme is not None):
            raise _unmeasured("The transparency of a pattern with colours of its own")
        if (before.line_align is not None or before.line_open or before.line_index is not None
                or any(before.line_head[1:]) or any(before.line_tail[1:])):
            raise _unmeasured("The transparency of a line with an InsetPen, no colour, or arrowheads sized,")
        look = self.change()
        look.line_ends = True
        look.line_alpha = None if wanted == 0 else round((1 - wanted) * 100000)


class ColorFormat(ShapePart):
    """A fill's or line's colour: its own once RGB, ObjectThemeColor or SchemeColor is set, else its style's theme
    colour, either under what Brightness and TintAndShade add. A back colour shows only on a pattern; elsewhere setting
    it just writes the fill or line out (tests/fixtures/shape_line_ends.json)."""

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

    def scheme(self) -> str | None:
        look = self.owner.look()
        return look.fill_scheme if isinstance(self.owner, FillFormat) else look.line_scheme

    def mods(self) -> tuple[tuple[str, int], ...]:
        look = self.owner.look()
        return look.fill_mods if isinstance(self.owner, FillFormat) else look.line_mods

    def index(self) -> int | None:
        look = self.owner.look()
        return look.fill_index if isinstance(self.owner, FillFormat) else look.line_index

    def patterned(self) -> bool:
        return isinstance(self.owner, LineFormat) and self.owner.look().line_pattern is not None

    def has_pattern(self) -> bool:
        look = self.owner.look()
        if isinstance(self.owner, FillFormat):
            return look.fill_pattern is not None or look.fill_gradient is not None
        return look.line_pattern is not None

    def regraded(self, look: FillLine, *, fore: bool) -> None:
        """A fill's fore or back colour set on a gradient is its first or last stop, which then has no transparency,
        and the gradient is no longer the preset or one-colour one it was made as."""
        found = look.fill_gradient
        if found is None:
            return
        stops = list(found.stops)
        if fore:
            stops[0] = GradientPoint(stops[0].position, look.fill_rgb, look.fill_scheme, look.fill_mods)
        else:
            stops[-1] = GradientPoint(stops[-1].position, look.fill_back_rgb, look.fill_back_scheme, look.fill_back_mods)
        look.fill_gradient = replace(found, stops=tuple(stops), degree=None, preset=None, mirror=None)

    def back(self) -> tuple[str | None, str | None]:
        """A pattern's back colour: its RGB, and its theme colour."""
        look = self.owner.look()
        if isinstance(self.owner, FillFormat):
            return look.fill_back_rgb, look.fill_back_scheme
        return look.line_back_rgb, look.line_back_scheme

    def back_mods(self) -> tuple[tuple[str, int], ...]:
        return self.owner.look().fill_back_mods if isinstance(self.owner, FillFormat) else ()

    def style_shade(self) -> int:
        """The shade the style still gives this colour."""
        if (not isinstance(self.owner, LineFormat) or self.own() is not None or self.scheme() is not None
                or self.patterned() or self.owner.look().line_plain):
            return 0
        return _LINE_SHADES[self.owner.shape.kind]

    def theme_hex(self, scheme: str) -> str:
        return self.owner.owner.sheet.book.stylesheet.colors.theme[_THEME_SLOT[scheme]]

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
        colour, theme = self.styled()
        if not self.fore:
            back_rgb, back_scheme = self.back()
            if back_rgb is not None:
                return VBAInt(painted(back_rgb, self.back_mods()), "Long")
            if back_scheme is not None:
                return VBAInt(painted(self.theme_hex(back_scheme), self.back_mods()), "Long")
            # A new shape's back colour is white; a line, with no fill, reads 0.
            return VBAInt(0 if colour is None else 16777215, "Long")
        if self.lost() or self.bare():
            return VBAInt(16777215, "Long")
        if colour is None:
            return VBAInt(0, "Long")
        own, scheme, mods = self.own(), self.scheme(), self.mods()
        if own is not None:
            # A SchemeColor's brightness is already in the RGB the file holds.
            return VBAInt(painted(own, () if self.index() is not None else mods), "Long")
        if scheme is not None:
            return VBAInt(painted(self.theme_hex(scheme), mods), "Long")
        if self.patterned():
            return VBAInt(0, "Long")
        shade = self.style_shade()
        if not mods and (shade or not isinstance(self.owner, LineFormat) or not self.owner.look().line_plain):
            return VBAInt(colour, "Long")
        chain = ((("shade", shade),) if shade else ()) + mods
        return VBAInt(painted(self.theme_hex(SCHEMES[theme]), chain), "Long")

    def set_back(self, rgb: str | None, scheme: str | None) -> None:
        """A back colour: a pattern's own, else it writes the fill or line out as it stands, and a text box's, already
        written out, does not change."""
        colour, _ = self.styled()
        before = self.owner.look()
        if isinstance(self.owner, FillFormat):
            if colour is None:
                raise error(*_INVALID)
            if before.fill_pattern is not None or before.fill_gradient is not None:
                if before.fill_hidden:
                    raise _unmeasured("The back colour of a hidden pattern or gradient")
                look = self.owner.change()
                look.fill_back_rgb, look.fill_back_scheme, look.fill_back_mods = rgb, scheme, ()
                self.regraded(look, fore=False)
            elif self.owner.shape.kind == "shape":
                self.owner.vba_get("Solid")
            return
        if self.patterned():
            if before.line_hidden:
                raise _unmeasured("The back colour of a hidden pattern")
            look = self.owner.change()
            look.line_back_rgb, look.line_back_scheme = rgb, scheme
            return
        if self.owner.shape.kind == "textBox":
            return
        if before.line_hidden or before.line_lost or before.line_rgb is not None or before.line_scheme is not None:
            raise _unmeasured("A back colour on a line that is hidden or has a colour of its own")
        self.owner.change().line_open = True

    def set_own(self, rgb: str | None, scheme: str | None, index: int | None) -> None:
        """A fore colour that is not plain RGB, which takes the place of what the colour had. On a fill it shows a
        hidden fill; on a line it goes on being hidden."""
        before = self.owner.look()
        if before.line_lost or before.fill_bare:
            raise _unmeasured("Changing a colour once a hidden line or a bare fill has lost it")
        if isinstance(self.owner, LineFormat):
            if before.line_hidden and self.patterned():
                raise _unmeasured("The colour of a hidden pattern")
            look = self.owner.change()
            look.line_rgb, look.line_scheme, look.line_index = rgb, scheme, index
            look.line_mods, look.line_plain = (), False
            return
        if before.line_hidden:
            raise _unmeasured("A fill's theme colour while the line is hidden")
        if index is not None and before.fill_gradient is not None:
            raise _unmeasured("A SchemeColor on a gradient")
        look = self.owner.change()
        look.fill_rgb, look.fill_scheme, look.fill_index, look.fill_mods = rgb, scheme, index, ()
        look.fill_hidden = False
        self.regraded(look, fore=True)

    def set_mods(self, mods: tuple[tuple[str, int], ...]) -> None:
        """A brightness or tint, which on a fill shows a hidden fill as a colour does."""
        before = self.owner.look()
        if before.line_lost or before.fill_bare or self.patterned():
            raise _unmeasured("Brightness or TintAndShade on this colour")
        if self.index() is not None:
            # A palette colour's brightness goes into the RGB too, as well as beside it.
            if self.mods() or not isinstance(self.owner, FillFormat) or before.line_hidden or (
                    {name for name, _ in mods} - {"lumMod", "lumOff"}):
                raise _unmeasured("Anything but one brightness on a fill's SchemeColor")
            own = self.own()
            assert own is not None
            look = self.owner.change()
            look.fill_mods, look.fill_rgb, look.fill_hidden = mods, _rgb(painted(own, mods)), False
            return
        if isinstance(self.owner, LineFormat):
            self.owner.change().line_mods = mods
            return
        if before.line_hidden:
            raise _unmeasured("A fill's brightness or tint while the line is hidden")
        look = self.owner.change()
        look.fill_mods, look.fill_hidden, look.fill_solid = mods, False, True
        self.regraded(look, fore=True)

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        self.owner.alone()
        wanted = int(to_integer(value, "Long")) & 0xFFFFFF
        colour, _ = self.styled()
        if not self.fore:
            self.set_back(_rgb(wanted), None)
            return
        if colour is None:
            # A line has no fill to colour.
            raise error(*_INVALID)
        before = self.owner.look()
        if before.line_lost:
            raise _unmeasured("Changing a colour once a hidden line is lost")
        if isinstance(self.owner, LineFormat):
            if self.owner.shape.kind == "textBox":
                raise VBAUnsupportedError("Setting a text box's line colour is not implemented")
            if before.line_hidden and self.patterned():
                raise _unmeasured("The colour of a hidden pattern")
            look = self.owner.change()
            look.line_rgb, look.line_scheme, look.line_index = _rgb(wanted), None, None
            look.line_mods, look.line_plain = (), False
            return
        if before.line_hidden and (self.owner.shape.kind != "shape" or before.fill_gradient is not None):
            raise _unmeasured("A fill colour for this kind of shape, or on a gradient, while its line is hidden")
        # Setting it shows a hidden fill, and loses a hidden line.
        look = self.owner.change()
        look.fill_rgb, look.fill_hidden, look.fill_bare, look.line_lost = _rgb(wanted), False, False, look.line_hidden
        look.fill_scheme, look.fill_index, look.fill_mods = None, None, ()
        self.regraded(look, fore=True)

    @member
    def ObjectThemeColor(self) -> object:
        colour, theme = self.styled()
        if colour is None:
            return VBAInt(-2 if isinstance(self.owner, FillFormat) else theme, "Long")
        if not self.fore:
            scheme = self.back()[1]
            return VBAInt(0 if scheme is None else _scheme_number(scheme), "Long")
        scheme = self.scheme()
        if scheme is not None:
            return VBAInt(_scheme_number(scheme), "Long")
        if self.own() is not None or self.lost() or self.bare() or self.patterned():
            return VBAInt(0, "Long")
        return VBAInt(theme, "Long")

    @setter("ObjectThemeColor")
    def _object_theme_color(self, value: object) -> None:
        self.owner.alone()
        wanted = int(to_integer(value, "Long"))
        colour, _ = self.styled()
        if wanted not in SCHEMES or (colour is None and self.fore):
            raise error(*_INVALID)
        if self.fore:
            self.set_own(None, SCHEMES[wanted], None)
        else:
            self.set_back(None, SCHEMES[wanted])

    @member
    def Type(self) -> object:
        # msoColorTypeRGB once set, msoColorTypeScheme from the style, mixed where there is no fill.
        colour, _ = self.styled()
        if colour is None:
            return VBAInt(-2, "Long")
        if not self.fore:
            return VBAInt(1 if self.back()[1] is None else 2, "Long")
        if self.scheme() is not None or self.index() is not None:
            return VBAInt(2, "Long")
        return VBAInt(1 if self.own() is not None or self.lost() or self.bare() or self.patterned() else 2, "Long")

    @member
    def Brightness(self) -> object:
        self.owner.alone()
        self.styled()
        found = dict(self.mods()) if self.fore else {}
        if "lumOff" in found:
            return VBASingle(found["lumOff"] / 100000)
        return VBASingle(found["lumMod"] / 100000 - 1 if "lumMod" in found else 0.0)

    @setter("Brightness")
    def _brightness(self, value: object) -> None:
        self.owner.alone()
        wanted = float(VBASingle(float(to_number(value))))
        colour, _ = self.styled()
        if not -1 <= wanted <= 1:
            raise error(*_INVALID)
        if not self.fore:
            self.set_back_only()
            return
        if colour is None:
            # A line has no fill to brighten.
            raise error(-2147467259, "Method 'Brightness' of object 'ColorFormat' failed")
        if any(name in ("tint", "shade") for name, _ in self.mods()) and any(
                name == "lumMod" for name, _ in self.mods()):
            raise _unmeasured("A second brightness over a tint")
        kept = tuple(one for one in self.mods() if one[0] not in ("lumMod", "lumOff"))
        lum = (("lumMod", round((1 - wanted) * 100000)), ("lumOff", round(wanted * 100000))) if wanted > 0 else (
            ("lumMod", round((1 + wanted) * 100000)),)
        self.set_mods(kept + lum)

    @member
    def TintAndShade(self) -> object:
        self.owner.alone()
        self.styled()
        found = dict(self.mods()) if self.fore else {}
        if "tint" in found:
            return VBASingle(1 - found["tint"] / 100000)
        if "shade" in found:
            return VBASingle(found["shade"] / 100000 - 1)
        shade = self.style_shade() if self.fore else 0
        return VBASingle(shade / 100000 - 1 if shade else 0.0)

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        self.owner.alone()
        wanted = float(VBASingle(float(to_number(value))))
        colour, _ = self.styled()
        if not -1 <= wanted <= 1:
            raise error(*_INVALID)
        if not self.fore:
            self.set_back_only()
            return
        if colour is None:
            # A line has no fill to tint, and nothing happens.
            return
        mods = self.mods()
        if any(name in ("tint", "shade") for name, _ in mods) and any(name == "lumMod" for name, _ in mods):
            raise _unmeasured("A second tint over a brightness")
        if self.style_shade() and mods:
            raise _unmeasured("A tint over a brightness on the shade a style gives")
        kept = tuple(one for one in mods if one[0] not in ("tint", "shade"))
        tint = () if wanted == 0 else (
            ("tint", round((1 - wanted) * 100000)),) if wanted > 0 else (("shade", round((1 + wanted) * 100000)),)
        self.set_mods(kept + tint)
        if isinstance(self.owner, LineFormat):
            self.owner.change().line_plain = True

    def set_back_only(self) -> None:
        """A back colour's brightness or tint, which only writes the fill or line out."""
        colour, _ = self.styled()
        if colour is None or self.has_pattern():
            raise _unmeasured("The brightness or tint of this back colour")
        self.set_back(None, None)

    @member
    def SchemeColor(self) -> object:
        self.owner.alone()
        colour, _ = self.styled()
        index = self.index() if self.fore else None
        if index is not None:
            return VBAInt(index, "Long")
        if colour is None:
            return VBAInt(-2, "Long")
        # Excel reads the palette's nearest colour by a measure no fit has found: only an exact match is known.
        shown = _rgb(int(to_integer(self.vba_get("RGB"), "Long")))
        palette = self.owner.owner.sheet.book.stylesheet.colors.indexed[8:64]
        if shown not in palette:
            raise VBAUnsupportedError("SchemeColor of a colour not in the palette is not implemented")
        return VBAInt(palette.index(shown) + 8, "Long")

    @setter("SchemeColor")
    def _scheme_color(self, value: object) -> None:
        self.owner.alone()
        wanted = int(to_integer(value, "Long"))
        colour, _ = self.styled()
        if not self.fore:
            if colour is None:
                raise _unmeasured("The back SchemeColor of a line's fill")
            self.set_back_only()
            return
        if colour is None or not 0 <= wanted <= 80:
            raise error(*_INVALID)
        if self.owner.shape.kind == "textBox" or 66 <= wanted <= 79:
            raise _unmeasured("The SchemeColor of a text box, or one of Windows' own colours,")
        palette = self.owner.owner.sheet.book.stylesheet.colors.indexed
        self.set_own(_SCHEME_EXTRAS.get(wanted) or palette[wanted], None, wanted)


class GradientStops(ShapePart):
    """A gradient's stops, in the order the file lists them, which their positions need not follow. A fill with no
    gradient gives this object, but every use of it fails. It has no Parent."""

    vba_type_name = "GradientStops"

    def __init__(self, fill: FillFormat) -> None:
        self.fill = fill

    def vba_iterate(self) -> Iterator[object]:
        try:
            stops = self.fill.gradient().stops
        except VBARuntimeError as failure:
            # Walking the stops of a fill with no gradient fails with its own number (tests/fixtures/shape_gradient_stops.json).
            raise error(-2147352567, "Exception occurred.") from failure
        return iter([GradientStop(self.fill, index) for index in range(len(stops))])

    @member
    def Parent(self) -> object:
        raise error(438, "Object doesn't support this property or method")

    @member
    def Count(self) -> object:
        return VBAInt(len(self.fill.gradient().stops), "Long")

    @member(default=True, kind="method")
    def Item(self, index: object) -> object:
        chosen = int(to_integer(index, "Long"))
        if not 1 <= chosen <= len(self.fill.gradient().stops):
            raise error(*_INVALID)
        return GradientStop(self.fill, chosen - 1)

    @member
    def Application(self) -> object:
        return self.fill.owner.sheet.book.application

    @method
    def Insert2(self, rgb: object = MISSING, position: object = MISSING, transparency: object = MISSING, index: object = MISSING, brightness: object = MISSING) -> object:
        """Insert with a Brightness (tests/fixtures/shape_gradient_stops.json). Excel writes its alpha before the Brightness
        here, after it elsewhere, which the model's stops do not hold, so a stop with both reports itself."""
        if rgb is MISSING or position is MISSING:
            raise error(450, "Wrong number of arguments or invalid property assignment")
        wanted = 0.0 if brightness is MISSING else float(VBASingle(float(to_number(brightness))))
        if not -1 <= wanted <= 1:
            self.fill.gradient()
            raise error(*_INVALID)
        if wanted and transparency is not MISSING and to_number(transparency):
            raise _unmeasured("A new gradient stop with both a transparency and a brightness")
        self.Insert(rgb, position, transparency, index)
        if wanted:
            stops = self.fill.gradient().stops
            at = len(stops) - 1 if index is MISSING else int(to_integer(index, "Long")) - 1
            changed(self.fill, stops[:at] + (replace(stops[at], mods=_lum(wanted)),) + stops[at + 1:])
        return EMPTY

    @method
    def Insert(self, rgb: object, position: object, transparency: object = MISSING, index: object = MISSING) -> object:
        found = self.fill.gradient()
        where = float(VBASingle(float(to_number(position))))
        see_through = 0.0 if transparency is MISSING else float(VBASingle(float(to_number(transparency))))
        chosen = -1 if index is MISSING else int(to_integer(index, "Long"))
        if not 0 <= where <= 1 or not 0 <= see_through <= 1 or not (chosen == -1 or 1 <= chosen <= len(found.stops) + 1):
            raise error(*_INVALID)
        added = GradientPoint(round(where * 100000), _rgb(int(to_integer(rgb, "Long")) & 0xFFFFFF), None, (),
                              None if see_through == 0 else round((1 - see_through) * 100000))
        at = len(found.stops) if chosen == -1 else chosen - 1
        changed(self.fill, found.stops[:at] + (added,) + found.stops[at:])
        return EMPTY

    @method
    def Delete(self, index: object = MISSING) -> object:
        found = self.fill.gradient()
        chosen = -1 if index is MISSING else int(to_integer(index, "Long"))
        # Two stops are as few as a gradient has.
        if len(found.stops) <= 2 or not (chosen == -1 or 1 <= chosen <= len(found.stops)):
            raise error(*_INVALID)
        at = len(found.stops) - 1 if chosen == -1 else chosen - 1
        changed(self.fill, found.stops[:at] + found.stops[at + 1:])
        return EMPTY


class GradientStop(ShapePart):
    vba_type_name = "GradientStop"

    @member
    def Application(self) -> object:
        return self.fill.owner.sheet.book.application

    def __init__(self, fill: FillFormat, index: int) -> None:
        self.fill, self.index = fill, index

    def stop(self) -> GradientPoint:
        return self.fill.gradient().stops[self.index]

    def put(self, stop: GradientPoint) -> None:
        stops = self.fill.gradient().stops
        changed(self.fill, stops[:self.index] + (stop,) + stops[self.index + 1:])

    @member
    def Color(self) -> object:
        return StopColor(self)

    @member
    def Position(self) -> object:
        return VBASingle(self.stop().position / 100000)

    @setter("Position")
    def _position(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        self.put(replace(self.stop(), position=round(wanted * 100000)))

    @member
    def Transparency(self) -> object:
        alpha = self.stop().alpha
        return VBASingle(0.0 if alpha is None else (100000 - alpha) / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        self.put(replace(self.stop(), alpha=None if wanted == 0 else round((1 - wanted) * 100000)))


class StopColor(ShapePart):
    """A gradient stop's colour, with what a fill's ForeColor has."""

    vba_type_name = "ColorFormat"

    def __init__(self, stop: GradientStop) -> None:
        self.stop = stop

    def put(self, stop: GradientPoint) -> None:
        if self.stop.stop().alpha is not None:
            raise _unmeasured("Changing the colour of a stop with a transparency")
        self.stop.put(stop)

    @member
    def RGB(self) -> object:
        stop = self.stop.stop()
        colours = self.stop.fill.owner.sheet.book.stylesheet.colors.theme
        base = stop.rgb if stop.rgb is not None else colours[_THEME_SLOT[stop.scheme or "accent1"]]
        return VBAInt(painted(base, stop.mods), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        wanted = _rgb(int(to_integer(value, "Long")) & 0xFFFFFF)
        self.put(replace(self.stop.stop(), rgb=wanted, scheme=None, mods=()))

    @member
    def ObjectThemeColor(self) -> object:
        scheme = self.stop.stop().scheme
        return VBAInt(0 if scheme is None else _scheme_number(scheme), "Long")

    @setter("ObjectThemeColor")
    def _object_theme_color(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in SCHEMES:
            raise error(*_INVALID)
        self.put(replace(self.stop.stop(), rgb=None, scheme=SCHEMES[wanted], mods=()))

    @member
    def Type(self) -> object:
        return VBAInt(2 if self.stop.stop().scheme is not None else 1, "Long")

    @member
    def Brightness(self) -> object:
        found = dict(self.stop.stop().mods)
        if "lumOff" in found:
            return VBASingle(found["lumOff"] / 100000)
        return VBASingle(found["lumMod"] / 100000 - 1 if "lumMod" in found else 0.0)

    @setter("Brightness")
    def _brightness(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not -1 <= wanted <= 1:
            raise error(*_INVALID)
        stop = self.stop.stop()
        kept = tuple(one for one in stop.mods if one[0] not in ("lumMod", "lumOff"))
        self.put(replace(stop, mods=kept + _lum(wanted)))

    @member
    def TintAndShade(self) -> object:
        found = dict(self.stop.stop().mods)
        if "tint" in found:
            return VBASingle(1 - found["tint"] / 100000)
        return VBASingle(found["shade"] / 100000 - 1 if "shade" in found else 0.0)

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not -1 <= wanted <= 1:
            raise error(*_INVALID)
        stop = self.stop.stop()
        kept = tuple(one for one in stop.mods if one[0] not in ("tint", "shade"))
        self.put(replace(stop, mods=kept + _tint(wanted)))


def laid_out(stops: tuple[GradientPoint, ...], style: int, variant: int) -> tuple[GradientPoint, ...]:
    """Where a new gradient puts its stops. A corner gradient keeps them. Variant 2 from the centre runs them backwards.
    Variants 3 and 4 fold them into each half, with the last stop or the first not repeated (tests/fixtures/shape_fill_gradients.json)."""
    if style == 5 or variant == 1 or (variant == 2 and style != 7):
        return stops
    if variant == 2:
        return tuple(replace(stop, position=100000 - stop.position) for stop in stops)
    if variant == 3:
        return (tuple(replace(stop, position=_folded(0, 1, stop)) for stop in stops)
                + tuple(replace(stop, position=_folded(100000, -1, stop)) for stop in stops[:-1]))
    return (tuple(replace(stop, position=_folded(50000, -1, stop)) for stop in stops)
            + tuple(replace(stop, position=_folded(50000, 1, stop)) for stop in stops[1:]))


def _single(value: float) -> float:
    return float(VBASingle(value))


#: The folds that round a half down: position, the fold's start, and which way it runs. Every other odd position of all 24
#: presets rounds up at all four folds. No rounding rule fits these, so they are as measured.
_FOLDS_DOWN = {(32001, 0, 1), (64999, 0, 1), (35001, 50000, -1), (71001, 50000, -1), (82001, 50000, -1),
               (85001, 50000, -1)}


def _folded(start: int, sign: int, stop: GradientPoint) -> int:
    """A stop's place in half of a folded gradient, half its position from one end of that half."""
    exact = start * 2 + sign * stop.position
    return (exact - 1) // 2 if (stop.position, start, sign) in _FOLDS_DOWN and exact % 2 else (exact + 1) // 2


def _degree_read(wanted: float) -> float:
    """The degree a one-colour gradient reads, which comes back through Excel's Singles below the middle."""
    return wanted if wanted >= 0.5 else _single(_single(1 - _single(1 - _single(2 * wanted))) / 2)


def synced(look: FillLine) -> None:
    """The fore and back colours are a gradient's first and last stops."""
    found = look.fill_gradient
    if found is None:
        return
    first, last = found.stops[0], found.stops[-1]
    look.fill_rgb, look.fill_scheme, look.fill_mods = first.rgb, first.scheme, first.mods
    look.fill_back_rgb, look.fill_back_scheme, look.fill_back_mods = last.rgb, last.scheme, last.mods


def changed(fill: FillFormat, stops: tuple[GradientPoint, ...]) -> None:
    """New stops: it is no longer the preset or one-colour gradient it was made as."""
    look = fill.change()
    found = look.fill_gradient
    assert found is not None
    look.fill_gradient = replace(found, stops=stops, degree=None, preset=None, mirror=None)
    synced(look)


def _lum(wanted: float) -> tuple[tuple[str, int], ...]:
    if wanted > 0:
        return (("lumMod", round((1 - wanted) * 100000)), ("lumOff", round(wanted * 100000)))
    return (("lumMod", round((1 + wanted) * 100000)),)


def _tint(wanted: float) -> tuple[tuple[str, int], ...]:
    if wanted == 0:
        return ()
    return (("tint", round((1 - wanted) * 100000)),) if wanted > 0 else (("shade", round((1 + wanted) * 100000)),)


class RangeStops(Created):
    """A ShapeRange's GradientStops (tests/fixtures/shape_range_more_format.json): the first gradient's count, and a
    stop added or deleted on each. A single stop through a range was not measured."""

    vba_type_name = "GradientStops"

    def __init__(self, fill: RangeFill) -> None:
        self.fill = fill

    def each(self) -> list[GradientStops]:
        return [GradientStops(part) for part in self.fill.bearers()]

    @member(default=True, kind="method")
    def Item(self, index: object) -> object:
        raise VBAUnsupportedError("A single gradient stop through a ShapeRange is not implemented")

    @member
    def Count(self) -> object:
        graded = self.fill.graded()
        if not graded:
            raise error(*_INVALID)
        return GradientStops(graded[0]).vba_get("Count")

    @method
    def Insert(self, rgb: object, position: object, transparency: object = MISSING, index: object = MISSING) -> object:
        for stops in self.each():
            stops.Insert(rgb, position, transparency, index)
        return EMPTY


def _scheme_number(scheme: str) -> int:
    return next(number for number, name in SCHEMES.items() if name == scheme)


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

    def alone(self) -> None:
        raise VBAUnsupportedError("This member of a ShapeRange's fill is not implemented")

    def each(self, name: str, *args: object) -> object:
        """A method on each shape with a fill. A range of lines alone was not measured."""
        if not any(part.filled() for part in self.parts()):
            raise _unmeasured("This fill method on a range of lines")
        for part in self.bearers():
            part.vba_call(name, args)
        return EMPTY

    @method
    def Patterned(self, pattern: object) -> object:
        return self.each("Patterned", pattern)

    @member
    def Pattern(self) -> object:
        return _alike([part.vba_get("Pattern") for part in self.bearers()], VBAInt(-2, "Long"))

    @method
    def TwoColorGradient(self, style: object, variant: object) -> object:
        self.each("TwoColorGradient", style, variant)
        # A line among them takes the gradient too, though it reads as having no fill.
        made = self.bearers()[0].look().fill_gradient
        for part in self.parts():
            if not part.filled():
                self.lined_with(part, made)
        return EMPTY

    def lined_with(self, part: FillFormat, made: FillGradient | None) -> None:
        part.owner.guard("Fill")
        if part.shape.fill_line is None:
            part.shape.fill_line = FillLine()
        part.shape.fill_line.fill_gradient = made
        part.owner.restyled()

    def only_filled(self, what: str) -> None:
        if any(not part.filled() for part in self.parts()):
            raise _unmeasured(f"{what} on a range with a line")

    @method
    def OneColorGradient(self, style: object, variant: object, degree: object) -> object:
        self.only_filled("OneColorGradient")
        return self.each("OneColorGradient", style, variant, degree)

    @method
    def PresetGradient(self, style: object, variant: object, preset: object) -> object:
        self.only_filled("PresetGradient")
        return self.each("PresetGradient", style, variant, preset)

    def graded(self) -> list[FillFormat]:
        return [part for part in self.bearers() if part.look().fill_gradient is not None]

    @member
    def GradientColorType(self) -> object:
        # Shapes with no gradient do not count here, as they do in the rest.
        return _alike([part.vba_get("GradientColorType") for part in self.graded()] or [VBAInt(-2, "Long")],
                      VBAInt(-2, "Long"))

    @member
    def GradientStyle(self) -> object:
        return _alike([part.vba_get("GradientStyle") for part in self.bearers()], VBAInt(-2, "Long"))

    @member
    def GradientVariant(self) -> object:
        if self.vba_get("GradientStyle") == -2:
            return VBAInt(0, "Long")
        return _alike([part.vba_get("GradientVariant") for part in self.bearers()], VBAInt(0, "Long"))

    @member
    def GradientDegree(self) -> object:
        found = [part.vba_get("GradientDegree") for part in self.bearers()]
        if any(one != found[0] for one in found[1:]):
            raise error(70, "Permission denied")
        return found[0]

    @member
    def PresetGradientType(self) -> object:
        return _alike([part.vba_get("PresetGradientType") for part in self.bearers()], VBAInt(-2, "Long"))

    @member
    def GradientAngle(self) -> object:
        found = [part.vba_get("GradientAngle") for part in self.bearers()]
        if any(one != found[0] for one in found[1:]):
            raise error(-2147467259, "Method 'GradientAngle' of object 'FillFormat' failed")
        return found[0]

    @setter("GradientAngle")
    def _gradient_angle(self, value: object) -> None:
        for part in self.bearers():
            part.vba_set("GradientAngle", value)

    @member
    def RotateWithObject(self) -> object:
        if not self.agreed():
            return VBAInt(-2, "Long")
        return _alike([part.vba_get("RotateWithObject") for part in self.bearers()], VBAInt(-2, "Long"))

    @setter("RotateWithObject")
    def _rotate_with_object(self, value: object) -> None:
        for part in self.bearers():
            part.vba_set("RotateWithObject", value)

    @member
    def GradientStops(self) -> object:
        return RangeStops(self)

    def parts(self) -> list[FillFormat]:
        on_sheet = self.owner.sheet.shapes_
        return [FillFormat(one) for one in self.range.shapes if any(shape is one.shape for shape in on_sheet)]

    def bearers(self) -> list[FillFormat]:
        parts = self.parts()
        return [part for part in parts if part.filled()] or parts

    def agreed(self) -> bool:
        return len({(part.look().fill_hidden, part.look().fill_bare, int(to_integer(part.vba_get("Type"), "Long")))
                    for part in self.bearers()}) == 1

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

    def alone(self) -> None:
        raise VBAUnsupportedError("This member of a ShapeRange's line is not implemented")

    def arrow(self, end: str, part: int) -> object:
        return _alike([one.arrow(end, part) for one in self.parts()], VBAInt(-2, "Long"))

    def set_arrow(self, end: str, part: int, value: object) -> None:
        """Every shape takes it, a closed outline too, so long as one is a line."""
        table = ARROW_STYLES if part == 0 else ARROW_SIZES
        wanted = int(to_integer(value, "Long"))
        parts = self.parts()
        if not any(one.shape.kind == "line" for one in parts) or wanted not in table:
            raise error(*_INVALID)
        for one in parts:
            one.put_arrow(end, part, table[wanted])

    @member
    def BackColor(self) -> object:
        return RangeColor(self, fore=False)

    @member
    def Pattern(self) -> object:
        # A line with no pattern does not count against one with a pattern.
        found = [value for value in (part.vba_get("Pattern") for part in self.parts()) if value != -2]
        return _alike(found or [VBAInt(-2, "Long")], VBAInt(-2, "Long"))

    @setter("Pattern")
    def _pattern(self, value: object) -> None:
        self.spread("Pattern", value)

    @member
    def InsetPen(self) -> object:
        return self.alike("InsetPen", VBAInt(-2, "Long"))

    @setter("InsetPen")
    def _inset_pen(self, value: object) -> None:
        self.spread("InsetPen", value)

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
        if self.fore or not isinstance(self.whole, RangeFill) or any(
                part.look().fill_pattern is not None or part.look().fill_gradient is not None for part in self.whole.bearers()):
            for part in self.parts():
                part.vba_set("RGB", value)
            return
        # A range's back colour changes no colour: it writes the fill out, as Solid does.
        if any(part.shape.kind != "shape" for part in self.whole.bearers()):
            raise _unmeasured("Setting the back colour of this range")
        for part in self.whole.bearers():
            part.vba_get("Solid")

    def spread(self, name: str, value: object) -> None:
        if isinstance(self.whole, RangeFill):
            self.whole.still(name)
        for part in self.parts():
            part.vba_set(name, value)

    def unit_read(self, name: str) -> object:
        if self.unit() is None:
            return VBASingle(_MIXED)
        return _alike([part.vba_get(name) for part in self.parts()], VBASingle(_MIXED))

    @setter("ObjectThemeColor")
    def _object_theme_color(self, value: object) -> None:
        self.spread("ObjectThemeColor", value)

    @member
    def Brightness(self) -> object:
        return self.unit_read("Brightness")

    @setter("Brightness")
    def _brightness(self, value: object) -> None:
        self.spread("Brightness", value)

    @member
    def TintAndShade(self) -> object:
        return self.unit_read("TintAndShade")

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        self.spread("TintAndShade", value)

    @setter("SchemeColor")
    def _scheme_color(self, value: object) -> None:
        self.spread("SchemeColor", value)

    @member
    def ObjectThemeColor(self) -> object:
        unit = self.unit()
        return VBAInt(-2, "Long") if unit is None else unit[1]

    @member
    def Type(self) -> object:
        unit = self.unit()
        return VBAInt(-2, "Long") if unit is None else unit[2]
