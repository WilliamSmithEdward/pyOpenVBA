"""A workbook's cell formats: its stylesheet, its colours, and the format each cell carries.

``xl/styles.xml`` holds a cell's format as an index into ``cellXfs``, and
each xf points at a font, a fill, a border and a number format. Here that
becomes one frozen :class:`Style` per xf, made of frozen parts, so a cell
holds its format as a value and two cells with the same format hold equal
values. The parts keep the XML they were read from, which is what lets an
untouched style write back byte for byte; a changed cell looks its new
format up with :meth:`Stylesheet.index_of`, which reuses an equal xf when
there is one and appends only the parts the stylesheet lacks, spelled the
way Excel spells them.

Colours follow what live Excel answers (tests/fixtures/range_format.json):

* a theme colour is tinted in Windows' integer HLS space (``ColorRGBToHLS``
  and ``ColorHLSToRGB``, 240 steps), with the tint held as ``n / 32767``
  and the luminance products truncated one by one;
* ``ColorIndex`` is the palette entry nearest by summed channel distance,
  the first one on a tie;
* the file's theme index counts light 1 before dark 1, so VBA's
  ``ThemeColor`` is that index plus one.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field, replace
from fractions import Fraction
from typing import Final, TypeVar

from pyopenvba._xml import attributes as _attributes
from pyopenvba._xml import escape as _escape
from pyopenvba._xml import unescape as _unescape
from pyopenvba.apps.excel._number_format import BUILTIN, SPELLED_OUT, from_file, to_file
from pyopenvba.exceptions import VBAUnsupportedError

_T = TypeVar("_T")

# --- colours -----------------------------------------------------------------------------

#: Excel's default 56-colour palette, ColorIndex 1 to 56, as RRGGBB.
PALETTE: Final = (
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
)
#: indexed="64" and "65": the system's window text and window colours, Excel's automatic.
SYSTEM_FOREGROUND: Final = 64
SYSTEM_BACKGROUND: Final = 65
#: The new Office theme, for a workbook that carries no theme part: light 1, dark 1,
#: light 2, dark 2, six accents, hyperlink, followed hyperlink, in file index order.
DEFAULT_THEME: Final = ("FFFFFF", "000000", "E8E8E8", "0E2841", "156082", "E97132", "196B24", "0F9ED5",
                        "A02B93", "4EA72E", "467886", "96607D")
TINT_SCALE: Final = 32767
XL_AUTOMATIC: Final = -4105
XL_NONE: Final = -4142


@dataclass(frozen=True, slots=True)
class Color:
    """One colour the way a stylesheet records it.

    ``kind`` is ``rgb`` (``value`` is AARRGGBB), ``theme`` (the file's theme
    index), ``indexed`` (a palette slot; 64 and 65 are the system colours)
    or ``auto``. ``tint`` is n, where the tint is ``n / 32767``.
    """

    kind: str
    value: str | int = ""
    tint: int = 0

    @classmethod
    def rgb(cls, red: int, green: int, blue: int) -> Color:
        return cls("rgb", f"FF{red:02X}{green:02X}{blue:02X}")

    @classmethod
    def from_bgr(cls, value: int) -> Color:
        """A VBA colour, which packs blue, green and red from the high byte down."""
        return cls.rgb(value & 255, (value >> 8) & 255, (value >> 16) & 255)


FOREGROUND: Final = Color("indexed", SYSTEM_FOREGROUND)
BACKGROUND: Final = Color("indexed", SYSTEM_BACKGROUND)


def tint_n(text: str) -> int:
    """The n of a tint as a file writes it, where Excel stores ``n / 32767``."""
    value = float(text) * TINT_SCALE
    nearest = round(value)
    return nearest if abs(value - nearest) < 1e-6 else int(value)


def tint_of(value: float) -> int:
    """The n Excel keeps when a macro sets TintAndShade: it truncates toward zero."""
    return int(value * TINT_SCALE)


#: How far a double may sit from its 15-digit spelling, relative to its
#: size, for Excel to write those 15 digits: 5/16 of DBL_EPSILON.  All
#: 2,184 row heights Excel can hold on a 96-DPI display are written by
#: this rule, and no other threshold separates them.
_SHORT_ENOUGH = Fraction(5, 16) * Fraction(2) ** -52


def excel_number(value: float) -> str:
    """A measurement as Excel writes one into XML: a row height, a tint.

    Fifteen significant digits when the double is within 5/16 of an
    epsilon of them, relative to its size, and seventeen otherwise.  That
    is not a round-trip test: 20.1 and 19.9 round-trip at 15 digits, yet
    Excel writes ``20.100000000000001`` and ``19.899999999999999``, because
    near the bottom of a binade the same half-ulp error is relatively
    larger.  15.2, further up its binade, stays ``15.2``.
    """
    short = f"{value:.15g}"
    if abs(Fraction(value) - Fraction(short)) <= abs(Fraction(value)) * _SHORT_ENOUGH:
        return short
    return f"{value:.17g}"


def excel_double(value: float) -> str:
    """A tint as Excel writes one into XML.

    Below 0.1 always 17 digits, in scientific notation; from 0.1 up the
    way :func:`excel_number` spells any measurement: ``0.249977111117893``,
    ``0.39997558519241921``, ``9.9978637043366805E-2``.
    """
    if value != 0 and abs(value) < 0.1:
        mantissa, exponent = f"{value:.16e}".split("e")
        mantissa = mantissa.rstrip("0").rstrip(".")
        return f"{mantissa}E{int(exponent)}"
    return excel_number(value)


def _color_attributes(color: Color) -> str:
    if color.kind == "auto":
        text = 'auto="1"'
    elif color.kind == "rgb":
        text = f'rgb="{color.value}"'
    elif color.kind == "theme":
        text = f'theme="{color.value}"'
    else:
        text = f'indexed="{color.value}"'
    if color.tint:
        text += f' tint="{excel_double(color.tint / TINT_SCALE)}"'
    return text


def color_xml(tag: str, color: Color) -> str:
    return f"<{tag} {_color_attributes(color)}/>"


def parse_color(element: str) -> Color | None:
    """The colour a ``<color>``, ``<fgColor>`` or ``<bgColor>`` element names."""
    attributes = _attributes(element)
    tint = tint_n(attributes["tint"]) if "tint" in attributes else 0
    if attributes.get("auto") in ("1", "true"):
        return Color("auto", "", tint)
    if "rgb" in attributes:
        return Color("rgb", attributes["rgb"].upper(), tint)
    if "theme" in attributes:
        return Color("theme", int(attributes["theme"]), tint)
    if "indexed" in attributes:
        return Color("indexed", int(attributes["indexed"]), tint)
    return None


def _hls(red: int, green: int, blue: int) -> tuple[int, int, int]:
    """Windows' ColorRGBToHLS: hue, luminance and saturation on a 0..240 scale."""
    high, low = max(red, green, blue), min(red, green, blue)
    luminance = ((high + low) * 240 + 255) // 510
    if high == low:
        return 160, luminance, 0
    delta = high - low
    if luminance <= 120:
        saturation = ((high + low) // 2 + delta * 240) // (high + low)
    else:
        saturation = ((510 - high - low) // 2 + delta * 240) // (510 - high - low)
    red_norm = (delta // 2 + high * 40 - red * 40) // delta
    green_norm = (delta // 2 + high * 40 - green * 40) // delta
    blue_norm = (delta // 2 + high * 40 - blue * 40) // delta
    if red == high:
        hue = blue_norm - green_norm
    elif green == high:
        hue = 80 + red_norm - blue_norm
    else:
        hue = 160 + green_norm - red_norm
    if hue < 0:
        hue += 240
    elif hue > 240:
        hue -= 240
    return hue, luminance, saturation


def _rgb(hue: int, luminance: int, saturation: int) -> tuple[int, int, int]:
    """Windows' ColorHLSToRGB, except that a grey rounds where Windows truncates."""
    if not saturation:
        grey = (luminance * 255 + 120) // 240
        return grey, grey, grey
    if luminance > 120:
        mid2 = saturation + luminance - (saturation * luminance + 120) // 240
    else:
        mid2 = ((saturation + 240) * luminance + 120) // 240
    mid1 = luminance * 2 - mid2

    def channel(value: int) -> int:
        value = value - 240 if value > 240 else value + 240 if value < 0 else value
        if value > 160:
            level = mid1
        elif value > 120:
            level = ((160 - value) * (mid2 - mid1) + 20) // 40 + mid1
        elif value > 40:
            level = mid2
        else:
            level = (value * (mid2 - mid1) + 20) // 40 + mid1
        return (level * 255 + 120) // 240

    return channel(hue + 80), channel(hue), channel(hue - 80)


def tinted(rrggbb: str, n: int) -> str:
    """A colour lightened (n > 0) or darkened (n < 0) the way Excel does it."""
    if n == 0:
        return rrggbb
    red, green, blue = (int(rrggbb[i:i + 2], 16) for i in (0, 2, 4))
    hue, luminance, saturation = _hls(red, green, blue)
    if n < 0:
        luminance = luminance * (TINT_SCALE + n) // TINT_SCALE
    else:
        luminance = luminance * (TINT_SCALE - n) // TINT_SCALE + (240 - 240 * (TINT_SCALE - n) // TINT_SCALE)
    return "".join(f"{channel:02X}" for channel in _rgb(hue, luminance, saturation))


def bgr(rrggbb: str) -> int:
    """RRGGBB as the number VBA's Color properties answer."""
    return int(rrggbb[0:2], 16) + int(rrggbb[2:4], 16) * 256 + int(rrggbb[4:6], 16) * 65536


class Colors:
    """What the colours of one workbook resolve to: its theme and its palette."""

    def __init__(self, theme: tuple[str, ...] = DEFAULT_THEME, palette: tuple[str, ...] = ()) -> None:
        self.theme = theme
        #: indexed 0..63; 8..63 are ColorIndex 1..56.
        self.indexed = palette or (PALETTE[:8] + PALETTE)

    def rrggbb(self, color: Color | None, automatic: str) -> str:
        """The colour a macro reads, with ``automatic`` standing in for Excel's automatic."""
        if color is None or color.kind == "auto":
            base = automatic
        elif color.kind == "rgb":
            base = str(color.value)[-6:]
        elif color.kind == "theme":
            index = int(color.value)
            base = self.theme[index] if 0 <= index < len(self.theme) else automatic
        else:
            index = int(color.value)
            if index == SYSTEM_FOREGROUND:
                base = "000000"
            elif index == SYSTEM_BACKGROUND:
                base = "FFFFFF"
            else:
                base = self.indexed[index] if 0 <= index < len(self.indexed) else automatic
        return tinted(base, color.tint) if color is not None else base

    def color_index(self, color: Color | None, automatic: str) -> int:
        """VBA's ColorIndex: a palette slot, or xlColorIndexAutomatic."""
        if color is None or color.kind == "auto":
            return XL_AUTOMATIC
        if color.kind == "indexed":
            # A tinted palette colour still answers its own slot.
            index = int(color.value)
            if index in (SYSTEM_FOREGROUND, SYSTEM_BACKGROUND):
                return XL_AUTOMATIC
            if 8 <= index < 64:
                return index - 7
            if 0 <= index < 8:
                return index + 1
        return self.nearest(self.rrggbb(color, automatic))

    def nearest(self, rrggbb: str) -> int:
        """The palette slot nearest by summed channel distance, the first on a tie."""
        wanted = [int(rrggbb[i:i + 2], 16) for i in (0, 2, 4)]
        best, found = 1 << 30, 1
        for index, entry in enumerate(self.indexed[8:64], start=1):
            distance = sum(abs(int(entry[i * 2:i * 2 + 2], 16) - wanted[i]) for i in range(3))
            if distance < best:
                best, found = distance, index
        return found

    def palette_color(self, index: int) -> Color:
        """The colour a macro means by ColorIndex n, written the way Excel writes it."""
        return Color("indexed", index + 7)


def parse_theme(xml: str) -> tuple[str, ...]:
    """The twelve theme colours in the order a stylesheet's theme index counts them."""
    scheme = re.search(r"<a:clrScheme\b.*?</a:clrScheme>", xml, re.DOTALL)
    if scheme is None:
        return DEFAULT_THEME
    found: dict[str, str] = {}
    for name, body in re.findall(r"<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>(.*?)</a:\1>",
                                 scheme.group(0), re.DOTALL):
        value = re.search(r'(?:srgbClr val|lastClr)="([0-9A-Fa-f]{6})"', body)
        if value:
            found[name] = value.group(1).upper()
    order = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
             "hlink", "folHlink")
    return tuple(found.get(name, DEFAULT_THEME[index]) for index, name in enumerate(order))


# --- the parts of a format ---------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Font:
    name: str = "Aptos Narrow"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    strike: bool = False
    #: "", "single", "double", "singleAccounting" or "doubleAccounting".
    underline: str = ""
    #: "", "superscript" or "subscript".
    vert_align: str = ""
    color: Color | None = None
    family: str = ""
    charset: str = ""
    scheme: str = ""
    outline: bool = False
    shadow: bool = False
    condense: bool = False
    extend: bool = False


@dataclass(frozen=True, slots=True)
class Fill:
    """A fill in Excel's own terms: a pattern, its foreground and its background.

    A solid fill shows its foreground, so the cell colour a macro sets is
    the foreground there and the background of every other pattern.
    ``gradient`` keeps a gradient fill's XML, which nothing here edits.
    """

    pattern: str = "none"
    foreground: Color = FOREGROUND
    background: Color = BACKGROUND
    gradient: str = ""
    #: Whether the file writes it without a bgColor, as Excel writes most of its built-in cell styles' fills;
    #: only a fill with the default background can be, and Excel writes every other one with its bgColor.
    bare: bool = False

    def __post_init__(self) -> None:
        if self.bare and (self.background != BACKGROUND or self.pattern in ("none", "gradient")):
            object.__setattr__(self, "bare", False)


@dataclass(frozen=True, slots=True)
class Side:
    #: The file's style name ("thin", "medium", "dashed", ...); "" is no border.
    style: str = ""
    color: Color | None = None


NO_SIDE: Final = Side()


@dataclass(frozen=True, slots=True)
class Border:
    left: Side = NO_SIDE
    right: Side = NO_SIDE
    top: Side = NO_SIDE
    bottom: Side = NO_SIDE
    #: One line style for both diagonals; the flags say which of them show it.
    diagonal: Side = NO_SIDE
    diagonal_up: bool = False
    diagonal_down: bool = False
    #: A dxf's vertical and horizontal sides, kept as read.
    extra: str = ""


@dataclass(frozen=True, slots=True)
class Alignment:
    horizontal: str = ""
    vertical: str = ""
    text_rotation: int = 0
    wrap: bool = False
    indent: int = 0
    relative_indent: int = 0
    justify_last_line: bool = False
    shrink: bool = False
    reading_order: int = 0


@dataclass(frozen=True, slots=True)
class Protection:
    locked: bool = True
    hidden: bool = False


@dataclass(frozen=True, slots=True)
class Style:
    """One cell format: what an xf in ``cellXfs`` describes."""

    number_format: str = "General"
    font: Font = field(default_factory=Font)
    fill: Fill = field(default_factory=Fill)
    border: Border = field(default_factory=Border)
    alignment: Alignment = field(default_factory=Alignment)
    protection: Protection = field(default_factory=Protection)
    #: The cell style it derives from: the index of its xf in cellStyleXfs, or for a style the file does not
    #: hold yet the place the stylesheet keeps it until a save writes it (Stylesheet.cell_styles).
    base: int = 0
    quote_prefix: bool = False
    #: The parts of the format it applies itself, the xf's apply flags: each
    #: part it holds differently from its cell style, and every part a macro
    #: has set, even to what it already was. Excel keeps a format a macro
    #: made bold and then not bold apart from the default one, so a cell
    #: that went through that stays; a file's flags are not read, since
    #: Excel takes a format that differs only in them as the same one.
    applied: frozenset[str] = frozenset()
    #: The parts it takes from its cell style, which its xf points at by the
    #: style's own entries: Excel keeps a second copy of the Normal font that
    #: some styles use, and a cell given one of them points at the copy.
    styled: frozenset[str] = frozenset()


#: The parts of a format, in the order an xf's apply flags are written.
APPLY_FLAGS: Final = {"number_format": "applyNumberFormat", "font": "applyFont", "fill": "applyFill",
                      "border": "applyBorder", "alignment": "applyAlignment", "protection": "applyProtection"}
PARTS: Final = frozenset(APPLY_FLAGS)


def applying(style: Style, part: str, **changes: object) -> Style:
    """The format with one part set by a macro: changed as asked, and applied from now on."""
    return replace(style, applied=style.applied | {part}, styled=style.styled - {part},
                   **changes)  # type: ignore[arg-type]


# --- reading a stylesheet -------------------------------------------------------------------------

_BOOL_OFF = ("0", "false")


def _flag(body: str, tag: str) -> bool:
    match = re.search(rf"<{tag}\b([^>]*)/?>", body)
    return match is not None and _attributes(f"<{tag}{match.group(1)}>").get("val", "1") not in _BOOL_OFF


def _value(body: str, tag: str) -> str | None:
    match = re.search(rf"<{tag}\b[^>]*/?>", body)
    return None if match is None else _attributes(match.group(0)).get("val", "")


def parse_font(xml: str) -> Font:
    underline = _value(xml, "u")
    color = re.search(r"<color\b[^>]*/>", xml)
    size = _value(xml, "sz")
    return Font(
        name=_unescape(_value(xml, "name") or ""),
        size=float(size) if size else 11.0,
        bold=_flag(xml, "b"),
        italic=_flag(xml, "i"),
        strike=_flag(xml, "strike"),
        underline="" if underline is None or underline == "none" else (underline or "single"),
        vert_align="" if (_value(xml, "vertAlign") or "baseline") == "baseline" else str(_value(xml, "vertAlign")),
        color=parse_color(color.group(0)) if color else None,
        family=_value(xml, "family") or "",
        charset=_value(xml, "charset") or "",
        scheme=_value(xml, "scheme") or "",
        outline=_flag(xml, "outline"),
        shadow=_flag(xml, "shadow"),
        condense=_flag(xml, "condense"),
        extend=_flag(xml, "extend"),
    )


def parse_fill(xml: str) -> Fill:
    if "<gradientFill" in xml:
        return Fill(pattern="gradient", gradient=xml)
    pattern = re.search(r"<patternFill\b[^>]*>", xml)
    kind = _attributes(pattern.group(0)).get("patternType", "none") if pattern else "none"
    foreground = re.search(r"<fgColor\b[^>]*/>", xml)
    background = re.search(r"<bgColor\b[^>]*/>", xml)
    return Fill(
        pattern=kind,
        foreground=(parse_color(foreground.group(0)) if foreground else None) or FOREGROUND,
        background=(parse_color(background.group(0)) if background else None) or BACKGROUND,
        bare=foreground is not None and background is None,
    )


def _parse_side(xml: str, tag: str) -> Side:
    match = re.search(rf"<{tag}\b([^>]*?)(?:/>|>(.*?)</{tag}>)", xml, re.DOTALL)
    if match is None:
        return NO_SIDE
    style = _attributes(f"<{tag}{match.group(1)}>").get("style", "")
    if not style or style == "none":
        return NO_SIDE
    color = re.search(r"<color\b[^>]*/>", match.group(2) or "")
    return Side(style, parse_color(color.group(0)) if color else None)


def parse_border(xml: str) -> Border:
    opening = re.match(r"<border\b[^>]*>", xml)
    head = _attributes(opening.group(0)) if opening else {}
    extra = "".join(re.findall(r"<(?:vertical|horizontal)\b.*?(?:/>|</(?:vertical|horizontal)>)", xml, re.DOTALL))
    return Border(
        left=_parse_side(xml, "left") if "<left" in xml else _parse_side(xml, "start"),
        right=_parse_side(xml, "right") if "<right" in xml else _parse_side(xml, "end"),
        top=_parse_side(xml, "top"),
        bottom=_parse_side(xml, "bottom"),
        diagonal=_parse_side(xml, "diagonal"),
        diagonal_up=head.get("diagonalUp", "0") in ("1", "true"),
        diagonal_down=head.get("diagonalDown", "0") in ("1", "true"),
        extra=extra,
    )


def parse_alignment(xml: str) -> Alignment:
    match = re.search(r"<alignment\b[^>]*/?>", xml)
    if match is None:
        return Alignment()
    found = _attributes(match.group(0))

    def number(name: str) -> int:
        try:
            return int(found.get(name, "0"))
        except ValueError:
            return 0

    return Alignment(
        horizontal="" if found.get("horizontal", "general") == "general" else found["horizontal"],
        vertical="" if found.get("vertical", "bottom") == "bottom" else found["vertical"],
        text_rotation=number("textRotation"),
        wrap=found.get("wrapText", "0") in ("1", "true"),
        indent=number("indent"),
        relative_indent=number("relativeIndent"),
        justify_last_line=found.get("justifyLastLine", "0") in ("1", "true"),
        shrink=found.get("shrinkToFit", "0") in ("1", "true"),
        reading_order=number("readingOrder"),
    )


def parse_protection(xml: str) -> Protection:
    match = re.search(r"<protection\b[^>]*/?>", xml)
    if match is None:
        return Protection()
    found = _attributes(match.group(0))
    return Protection(locked=found.get("locked", "1") not in _BOOL_OFF, hidden=found.get("hidden", "0") in ("1", "true"))


# --- writing the parts, Excel's way ------------------------------------------------------------------

#: Excel's font families for the fonts a macro commonly names; others keep the family they replace.
_FAMILIES: Final = {"arial": "2", "calibri": "2", "aptos": "2", "aptos narrow": "2", "verdana": "2",
                    "tahoma": "2", "segoe ui": "2", "times new roman": "1", "georgia": "1", "cambria": "1",
                    "courier new": "3", "consolas": "3", "lucida console": "3", "comic sans ms": "4"}


def font_family(name: str, fallback: str) -> str:
    return _FAMILIES.get(name.lower(), fallback)


def _number(value: float) -> str:
    return str(int(value)) if value == int(value) else repr(value)


def font_xml(font: Font) -> str:
    parts = ["<font>"]
    for flag, tag in ((font.bold, "b"), (font.italic, "i"), (font.strike, "strike"), (font.condense, "condense"),
                      (font.extend, "extend"), (font.outline, "outline"), (font.shadow, "shadow")):
        if flag:
            parts.append(f"<{tag}/>")
    if font.underline:
        parts.append("<u/>" if font.underline == "single" else f'<u val="{font.underline}"/>')
    if font.vert_align:
        parts.append(f'<vertAlign val="{font.vert_align}"/>')
    parts.append(f'<sz val="{_number(font.size)}"/>')
    # An automatic colour is written as no colour at all, so a tint a macro gave it
    # does not survive a save; Excel still keeps the font as an entry of its own.
    if font.color is not None and font.color.kind != "auto":
        parts.append(color_xml("color", font.color))
    parts.append(f'<name val="{_escape(font.name)}"/>')
    if font.family:
        parts.append(f'<family val="{font.family}"/>')
    if font.charset:
        parts.append(f'<charset val="{font.charset}"/>')
    if font.scheme:
        parts.append(f'<scheme val="{font.scheme}"/>')
    parts.append("</font>")
    return "".join(parts)


def fill_xml(fill: Fill) -> str:
    if fill.gradient:
        return fill.gradient
    if fill.pattern == "none":
        return '<fill><patternFill patternType="none"/></fill>'
    foreground = fill.foreground != FOREGROUND
    background = fill.background != BACKGROUND
    if fill.bare:
        colors = color_xml("fgColor", fill.foreground)
    elif fill.pattern == "solid" or foreground:
        colors = color_xml("fgColor", fill.foreground) + color_xml("bgColor", fill.background)
    elif background:
        colors = color_xml("bgColor", fill.background)
    else:
        return f'<fill><patternFill patternType="{fill.pattern}"/></fill>'
    return f'<fill><patternFill patternType="{fill.pattern}">{colors}</patternFill></fill>'


def _side_xml(tag: str, side: Side) -> str:
    if not side.style:
        return f"<{tag}/>"
    color = color_xml("color", side.color) if side.color is not None else ""
    return f'<{tag} style="{side.style}">{color}</{tag}>'


def border_xml(border: Border) -> str:
    head = "<border"
    if border.diagonal_up:
        head += ' diagonalUp="1"'
    if border.diagonal_down:
        head += ' diagonalDown="1"'
    return (head + ">" + _side_xml("left", border.left) + _side_xml("right", border.right)
            + _side_xml("top", border.top) + _side_xml("bottom", border.bottom)
            + _side_xml("diagonal", border.diagonal) + border.extra + "</border>")


def alignment_xml(alignment: Alignment) -> str:
    if alignment == Alignment():
        return ""
    found: list[str] = []
    if alignment.horizontal:
        found.append(f'horizontal="{alignment.horizontal}"')
    if alignment.vertical:
        found.append(f'vertical="{alignment.vertical}"')
    if alignment.text_rotation:
        found.append(f'textRotation="{alignment.text_rotation}"')
    if alignment.wrap:
        found.append('wrapText="1"')
    if alignment.indent:
        found.append(f'indent="{alignment.indent}"')
    if alignment.relative_indent:
        found.append(f'relativeIndent="{alignment.relative_indent}"')
    if alignment.justify_last_line:
        found.append('justifyLastLine="1"')
    if alignment.shrink:
        found.append('shrinkToFit="1"')
    if alignment.reading_order:
        found.append(f'readingOrder="{alignment.reading_order}"')
    return "<alignment " + " ".join(found) + "/>"


def protection_xml(protection: Protection) -> str:
    if protection == Protection():
        return ""
    found: list[str] = []
    if not protection.locked:
        found.append('locked="0"')
    if protection.hidden:
        found.append('hidden="1"')
    return "<protection " + " ".join(found) + "/>"


# --- the stylesheet -----------------------------------------------------------------------------------

_SECTION = {"numFmts": "numFmt", "fonts": "font", "fills": "fill", "borders": "border",
            "cellStyleXfs": "xf", "cellXfs": "xf", "cellStyles": "cellStyle"}
#: The order a stylesheet's children come in, for adding a section a file lacks.
_ORDER: Final = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
                 "tableStyles", "colors", "extLst")
#: The currency and accounting formats Excel spells out, in the order its own table holds them: a save that
#: gives the file several new cell styles using them writes them in this order (tests/fixtures/cell_styles).
_SPELLED_ORDER: Final = (5, 6, 7, 8, 42, 41, 44, 43)
#: The theme's heading and body fonts where a workbook carries no theme part.
DEFAULT_THEME_FONTS: Final = ("Aptos Display", "Aptos Narrow")


def parse_theme_fonts(xml: str) -> tuple[str, str]:
    """The theme's heading and body fonts: the Latin typefaces of its major and minor fonts."""
    found: list[str] = []
    for kind, default in zip(("major", "minor"), DEFAULT_THEME_FONTS, strict=True):
        latin = re.search(rf'<a:{kind}Font>\s*<a:latin typeface="([^"]*)"', xml)
        found.append(_unescape(latin.group(1)) if latin else default)
    return found[0], found[1]


def _children(xml: str, section: str) -> list[str]:
    match = re.search(rf"<{section}\b[^>]*?(?:/>|>(.*?)</{section}>)", xml, re.DOTALL)
    if match is None or match.group(1) is None:
        return []
    tag = _SECTION[section]
    return re.findall(rf"<{tag}\b[^>]*?/>|<{tag}\b[^>]*?>.*?</{tag}>", match.group(1), re.DOTALL)


@dataclass(slots=True)
class _BaseStyle:
    """A cell style's xf as the file has it: its parts by index, which decide the apply flags of the xfs
    that derive from it."""

    number_format: int = 0
    font: int = 0
    fill: int = 0
    border: int = 0
    alignment: Alignment = Alignment()
    protection: Protection = Protection()


@dataclass(slots=True)
class CellStyle:
    """A cell style: the format it gives a cell, the parts of a format it gives, its name, and its xf.

    Every xf in the file's cellStyleXfs is one, named where the file's
    cellStyles name it. One the file does not hold yet -- a built-in style
    a cell was just given, or one a macro added -- waits with no xf until a
    save writes it.
    """

    #: The format a cell takes from it; its base is its own place in Stylesheet.cell_styles.
    format: Style
    #: The parts of a cell's format it sets, IncludeNumber to IncludeProtection.
    includes: frozenset[str] = PARTS
    #: Its name; "" for an xf the file's cellStyles do not name.
    name: str = ""
    #: Excel's builtinId, None for a style a user made.
    builtin: int | None = None
    #: Its xf in the file's cellStyleXfs, once the file holds it.
    xf: int | None = None
    #: Where it comes in Excel's own order of its standard built-in styles, which a save writes new ones in;
    #: -1 for any other style.
    rank: int = -1
    #: The order the styles made while the workbook was open came in, which a save writes the rest in.
    made: int = 0


class Stylesheet:
    """``xl/styles.xml`` read into formats and cell styles, and the entries a save has to add to it."""

    def __init__(self, xml: str, theme_xml: str = "") -> None:
        self.xml = xml
        palette = tuple(
            _attributes(one).get("rgb", "FF000000")[-6:].upper()
            for one in re.findall(r"<rgbColor\b[^>]*/>", xml)
        )
        self.colors = Colors(parse_theme(theme_xml) if theme_xml else DEFAULT_THEME,
                             palette if len(palette) >= 64 else ())
        #: The theme's heading and body typefaces, which Excel's built-in cell styles are written in.
        self.theme_fonts = parse_theme_fonts(theme_xml) if theme_xml else DEFAULT_THEME_FONTS
        #: Every format by id, as Range.NumberFormat spells it; the file's own spelling is read through from_file.
        self.number_formats: dict[int, str] = dict(BUILTIN)
        #: The ids the file's numFmts spell out, which a save does not spell out again.
        self._declared: set[int] = set()
        for element in _children(xml, "numFmts"):
            found = _attributes(element)
            try:
                identifier = int(found.get("numFmtId", "0"))
            except ValueError:
                continue
            self.number_formats[identifier] = from_file(_unescape(found.get("formatCode", "")))
            self._declared.add(identifier)
        self.fonts = [parse_font(one) for one in _children(xml, "fonts")] or [Font()]
        self.fills = [parse_fill(one) for one in _children(xml, "fills")] or [Fill()]
        self.borders = [parse_border(one) for one in _children(xml, "borders")] or [Border()]
        style_xfs = _children(xml, "cellStyleXfs")
        self._file_style_xfs = len(style_xfs)
        self.bases = [self._base(one) for one in style_xfs] or [_BaseStyle()]
        #: Every cell style: the file's xfs in cellStyleXfs order, then those waiting for a save.
        self.cell_styles = [self._cell_style(one, index) for index, one in enumerate(style_xfs)] or [
            CellStyle(Style(font=self.fonts[0], styled=PARTS), name="Normal", builtin=0, xf=0)]
        #: The file's cellStyles elements in order, with those a save adds put in by name.
        self._named = _children(xml, "cellStyles")
        for element in self._named:
            found = _attributes(element)
            try:
                xf = int(found.get("xfId", ""))
            except ValueError:
                continue
            if 0 <= xf < len(self.cell_styles) and not self.cell_styles[xf].name:
                entry = self.cell_styles[xf]
                entry.name = _unescape(found.get("name", ""))
                builtin = found.get("builtinId", "")
                entry.builtin = int(builtin) if builtin.isdigit() else None
        self._named_changed = False
        self._made = 0
        self.styles = [self._style(one) for one in _children(xml, "cellXfs")] or [Style(font=self.fonts[0])]
        self._index: dict[Style, int] = {}
        for index, style in enumerate(self.styles):
            self._index.setdefault(style, index)
        self._added: dict[str, list[str]] = {
            name: [] for name in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs")}

    # -- reading

    def _part(self, items: list[_T], index: int, fallback: _T) -> _T:
        return items[index] if 0 <= index < len(items) else fallback

    def _base(self, element: str) -> _BaseStyle:
        found = _attributes(element)
        return _BaseStyle(
            number_format=int(found.get("numFmtId", "0") or 0),
            font=int(found.get("fontId", "0") or 0),
            fill=int(found.get("fillId", "0") or 0),
            border=int(found.get("borderId", "0") or 0),
            alignment=parse_alignment(element),
            protection=parse_protection(element),
        )

    def _cell_style(self, element: str, index: int) -> CellStyle:
        """A cellStyleXfs xf: the format it gives, and the parts it gives, those its apply flags do not turn off."""
        base = self.bases[index]
        found = _attributes(element)
        includes = frozenset(part for part, flag in APPLY_FLAGS.items() if found.get(flag, "1") not in _BOOL_OFF)
        own = Style(
            number_format=self.number_formats.get(base.number_format, "General"),
            font=self._part(self.fonts, base.font, self.fonts[0]),
            fill=self._part(self.fills, base.fill, Fill()),
            border=self._part(self.borders, base.border, Border()),
            alignment=base.alignment,
            protection=base.protection,
            base=index,
            styled=PARTS,
        )
        return CellStyle(own, includes, xf=index)

    def _style(self, element: str) -> Style:
        found = _attributes(element)
        number = int(found.get("numFmtId", "0") or 0)
        font, fill, border = (int(found.get(key, "0") or 0) for key in ("fontId", "fillId", "borderId"))
        base_index = int(found.get("xfId", "0") or 0)
        alignment, protection = parse_alignment(element), parse_protection(element)
        base = self.bases[base_index] if 0 <= base_index < len(self.bases) else _BaseStyle()
        differs = {"number_format": number != base.number_format, "font": font != base.font,
                   "fill": fill != base.fill, "border": border != base.border,
                   "alignment": alignment != base.alignment, "protection": protection != base.protection}
        applied = frozenset(part for part, yes in differs.items() if yes)
        return Style(
            number_format=self.number_formats.get(number, "General"),
            font=self._part(self.fonts, font, self.fonts[0]),
            fill=self._part(self.fills, fill, Fill()),
            border=self._part(self.borders, border, Border()),
            alignment=alignment,
            protection=protection,
            base=base_index,
            quote_prefix=found.get("quotePrefix", "0") in ("1", "true"),
            applied=applied,
            styled=PARTS - applied,
        )

    @property
    def default(self) -> Style:
        return self.styles[0]

    def style(self, index: int) -> Style:
        return self.styles[index] if 0 <= index < len(self.styles) else self.styles[0]

    # -- cell styles

    def named(self, name: str) -> int | None:
        """The place of the cell style a name names, in any case, among the ones the stylesheet holds."""
        key = name.casefold()
        for index, entry in enumerate(self.cell_styles):
            if entry.name and entry.name.casefold() == key:
                return index
        return None

    def add_cell_style(self, entry: CellStyle) -> int:
        """Hold a cell style the file does not have yet, which a save writes when it has to; its place."""
        self._made += 1
        entry.made = self._made
        self.cell_styles.append(entry)
        index = len(self.cell_styles) - 1
        entry.format = replace(entry.format, base=index, styled=PARTS)
        return index

    def write_cell_styles(self, in_use: set[int], lead_font: Font | None = None) -> None:
        """Give the file the cell styles a save writes that it lacks.

        Each style a user made is written, used or not; a built-in one only
        while a cell uses it. The standard built-in styles come in Excel's
        own order, then the rest in the order they were made. ``lead_font``
        is the one Excel's own table holds ahead of every built-in style's,
        the theme's body font, written first where a style uses it as it is.
        """
        waiting = [entry for index, entry in enumerate(self.cell_styles)
                   if entry.xf is None and (entry.builtin is None or index in in_use)]
        waiting.sort(key=lambda entry: (entry.rank < 0, entry.rank, entry.made))
        if lead_font is not None and any(entry.format.font == lead_font for entry in waiting):
            self._entry("fonts", self.fonts, lead_font, font_xml, start=1)
        codes = {entry.format.number_format for entry in waiting}
        for identifier in _SPELLED_ORDER:
            code = self.number_formats.get(identifier, "")
            if identifier not in self._declared and code in codes:
                self._declare(identifier, code)
        for entry in waiting:
            self._register(entry)

    def _register(self, entry: CellStyle) -> None:
        """Write a cell style's xf, its parts and its name.

        A style's font is never the file's first one: where the Normal font
        is what it wants, Excel gives it a copy of that font.
        """
        if not self._file_style_xfs:
            raise VBAUnsupportedError("adding a cell style to a stylesheet with no cellStyleXfs is not implemented")
        own = entry.format
        number = self._number_format_id(own.number_format)
        font = self._entry("fonts", self.fonts, own.font, font_xml, start=1)
        fill = self._entry("fills", self.fills, own.fill, fill_xml)
        border = self._entry("borders", self.borders, own.border, border_xml)
        head = f'<xf numFmtId="{number}" fontId="{font}" fillId="{fill}" borderId="{border}"'
        for part, name in APPLY_FLAGS.items():
            if part not in entry.includes:
                head += f' {name}="0"'
        inner = alignment_xml(own.alignment) + protection_xml(own.protection)
        self._added["cellStyleXfs"].append(head + (f">{inner}</xf>" if inner else "/>"))
        self.bases.append(_BaseStyle(number, font, fill, border, own.alignment, own.protection))
        entry.xf = len(self.bases) - 1
        element = (f'<cellStyle name="{_escape(entry.name)}" xfId="{entry.xf}"'
                   + (f' builtinId="{entry.builtin}"' if entry.builtin is not None else "") + "/>")
        self._name(element, entry.name)

    def _name(self, element: str, name: str) -> None:
        """Put a new cellStyle among the file's, which Excel keeps in the order the Styles collection lists."""
        from pyopenvba.apps.excel._sort import text_key

        key = text_key(name, False)
        for position, existing in enumerate(self._named):
            if text_key(_unescape(_attributes(existing).get("name", "")), False) > key:
                self._named.insert(position, element)
                break
        else:
            self._named.append(element)
        self._named_changed = True

    # -- adding

    def index_as_read(self, style: Style, xf: int) -> int:
        """The xf a format is written with: the one it was read with while that still says the same."""
        if 0 <= xf < len(self.styles) and self.styles[xf] == style:
            return xf
        return self.index_of(style)

    def index_of(self, style: Style) -> int:
        """The xf for ``style``, reusing an equal one or adding what the stylesheet lacks.

        A part the format takes from its cell style, and still holds as the
        style has it, points at the style's own entry.
        """
        found = self._index.get(style)
        if found is not None:
            return found
        entry = self.cell_styles[style.base] if 0 <= style.base < len(self.cell_styles) else self.cell_styles[0]
        if entry.xf is None:
            self._register(entry)
        assert entry.xf is not None
        base = self.bases[entry.xf]

        def taken(part: str) -> bool:
            return part in style.styled and getattr(style, part) == getattr(entry.format, part)

        number = base.number_format if taken("number_format") else self._number_format_id(style.number_format)
        font = base.font if taken("font") else self._entry("fonts", self.fonts, style.font, font_xml)
        fill = base.fill if taken("fill") else self._entry("fills", self.fills, style.fill, fill_xml)
        border = base.border if taken("border") else self._entry("borders", self.borders, style.border, border_xml)
        head = (f'<xf numFmtId="{number}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="{entry.xf}"'
                + (' quotePrefix="1"' if style.quote_prefix else ""))
        differs = {"number_format": number != base.number_format, "font": font != base.font,
                   "fill": fill != base.fill, "border": border != base.border,
                   "alignment": style.alignment != base.alignment, "protection": style.protection != base.protection}
        for part, name in APPLY_FLAGS.items():
            if part in style.applied or differs[part]:
                head += f' {name}="1"'
        inner = alignment_xml(style.alignment) + protection_xml(style.protection)
        self._added["cellXfs"].append(head + (f">{inner}</xf>" if inner else "/>"))
        self.styles.append(style)
        index = len(self.styles) - 1
        self._index[style] = index
        return index

    def _entry(self, section: str, items: list[_T], item: _T, render: Callable[[_T], str], *, start: int = 0) -> int:
        """The index of the first entry from ``start`` on equal to ``item``, adding it when there is none."""
        for index in range(start, len(items)):
            if items[index] == item:
                return index
        items.append(item)
        self._added[section].append(render(item))
        return len(items) - 1

    def _number_format_id(self, code: str) -> int:
        """The id a format is written with: a built-in one where it has one, else the file's own or a new one.

        Excel spells out the currency and accounting built-ins in numFmts
        as well, the first time a cell uses one.
        """
        for identifier, known in self.number_formats.items():
            if known == code and (identifier < 164 or identifier in self._custom_ids()):
                if identifier in SPELLED_OUT and identifier not in self._declared:
                    self._declare(identifier, code)
                return identifier
        identifier = max([163, *self.number_formats]) + 1
        self.number_formats[identifier] = code
        self._declare(identifier, code)
        return identifier

    def _declare(self, identifier: int, code: str) -> None:
        self._declared.add(identifier)
        self._added["numFmts"].append(f'<numFmt numFmtId="{identifier}" formatCode="{_escape(to_file(code))}"/>')

    def _custom_ids(self) -> set[int]:
        return {one for one in self.number_formats if one >= 164}

    @property
    def dirty(self) -> bool:
        return any(self._added.values()) or self._named_changed

    def written(self) -> str:
        """The stylesheet XML with every added entry in its section."""
        text = self.xml
        for section, entries in self._added.items():
            if entries:
                text = _append(text, section, entries)
        if self._named_changed:
            text = _replaced(text, "cellStyles", self._named)
        return text

    def saved(self) -> None:
        """The file now holds the added entries."""
        self.xml = self.written()
        for entries in self._added.values():
            entries.clear()
        self._named_changed = False


def _append(text: str, section: str, entries: list[str]) -> str:
    match = re.search(rf"<{section}\b([^>]*?)(/?)>", text)
    if match is None:
        # Put the missing section before the first one that follows it.
        later = _ORDER[_ORDER.index(section) + 1:]
        anchor = re.search(r"<(?:" + "|".join(later) + r")\b|</styleSheet>", text)
        position = anchor.start() if anchor else len(text)
        block = f'<{section} count="{len(entries)}">{"".join(entries)}</{section}>'
        return text[:position] + block + text[position:]
    head = match.group(0)
    count = int(_attributes(f"<x{match.group(1)}>").get("count", "0") or 0) + len(entries)
    opened = re.sub(r'count="\d+"', f'count="{count}"', head) if 'count="' in head else head
    if match.group(2):
        opened = opened[:-2].rstrip() + ">"
        if 'count="' not in opened:
            opened = opened[:-1] + f' count="{count}">'
        return text[:match.start()] + opened + "".join(entries) + f"</{section}>" + text[match.end():]
    closing = text.index(f"</{section}>", match.end())
    return text[:match.start()] + opened + text[match.end():closing] + "".join(entries) + text[closing:]


def _replaced(text: str, section: str, entries: list[str]) -> str:
    """The stylesheet with a section's entries replaced by ``entries``, its count following."""
    match = re.search(rf"<{section}\b[^>]*?(?:/>|>.*?</{section}>)", text, re.DOTALL)
    if match is None:
        return _append(text, section, entries)
    block = f'<{section} count="{len(entries)}">{"".join(entries)}</{section}>'
    return text[:match.start()] + block + text[match.end():]
