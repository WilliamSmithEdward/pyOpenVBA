"""A workbook's cell formats: its stylesheet, its colours, and the format each cell carries.

``xl/styles.xml`` holds a cell's format as an index into ``cellXfs``, and
each xf points at a font, a fill, a border and a number format. Here that
becomes one frozen :class:`Style` per xf, made of frozen parts, so a cell
holds its format as a value and two cells with the same format hold equal
values. :class:`Stylesheet` keeps them as Excel's tables do, in the order
they came, and a save that follows a change writes what is in use in that
order, as Excel writes it (Stylesheet.prepare); the file's own parts keep
the XML they were read from.

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
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field, replace
from fractions import Fraction
from typing import Final, TypeVar

from pyopenvba._xml import attributes as _attributes
from pyopenvba._xml import escape as _escape
from pyopenvba._xml import unescape as _unescape
from pyopenvba.apps.excel._number_format import BUILTIN, SPELLED_OUT, from_file, to_file

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
    #: The parts of the format it holds of its own, the xf's apply flags:
    #: each part a macro set, even to what it already was, and each part a
    #: style it was given leaves alone and holds otherwise. Every other part
    #: follows its cell style: a change to the style reaches it while the
    #: style includes the part (tests/fixtures/cell_styles/
    #: style_changes.json). Excel keeps a format a macro made bold and then
    #: not bold apart from the default one; a file's flags are worked out
    #: afresh from what differs from its style, as Excel works them out on
    #: opening it, so two that differ only in them are one.
    applied: frozenset[str] = frozenset()


#: The parts of a format, in the order an xf's apply flags are written.
APPLY_FLAGS: Final = {"number_format": "applyNumberFormat", "font": "applyFont", "fill": "applyFill",
                      "border": "applyBorder", "alignment": "applyAlignment", "protection": "applyProtection"}
PARTS: Final = frozenset(APPLY_FLAGS)


def applying(style: Style, part: str, **changes: object) -> Style:
    """The format with one part set by a macro: changed as asked, and applied from now on."""
    return replace(style, applied=style.applied | {part}, **changes)  # type: ignore[arg-type]


def _taken(style: Style, entry: CellStyle, part: str) -> bool:
    """Whether a format takes a part from its cell style as the style has it, pointing at the style's own entry:
    Excel keeps a copy of the Normal font for the styles that want it, and a cell given one of them points at
    the copy."""
    return part not in style.applied and getattr(style, part) == getattr(entry.format, part)


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


@dataclass(frozen=True, slots=True)
class Latent:
    """The fonts, fills and borders of Excel's standard built-in styles, in its own order.

    Excel holds them in its tables from the start, ahead of anything a
    macro makes, so a new part equal to one of them -- a red font, which
    Warning Text has -- comes out in that place (tests/fixtures/cell_styles/
    order.xlsx). The fonts start with the theme's body font, which a style
    that wants it as it is takes ahead of the others.
    """

    fonts: tuple[Font, ...] = ()
    fills: tuple[Fill, ...] = ()
    borders: tuple[Border, ...] = ()


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
    #: A built-in style a macro changed, which the file marks customBuiltin and writes whether used or not.
    custom: bool = False
    #: Deleted: a built-in one stays in the file, hidden, a user's is gone; neither is in Styles any more.
    deleted: bool = False


class Stylesheet:
    """``xl/styles.xml`` read into formats and cell styles, held as Excel's tables hold them, and written out
    at each save as Excel writes them.

    Excel keeps a workbook's number formats, fonts, fills, borders, cell
    styles and cell xfs as tables in the order they came: the file's own,
    then the built-in styles' parts, then what the session made. On
    opening a file it numbers the file's custom formats from 164 in the
    order the file lists them, and takes an xf that points at what an
    earlier one points at, once its flags are worked out afresh, as that
    one. A save writes what is in use and nothing else, in table order,
    each index moved up over what it left out (tests/fixtures/cell_styles/:
    planted, duplicates and renumbered, each file and what Excel saved of
    it). An entry of the file's that nothing changed keeps the XML the
    file gave it; a cell xf keeps the parts it pointed at, so a font the
    file holds twice stays the one it was.
    """

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
        #: The numFmt elements the file spells out, by the id Excel gives each on opening.
        self._format_xml: dict[int, str] = {}
        #: The ids the file gives its custom formats, against the ones Excel numbers them with from 164.
        self._renumbered: dict[int, int] = {}
        for element in _children(xml, "numFmts"):
            found = _attributes(element)
            try:
                identifier = int(found.get("numFmtId", "0"))
            except ValueError:
                continue
            if identifier >= 164:
                self._renumbered[identifier] = identifier = 164 + len(self._renumbered)
            self.number_formats[identifier] = from_file(_unescape(found.get("formatCode", "")))
            self._format_xml[identifier] = element
        #: Each format's id, the first one that spells it: a built-in one ahead of the file's own.
        self._ids: dict[str, int] = {}
        for identifier, code in self.number_formats.items():
            self._ids.setdefault(code, identifier)
        fonts, fills, borders = (_children(xml, section) for section in ("fonts", "fills", "borders"))
        self.fonts = [parse_font(one) for one in fonts] or [Font()]
        self.fills = [parse_fill(one) for one in fills] or [Fill()]
        self.borders = [parse_border(one) for one in borders] or [Border()]
        #: The XML the file wrote each of its fonts, fills and borders in, which a save writes them in again.
        self._part_xml: dict[str, list[str]] = {"fonts": fonts, "fills": fills, "borders": borders}
        style_xfs = _children(xml, "cellStyleXfs")
        #: The parts of each cell style the file holds, by index, by the style's place in cell_styles.
        self._style_parts: dict[int, _BaseStyle] = {index: self._base(one) for index, one in enumerate(style_xfs)}
        if not style_xfs:
            self._style_parts[0] = _BaseStyle()
        #: The cell styles the file holds, by place, in the order it holds them.
        self._held: list[int] = list(self._style_parts)
        #: Every cell style: the file's in cellStyleXfs order, then those the session added.
        self.cell_styles = [self._cell_style(one, index) for index, one in enumerate(style_xfs)] or [
            CellStyle(Style(font=self.fonts[0]), name="Normal", builtin=0, xf=0)]
        #: The file's cellStyles elements in its order, each with the place of the style it names.
        self._named: list[tuple[int, str]] = []
        for element in _children(xml, "cellStyles"):
            found = _attributes(element)
            try:
                place = int(found.get("xfId", ""))
            except ValueError:
                continue
            if 0 <= place < len(self.cell_styles) and not self.cell_styles[place].name:
                entry = self.cell_styles[place]
                entry.name = _unescape(found.get("name", ""))
                builtin = found.get("builtinId", "")
                entry.builtin = int(builtin) if builtin.isdigit() else None
                entry.custom = found.get("customBuiltin", "0") in ("1", "true")
                entry.deleted = found.get("hidden", "0") in ("1", "true") and entry.builtin is not None
                self._named.append((place, element))
        self._made = 0
        #: The Normal font as the workbook opened, which the built-in styles that want it keep a copy of.
        normal = self.named("Normal")
        self.opening_font = self.cell_styles[0 if normal is None else normal].format.font
        cell_xfs = _children(xml, "cellXfs")
        #: Every cell xf as Excel's table holds them: the file's, then those the session made, in that order.
        self.styles = [self._style(one) for one in cell_xfs] or [Style(font=self.fonts[0])]
        #: The parts each of the file's cell xfs points at, by index: number format, font, fill and border.
        self._xf_parts: list[tuple[int, int, int, int]] = [self._parts(one) for one in cell_xfs] or [(0, 0, 0, 0)]
        self._index: dict[Style, int] = {}
        for index, style in enumerate(self.styles):
            self._index.setdefault(style, index)
        #: An xf of the file's that points at what an earlier one points at, flags aside: Excel takes it as that
        #: one on opening.
        self._alias: dict[int, int] = {}
        seen: dict[tuple[Style, tuple[int, int, int, int]], int] = {}
        for index, style in enumerate(self.styles):
            first = seen.setdefault((style, self._xf_parts[index]), index)
            if first != index:
                self._alias[index] = first
        #: The file's xfs and cell styles a change made in place, which a save writes by what they now say.
        self._changed_xfs: set[int] = set()
        self._changed_styles: set[int] = set()
        #: The fonts, fills and borders made while the workbook was open, keyed by kind and value, in the order
        #: they were made, which is the order a save writes those still in use in (see meet).
        self._born: dict[tuple[str, object], int] = {}
        #: What the save being made writes: each xf of the table's place in cellXfs, the sections, and the
        #: tables as they will stand once it is written.
        self._written_at: dict[int, int] = {}
        self._sections: dict[str, list[str]] = {}
        self._after: _Written | None = None
        #: Whether the session changed a format, a style or a part: a save leaves a stylesheet it did not touch
        #: as the file had it, as it leaves a sheet no macro wrote to.
        self._touched = False

    # -- reading

    def _part(self, items: list[_T], index: int, fallback: _T) -> _T:
        return items[index] if 0 <= index < len(items) else fallback

    def _number_id(self, text: str) -> int:
        """A number format's id as the file gives it, as Excel numbers it on opening."""
        identifier = int(text or 0)
        return self._renumbered.get(identifier, identifier)

    def _base(self, element: str) -> _BaseStyle:
        found = _attributes(element)
        return _BaseStyle(
            number_format=self._number_id(found.get("numFmtId", "0")),
            font=int(found.get("fontId", "0") or 0),
            fill=int(found.get("fillId", "0") or 0),
            border=int(found.get("borderId", "0") or 0),
            alignment=parse_alignment(element),
            protection=parse_protection(element),
        )

    def _parts(self, element: str) -> tuple[int, int, int, int]:
        found = _attributes(element)
        font, fill, border = (int(found.get(key, "0") or 0) for key in ("fontId", "fillId", "borderId"))
        return self._number_id(found.get("numFmtId", "0")), font, fill, border

    def _cell_style(self, element: str, index: int) -> CellStyle:
        """A cellStyleXfs xf: the format it gives, and the parts it gives, those its apply flags do not turn off."""
        base = self._style_parts[index]
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
        )
        return CellStyle(own, includes, xf=index)

    def _style(self, element: str) -> Style:
        found = _attributes(element)
        number, font, fill, border = self._parts(element)
        base_index = int(found.get("xfId", "0") or 0)
        alignment, protection = parse_alignment(element), parse_protection(element)
        base = self._style_parts.get(base_index, _BaseStyle())
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
            if entry.name and not entry.deleted and entry.name.casefold() == key:
                return index
        return None

    def add_cell_style(self, entry: CellStyle) -> int:
        """Hold a cell style the file does not have yet, which a save writes when it has to; its place."""
        self._made += 1
        entry.made = self._made
        self.cell_styles.append(entry)
        index = len(self.cell_styles) - 1
        entry.format = replace(entry.format, base=index)
        return index

    def _normal(self) -> int:
        found = self.named("Normal")
        return 0 if found is None else found

    # -- what a macro makes

    def meet(self, style: Style | None) -> int:
        """Note a format as a macro makes it, and answer its place in the table of cell xfs; -1 for none.

        Excel keeps every font, fill, border and cell xf in the order the
        session made it, after the file's own and the built-in styles'
        parts, and a save writes those still in use in that order
        (tests/fixtures/cell_styles/order.xlsx). A format equal to one the
        table holds is that one.
        """
        self._touched = True
        return -1 if style is None else self._entry_for(style)

    def _entry_for(self, style: Style) -> int:
        found = self._index.get(style)
        if found is not None:
            return found
        self._touched = True
        self.meet_parts(style)
        self.styles.append(style)
        index = self._index[style] = len(self.styles) - 1
        return index

    def touch(self) -> None:
        """Note that a macro added, merged or deleted a cell style, which the next save writes."""
        self._touched = True

    def meet_parts(self, style: Style) -> None:
        """Note the parts of a format made now: its font, fill and border, and its number format, which takes an
        id at once that stays taken though nothing uses it later."""
        born = self._born
        for key in (("font", style.font), ("fill", style.fill), ("border", style.border)):
            born.setdefault(key, len(born))
        self._code_id(style.number_format)

    def entry_of(self, style: Style, xf: int) -> int:
        """The table's xf a format is: the one it points at while that says the same, else an equal one."""
        if 0 <= xf < len(self.styles) and self.styles[xf] == style:
            return self._alias.get(xf, xf) if self._touched else xf
        return self._entry_for(style)

    def restyle_xfs(self, change: Callable[[Style], Style | None]) -> None:
        """Change the table's xfs in place, as Excel changes the xfs of a style that changes: ``change`` answers
        an xf's new format, or None to leave it. Two that come to say the same stay apart."""
        changed = False
        for index, style in enumerate(self.styles):
            new = change(style)
            if new is None or new == style:
                continue
            self.styles[index] = new
            self.meet_parts(new)
            self._changed_xfs.add(index)
            changed = self._touched = True
        if changed:
            self._index = {}
            for index, style in enumerate(self.styles):
                self._index.setdefault(style, index)

    def restyle_cell_style(self, entry: CellStyle) -> None:
        """A cell style changed in place: a save writes it by what it now says."""
        self._touched = True
        self.meet_parts(entry.format)
        self._changed_styles.add(next(index for index, one in enumerate(self.cell_styles) if one is entry))

    def renormal_font(self, font: Font) -> None:
        """The Normal font changed: the file's first font, which it is, changes in place."""
        self._touched = True
        self.fonts[0] = font
        if self._part_xml["fonts"]:
            self._part_xml["fonts"][0] = font_xml(font)

    # -- a save

    def prepare(self, holders: Iterable[tuple[Style, int]], latent: Latent) -> None:
        """Work out what a save writes, before the sheets are written with the places it gives their xfs.

        The cell xfs are the default and every one a cell, row or column
        points at, in table order. The cell styles are the file's and those
        the session added, the built-in ones while an xf derives from them,
        a macro changed them or deleted them (tests/fixtures/cell_styles/
        delete.xlsx): the file's in its order, then the standard built-in
        ones in Excel's own order, then the rest in the order they were
        made. A part they want that the table lacks goes where Excel's
        table holds it: a built-in style's in that style's place, else in
        the order the session made it. Each part a save writes keeps its
        order, and every index moves up over what it leaves out.
        """
        if not self._touched:
            # Nothing changed: the file's stylesheet stays as it is, every xf where it was.
            self._written_at, self._sections, self._after = {index: index for index in range(len(self.styles))}, {}, None
            return
        normal = self._normal()
        used = {0}
        for style, xf in holders:
            used.add(self.entry_of(style, xf))
        xfs = sorted(used)
        bases = {self.styles[index].base for index in xfs}
        held = list(self._held)
        waiting = sorted((place for place in range(len(self.cell_styles)) if place not in self._style_parts),
                         key=lambda place: (self.cell_styles[place].rank < 0, self.cell_styles[place].rank,
                                            self.cell_styles[place].made))
        places = [place for place in [*held, *waiting]
                  if place == normal or _written(self.cell_styles[place], place in bases)]
        self._add_wanted(places, xfs, normal, latent)
        style_parts = {place: self._style_parts_now(place, normal) for place in places}
        xf_parts = {index: self._xf_parts_now(index, style_parts) for index in xfs}
        fonts = sorted({0, *(parts.font for parts in style_parts.values()), *(parts[1] for parts in xf_parts.values())})
        fills = sorted({0, *((1,) if len(self.fills) > 1 else ()), *(parts.fill for parts in style_parts.values()),
                        *(parts[2] for parts in xf_parts.values())})
        borders = sorted({0, *(parts.border for parts in style_parts.values()),
                          *(parts[3] for parts in xf_parts.values())})
        numbers = {*(parts.number_format for parts in style_parts.values()), *(parts[0] for parts in xf_parts.values())}
        font_at, fill_at, border_at = ({index: place for place, index in enumerate(kept)}
                                       for kept in (fonts, fills, borders))
        style_at = {place: position for position, place in enumerate(places)}
        self._written_at = {index: position for position, index in enumerate(xfs)}
        for alias, first in self._alias.items():
            if first in self._written_at:
                self._written_at[alias] = self._written_at[first]
        formats = self._formats_written(numbers)
        named = self._cell_style_elements(places, style_at)
        self._sections = {
            "numFmts": [self._format_element(identifier) for identifier in formats],
            "fonts": [self._part_element("fonts", index, font_xml) for index in fonts],
            "fills": [self._part_element("fills", index, fill_xml) for index in fills],
            "borders": [self._part_element("borders", index, border_xml) for index in borders],
            "cellStyleXfs": [_style_xf_element(self.cell_styles[place], style_parts[place], font_at, fill_at,
                                                border_at) for place in places],
            "cellXfs": [_xf_element(self.styles[index], xf_parts[index], style_at[self.styles[index].base],
                                     font_at, fill_at, border_at) for index in xfs],
            "cellStyles": [element for _, element in named],
        }
        self._after = _Written(places, {place: _moved(style_parts[place], font_at, fill_at, border_at)
                                        for place in places},
                               xfs, [(parts[0], font_at[parts[1]], fill_at[parts[2]], border_at[parts[3]])
                                     for parts in (xf_parts[index] for index in xfs)],
                               fonts, fills, borders, formats, named)

    def _add_wanted(self, places: list[int], xfs: list[int], normal: int, latent: Latent) -> None:
        """Give the tables the parts the written styles and xfs want by what they say, in Excel's table order."""
        wanted: dict[str, list[tuple[object, int]]] = {"fonts": [], "fills": [], "borders": []}
        codes: list[str] = []
        for place in places:
            own = self.cell_styles[place].format
            if self._keeps_style(place):
                if place != normal and self._style_parts[place].font == 0:
                    wanted["fonts"].append((own.font, 1))
                continue
            wanted["fonts"].append((own.font, 0 if place == normal else 1))
            wanted["fills"].append((own.fill, 0))
            wanted["borders"].append((own.border, 0))
            codes.append(own.number_format)
        for index in xfs:
            if self._keeps_xf(index):
                continue
            style = self.styles[index]
            base = self.cell_styles[style.base] if 0 <= style.base < len(self.cell_styles) else self.cell_styles[0]
            for part, section in (("font", "fonts"), ("fill", "fills"), ("border", "borders")):
                if not _taken(style, base, part):
                    wanted[section].append((getattr(style, part), self._font_start(style) if part == "font" else 0))
            if not _taken(style, base, "number_format"):
                codes.append(style.number_format)
        self._add_in_order("font", self.fonts, wanted["fonts"], latent.fonts)
        self._add_in_order("fill", self.fills, wanted["fills"], latent.fills)
        self._add_in_order("border", self.borders, wanted["borders"], latent.borders)
        for code in codes:
            self._code_id(code)

    def _keeps_style(self, place: int) -> bool:
        """Whether a cell style is written with the parts the file gave it: one the file holds and nothing
        changed."""
        return place in self._style_parts and place not in self._changed_styles

    def _keeps_xf(self, index: int) -> bool:
        return index < len(self._xf_parts) and index not in self._changed_xfs

    def _style_parts_now(self, place: int, normal: int) -> _BaseStyle:
        """A written cell style's parts by the table's indexes. A style's font is never the file's first:
        where the Normal font is what it wants, Excel gives it a copy (tests/fixtures/cell_styles/
        planted_saved.xlsx)."""
        own = self.cell_styles[place].format
        if self._keeps_style(place):
            parts = self._style_parts[place]
            if place != normal and parts.font == 0:
                parts = replace(parts, font=_find(self.fonts, own.font, 1))
            return parts
        return _BaseStyle(self._code_id(own.number_format), _find(self.fonts, own.font, 0 if place == normal else 1),
                          _find(self.fills, own.fill, 0), _find(self.borders, own.border, 0), own.alignment,
                          own.protection)

    def _xf_parts_now(self, index: int, style_parts: dict[int, _BaseStyle]) -> tuple[int, int, int, int]:
        """A written cell xf's parts by the table's indexes: the file's own where nothing changed it, else a
        part it takes from its style as the style has it at the style's entry, and the rest by what it is."""
        if self._keeps_xf(index):
            return self._xf_parts[index]
        style = self.styles[index]
        entry = self.cell_styles[style.base] if 0 <= style.base < len(self.cell_styles) else self.cell_styles[0]
        base = style_parts[style.base]
        return (base.number_format if _taken(style, entry, "number_format") else self._code_id(style.number_format),
                base.font if _taken(style, entry, "font") else _find(self.fonts, style.font, self._font_start(style)),
                base.fill if _taken(style, entry, "fill") else _find(self.fills, style.fill, 0),
                base.border if _taken(style, entry, "border") else _find(self.borders, style.border, 0))

    def _font_start(self, style: Style) -> int:
        """Where a format's font is looked up from: a font it took from a style other than Normal is one of that
        style's entries, never the file's first, even once the style's font changed and left it behind
        (tests/fixtures/cell_styles/reach.xlsx)."""
        return 1 if "font" not in style.applied and style.base != self._normal() else 0

    def _add_in_order(self, kind: str, items: list[_T], wanted: list[tuple[object, int]],
                      latent: tuple[object, ...]) -> None:
        """Add the parts a save needs that ``items`` lacks from each one's start on, in Excel's table order."""
        missing: list[object] = []
        for value, start in wanted:
            if value not in missing and not any(items[index] == value for index in range(start, len(items))):
                missing.append(value)

        def place(value: object) -> tuple[int, int]:
            if value in latent:
                return 0, latent.index(value)
            born = self._born.get((kind, value))
            return (1, born) if born is not None else (2, missing.index(value))

        for value in sorted(missing, key=place):
            item: _T = value  # type: ignore[assignment]
            items.append(item)

    def _formats_written(self, numbers: set[int]) -> list[int]:
        """The number formats a save spells out, in Excel's order: the file's own it still uses, as the file
        listed them; then the built-in ones Excel spells, in its own table's order; then the custom ones by id,
        the order they were made in (tests/fixtures/cell_styles/spelled.xlsx)."""
        spelled = [identifier for identifier in self._format_xml if identifier in numbers
                   and (identifier in SPELLED_OUT or identifier >= 164)]
        spelled += [identifier for identifier in _SPELLED_ORDER if identifier in numbers and identifier not in spelled]
        return spelled + sorted(identifier for identifier in numbers if identifier >= 164 and identifier not in spelled)

    def _format_element(self, identifier: int) -> str:
        element = self._format_xml.get(identifier)
        if element is None:
            return f'<numFmt numFmtId="{identifier}" formatCode="{_escape(to_file(self.number_formats[identifier]))}"/>'
        return re.sub(r'\bnumFmtId="[^"]*"', f'numFmtId="{identifier}"', element, count=1)

    def _part_element(self, section: str, index: int, render: Callable[[_T], str]) -> str:
        xml = self._part_xml[section]
        if index < len(xml):
            return xml[index]
        items: list[_T] = getattr(self, section)
        return render(items[index])

    def _cell_style_elements(self, places: list[int], style_at: dict[int, int]) -> list[tuple[int, str]]:
        """The cellStyles elements a save writes: the file's in its order, each naming its style's new place, and
        a new one put in by name, which is the order the Styles collection lists them in."""
        from pyopenvba.apps.excel._sort import text_key

        elements: list[tuple[int, str]] = []
        named: set[int] = set()
        for place, element in self._named:
            entry = self.cell_styles[place]
            named.add(place)
            if place not in style_at:
                continue
            if entry.custom or entry.deleted:
                elements.append((place, _cell_style_element(entry, style_at[place])))
            else:
                elements.append((place, re.sub(r'\bxfId="[^"]*"', f'xfId="{style_at[place]}"', element, count=1)))
        for place in places:
            entry = self.cell_styles[place]
            if place in named or not entry.name:
                continue
            key = text_key(entry.name, False)
            position = next((at for at, (other, _) in enumerate(elements)
                             if text_key(self.cell_styles[other].name, False) > key), len(elements))
            elements.insert(position, (place, _cell_style_element(entry, style_at[place])))
        return elements

    def _code_id(self, code: str) -> int:
        """A format's id, a custom one taking the next when first made (tests/fixtures/cell_styles/order.xlsx)."""
        identifier = self._ids.get(code)
        if identifier is None:
            identifier = max([163, *self.number_formats]) + 1
            self.number_formats[identifier] = code
            self._ids[code] = identifier
        return identifier

    def index_as_read(self, style: Style, xf: int) -> int:
        """The place in cellXfs a save writes a format's xf at: the table's one it points at while that says the
        same, else an equal one."""
        return self._written_at[self.entry_of(style, xf)]

    def moved_from_file(self) -> dict[int, int]:
        """Where the save puts each of the file's own cell xfs that it moves, which a sheet no macro changed
        points at by the file's numbers."""
        return {index: self._written_at[index] for index in range(len(self._xf_parts))
                if index in self._written_at and self._written_at[index] != index}

    @property
    def dirty(self) -> bool:
        return self.written() != self.xml

    def written(self) -> str:
        """The stylesheet XML as the save writes it: each section the tables fill, in Excel's order."""
        text = self.xml
        for section, entries in self._sections.items():
            text = _with_section(text, section, entries)
        return text

    def saved(self) -> dict[int, int]:
        """The file now holds what the save wrote: the tables become that, and each xf of the table that it
        wrote answers where it went, for the cells, rows and columns to point there."""
        after = self._after
        moved = dict(self._written_at)
        if after is None:
            return moved
        self.xml = self.written()
        for section, kept in (("fonts", after.fonts), ("fills", after.fills), ("borders", after.borders)):
            items: list[object] = getattr(self, section)
            setattr(self, section, [items[index] for index in kept])
            self._part_xml[section] = list(self._sections[section])
        self._format_xml = {identifier: self._format_element(identifier) for identifier in after.formats}
        self._renumbered = {}
        self._style_parts = after.styles
        self._held = list(after.places)
        written_places = {place: position for position, place in enumerate(after.places)}
        for place, entry in enumerate(self.cell_styles):
            entry.xf = written_places.get(place)
        self._named = after.named
        self.styles = [self.styles[index] for index in after.xfs]
        self._xf_parts = after.xf_parts
        self._index = {}
        for index, style in enumerate(self.styles):
            self._index.setdefault(style, index)
        self._alias = {}
        self._changed_xfs.clear()
        self._changed_styles.clear()
        self._written_at = {index: index for index in range(len(self.styles))}
        self._after = None
        self._touched = False
        return moved


@dataclass(slots=True)
class _Written:
    """The tables as a save leaves them: the cell styles and cell xfs it wrote, with their parts by the indexes
    it gave them, and the parts and number formats it kept."""

    places: list[int]
    styles: dict[int, _BaseStyle]
    xfs: list[int]
    xf_parts: list[tuple[int, int, int, int]]
    fonts: list[int]
    fills: list[int]
    borders: list[int]
    formats: list[int]
    named: list[tuple[int, str]]


def _find(items: list[_T], item: _T, start: int) -> int:
    """The first index from ``start`` on of an entry equal to ``item``, which the table has."""
    return next(index for index in range(start, len(items)) if items[index] == item)


def _moved(parts: _BaseStyle, fonts: dict[int, int], fills: dict[int, int], borders: dict[int, int]) -> _BaseStyle:
    return replace(parts, font=fonts[parts.font], fill=fills[parts.fill], border=borders[parts.border])


def _style_xf_element(entry: CellStyle, parts: _BaseStyle, fonts: dict[int, int], fills: dict[int, int],
                      borders: dict[int, int]) -> str:
    """A cell style's xf as a save writes it: its parts, and each part it leaves out turned off."""
    head = (f'<xf numFmtId="{parts.number_format}" fontId="{fonts[parts.font]}" fillId="{fills[parts.fill]}" '
            f'borderId="{borders[parts.border]}"')
    for part, name in APPLY_FLAGS.items():
        if part not in entry.includes:
            head += f' {name}="0"'
    inner = alignment_xml(parts.alignment) + protection_xml(parts.protection)
    return head + (f">{inner}</xf>" if inner else "/>")


def _xf_element(style: Style, parts: tuple[int, int, int, int], xf_id: int, fonts: dict[int, int],
                fills: dict[int, int], borders: dict[int, int]) -> str:
    """A cell xf as a save writes it. Its apply flags are the parts it holds of its own, as Excel noted them when
    they were set (Style.applied), not what differs from its style now."""
    number, font, fill, border = parts
    head = (f'<xf numFmtId="{number}" fontId="{fonts[font]}" fillId="{fills[fill]}" borderId="{borders[border]}" '
            f'xfId="{xf_id}"' + (' quotePrefix="1"' if style.quote_prefix else ""))
    for part, name in APPLY_FLAGS.items():
        if part in style.applied:
            head += f' {name}="1"'
    inner = alignment_xml(style.alignment) + protection_xml(style.protection)
    return head + (f">{inner}</xf>" if inner else "/>")


def _written(entry: CellStyle, used: bool) -> bool:
    """Whether a save writes a cell style: a user's while it lasts; a built-in one while an xf derives from it,
    or once a macro changed or deleted it (tests/fixtures/cell_styles/delete.xlsx)."""
    if entry.builtin is None:
        return not entry.deleted
    return used or entry.custom or entry.deleted


def _cell_style_element(entry: CellStyle, place: int) -> str:
    """A cell style's cellStyles element: a deleted built-in one hidden, a changed one customBuiltin."""
    element = f'<cellStyle name="{_escape(entry.name)}" xfId="{place}"'
    if entry.builtin is not None:
        element += f' builtinId="{entry.builtin}"'
        if entry.deleted:
            element += ' hidden="1"'
        if entry.custom:
            element += ' customBuiltin="1"'
    return element + "/>"


def _with_section(text: str, section: str, entries: list[str]) -> str:
    """The stylesheet with a section holding ``entries``, its count following; one with none left out, and
    one it lacks put in before the first section that follows it."""
    match = re.search(rf"<{section}\b([^>]*?)(?:/>|>.*?</{section}>)", text, re.DOTALL)
    if match is not None:
        if not entries:
            return text[:match.start()] + text[match.end():]
        opened = match.group(0)[:match.group(0).index(">") + 1]
        opened = opened[:-2].rstrip() + ">" if opened.endswith("/>") else opened
        opened = re.sub(r'\bcount="\d+"', f'count="{len(entries)}"', opened) if 'count="' in opened \
            else opened[:-1] + f' count="{len(entries)}">'
        return text[:match.start()] + opened + "".join(entries) + f"</{section}>" + text[match.end():]
    if not entries:
        return text
    later = _ORDER[_ORDER.index(section) + 1:]
    anchor = re.search(r"<(?:" + "|".join(later) + r")\b|</styleSheet>", text)
    position = anchor.start() if anchor else len(text)
    return text[:position] + f'<{section} count="{len(entries)}">{"".join(entries)}</{section}>' + text[position:]
