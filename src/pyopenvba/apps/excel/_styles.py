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
from dataclasses import dataclass, field
from typing import Final, TypeVar

from pyopenvba._xml import attributes as _attributes
from pyopenvba._xml import escape as _escape
from pyopenvba._xml import unescape as _unescape

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


def excel_double(value: float) -> str:
    """A tint as Excel writes one into XML.

    From 0.1 up, 15 significant digits when they round-trip and 17 when they
    do not; below 0.1 always 17, in scientific notation: ``0.249977111117893``,
    ``0.39997558519241921``, ``9.9978637043366805E-2``.
    """
    if value != 0 and abs(value) < 0.1:
        mantissa, exponent = f"{value:.16e}".split("e")
        mantissa = mantissa.rstrip("0").rstrip(".")
        return f"{mantissa}E{int(exponent)}"
    digits = 15 if float(f"{value:.15g}") == value else 17
    return f"{value:.{digits}g}"


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
    #: The cell style it derives from, the xf's xfId.
    base: int = 0
    quote_prefix: bool = False


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
    if fill.pattern == "solid" or foreground:
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


# --- number formats -------------------------------------------------------------------------------

BUILTIN_FORMAT_CODES: Final[dict[int, str]] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    49: "@",
}
#: The built-in formats that show a date or a time.
BUILTIN_DATE_FORMATS: Final = frozenset({14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47})


# --- the stylesheet -----------------------------------------------------------------------------------

_SECTION = {"numFmts": "numFmt", "fonts": "font", "fills": "fill", "borders": "border",
            "cellStyleXfs": "xf", "cellXfs": "xf"}
#: The order a stylesheet's children come in, for adding a section a file lacks.
_ORDER: Final = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
                 "tableStyles", "colors", "extLst")


def _children(xml: str, section: str) -> list[str]:
    match = re.search(rf"<{section}\b[^>]*?(?:/>|>(.*?)</{section}>)", xml, re.DOTALL)
    if match is None or match.group(1) is None:
        return []
    tag = _SECTION[section]
    return re.findall(rf"<{tag}\b[^>]*?/>|<{tag}\b[^>]*?>.*?</{tag}>", match.group(1), re.DOTALL)


@dataclass(slots=True)
class _BaseStyle:
    """A cell style's xf, which decides the apply flags of the xfs that derive from it."""

    number_format: int = 0
    font: int = 0
    fill: int = 0
    border: int = 0
    alignment: Alignment = Alignment()
    protection: Protection = Protection()


class Stylesheet:
    """``xl/styles.xml`` read into formats, and the entries a save has to add to it."""

    def __init__(self, xml: str, theme_xml: str = "") -> None:
        self.xml = xml
        palette = tuple(
            _attributes(one).get("rgb", "FF000000")[-6:].upper()
            for one in re.findall(r"<rgbColor\b[^>]*/>", xml)
        )
        self.colors = Colors(parse_theme(theme_xml) if theme_xml else DEFAULT_THEME,
                             palette if len(palette) >= 64 else ())
        self.number_formats: dict[int, str] = dict(BUILTIN_FORMAT_CODES)
        for element in _children(xml, "numFmts"):
            found = _attributes(element)
            try:
                self.number_formats[int(found.get("numFmtId", "0"))] = _unescape(found.get("formatCode", ""))
            except ValueError:
                continue
        self.fonts = [parse_font(one) for one in _children(xml, "fonts")] or [Font()]
        self.fills = [parse_fill(one) for one in _children(xml, "fills")] or [Fill()]
        self.borders = [parse_border(one) for one in _children(xml, "borders")] or [Border()]
        self.bases = [self._base(one) for one in _children(xml, "cellStyleXfs")] or [_BaseStyle()]
        self.styles = [self._style(one) for one in _children(xml, "cellXfs")] or [Style(font=self.fonts[0])]
        self._index: dict[Style, int] = {}
        for index, style in enumerate(self.styles):
            self._index.setdefault(style, index)
        self._added: dict[str, list[str]] = {name: [] for name in ("numFmts", "fonts", "fills", "borders", "cellXfs")}

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

    def _style(self, element: str) -> Style:
        found = _attributes(element)
        number = int(found.get("numFmtId", "0") or 0)
        return Style(
            number_format=self.number_formats.get(number, "General"),
            font=self._part(self.fonts, int(found.get("fontId", "0") or 0), self.fonts[0]),
            fill=self._part(self.fills, int(found.get("fillId", "0") or 0), Fill()),
            border=self._part(self.borders, int(found.get("borderId", "0") or 0), Border()),
            alignment=parse_alignment(element),
            protection=parse_protection(element),
            base=int(found.get("xfId", "0") or 0),
            quote_prefix=found.get("quotePrefix", "0") in ("1", "true"),
        )

    @property
    def default(self) -> Style:
        return self.styles[0]

    def style(self, index: int) -> Style:
        return self.styles[index] if 0 <= index < len(self.styles) else self.styles[0]

    def is_date_style(self, index: int) -> bool:
        return self.style(index).number_format in {BUILTIN_FORMAT_CODES[one] for one in BUILTIN_DATE_FORMATS}

    # -- adding

    def index_of(self, style: Style) -> int:
        """The xf for ``style``, reusing an equal one or adding what the stylesheet lacks."""
        found = self._index.get(style)
        if found is not None:
            return found
        base = self.bases[style.base] if 0 <= style.base < len(self.bases) else _BaseStyle()
        number = self._number_format_id(style.number_format)
        font = self._entry("fonts", self.fonts, style.font, font_xml)
        fill = self._entry("fills", self.fills, style.fill, fill_xml)
        border = self._entry("borders", self.borders, style.border, border_xml)
        head = (f'<xf numFmtId="{number}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="{style.base}"'
                + (' quotePrefix="1"' if style.quote_prefix else ""))
        for applied, name in ((number != base.number_format, "applyNumberFormat"), (font != base.font, "applyFont"),
                              (fill != base.fill, "applyFill"), (border != base.border, "applyBorder"),
                              (style.alignment != base.alignment, "applyAlignment"),
                              (style.protection != base.protection, "applyProtection")):
            if applied:
                head += f' {name}="1"'
        inner = alignment_xml(style.alignment) + protection_xml(style.protection)
        self._added["cellXfs"].append(head + (f">{inner}</xf>" if inner else "/>"))
        self.styles.append(style)
        index = len(self.styles) - 1
        self._index[style] = index
        return index

    def _entry(self, section: str, items: list[_T], item: _T, render: Callable[[_T], str]) -> int:
        try:
            return items.index(item)
        except ValueError:
            items.append(item)
            self._added[section].append(render(item))
            return len(items) - 1

    def _number_format_id(self, code: str) -> int:
        for identifier, known in self.number_formats.items():
            if known == code and (identifier < 164 or identifier in self._custom_ids()):
                return identifier
        identifier = max([163, *self.number_formats]) + 1
        self.number_formats[identifier] = code
        self._added["numFmts"].append(f'<numFmt numFmtId="{identifier}" formatCode="{_escape(code)}"/>')
        return identifier

    def _custom_ids(self) -> set[int]:
        return {one for one in self.number_formats if one >= 164}

    @property
    def dirty(self) -> bool:
        return any(self._added.values())

    def written(self) -> str:
        """The stylesheet XML with every added entry in its section."""
        text = self.xml
        for section, entries in self._added.items():
            if entries:
                text = _append(text, section, entries)
        return text

    def saved(self) -> None:
        """The file now holds the added entries."""
        self.xml = self.written()
        for entries in self._added.values():
            entries.clear()


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
