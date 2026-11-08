"""Font on a chart's text, as tests/fixtures/chart_fonts.json measured it on the chart and axis titles, tick labels,
legend, data labels and chart area, which all behave alike, and tests/fixtures/chart_font_gaps.json the rest of what they take.

A title reads 18 points or 10, and bold; the rest read 10 and not bold, each in the theme's body font. A save writes
what was set in the paragraph's defRPr: a typed title's run is left alone, and the rest take a txPr. Superscript reads
back and writes nothing. Excel's Subscript writes one of two things at random, the same statement on the same element reading and
saving a baseline of -25000 or 30000 (tests/fixtures/chart_subscript.json); the model writes what was asked, which is always one
of the two. ThemeColor, TintAndShade and ThemeFont answer error 445 either way, and Background takes anything
and stays automatic.

The chart area's font is what the rest read until they are set themselves, a title's size at 1.2 times the area's. Any setting
on it writes only the area's text and takes everything set off every element, whichever setting that was, so what an element is
given after that lasts until the area is set again.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pyopenvba.apps.excel._chart_styled import font_read, is_styled, theme_colour
from pyopenvba.apps.excel._chart_xml import TextLook
from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import VBAInt, error, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Chart

_AUTOMATIC = -4105
_NONE = -4142
_SINGLE = 2
_DOUBLE = -4119
_LOWERED, _RAISED = -25000, 30000
#: Underline and how the file spells it; True is single and False none.
_UNDERLINES = {_NONE: "none", 0: "none", _SINGLE: "sng", _DOUBLE: "dbl", -1: "sng"}
#: FontStyle and what it sets.
_STYLES = {"regular": (False, False), "bold": (True, False), "italic": (False, True), "bold italic": (True, True)}


def _flag(value: object) -> bool:
    """A chart font's True or False: text that is neither is error 1004, as Excel's setters answer here."""
    try:
        return to_bool(value)
    except VBARuntimeError:
        raise error(1004, "Unable to set the property of the Font class") from None


def _number(value: object) -> float:
    try:
        return float(to_number(value))
    except VBARuntimeError:
        raise error(1004, "Unable to set the property of the Font class") from None


def read_look(owner: Any, ns: dict[str, str]) -> TextLook | None:
    """What a file's element says of its text, in the defRPr of its rich title or its txPr; None where it holds neither."""
    run = owner.find("c:tx/c:rich/a:p/a:pPr/a:defRPr", ns) if owner.find("c:tx/c:rich", ns) is not None else owner.find("c:txPr/a:p/a:pPr/a:defRPr", ns)
    if run is None:
        return None
    look = TextLook()
    size = run.get("sz")
    look.size = int(size) / 100 if size and size.isdigit() else None
    for name, field in (("b", "bold"), ("i", "italic")):
        if run.get(name) is not None:
            setattr(look, field, run.get(name) in ("1", "true"))
    look.underline, look.strike = run.get("u"), run.get("strike")
    baseline = run.get("baseline")
    if baseline is not None and baseline.lstrip("-").isdigit():
        look.baseline = int(baseline)
        look.subscript = look.baseline < 0
    solid = run.find("a:solidFill", ns)
    if solid is not None:
        plain, themed = solid.find("a:srgbClr", ns), solid.find("a:schemeClr", ns)
        if plain is not None and plain.find("a:alpha", ns) is not None and plain.find("a:alpha", ns).get("val") == "0":
            look.hidden = True
        elif plain is not None:
            look.rgb = plain.get("val")
        elif themed is not None:
            parts = "".join(f'<a:{child.tag.rpartition("}")[2]} val="{child.get("val")}"/>' for child in themed)
            look.scheme = f'<a:schemeClr val="{themed.get("val")}">{parts}</a:schemeClr>'
    latin = run.find("a:latin", ns)
    if latin is not None and not latin.get("typeface", "").startswith("+"):
        look.name = latin.get("typeface")
    return look


def _rgb(colour: int) -> str:
    return f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


class ChartFont(Created):
    vba_type_name = "Font"

    def __init__(self, element: object, chart: Chart, key: str, size: float, bold: bool,
                 labels: Any = None) -> None:
        self.element, self.chart, self.key, self.size, self.bold = element, chart, key, size, bold
        #: The colour a style from 201 gives the text, which also says the size and whether it is bold.
        self.colour = "000000"
        style = chart.state.chart_style
        if is_styled(style):
            self.size, self.bold, self.colour = font_read(style, key, chart.sheet.book.stylesheet.colors.theme)
        #: A data labels' font is kept with their look; the rest in the chart's fonts.
        self.labels = labels

    def look(self) -> TextLook:
        self.chart.admit_read()
        return self.stored() or TextLook()

    def stored(self) -> TextLook | None:
        return self.chart.state.element_fonts.get(self.key) if self.labels is None else self.labels.font

    def shown(self, field: str) -> Any:
        """A setting as this element reads it: its own, or a legend entry's legend's, or the chart area's."""
        own = getattr(self.look(), field)
        if own is not None or self.key == "area":
            return own
        fonts = self.chart.state.element_fonts
        for key in (("legend", "area") if self.key[:6] == "legend" and self.key[6:].isdigit() else ("area",)):
            found = fonts.get(key)
            if found is not None and getattr(found, field) is not None:
                return getattr(found, field)
        return None

    def change(self, member: str = "") -> TextLook:
        self.look()
        try:
            self.chart.changed()
        except VBAUnsupportedError:
            # On a loaded chart with a style, as a new chart of its style takes it, where that was measured.
            if not (self.chart.state.original and member and self.chart.whole_changed(member)):
                raise
        look = self.stored()
        if look is None:
            look = TextLook()
            if self.labels is None:
                self.chart.state.element_fonts[self.key] = look
            else:
                self.labels.font = look
        return look

    def put(self, **fields: object) -> None:
        """A setting set; on the chart area, everything set of every element goes."""
        look = self.change(f"{self.key}.Font.{','.join(sorted(fields))}" if self.labels is None else "")
        for name, value in fields.items():
            setattr(look, name, value)
        if self.key != "area":
            return
        state = self.chart.state
        state.element_fonts = {"area": look}
        for entry in state.series:
            entry.label_look.font = None

    @member
    def Parent(self) -> object:
        return self.element

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Name(self) -> object:
        name = self.shown("name")
        return self.chart.sheet.book.stylesheet.theme_fonts[1] if name is None else name

    @setter("Name")
    def _name(self, value: object) -> None:
        self.put(name=to_text(value))

    @member
    def Size(self) -> object:
        size = self.shown("size")
        if self.look().size is not None:
            return size
        if size is None:
            return self.size
        return round(size * 1.2, 2) if self.key == "title" else size

    @setter("Size")
    def _size(self, value: object) -> None:
        wanted = _number(value)
        self.look()
        if not 1 <= wanted <= 409:
            raise error(1004, "Unable to set the Size property of the Font class")
        self.put(size=wanted)

    def bolded(self) -> bool:
        bold = self.shown("bold")
        return self.bold if bold is None else bold

    def slanted(self) -> bool:
        return bool(self.shown("italic"))

    @member
    def Bold(self) -> object:
        return self.bolded()

    @setter("Bold")
    def _bold(self, value: object) -> None:
        self.put(bold=_flag(value))

    @member
    def Italic(self) -> object:
        return self.slanted()

    @setter("Italic")
    def _italic(self, value: object) -> None:
        self.put(italic=_flag(value))

    @member
    def FontStyle(self) -> object:
        bold, italic = self.bolded(), self.slanted()
        return {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic",
                (True, True): "Bold Italic"}[(bold, italic)]

    @setter("FontStyle")
    def _font_style(self, value: object) -> None:
        # Any other text is the regular style (tests/fixtures/chart_font_gaps.json).
        self.look()
        bold, italic = _STYLES.get(to_text(value).lower(), (False, False))
        self.put(bold=bold, italic=italic)

    @member
    def Underline(self) -> object:
        return VBAInt({"sng": _SINGLE, "dbl": _DOUBLE}.get(self.shown("underline") or "", _NONE), "Long")

    @setter("Underline")
    def _underline(self, value: object) -> None:
        self.look()
        wanted = _UNDERLINES.get(int(_number(value)))
        if wanted is None:
            raise error(1004, "Unable to set the Underline property of the Font class")
        self.put(underline=wanted)

    @member
    def Strikethrough(self) -> object:
        return self.shown("strike") == "sngStrike"

    @setter("Strikethrough")
    def _strike(self, value: object) -> None:
        self.put(strike="sngStrike" if _flag(value) else "noStrike")

    @member
    def Color(self) -> object:
        if self.shown("hidden"):
            return VBAInt(-2, "Long")
        rgb = self.shown("rgb") or self.from_theme(self.shown("scheme"))
        return VBAInt(bgr(self.colour if rgb is None else rgb), "Long")

    def from_theme(self, spelled: str | None) -> str | None:
        """A theme colour a file gave the text, as the colour it comes to."""
        return None if spelled is None else theme_colour(spelled, self.chart.sheet.book.stylesheet.colors.theme)

    @setter("Color")
    def _color(self, value: object) -> None:
        try:
            wanted = int(to_integer(value, "Long"))
        except VBARuntimeError:
            raise error(-2147467259, "Automation error\nUnspecified error") from None
        if wanted == -1:
            self.put(rgb=None, hidden=None)
            return
        self.put(rgb=_rgb(wanted if 0 <= wanted <= 0xFFFFFF else 0), hidden=None)

    @member
    def ColorIndex(self) -> object:
        # An element reads automatic while its colour is the chart area's, though a legend entry reads its legend's.
        look = self.look()
        if self.key[:6] == "legend" and self.key[6:].isdigit() and look.rgb is None and look.scheme is None:
            look = self.chart.state.element_fonts.get("legend") or look
        if look.hidden:
            return VBAInt(_NONE, "Long")
        rgb = look.rgb or self.from_theme(look.scheme)
        if rgb is None:
            return VBAInt(_AUTOMATIC, "Long")
        return VBAInt(self.chart.sheet.book.stylesheet.colors.nearest(rgb), "Long")

    @setter("ColorIndex")
    def _color_index(self, value: object) -> None:
        wanted = int(_number(value))
        self.look()
        if wanted in (0, _AUTOMATIC):
            self.put(rgb=None, hidden=None)
        elif wanted in (-1, _NONE):
            self.put(rgb=None, hidden=True)
        elif 1 <= wanted <= 57:
            # One past the palette is its first (tests/fixtures/chart_font_gaps.json).
            self.put(rgb=self.chart.sheet.book.stylesheet.colors.indexed[7 + (wanted if wanted < 57 else 1)], hidden=None)
        else:
            raise VBAUnsupportedError("This chart font ColorIndex is not implemented")

    @member
    def Superscript(self) -> object:
        return bool(self.look().superscript) or self.shown("baseline") == _RAISED

    @setter("Superscript")
    def _superscript(self, value: object) -> None:
        wanted = _flag(value)
        if wanted and self.look().baseline == _LOWERED:
            # Over a subscript this does write, and the subscript goes (tests/fixtures/chart_font_gaps.json).
            self.put(baseline=_RAISED, subscript=False, superscript=True)
            return
        self.put(superscript=wanted)

    @member
    def Subscript(self) -> object:
        return bool(self.look().subscript)

    @setter("Subscript")
    def _subscript(self, value: object) -> None:
        wanted = _flag(value)
        self.put(baseline=_LOWERED if wanted else 0, subscript=wanted)

    @member
    def Shadow(self) -> object:
        self.look()
        return False

    @member
    def OutlineFont(self) -> object:
        self.look()
        return False

    @member
    def Background(self) -> object:
        self.look()
        return VBAInt(_AUTOMATIC, "Long")

    @setter("Background")
    def _background(self, value: object) -> None:
        """It takes anything, and stays automatic."""
        self.look()

    def no_theme(self) -> object:
        self.look()
        raise error(445, "Object doesn't support this action")

    @member
    def ThemeColor(self) -> object:
        return self.no_theme()

    @setter("ThemeColor")
    def _theme_color(self, value: object) -> None:
        self.no_theme()

    @member
    def TintAndShade(self) -> object:
        return self.no_theme()

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        self.no_theme()

    @member
    def ThemeFont(self) -> object:
        return self.no_theme()

    @setter("ThemeFont")
    def _theme_font(self, value: object) -> None:
        to_integer(value, "Long")
        self.no_theme()
