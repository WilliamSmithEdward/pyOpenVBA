"""Font on a chart's text, as tests/fixtures/chart_fonts.json measured it on the chart and axis titles, tick labels,
legend, data labels and chart area, which all behave alike.

A title reads 18 points or 10, and bold; the rest read 10 and not bold, each in the theme's body font. A save writes
what was set in the paragraph's defRPr: a typed title's run is left alone, and the rest take a txPr. Superscript reads
back and writes nothing, so setting it is refused, as are the settings no case measured. Whether the chart area's font
reaches the other elements was not measured, so reading or setting one beside the other is refused."""
from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pyopenvba.apps.excel._chart_xml import TextLook
from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import VBAInt, error, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Chart

_AUTOMATIC = -4105
_NONE = -4142
_SINGLE = 2
#: FontStyle and what it sets.
_STYLES = {"regular": (False, False), "bold": (True, False), "italic": (False, True), "bold italic": (True, True)}


def _rgb(colour: int) -> str:
    return f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


class ChartFont(Created):
    vba_type_name = "Font"

    def __init__(self, element: object, chart: Chart, key: str, size: float, bold: bool,
                 labels: Any = None) -> None:
        self.element, self.chart, self.key, self.size, self.bold = element, chart, key, size, bold
        #: A data labels' font is kept with their look; the rest in the chart's fonts.
        self.labels = labels

    def look(self) -> TextLook:
        self.chart.admit_read()
        state = self.chart.state
        if state.original:
            raise VBAUnsupportedError("The font of a part of an imported chart is not implemented")
        fonts = state.element_fonts
        if "area" in fonts and self.key != "area":
            raise VBAUnsupportedError("The font of a chart element beside a chart area's font is not implemented")
        return self.stored() or TextLook()

    def stored(self) -> TextLook | None:
        return self.chart.state.element_fonts.get(self.key) if self.labels is None else self.labels.font

    def change(self) -> TextLook:
        self.look()
        state = self.chart.state
        if self.key == "area" and (any(key != "area" for key in state.element_fonts)
                                   or any(one.label_look.font is not None for one in state.series)):
            raise VBAUnsupportedError("A chart area's font beside a chart element's font is not implemented")
        self.chart.changed()
        look = self.stored()
        if look is None:
            look = TextLook()
            if self.labels is None:
                state.element_fonts[self.key] = look
            else:
                self.labels.font = look
        return look

    @member
    def Parent(self) -> object:
        return self.element

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Name(self) -> object:
        name = self.look().name
        return self.chart.sheet.book.stylesheet.theme_fonts[1] if name is None else name

    @setter("Name")
    def _name(self, value: object) -> None:
        wanted = to_text(value)
        self.change().name = wanted

    @member
    def Size(self) -> object:
        size = self.look().size
        return self.size if size is None else size

    @setter("Size")
    def _size(self, value: object) -> None:
        wanted = float(to_number(value))
        self.look()
        if wanted <= 0 or wanted >= 500:
            raise error(1004, "Unable to set the Size property of the Font class")
        if not 1 <= wanted <= 409 or (wanted * 2) % 1:
            raise VBAUnsupportedError("This chart font size is not implemented")
        self.change().size = wanted

    def bolded(self) -> bool:
        bold = self.look().bold
        return self.bold if bold is None else bold

    def slanted(self) -> bool:
        return bool(self.look().italic)

    @member
    def Bold(self) -> object:
        return self.bolded()

    @setter("Bold")
    def _bold(self, value: object) -> None:
        wanted = to_bool(value)
        self.change().bold = wanted

    @member
    def Italic(self) -> object:
        return self.slanted()

    @setter("Italic")
    def _italic(self, value: object) -> None:
        wanted = to_bool(value)
        self.change().italic = wanted

    @member
    def FontStyle(self) -> object:
        bold, italic = self.bolded(), self.slanted()
        return {(False, False): "Regular", (True, False): "Bold", (False, True): "Italic",
                (True, True): "Bold Italic"}[(bold, italic)]

    @setter("FontStyle")
    def _font_style(self, value: object) -> None:
        wanted = _STYLES.get(to_text(value).lower())
        self.look()
        if wanted is None:
            raise VBAUnsupportedError("This chart font style is not implemented")
        look = self.change()
        look.bold, look.italic = wanted

    @member
    def Underline(self) -> object:
        return VBAInt(_SINGLE if self.look().underline else _NONE, "Long")

    @setter("Underline")
    def _underline(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        self.look()
        if wanted != _SINGLE:
            raise VBAUnsupportedError("This chart font underline is not implemented")
        self.change().underline = True

    @member
    def Strikethrough(self) -> object:
        return self.look().strike

    @setter("Strikethrough")
    def _strike(self, value: object) -> None:
        self.look()
        if not to_bool(value):
            raise VBAUnsupportedError("Taking a chart font's strikethrough off is not implemented")
        self.change().strike = True

    @member
    def Color(self) -> object:
        rgb = self.look().rgb
        return VBAInt(0 if rgb is None else bgr(rgb), "Long")

    @setter("Color")
    def _color(self, value: object) -> None:
        wanted = int(to_integer(value, "Long")) & 0xFFFFFF
        self.change().rgb = _rgb(wanted)

    @member
    def ColorIndex(self) -> object:
        rgb = self.look().rgb
        if rgb is None:
            return VBAInt(_AUTOMATIC, "Long")
        return VBAInt(self.chart.sheet.book.stylesheet.colors.nearest(rgb), "Long")

    @setter("ColorIndex")
    def _color_index(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        self.look()
        if not 1 <= wanted <= 56:
            raise VBAUnsupportedError("This chart font ColorIndex is not implemented")
        self.change().rgb = self.chart.sheet.book.stylesheet.colors.indexed[7 + wanted]

    # -- what a chart's text reads with no way the model sets it

    @member
    def Superscript(self) -> object:
        self.look()
        return False

    @member
    def Subscript(self) -> object:
        self.look()
        return False

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

    def no_theme(self) -> object:
        self.look()
        raise error(445, "Object doesn't support this action")

    @member
    def ThemeColor(self) -> object:
        return self.no_theme()

    @member
    def TintAndShade(self) -> object:
        return self.no_theme()

    @member
    def ThemeFont(self) -> object:
        return self.no_theme()

    # -- settings refused: Superscript writes nothing a save keeps, and no case measured the rest

    def refused(self, value: object) -> None:
        raise VBAUnsupportedError("Setting this on a chart's font is not implemented")

    @setter("Superscript")
    def _superscript(self, value: object) -> None:
        self.refused(value)

    @setter("Subscript")
    def _subscript(self, value: object) -> None:
        self.refused(value)

    @setter("Background")
    def _background(self, value: object) -> None:
        self.refused(value)

    @setter("ThemeColor")
    def _theme_color(self, value: object) -> None:
        self.refused(value)

    @setter("TintAndShade")
    def _tint_and_shade(self, value: object) -> None:
        self.refused(value)

    @setter("ThemeFont")
    def _theme_font(self, value: object) -> None:
        self.refused(value)
