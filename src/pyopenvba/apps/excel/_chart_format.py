"""Series.Format, as tests/fixtures/series_format.json measured it on each kind of chart.

A series the model made has no look of its own until one is set, and reads what its kind and place give it:

- A column, bar, pie, doughnut or area series fills with the theme's accent for its place, which Excel writes as a
  fixed colour once it spells it out. Its transparency reads mixed, and setting it does nothing, until then. It has no
  line: a weight or dash alone is written but reads as hidden, and a colour makes the line real, at 0.75 points. Hiding
  the line spells out the fill and leaves a noFill line for good, so it goes on reading hidden once shown again.
- A line series draws in that accent at 3 points, and takes no fill. A setting makes its transparency read 0. Hiding
  it writes it out in full with an effectLst, which stays.
- A line with markers or a scatter series gives a fill colour to its markers, as MarkerBackgroundColor does.

Combinations no case measured, a series past the sixth accent, and an imported chart report themselves unsupported."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._xml import escape
from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type
from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState, Series, SeriesState

_FILLED = ("column", "bar", "pie", "doughnut", "area")
_MIXED = -2147483648.0
_MIXED_COLOUR = -2147483648
_WHITE = 16777215
_INVALID = (-2147024809, "The specified value is out of range.")
_A14 = 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'
#: LineFormat.DashStyle and how the file writes it: its prstDash and the cap it takes.
_DASHES = {4: ("dash", None), 3: ("dot", "rnd")}


def _unmeasured(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is not implemented")


def _rgb(colour: int) -> str:
    return f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


def _colour(rgb: str, alpha: int | None = None) -> str:
    code = escape(rgb)
    return f'<a:srgbClr val="{code}"/>' if alpha is None else (
        f'<a:srgbClr val="{code}"><a:alpha val="{alpha}"/></a:srgbClr>')


def _solid(rgb: str, alpha: int | None = None) -> str:
    return f"<a:solidFill>{_colour(rgb, alpha)}</a:solidFill>"


def _accent(theme: tuple[str, ...], place: int) -> str:
    """The theme's accent for a series' place, of the six a chart goes through before it varies them."""
    if not 1 <= place <= 6:
        raise _unmeasured("The look of a series past the sixth")
    return theme[3 + place]


def _plain_line(look: FillLine) -> str:
    """A line as Excel writes what was set of it."""
    if look.line_rgb is None and look.line_width is None and look.line_dash is None:
        return ""
    attributes = ("" if look.line_width is None else f' w="{look.line_width}"') + (
        "" if look.line_cap is None else f' cap="{look.line_cap}"')
    inside = ("" if look.line_rgb is None else _solid(look.line_rgb)) + (
        "" if look.line_dash is None else f'<a:prstDash val="{look.line_dash}"/>')
    return f"<a:ln{attributes}>{inside}</a:ln>" if inside else f"<a:ln{attributes}/>"


def markup(entry: SeriesState, kind: str, place: int, theme: tuple[str, ...]) -> str:
    """The series' spPr, none while nothing was set of it."""
    look = entry.format_look
    if look is None or kind not in (*_FILLED, "line"):
        return ""
    accent = _accent(theme, place)
    extensions = ""
    if kind == "line":
        if look.line_full:
            paint = "<a:noFill/>" if look.line_hidden else _solid(accent)
            stroke = f'w="38100" cap="rnd" cmpd="sng" algn="ctr">{{}}<a:prstDash val="solid"/><a:round/>'
            body = f"<a:ln {stroke.format(paint)}</a:ln><a:effectLst/>"
            if look.line_hidden:
                extensions = ('<a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
                              f"<a14:hiddenLine {_A14} {stroke.format(_solid(accent))}</a14:hiddenLine></a:ext>")
        else:
            body = _plain_line(look)
    else:
        colour = look.fill_rgb or accent
        fill = ""
        if look.fill_hidden:
            fill = "<a:noFill/>"
            extensions = ('<a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">'
                          f"<a14:hiddenFill {_A14}>{_solid(colour, look.fill_alpha)}</a14:hiddenFill></a:ext>")
        elif look.fill_rgb is not None or look.fill_solid:
            fill = _solid(colour, look.fill_alpha)
        line = "<a:ln><a:noFill/></a:ln>" if look.line_full else _plain_line(look)
        body = fill + line + ("<a:effectLst/>" if look.effects else "")
        if look.line_hidden:
            extensions += ('<a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
                           f"<a14:hiddenLine {_A14}><a:noFill/></a14:hiddenLine></a:ext>")
    if extensions:
        body += f"<a:extLst>{extensions}</a:extLst>"
    return f"<c:spPr>{body}</c:spPr>" if body else ""


class _Part(Created):
    """What a series' Fill and Line share."""

    vba_type_name = ""

    def __init__(self, owner: ChartFormat) -> None:
        self.owner = owner

    @property
    def series(self) -> Series:
        return self.owner.series

    def state(self) -> ChartState:
        self.series.position()
        state = self.series.chart.state
        if state.original:
            raise _unmeasured("The look of a series in an imported chart")
        return state

    def kind(self) -> str:
        return CHART_TYPES.get(shown_type(self.state()), "")

    def look(self) -> FillLine:
        return self.series.entry.format_look or FillLine()

    def change(self) -> FillLine:
        self.series.chart.changed()
        if self.series.entry.format_look is None:
            self.series.entry.format_look = FillLine()
        return self.series.entry.format_look

    def accent(self) -> int:
        """The series' accent, as a BGR Long."""
        theme = self.series.chart.sheet.book.stylesheet.colors.theme
        return bgr(_accent(theme, self.series.position()))

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application


class SeriesFill(_Part):
    vba_type_name = "FillFormat"

    def filled(self) -> bool:
        return self.kind() in _FILLED

    def written(self) -> bool:
        look = self.look()
        return look.fill_rgb is not None or look.fill_solid

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.filled() and self.look().fill_hidden else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in (0, -1):
            raise _unmeasured("Visible set to anything but msoTrue or msoFalse")
        kind = self.kind()
        if kind == "line" and wanted == 0:
            return
        if not self.filled():
            raise _unmeasured("The visibility of this kind of series' fill")
        if wanted == 0 and self.look().line_full:
            raise _unmeasured("Hiding a series' fill once its line was hidden")
        look = self.change()
        look.fill_hidden = wanted == 0
        look.fill_solid = True

    @member
    def ForeColor(self) -> object:
        return SeriesColor(self, fore=True)

    @member
    def BackColor(self) -> object:
        return SeriesColor(self, fore=False)

    @member
    def Type(self) -> object:
        return VBAInt(1 if self.filled() and self.written() else -2, "Long")

    @member
    def Transparency(self) -> object:
        if not (self.filled() and self.written()):
            return VBASingle(_MIXED)
        alpha = self.look().fill_alpha
        return VBASingle(0.0 if alpha is None else (100000 - alpha) / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        if not self.filled():
            self.state()
            raise _unmeasured("The fill transparency of this kind of series")
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        # A fill with no colour written takes none.
        if self.written():
            self.change().fill_alpha = round((1 - wanted) * 100000)

    @method
    def Solid(self) -> object:
        if not self.filled() or self.look().fill_hidden:
            raise _unmeasured("Solid on this kind of series, or on a hidden fill,")
        self.change().fill_solid = True
        return EMPTY


class SeriesLine(_Part):
    vba_type_name = "LineFormat"

    def reads_hidden(self) -> bool:
        """A column-like series' line that is no real line, or a scatter series': hidden, white and mixed."""
        look = self.look()
        kind = self.kind()
        if kind == "scatter":
            return True
        return kind in _FILLED and look.line_rgb is None and (
            look.line_full or look.line_width is not None or look.line_dash is not None)

    def absent(self) -> bool:
        """A column-like series' line with nothing set of it."""
        look = self.look()
        return (self.kind() in _FILLED and look.line_rgb is None and look.line_width is None
                and look.line_dash is None and not look.line_full)

    def touched(self) -> bool:
        look = self.look()
        return (look.line_rgb is not None or look.line_width is not None or look.line_dash is not None
                or look.line_full)

    def settable(self, what: str) -> None:
        if self.kind() not in (*_FILLED, "line"):
            raise _unmeasured(f"The line {what} of this kind of series")
        if self.look().line_full:
            raise _unmeasured(f"The line {what} of a series whose line was hidden")

    @member
    def Visible(self) -> object:
        if self.kind() == "line":
            return VBAInt(0 if self.look().line_hidden else -1, "Long")
        return VBAInt(0 if self.reads_hidden() else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in (0, -1):
            raise _unmeasured("Visible set to anything but msoTrue or msoFalse")
        kind = self.kind()
        if kind not in (*_FILLED, "line"):
            raise _unmeasured("The line visibility of this kind of series")
        before = self.look()
        if wanted == -1:
            if before.line_hidden:
                self.change().line_hidden = False
            return
        if before.line_hidden:
            return
        if before != FillLine():
            raise _unmeasured("Hiding the line of a series with a look set")
        look = self.change()
        look.line_hidden = look.line_full = look.effects = True
        look.fill_solid = kind in _FILLED

    @member
    def ForeColor(self) -> object:
        return SeriesColor(self, fore=True)

    @member
    def Weight(self) -> object:
        if self.absent() or self.reads_hidden():
            return VBASingle(_MIXED)
        width = self.look().line_width
        if width is not None:
            return VBASingle(points(width))
        return VBASingle(0.75 if self.kind() in _FILLED else 3.0)

    @setter("Weight")
    def _weight(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        self.settable("weight")
        if wanted < 0:
            raise error(*_INVALID)
        self.change().line_width = emu(wanted)

    @member
    def DashStyle(self) -> object:
        if self.absent() or self.reads_hidden():
            return VBAInt(-2, "Long")
        dash = self.look().line_dash
        return VBAInt(next((key for key, (name, _) in _DASHES.items() if name == dash), 1), "Long")

    @setter("DashStyle")
    def _dash_style(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        self.settable("dash")
        if wanted not in _DASHES or (_DASHES[wanted][1] is not None and self.kind() != "line"):
            raise _unmeasured("This line dash style")
        look = self.change()
        look.line_dash, look.line_cap = _DASHES[wanted]

    @member
    def Style(self) -> object:
        return VBAInt(-2 if self.absent() or self.reads_hidden() else 1, "Long")

    @member
    def Transparency(self) -> object:
        if self.absent() or self.reads_hidden():
            return VBASingle(_MIXED)
        return VBASingle(0.0 if self.touched() else _MIXED)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        # With nothing set of the line it takes none.
        self.settable("transparency")
        if self.touched():
            raise _unmeasured("The transparency of a series' line with a look set")


class SeriesColor(Created):
    """A series' fill or line colour."""

    vba_type_name = "ColorFormat"

    def __init__(self, owner: SeriesFill | SeriesLine, *, fore: bool) -> None:
        self.owner, self.fore = owner, fore

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.Application()

    def colourless(self) -> bool:
        """Whether the part has no colour to read: a line series' fill, or a column-like series' absent line."""
        if isinstance(self.owner, SeriesFill):
            return not self.owner.filled()
        return self.owner.absent()

    @member
    def RGB(self) -> object:
        if not self.fore:
            self.owner.state()
            return VBAInt(_WHITE, "Long")
        if self.colourless():
            return VBAInt(0, "Long")
        if isinstance(self.owner, SeriesLine) and self.owner.reads_hidden():
            return VBAInt(_WHITE, "Long")
        look = self.owner.look()
        own = look.fill_rgb if isinstance(self.owner, SeriesFill) else look.line_rgb
        return VBAInt(self.owner.accent() if own is None else bgr(own), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        wanted = _rgb(int(to_integer(value, "Long")) & 0xFFFFFF)
        kind = self.owner.kind()
        if not self.fore:
            raise _unmeasured("Setting a series' back colour")
        if isinstance(self.owner, SeriesLine):
            self.owner.settable("colour")
            self.owner.change().line_rgb = wanted
            return
        if kind == "line":
            return  # A line series takes no fill.
        if kind in ("lineMarkers", "scatter"):
            # Its markers take it, as MarkerBackgroundColor does.
            self.owner.series.chart.changed()
            self.owner.series.entry.marker_back = wanted
            return
        if self.owner.look().fill_hidden or self.owner.look().line_full:
            raise _unmeasured("A colour for a series' hidden fill, or once its line was hidden,")
        self.owner.change().fill_rgb = wanted

    @member
    def ObjectThemeColor(self) -> object:
        if not self.fore:
            raise _unmeasured("The theme colour of a series' back colour")
        return VBAInt(-2 if self.colourless() else 0, "Long")

    @member
    def Type(self) -> object:
        if not self.fore:
            raise _unmeasured("The type of a series' back colour")
        return VBAInt(-2 if self.colourless() else 1, "Long")


class ChartFormat(Created):
    """A chart element's look; a series' is what the model has."""

    vba_type_name = "ChartFormat"

    def __init__(self, series: Series) -> None:
        self.series = series

    @member
    def Parent(self) -> object:
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Fill(self) -> object:
        return SeriesFill(self)

    @member
    def Line(self) -> object:
        return SeriesLine(self)
