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

from typing import TYPE_CHECKING, Any, cast

from pyopenvba._xml import escape
from pyopenvba.apps.excel._chart_styled import is_styled, line_read, point_line_weight, series_line_weight
from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type
from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, VBASingle, error, to_integer, to_number
from pyopenvba.shapes._values import FillLine, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Chart, ChartState, PointState, Series, SeriesState

FILLED = ("column", "bar", "pie", "doughnut", "area")
_MIXED = -2147483648.0
_MIXED_COLOUR = -2147483648
_WHITE = 16777215
_INVALID = (-2147024809, "The specified value is out of range.")
_A14 = 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'
#: LineFormat.DashStyle and how the file writes it: its prstDash and the cap it takes.
_DASHES = {4: ("dash", None), 3: ("dot", "rnd"), 7: ("lgDash", None), 10: ("sysDash", None)}


def unmeasured(what: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f"{what} is not implemented")


def _rgb(colour: int) -> str:
    return f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


def _colour(rgb: str, alpha: int | None = None) -> str:
    code = escape(rgb)
    return f'<a:srgbClr val="{code}"/>' if alpha is None else (
        f'<a:srgbClr val="{code}"><a:alpha val="{alpha}"/></a:srgbClr>')


def _solid(rgb: str, alpha: int | None = None) -> str:
    return f"<a:solidFill>{_colour(rgb, alpha)}</a:solidFill>"


def accent_colour(theme: tuple[str, ...], place: int) -> str:
    """The theme's accent for a series' place, of the six a chart goes through before it varies them."""
    if not 1 <= place <= 6:
        raise unmeasured("The look of a series past the sixth")
    return theme[3 + place]


def solid_fill(rgb: str, alpha: int | None = None) -> str:
    return _solid(rgb, alpha)


def effect_list(look: FillLine) -> str:
    """The effect set on a chart element, as Excel spells it (tests/fixtures/chart_protection.json)."""
    if look.chart_effect is None:
        return ""
    name, value = look.chart_effect
    black = '<a:scrgbClr r="0" g="0" b="0"/>'
    if name == "Glow.Radius":
        inside = f'<a:glow rad="{round(float(str(value)) * 12700)}">{black}</a:glow>'
    elif name == "Glow.RGB":
        inside = f'<a:glow><a:srgbClr val="{value}"/></a:glow>'
    elif name == "Shadow.Visible":
        inside = f'<a:outerShdw blurRad="63500" dist="37357" dir="2700000" rotWithShape="0">{black}</a:outerShdw>'
    elif name == "Shadow.Type":
        inside = ('<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl">'
                  '<a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw>')
    else:
        inside = f'<a:softEdge rad="{round(float(str(value)) * 12700)}"/>'
    return f"<a:effectLst>{inside}</a:effectLst>"


def with_effect(properties: str, look: FillLine | None) -> str:
    """A spPr with the effect set on the element in its place, before any extensions."""
    effect = "" if look is None else effect_list(look)
    if not effect:
        return properties
    if not properties:
        return f"<c:spPr>{effect}</c:spPr>"
    if "<a:effectLst" in properties:
        raise VBAUnsupportedError("An effect on a chart element whose line or fill is hidden is not implemented")
    place = properties.find("<a:extLst>")
    if place < 0:
        place = properties.rindex("</c:spPr>")
    return properties[:place] + effect + properties[place:]


def plain_line(look: FillLine) -> str:
    """A line as Excel writes what was set of it."""
    if look.line_rgb is None and look.line_width is None and look.line_dash is None and look.line_compound is None:
        return ""
    attributes = ("" if look.line_width is None else f' w="{look.line_width}"') + (
        "" if look.line_cap is None else f' cap="{look.line_cap}"') + (
        "" if look.line_compound is None else f' cmpd="{look.line_compound}"')
    inside = ("<a:noFill/>" if look.line_off else "" if look.line_rgb is None else _solid(look.line_rgb)) + (
        "" if look.line_dash is None else f'<a:prstDash val="{look.line_dash}"/>')
    return f"<a:ln{attributes}>{inside}</a:ln>" if inside else f"<a:ln{attributes}/>"


def hidden_fill(colour: str, alpha: int | None = None) -> str:
    """The extension that keeps a hidden fill's colour."""
    return ('<a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">'
            f"<a14:hiddenFill {_A14}>{_solid(colour, alpha)}</a14:hiddenFill></a:ext>")


def markup(entry: SeriesState, kind: str, place: int, theme: tuple[str, ...]) -> str:
    """The series' spPr, none while nothing was set of it."""
    look = entry.format_look
    if look is None:
        return ""
    if kind in ("lineMarkers", "scatter"):
        # Only Border reaches the line of these.
        return look_markup(look, "line", accent_colour(theme, place)) if plain_line(look) else with_effect("", look)
    if kind not in (*FILLED, "line"):
        return with_effect("", look)
    return look_markup(look, kind, accent_colour(theme, place))


def look_markup(look: FillLine, kind: str, accent: str) -> str:
    """The spPr of a series or point of this kind of chart, whose accent is given."""
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
            body = plain_line(look)
    else:
        colour = look.fill_rgb or accent
        fill = ""
        if look.fill_hidden:
            fill = "<a:noFill/>"
            extensions = hidden_fill(colour, look.fill_alpha)
        elif look.fill_rgb is not None or look.fill_solid:
            fill = _solid(colour, look.fill_alpha)
        line = "<a:ln><a:noFill/></a:ln>" if look.line_full else plain_line(look)
        body = fill + line + ("<a:effectLst/>" if look.effects else "")
        if look.line_hidden:
            extensions += ('<a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
                           f"<a14:hiddenLine {_A14}><a:noFill/></a14:hiddenLine></a:ext>")
    if extensions:
        body += f"<a:extLst>{extensions}</a:extLst>"
    return with_effect(f"<c:spPr>{body}</c:spPr>" if body else "", look)


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
        if state.original and not self.series.entry.look_read:
            raise unmeasured("The look of a series in an imported chart")
        return state

    def kind(self) -> str:
        return CHART_TYPES.get(shown_type(self.state()), "")

    def look(self) -> FillLine:
        return self.series.entry.format_look or FillLine()

    def admit(self) -> None:
        """What a change must pass before it is looked at; a point adds to it."""

    def wipe(self, point: PointState) -> None:
        """What a change to the series takes from a point set apart (tests/fixtures/series_points.json)."""

    def change(self) -> FillLine:
        self.series.chart.changed()
        if self.series.entry.format_look is None:
            self.series.entry.format_look = FillLine()
        points = self.series.entry.points
        for place, point in list(points.items()):
            self.wipe(point)
            if point.look == FillLine():
                point.look = None
            if (point.look is None and point.invert is None and point.explosion is None and point.label is None
                    and not point.has_marker()):
                del points[place]
        return self.series.entry.format_look

    def accent(self) -> int:
        """The series' accent, as a BGR Long: of its idx, which it keeps when a series before it goes."""
        theme = self.series.chart.sheet.book.stylesheet.colors.theme
        place = self.series.position()
        identity = self.series.entry.series_id
        return bgr(accent_colour(theme, identity + 1 if identity >= 0 else place))

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application


class SeriesFill(_Part):
    vba_type_name = "FillFormat"

    def filled(self) -> bool:
        return self.kind() in FILLED

    def written(self) -> bool:
        look = self.look()
        return look.fill_rgb is not None or look.fill_solid

    def wipe(self, point: PointState) -> None:
        if point.look is not None:
            point.look.fill_rgb, point.look.fill_alpha = None, None
            point.look.fill_solid = point.look.fill_hidden = False

    def marker_took(self, wanted: str) -> bool:
        """Whether a point's marker took this colour, which a series' setter has done with."""
        del wanted
        return False

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.filled() and self.look().fill_hidden else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        self.admit()
        if wanted not in (0, -1):
            raise unmeasured("Visible set to anything but msoTrue or msoFalse")
        kind = self.kind()
        if kind == "line" and wanted == 0:
            return
        if not self.filled():
            raise unmeasured("The visibility of this kind of series' fill")
        if wanted == 0 and self.look().line_full:
            raise unmeasured("Hiding a series' fill once its line was hidden")
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
        self.admit()
        if not self.filled():
            self.state()
            raise unmeasured("The fill transparency of this kind of series")
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        # A fill with no colour written takes none.
        if self.written():
            self.change().fill_alpha = round((1 - wanted) * 100000)

    @method
    def Solid(self) -> object:
        self.admit()
        if not self.filled() or self.look().fill_hidden:
            raise unmeasured("Solid on this kind of series, or on a hidden fill,")
        self.change().fill_solid = True
        return EMPTY


class SeriesLine(_Part):
    vba_type_name = "LineFormat"

    def reads_hidden(self) -> bool:
        """A column-like series' line that is no real line, or a scatter series': hidden, white and mixed."""
        look = self.look()
        kind = self.kind()
        if kind == "scatter":
            return not self.series.entry.scatter_line
        return kind in FILLED and look.line_rgb is None and (
            look.line_full or look.line_width is not None or look.line_dash is not None)

    def absent(self) -> bool:
        """A column-like series' line with nothing set of it."""
        look = self.look()
        return (self.kind() in FILLED and look.line_rgb is None and look.line_width is None
                and look.line_dash is None and not look.line_full)

    def touched(self) -> bool:
        look = self.look()
        return (look.line_rgb is not None or look.line_width is not None or look.line_dash is not None
                or look.line_full)

    def wipe(self, point: PointState) -> None:
        if point.look is not None:
            look = point.look
            look.line_rgb, look.line_width, look.line_dash, look.line_cap = None, None, None, None
            point.look.line_full = point.look.line_hidden = point.look.effects = False
        if self.kind() not in FILLED:
            point.marker_fore, point.marker_line, point.marker_solid = None, None, False

    def settable(self, what: str) -> None:
        self.admit()
        if self.kind() not in (*FILLED, "line"):
            raise unmeasured(f"The line {what} of this kind of series")
        if self.look().line_full:
            raise unmeasured(f"The line {what} of a series whose line was hidden")

    @member
    def Visible(self) -> object:
        if self.kind() == "line":
            return VBAInt(0 if self.look().line_hidden else -1, "Long")
        return VBAInt(0 if self.reads_hidden() else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        self.admit()
        if wanted not in (0, -1):
            raise unmeasured("Visible set to anything but msoTrue or msoFalse")
        kind = self.kind()
        if kind not in (*FILLED, "line"):
            raise unmeasured("The line visibility of this kind of series")
        before = self.look()
        if wanted == -1:
            if before.line_hidden:
                self.change().line_hidden = False
            return
        if before.line_hidden:
            return
        if before != FillLine():
            raise unmeasured("Hiding the line of a series with a look set")
        look = self.change()
        look.line_hidden = look.line_full = look.effects = True
        look.fill_solid = kind in FILLED

    @member
    def ForeColor(self) -> object:
        return SeriesColor(self, fore=True)

    @member
    def Weight(self) -> object:
        style = self.series.chart.state.chart_style
        if is_styled(style) and self.look().line_width is None and self.kind() in FILLED and point_line_weight(style):
            return VBASingle(point_line_weight(style))
        if self.absent() or self.reads_hidden():
            return VBASingle(_MIXED)
        width = self.look().line_width
        if width is not None:
            return VBASingle(points(width))
        style = self.series.chart.state.chart_style
        if is_styled(style) and self.kind() not in FILLED:
            return VBASingle(series_line_weight(style))
        return VBASingle(0.75 if self.kind() in FILLED else 3.0)

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
            raise unmeasured("This line dash style")
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
            raise unmeasured("The transparency of a series' line with a look set")


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
        self.owner.admit()
        kind = self.owner.kind()
        if not self.fore:
            raise unmeasured("Setting a series' back colour")
        if isinstance(self.owner, SeriesLine):
            entry = self.owner.series.entry
            if entry.marker_line is not None and type(self.owner) is not SeriesLine:
                raise unmeasured("A line on a point of a series whose line colour reached its marker")
            if type(self.owner) is SeriesLine and kind in ("lineMarkers", "scatter") and entry.points:
                raise unmeasured("A series' line colour beside a look on one of its points")
            if type(self.owner) is SeriesLine and kind == "scatter" and self.owner.look() == FillLine():
                # A scatter series keeps no line, and its marker an empty one (tests/fixtures/series_format.json).
                self.owner.series.chart.changed()
                entry.marker_line = ""
                return
            if type(self.owner) is SeriesLine and kind == "lineMarkers" and not self.owner.look().line_full:
                entry.marker_line = wanted
            else:
                self.owner.settable("colour")
            self.owner.change().line_rgb = wanted
            return
        if self.owner.marker_took(wanted):
            return
        if kind == "line":
            return  # A line series takes no fill.
        if kind in ("lineMarkers", "scatter"):
            # Its markers take it, as MarkerBackgroundColor does.
            self.owner.series.chart.changed()
            self.owner.series.entry.marker_back = wanted
            self.owner.series.taken_from_points("marker_back")
            return
        if self.owner.look().fill_hidden or self.owner.look().line_full:
            raise unmeasured("A colour for a series' hidden fill, or once its line was hidden,")
        self.owner.change().fill_rgb = wanted

    @member
    def ObjectThemeColor(self) -> object:
        if not self.fore:
            raise unmeasured("The theme colour of a series' back colour")
        return VBAInt(-2 if self.colourless() else 0, "Long")

    @member
    def Type(self) -> object:
        if not self.fore:
            raise unmeasured("The type of a series' back colour")
        return VBAInt(-2 if self.colourless() else 1, "Long")


#: The elements whose Format has a 3-D look and a picture, and those with a text frame, by their type
#: (tests/fixtures/chart_protection.json).
_SOLID = ("Series", "Point", "ChartArea", "PlotArea", "Legend", "ChartTitle", "AxisTitle", "DataLabels")
_TEXTUAL = ("ChartArea", "Legend", "ChartTitle", "AxisTitle", "DataLabels", "DataTable")
_NO_MEMBER = (-2147467259, "Automation error\nUnspecified error")


class _EffectRead(Created):
    """An effect of a chart element as it reads with nothing set. Setting one is not built, so its setters report so."""

    vba_type_name = ""

    def __init__(self, owner: object) -> None:
        self.owner = owner

    @member
    def Parent(self) -> object:
        return self.owner

    def vba_set(self, name: str, value: object, *args: Any, **named: Any) -> None:
        del value, args, named
        raise unmeasured(f"Setting {self.vba_type_name}.{name} on a chart element")

    def set_effect(self) -> tuple[str, object] | None:
        owner = cast(EffectMembers, self.owner)
        return owner.effect_look().chart_effect

    def read(self, wanted: tuple[str, ...]) -> object:
        """The value of the effect set if it is one of these. Under another effect nothing was measured."""
        effect = self.set_effect()
        if effect is None:
            return None
        if effect[0] not in wanted:
            raise unmeasured(f"Reading {self.vba_type_name} with {effect[0]} set on a chart element")
        return effect[1]

    def write(self, name: str, value: object) -> None:
        effect = self.set_effect()
        if effect is not None and effect[0] != name:
            raise unmeasured(f"{name} with {effect[0]} set on a chart element")
        cast(EffectMembers, self.owner).effect_change().chart_effect = (name, value)


class EffectColorRead(_EffectRead):
    vba_type_name = "ColorFormat"

    @member
    def RGB(self) -> object:
        found = cast(_EffectRead, self.owner).read(("Glow.RGB",))
        return VBAInt(-2147483648 if found is None else bgr(str(found)), "Long")

    def vba_set(self, name: str, value: object, *args: Any, **named: Any) -> None:
        if name.lower() != "rgb":
            raise unmeasured(f"Setting ColorFormat.{name} of a chart element's glow")
        del args, named
        cast(_EffectRead, self.owner).write("Glow.RGB", _rgb(int(to_integer(value, "Long")) & 0xFFFFFF))

    def set_effect(self) -> tuple[str, object] | None:
        return cast(_EffectRead, self.owner).set_effect()


class GlowRead(_EffectRead):
    vba_type_name = "GlowFormat"

    @member
    def Radius(self) -> object:
        found = self.read(("Glow.Radius",))
        return VBASingle(0.0 if found is None else float(str(found)))

    @member
    def Color(self) -> object:
        return EffectColorRead(self)

    @member
    def Transparency(self) -> object:
        self.read(())
        return VBASingle(_MIXED)

    def vba_set(self, name: str, value: object, *args: Any, **named: Any) -> None:
        del args, named
        if name.lower() != "radius":
            raise unmeasured(f"Setting GlowFormat.{name} on a chart element")
        self.write("Glow.Radius", float(to_number(value)))


class ShadowRead(_EffectRead):
    vba_type_name = "ShadowFormat"

    @member
    def Visible(self) -> object:
        effect = self.read(("Shadow.Visible", "Shadow.Type"))
        return VBAInt(0 if effect is None else -1, "Long")

    @member
    def Type(self) -> object:
        effect = self.set_effect()
        if effect is not None and effect[0] == "Shadow.Type":
            return VBAInt(cast(int, effect[1]), "Long")
        self.read(("Shadow.Visible",))
        return VBAInt(-2, "Long")

    @member
    def Blur(self) -> object:
        self.read(("Shadow.Visible",))
        return VBASingle(5.0)

    @member
    def OffsetX(self) -> object:
        self.read(())
        return VBASingle(2.079952)

    @member
    def OffsetY(self) -> object:
        self.read(())
        return VBASingle(2.079952)

    def vba_set(self, name: str, value: object, *args: Any, **named: Any) -> None:
        del args, named
        wanted = int(to_integer(value, "Long"))
        if name.lower() == "visible" and wanted == -1:
            self.write("Shadow.Visible", True)
        elif name.lower() == "type" and wanted == 21:
            self.write("Shadow.Type", wanted)
        else:
            raise unmeasured(f"Setting ShadowFormat.{name} to {wanted} on a chart element")


class SoftEdgeRead(_EffectRead):
    vba_type_name = "SoftEdgeFormat"

    @member
    def Type(self) -> object:
        return VBAInt(0 if self.read(("SoftEdge.Type",)) is None else 1, "Long")

    @member
    def Radius(self) -> object:
        found = self.read(("SoftEdge.Type", "SoftEdge.Radius"))
        return VBASingle(0.0 if found is None else float(str(found)))

    def vba_set(self, name: str, value: object, *args: Any, **named: Any) -> None:
        del args, named
        if name.lower() == "type" and int(to_integer(value, "Long")) == 1:
            self.write("SoftEdge.Type", 1.0)
        elif name.lower() == "radius":
            self.write("SoftEdge.Radius", float(to_number(value)))
        else:
            raise unmeasured(f"Setting SoftEdgeFormat.{name} on a chart element")


class ThreeDRead(_EffectRead):
    vba_type_name = "ThreeDFormat"

    @member
    def Visible(self) -> object:
        return VBAInt(0, "Long")

    @member
    def BevelTopType(self) -> object:
        return VBAInt(1, "Long")


class PictureRead(_EffectRead):
    vba_type_name = "PictureFormat"


class TextFrameRead(_EffectRead):
    vba_type_name = "TextFrame2"

    def __init__(self, owner: object, titled: bool) -> None:
        super().__init__(owner)
        self.titled = titled

    @member
    def HasText(self) -> object:
        return VBAInt(-1 if self.titled else 0, "Long")


class EffectMembers(Created):
    """What every chart element's Format offers beyond its fill and line."""

    vba_type_name = ""

    def effect_look(self) -> FillLine:
        """The element's look as it stands, for its effect."""
        raise unmeasured("The effects of this chart element")

    def effect_change(self) -> FillLine:
        raise unmeasured("The effects of this chart element")

    def _element_type(self) -> str:
        element = self.vba_get("Parent", [])
        return str(getattr(element, "vba_type_name", ""))

    def _solid(self) -> None:
        if self._element_type() not in _SOLID:
            raise error(*_NO_MEMBER)

    @member
    def Glow(self) -> object:
        return GlowRead(self)

    @member
    def Shadow(self) -> object:
        return ShadowRead(self)

    @member
    def SoftEdge(self) -> object:
        return SoftEdgeRead(self)

    @member
    def ThreeD(self) -> object:
        self._solid()
        return ThreeDRead(self)

    @member
    def PictureFormat(self) -> object:
        if self._element_type() not in (*_SOLID, "TickLabels", "DataTable"):
            raise error(*_NO_MEMBER)
        return PictureRead(self)

    @member
    def TextFrame2(self) -> object:
        kind = self._element_type()
        if kind not in _TEXTUAL:
            raise error(*_NO_MEMBER)
        return TextFrameRead(self, kind in ("ChartTitle", "AxisTitle"))

    @member
    def Adjustments(self) -> object:
        raise error(*_NO_MEMBER)

    @member
    def AutoShapeType(self) -> object:
        raise error(*_NO_MEMBER)

    @setter("AutoShapeType")
    def _auto_shape_type(self, value: object) -> None:
        del value
        raise error(*_NO_MEMBER)


class ChartFormat(EffectMembers, Created):
    """A series' look; ElementFormat is the same type for the chart's other elements."""

    vba_type_name = "ChartFormat"

    def __init__(self, series: Series) -> None:
        self.series = series

    def effect_look(self) -> FillLine:
        return SeriesFill(self).look()

    def effect_change(self) -> FillLine:
        return SeriesFill(self).change()

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


# -- the chart's other elements (tests/fixtures/chart_element_format.json)
#
# Each is one of a few profiles. The chart and plot areas fill white, as the system's window colour; the rest have no
# fill until one is set, and gridlines refuse a fill altogether. The chart area, axes and gridlines draw a grey line of
# 1 point, the system's window text at a tint; the rest have no line until a colour is set, as a column series has not.

_WINDOW = '<a:sysClr val="window" lastClr="FFFFFF"/>'
_GREY = '<a:sysClr val="windowText" lastClr="000000"><a:tint val="75000"/></a:sysClr>'
_GREY_LONG = 9013641
_GREY_STROKE = ('w="12700" cap="flat" cmpd="sng" algn="ctr">{}<a:prstDash val="solid"/><a:round/>')


def _windowed(key: str) -> bool:
    return key in ("area", "plot")


def _grey(key: str) -> bool:
    return key == "area" or key.startswith(("axis", "grid"))


def element_markup(key: str, look: FillLine | None) -> str:
    """An element's spPr, none while nothing was set of it."""
    if look is None:
        return ""
    plain = _WINDOW if _windowed(key) else '<a:schemeClr val="accent1"/>'
    fill = ""
    extensions = ""
    if look.fill_hidden:
        fill = "<a:noFill/>"
        hidden = f"<a:solidFill>{_WINDOW}</a:solidFill>" if _windowed(key) else "<a:noFill/>"
        extensions = ('<a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">'
                      f"<a14:hiddenFill {_A14}>{hidden}</a14:hiddenFill></a:ext>")
    elif look.fill_rgb is not None:
        fill = _solid(look.fill_rgb, look.fill_alpha)
    elif look.fill_solid:
        fill = f"<a:solidFill>{plain}</a:solidFill>"
    elif look.no_fill:
        fill = "<a:noFill/>"
    if not look.line_full:
        line = plain_line(look)
    elif _grey(key):
        line = f"<a:ln {_GREY_STROKE.format('<a:noFill/>')}</a:ln>"
    else:
        line = "<a:ln><a:noFill/></a:ln>"
    if look.line_hidden:
        inside = (f"{_A14} {_GREY_STROKE.format(f'<a:solidFill>{_GREY}</a:solidFill>')}" if _grey(key)
                  else f"{_A14}><a:noFill/>")
        extensions += ('<a:ext uri="{91240B29-F687-4F45-9708-019B960494DF}">'
                       f"<a14:hiddenLine {inside}</a14:hiddenLine></a:ext>")
    body = fill + line + ("<a:effectLst/>" if look.effects else "") + (
        f"<a:extLst>{extensions}</a:extLst>" if extensions else "")
    return with_effect(f"<c:spPr>{body}</c:spPr>" if body else "", look)


class ElementFormat(EffectMembers, Created):
    """The Format of a chart's element, other than a series."""

    vba_type_name = "ChartFormat"

    def __init__(self, element: object, chart: Chart, key: str) -> None:
        self.element, self.chart, self.key = element, chart, key

    def effect_look(self) -> FillLine:
        return ElementFill(self).look()

    def effect_change(self) -> FillLine:
        return ElementFill(self).written()

    @member
    def Parent(self) -> object:
        return self.element

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Fill(self) -> object:
        return ElementFill(self)

    @member
    def Line(self) -> object:
        return ElementLine(self)


class _ElementPart(Created):
    vba_type_name = ""

    def __init__(self, owner: ElementFormat) -> None:
        self.owner = owner
        self.key = owner.key

    def look(self) -> FillLine:
        chart = self.owner.chart
        chart.admit_read()
        if chart.state.original and self.key not in chart.state.looks_read:
            raise unmeasured("The look of a part of an imported chart")
        return chart.state.element_looks.get(self.key) or FillLine()

    def change(self) -> FillLine:
        """The look to change, which one written out by hiding keeps as it is: no case changed it further."""
        before = self.look()
        if before.fill_hidden or before.line_full:
            raise unmeasured("Changing the look of a chart element with its fill or line hidden")
        return self.written()

    def written(self) -> FillLine:
        chart = self.owner.chart
        chart.changed()
        return chart.state.element_looks.setdefault(self.key, FillLine())

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.chart.sheet.book.application


class ElementFill(_ElementPart):
    vba_type_name = "FillFormat"

    def usable(self) -> FillLine:
        if self.key.startswith("grid"):
            # Gridlines are lines: they have no fill to read or set.
            raise error(-2147467259, "Automation error\nUnspecified error")
        return self.look()

    def hidden(self) -> bool:
        """Whether it reads as hidden, which an element with no fill does once its line is hidden."""
        look = self.usable()
        return look.fill_hidden or look.no_fill

    def spelled(self) -> bool:
        look = self.usable()
        return look.fill_rgb is not None or look.fill_solid or look.fill_hidden

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.hidden() else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        before = self.usable()
        if wanted not in (0, -1):
            raise unmeasured("Visible set to anything but msoTrue or msoFalse")
        if wanted == -1 and not before.fill_hidden:
            return
        if wanted == 0 and before.fill_hidden:
            return
        if before != FillLine():
            raise unmeasured("Hiding or showing the fill of a chart element with a look set")
        self.written().fill_hidden = True

    @member
    def ForeColor(self) -> object:
        return ElementColor(self, fore=True)

    @member
    def BackColor(self) -> object:
        return ElementColor(self, fore=False)

    @member
    def Type(self) -> object:
        if not self.spelled() or (self.hidden() and not _windowed(self.key)):
            return VBAInt(-2, "Long")
        return VBAInt(1, "Long")

    @member
    def Transparency(self) -> object:
        look = self.usable()
        if self.hidden() and not _windowed(self.key):
            return VBASingle(1.0)
        if not self.spelled():
            return VBASingle(_MIXED)
        return VBASingle(0.0 if look.fill_alpha is None else (100000 - look.fill_alpha) / 100000)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        look = self.usable()
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        if look.fill_rgb is not None:
            self.change().fill_alpha = round((1 - wanted) * 100000)
        elif look.fill_solid or look.fill_hidden:
            raise unmeasured("The transparency of a chart element's fill with no colour of its own")
        # With no colour written it takes none.

    @method
    def Solid(self) -> object:
        self.usable()
        self.change().fill_solid = True
        return EMPTY


class ElementLine(_ElementPart):
    vba_type_name = "LineFormat"

    def styled(self) -> tuple[float, str] | None:
        """The weight and colour one of AddChart2's styles gives the line."""
        chart = self.owner.chart
        style = chart.state.chart_style
        return line_read(style, self.key, chart.sheet.book.stylesheet.colors.theme) if is_styled(style) else None

    def reads_hidden(self) -> bool:
        """A line that is no real line: hidden, or a weight or dash on an element with no line."""
        look = self.look()
        if look.line_off:
            # Hidden through its Border (tests/fixtures/chart_table_look.json).
            return True
        return not _grey(self.key) and look.line_rgb is None and (
            look.line_full or look.line_width is not None or look.line_dash is not None)

    def absent(self) -> bool:
        return not _grey(self.key) and self.look().line_rgb is None and not self.reads_hidden()

    def touched(self) -> bool:
        look = self.look()
        return (look.line_rgb is not None or look.line_width is not None or look.line_dash is not None
                or look.line_full)

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.look().line_hidden or self.reads_hidden() else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        before = self.look()
        if wanted not in (0, -1):
            raise unmeasured("Visible set to anything but msoTrue or msoFalse")
        if (wanted == -1) != before.line_hidden:
            return
        if before != FillLine():
            raise unmeasured("Hiding or showing the line of a chart element with a look set")
        look = self.written()
        look.line_hidden = look.line_full = look.effects = True
        if _windowed(self.key):
            look.fill_solid = True
        elif not self.key.startswith("grid"):
            look.no_fill = True

    @member
    def ForeColor(self) -> object:
        return ElementColor(self, fore=True)

    @member
    def Weight(self) -> object:
        look = self.look()
        if self.absent() or self.reads_hidden():
            return VBASingle(_MIXED)
        if look.line_width is not None:
            return VBASingle(points(look.line_width))
        styled = self.styled()
        if styled is not None:
            return VBASingle(styled[0])
        return VBASingle(1.0 if _grey(self.key) else 0.75)

    @setter("Weight")
    def _weight(self, value: object) -> None:
        wanted = float(VBASingle(float(to_number(value))))
        self.look()
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
        self.look()
        if wanted != 4:
            raise unmeasured("This line dash style on a chart element")
        look = self.change()
        look.line_dash, look.line_cap = _DASHES[wanted]

    @member
    def Style(self) -> object:
        return VBAInt(-2 if self.absent() or self.reads_hidden() else 1, "Long")

    @member
    def Transparency(self) -> object:
        if self.absent() or self.reads_hidden() or not self.touched():
            return VBASingle(_MIXED)
        return VBASingle(0.0)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        """With nothing set of the line it takes none (tests/fixtures/chart_table_look.json)."""
        wanted = float(VBASingle(float(to_number(value))))
        if not 0 <= wanted <= 1:
            raise error(*_INVALID)
        if self.touched():
            raise unmeasured("The transparency of a chart element's line with a look")


class EntryFill(ElementFill):
    vba_type_name = "FillFormat"

    @setter("Visible")
    def _entry_visible(self, value: object) -> None:
        # Hiding it changes nothing.
        to_integer(value, "Long")


class EntryFormat(ElementFormat):
    """The Format of a legend entry: a fill that reads as set and reaches no file, and no line."""

    vba_type_name = "ChartFormat"

    @member
    def Fill(self) -> object:
        return EntryFill(self)

    @member
    def Line(self) -> object:
        raise error(-2147467259, "Automation error\nUnspecified error")

    @member
    def Glow(self) -> object:
        raise error(-2147467259, "Automation error\nUnspecified error")

    @member
    def TextFrame2(self) -> object:
        return TextFrameRead(self, False)


class ElementColor(Created):
    """A chart element's fill or line colour."""

    vba_type_name = "ColorFormat"

    def __init__(self, owner: ElementFill | ElementLine, *, fore: bool) -> None:
        self.owner, self.fore = owner, fore

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.owner.chart.sheet.book.application

    def colourless(self) -> bool:
        owner = self.owner
        if isinstance(owner, ElementLine):
            return owner.absent()
        return not _windowed(owner.key) and not owner.spelled() and not owner.hidden()

    @member
    def RGB(self) -> object:
        owner = self.owner
        if isinstance(owner, ElementFill):
            look = owner.usable()
            if not self.fore:
                return VBAInt(_WHITE, "Long")
            if self.colourless():
                return VBAInt(0, "Long")
            if look.fill_rgb is not None:
                return VBAInt(bgr(look.fill_rgb), "Long")
            if _windowed(owner.key) or owner.hidden():
                return VBAInt(_WHITE, "Long")
            theme = owner.owner.chart.sheet.book.stylesheet.colors.theme
            return VBAInt(bgr(theme[4]), "Long")
        look = owner.look()
        if self.colourless():
            return VBAInt(0, "Long")
        if owner.reads_hidden():
            return VBAInt(_WHITE, "Long")
        styled = owner.styled()
        if look.line_rgb is None and styled is not None:
            return VBAInt(bgr(styled[1]), "Long")
        return VBAInt(_GREY_LONG if look.line_rgb is None else bgr(look.line_rgb), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        wanted = _rgb(int(to_integer(value, "Long")) & 0xFFFFFF)
        owner = self.owner
        if not self.fore:
            raise unmeasured("Setting a chart element's back colour")
        if isinstance(owner, ElementFill):
            owner.usable()
            owner.change().fill_rgb = wanted
        else:
            owner.change().line_rgb = wanted

    @member
    def ObjectThemeColor(self) -> object:
        if not self.fore:
            raise unmeasured("The theme colour of a chart element's back colour")
        if isinstance(self.owner, ElementFill):
            self.owner.usable()
        else:
            # A loaded chart's line is not read yet (tests/fixtures/loaded_chart_looks.json).
            self.owner.look()
        return VBAInt(-2 if self.colourless() else 0, "Long")

    @member
    def Type(self) -> object:
        if not self.fore:
            raise unmeasured("The type of a chart element's back colour")
        if isinstance(self.owner, ElementFill):
            self.owner.usable()
        else:
            # A loaded chart's line is not read yet (tests/fixtures/loaded_chart_looks.json).
            self.owner.look()
        return VBAInt(-2 if self.colourless() else 1, "Long")
