"""Series.Points and Point, as tests/fixtures/series_points.json measured them on each kind of chart.

A point reads what its series gives until something is set of it, which the series then keeps a c:dPt for, in order of
the point's index, with a unique id from the chart's count in the order they were set. A pie's or doughnut's point
takes its own place's accent, where another chart's takes its series'. A change to the series' fill or line takes that
part from its points. On a chart of lines its look reaches its marker too (tests/fixtures/point_line_format.json), and a
marker keeps tests/fixtures/point_markers.json. Delete takes away its series, and a point held across that answers as
tests/fixtures/series_points.json has it."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._chart_format import (
    FILLED,
    SeriesFill,
    SeriesLine,
    accent_colour,
    hidden_fill,
    look_markup,
    plain_line,
    unmeasured,
)
from pyopenvba._xml import escape as escape_attribute, escape_text
from pyopenvba.apps.excel._chart_xml import LABEL_POSITIONS, UNIQUE_END, LabelLook, default_run
from pyopenvba.apps.excel._charts import (
    AUTOMATIC_MARKERS,
    CHART_TYPES,
    MARKER_ALIASES,
    MARKER_AUTOMATIC,
    MARKER_NONE,
    MARKER_SYMBOLS,
    PointLabel,
    PointState,
    marker_look,
    select_element,
    shown_type,
)
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.apps.excel._number_format import to_file
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_integer
from pyopenvba.shapes._values import FillLine

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Series, SeriesState

_PIES = ("pie", "doughnut")
_LINES = ("line", "lineMarkers", "scatter")
#: The kinds of chart whose points' markers were measured: a chart without lines keeps them to read back alone.
_MARKER_KINDS = (*_LINES, *_PIES, "column", "bar", "area")
#: ClearFormats leaves a point this line: black, 2 points and solid (tests/fixtures/series_points.json).
_CLEARED = FillLine(line_rgb="000000", line_width=25400, line_dash="solid")


def _unique_tail(number: int) -> str:
    return ('<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"'
            ' xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
            f'<c16:uniqueId val="{{{number:08X}{UNIQUE_END}}}"/></c:ext>')


def _marker_stroke(point: PointState) -> str:
    """The marker's copy of the point's line, as Excel writes what was set of it."""
    source = point.look if point.marker_line is None else point.marker_line
    if source is None:
        return ""
    dash = "solid" if point.marker_solid and source.line_dash is None else source.line_dash
    return plain_line(FillLine(line_rgb=source.line_rgb, line_width=source.line_width, line_dash=dash,
                               line_cap=source.line_cap))


def _marker_markup(point: PointState, entry: SeriesState, kind: str) -> str:
    """A point's c:marker: where it has a symbol or a size of its own a symbol, as its series' is written, else only its
    colours, which a series drawn without markers does not write. Other charts have nothing to write it in."""
    stroke = "" if kind not in _LINES else _marker_stroke(point)
    if kind not in _LINES or not (point.has_marker() or stroke):
        return ""
    shown = entry.marker if point.marker is None else point.marker
    tail = "" if point.marker_hidden is None else f"<a:extLst>{hidden_fill(point.marker_hidden)}</a:extLst>"
    look = "" if shown == MARKER_NONE else marker_look(point, stroke, tail)
    if point.marker is None and point.marker_size is None:
        return f"<c:marker>{look}</c:marker>" if look else "<c:marker/>"
    symbol = MARKER_SYMBOLS[shown]
    if not symbol:
        return f'<c:marker><c:symbol val="auto"/>{look}</c:marker>'
    if symbol in ("none", "picture"):
        return f'<c:marker><c:symbol val="{symbol}"/>{look}</c:marker>'
    size = point.marker_size or entry.marker_size or 5
    return f'<c:marker><c:symbol val="{symbol}"/><c:size val="{size}"/>{look}</c:marker>'


def points_markup(entry: SeriesState, kind: str, place: int, theme: tuple[str, ...]) -> str:
    """The series' dPt elements, by the points' index."""
    written = ""
    for index, point in sorted(entry.points.items()):
        explosion = point.explosion if kind in _PIES and point.explosion else None
        if point.look is None and point.invert is None and explosion is None and not point.has_marker():
            continue
        invert = ""
        if kind in ("column", "bar"):
            shown = entry.invert_if_negative if point.invert is None else point.invert
            invert = f'<c:invertIfNegative val="{int(shown)}"/>'
        burst = "" if explosion is None else f'<c:explosion val="{explosion}"/>'
        accent = accent_colour(theme, index + 1 if kind in _PIES else place)
        look = "" if point.look is None else look_markup(point.look, kind, accent)
        marker = _marker_markup(point, entry, kind)
        written += (f'<c:dPt><c:idx val="{index}"/>{invert}{marker}<c:bubble3D val="0"/>{burst}{look}'
                    + (f"<c:extLst>{_unique_tail(point.unique)}</c:extLst>" if point.unique >= 0 else "")
                    + "</c:dPt>")
    return written


def _label_body(point: PointState, entry: SeriesState) -> str:
    """What a point's dLbl says after its idx: its text, number format, look, place and flags, in the file's order."""
    own = point.label_own
    tx = ""
    if own.text is not None:
        run = '<a:endParaRPr lang="en-US"/>'
        if own.text:
            run = f'<a:r><a:rPr lang="en-US"/><a:t>{escape_text(own.text)}</a:t></a:r>'
        tx = f"<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>{run}</a:p></c:rich></c:tx>"
    number = "" if own.number_format is None else (
        f'<c:numFmt formatCode="{escape_attribute(to_file(own.number_format))}" sourceLinked="0"/>')
    look = ""
    if own.has_look():
        fill = "<a:noFill/>"
        if own.fill_rgb is not None:
            fill = f'<a:solidFill><a:srgbClr val="{own.fill_rgb}"/></a:solidFill>'
        effects = "<a:effectLst/>"
        if own.shadow:
            effects = ('<a:effectLst><a:outerShdw dist="35921" dir="2700000" algn="br"><a:srgbClr val="000000"/>'
                       "</a:outerShdw></a:effectLst>")
        turn = {-4171: 90, -4170: -90, -4128: 0}.get(own.orientation, own.orientation)
        rotated = f' rot="{-turn * 60000}" vert="horz"' if own.orientation != -4128 else ""
        paragraph = "" if own.horizontal == -4108 else f' algn="{ {-4131: "l", -4152: "r"}[own.horizontal] }"'
        if own.reading == -5004:
            paragraph += ' rtl="1"'
        look = (f"<c:spPr>{fill}<a:ln><a:noFill/></a:ln>{effects}</c:spPr><c:txPr>"
                f'<a:bodyPr{rotated} wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr">'
                f"<a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr{paragraph}>{default_run(own.font)}</a:pPr>"
                '<a:endParaRPr lang="en-US"/></a:p></c:txPr>')
    place = "" if own.position is None else f'<c:dLblPos val="{LABEL_POSITIONS[own.position]}"/>'
    series = entry.label_look
    base = {"value": series.value, "category": series.category, "series": series.series,
            "legend_key": series.legend_key, "percent": series.percent, "bubble": series.bubble}
    if not entry.data_labels:
        base = {"value": True, "category": False, "series": False, "legend_key": False, "percent": False,
                "bubble": False}
    flags = {**base, **own.shows}
    shows = "".join(f'<c:{tag} val="{int(flags[field])}"/>' for tag, field in (
        ("showLegendKey", "legend_key"), ("showVal", "value"), ("showCatName", "category"), ("showSerName", "series"),
        ("showPercent", "percent"), ("showBubbleSize", "bubble")))
    separator = "" if own.separator is None else f"<c:separator>{escape_text(own.separator)}</c:separator>"
    return f"{tx}{number}{look}{place}{shows}{separator}"


def _range_extension(own: PointLabel, deleted: bool) -> str:
    if deleted:
        return "/>"
    if own.show_range:
        return '><c15:dlblFieldTable/><c15:showDataLabelsRange val="1"/></c:ext>'
    return '><c15:showDataLabelsRange val="0"/></c:ext>' if own.text is not None else "/>"


def point_labels_markup(entry: SeriesState) -> str:
    """The series' dLbl elements: a point's own label, or where the series has labels, one it took away."""
    written = ""
    for index, point in sorted(entry.points.items()):
        if point.label_unique == -1 or (point.label is None and point.label_own.empty()):
            continue
        if point.label is False and not entry.data_labels:
            continue
        body = '<c:delete val="1"/>' if point.label is False else _label_body(point, entry)
        own = point.label_own
        # On a series with labels, one with a look of its own has no such extension; a typed one says it shows no range.
        extension = "" if entry.data_labels and point.label is not False and own.has_look() else (
            '<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"'
            ' xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"'
            + _range_extension(own, point.label is False))
        written += (f'<c:dLbl><c:idx val="{index}"/>{body}<c:extLst>{extension}'
                    f"{_unique_tail(point.label_unique) if point.label_unique >= 0 else ''}</c:extLst></c:dLbl>")
    return written


class Points(VBACollection, ItemDefault, Created):
    vba_type_name = "Points"

    def __init__(self, series: Series) -> None:
        self.series = series

    @member
    def Parent(self) -> object:
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    def vba_items(self) -> list[object]:
        self.series.position()
        return [Point(self.series, index) for index in range(len(self.series.values()))]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            raise error(13, "Type mismatch")
        place = int(to_integer(index, "Long"))
        if not 1 <= place <= len(items):
            raise error(1004, "Points method of Series class failed")
        return items[place - 1]


class Point(Created):
    vba_type_name = "Point"

    def __init__(self, series: Series, index: int) -> None:
        self.series, self.index = series, index

    def kind(self) -> str:
        self.series.position()
        state = self.series.chart.state
        if state.original:
            raise unmeasured("A point of an imported chart")
        return CHART_TYPES.get(shown_type(state), "")

    def own(self) -> PointState | None:
        return self.series.entry.points.get(self.series.point_key(self.index))

    def held(self) -> PointState:
        """The point's state, made if it had none."""
        self.series.chart.changed()
        point = self.series.entry.points.setdefault(self.series.point_key(self.index), PointState())
        if point.drawn_at < 0:
            point.drawn_at = self.index
        return point

    def gone(self) -> bool:
        """Whether its series has been taken away."""
        self.series.chart.admit_read()
        state = self.series.chart.state
        return not any(entry is self.series.entry for entry in [*state.series, *state.full_series])

    def numbered(self, point: PointState) -> None:
        """Its dPt's unique id, from the chart's count, once it has a dPt."""
        if point.unique < 0:
            state = self.series.chart.state
            point.unique, state.next_unique = state.next_unique, state.next_unique + 1

    @member
    def Parent(self) -> object:
        if self.gone():
            raise error(-2147417848, "Automation error\nThe object invoked has disconnected from its clients.")
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Name(self) -> object:
        return f"S{self.series.position()}P{self.index + 1}"

    @member
    def Format(self) -> object:
        if self.gone():
            raise error(-2147467259, "Automation error\nUnspecified error")
        self.kind()
        return PointFormat(self)

    @member
    def Border(self) -> object:
        from pyopenvba.apps.excel._chart_borders import SERIES, ChartBorder

        if self.gone():
            raise error(-2147467259, "Automation error\nUnspecified error ")
        self.kind()
        return ChartBorder(PointFormat(self).Line(), SERIES, self, self.series.chart.sheet.book.stylesheet.colors)

    @member
    def Explosion(self) -> object:
        self.kind()
        point = self.own()
        own = None if point is None else point.explosion
        return VBAInt(self.series.entry.explosion if own is None else own, "Long")

    @setter("Explosion")
    def _explosion(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        kind = self.kind()
        if wanted < 0:
            raise error(1004, "Unable to set the Explosion property of the Point class")
        if wanted == 0 and self.own() is None:
            return
        point = self.held()
        point.explosion = wanted
        if kind in _PIES and wanted:
            self.numbered(point)

    @member
    def InvertIfNegative(self) -> object:
        self.kind()
        point = self.own()
        own = None if point is None else point.invert
        return self.series.entry.invert_if_negative if own is None else own

    @setter("InvertIfNegative")
    def _invert(self, value: object) -> None:
        wanted = to_bool(value)
        if self.kind() not in FILLED:
            raise unmeasured("InvertIfNegative on a point of this kind of chart")
        point = self.held()
        point.invert = wanted
        # It makes the dPt anew, with the next id (tests/fixtures/series_points.json).
        point.unique = -1
        self.numbered(point)

    @member
    def HasDataLabel(self) -> object:
        if self.gone():
            return False
        self.kind()
        point = self.own()
        entry = self.series.entry
        if point is not None and point.label is False:
            return False
        if point is not None and (point.label is True or not point.label_own.empty()):
            return True
        return entry.data_labels and entry.label_look.shown()

    @setter("HasDataLabel")
    def _has_data_label(self, value: object) -> None:
        wanted = to_bool(value)
        self.kind()
        entry = self.series.entry
        point = self.held()
        if wanted and entry.data_labels:
            # Back to what its series' labels give, whatever it had of its own.
            point.label, point.label_own = None, PointLabel()
            entry.leader_lines = True
            return
        if wanted:
            # Labels on some points alone: the series' own show nothing.
            entry.label_look = LabelLook(value=False)
            point.label_own = PointLabel()
        point.label = wanted
        if point.label_unique < 0 and (wanted or entry.data_labels):
            state = self.series.chart.state
            if point.unique >= 0:
                # A point with a dPt shares its id with its dLbl (tests/fixtures/chart_copies.json).
                point.label_unique = point.unique
            else:
                point.label_unique, state.next_unique = state.next_unique, state.next_unique + 1

    @method
    def ApplyDataLabels(self, Type: object = MISSING, LegendKey: object = MISSING, AutoText: object = MISSING,
                        HasLeaderLines: object = MISSING, ShowSeriesName: object = MISSING,
                        ShowCategoryName: object = MISSING, ShowValue: object = MISSING,
                        ShowPercentage: object = MISSING, ShowBubbleSize: object = MISSING,
                        Separator: object = MISSING) -> object:
        from pyopenvba.apps.excel._chart_data_labels import applied_look, apply_to_point

        del ShowBubbleSize
        state = self.series.chart.state
        wanted = 2 if Type is MISSING else int(to_integer(Type, "Long"))
        if wanted in (3, 5) and state.chart_type != 5:
            # It takes the label away and then fails.
            apply_to_point(self, LabelLook(value=False), MISSING)
        look = applied_look(state.chart_type, Type, LegendKey, AutoText, ShowSeriesName, ShowCategoryName, ShowValue,
                            ShowPercentage, Separator)
        apply_to_point(self, look, HasLeaderLines)
        return True

    @member
    def DataLabel(self) -> object:
        from pyopenvba.apps.excel._chart_data_labels import DataLabel

        self.kind()
        if not self.series.has_labels():
            raise error(-2147467259, "Automation error\nUnspecified error")
        return DataLabel(self.series, self.index + 1, self.series.entry.label_generation)

    # -- its marker (tests/fixtures/point_markers.json)

    def marker_shown(self) -> int:
        point = self.own()
        self.kind()
        return self.series.entry.marker if point is None or point.marker is None else point.marker

    def marker_held(self) -> PointState:
        """The point's state, for a marker setting, which needs a kind of chart it was measured on."""
        if self.kind() not in _MARKER_KINDS:
            raise unmeasured("The marker of a point on this kind of chart")
        return self.held()

    @member
    def MarkerStyle(self) -> object:
        return VBAInt(self.marker_shown(), "Long")

    @setter("MarkerStyle")
    def _marker_style(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        wanted = MARKER_ALIASES.get(wanted, wanted)
        if wanted not in MARKER_SYMBOLS:
            raise error(1004, "Unable to set the MarkerStyle property of the Point class")
        point = self.marker_held()
        self._no_line_look(point)
        entry = self.series.entry
        if wanted == entry.marker and (point.marker is None or wanted == MARKER_AUTOMATIC):
            point.marker = None
            return
        point.marker = wanted
        self.numbered(point)

    @member
    def MarkerSize(self) -> object:
        point = self.own()
        if point is not None and point.marker_size is not None:
            return VBAInt(point.marker_size, "Long")
        automatic = self.marker_shown() == MARKER_AUTOMATIC and self.kind() in _LINES
        return VBAInt(self.series.entry.marker_size or (9 if automatic else 5), "Long")

    @setter("MarkerSize")
    def _marker_size(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if not 2 <= wanted <= 72:
            raise error(-2147467259, "Automation error\nUnspecified error")
        point = self.marker_held()
        self._no_line_look(point)
        entry = self.series.entry
        if self.kind() in _LINES and self.marker_shown() == MARKER_AUTOMATIC:
            # A size fixes the marker an automatic one shows, as it does for a series.
            identity = entry.series_id if entry.series_id >= 0 else self.series.position() - 1
            if identity >= len(AUTOMATIC_MARKERS):
                raise unmeasured("The automatic marker of a series past the ninth")
            point.marker = AUTOMATIC_MARKERS[identity]
        point.marker_size = wanted
        self.numbered(point)

    def _no_line_look(self, point: PointState) -> None:
        if self.kind() in _LINES and point.look is not None and _marker_stroke(point):
            raise unmeasured("A marker setting on a point whose line was set")

    def _coloured(self, field_name: str, wanted: str | None) -> None:
        point = self.marker_held()
        setattr(point, field_name, wanted)
        if field_name == "marker_back":
            point.marker_hidden = None
        # A background brings out the marker of a point drawn without one, even an automatic background.
        if field_name == "marker_back" and self.marker_shown() == MARKER_NONE:
            point.marker = MARKER_AUTOMATIC
        if point.has_marker():
            self.numbered(point)

    def _marker_colour(self, field_name: str) -> str | None:
        point = self.own()
        found: str | None = None if point is None else getattr(point, field_name)
        if found is not None:
            return found
        entry = self.series.entry
        if field_name == "marker_fore" and self.kind() in _LINES:
            # A line's colour is its marker's too, but for what ClearFormats took from the line.
            source = None if point is None else (point.look if point.marker_line is None else point.marker_line)
            if source is not None and source.line_rgb is not None:
                return source.line_rgb
            if entry.marker_fore is None and entry.format_look is not None:
                return entry.format_look.line_rgb
        return getattr(entry, field_name)

    def _read_colour(self, field_name: str) -> object:
        self.kind()
        found = self._marker_colour(field_name)
        if found is None or found == "none":
            return VBAInt(-1 if found is None else -2, "Long")
        return VBAInt(bgr(found), "Long")

    def _set_colour(self, field_name: str, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted < -2:
            raise unmeasured(f"Marker colour {wanted}")
        if wanted < 0:
            self._coloured(field_name, None if wanted == -1 else "none")
            return
        # Excel keeps the low 24 bits: 16777216 is black.
        blue, green, red = (wanted >> 16) & 255, (wanted >> 8) & 255, wanted & 255
        self._coloured(field_name, f"{red:02X}{green:02X}{blue:02X}")

    def _read_index(self, field_name: str) -> object:
        self.kind()
        found = self._marker_colour(field_name)
        if found is None or found == "none":
            # An automatic marker colour answers slot 2.
            return VBAInt(2 if found is None else -4142, "Long")
        return VBAInt(self.series.chart.sheet.book.stylesheet.colors.nearest(found), "Long")

    def _set_index(self, field_name: str, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted in (-4105, -4142):
            self._coloured(field_name, None if wanted == -4105 else "none")
            return
        palette = self.series.chart.sheet.book.stylesheet.colors.indexed
        if wanted == 57:
            # One past the 56 slots, Excel writes black: the system foreground.
            self._coloured(field_name, "000000")
            return
        if wanted > 57:
            raise error(-2147467259, "Automation error\nUnspecified error")
        if wanted < 1 or wanted + 7 >= len(palette):
            raise unmeasured(f"Marker colour index {wanted}")
        self._coloured(field_name, palette[wanted + 7])

    def _echoed(self) -> FillLine | None:
        """The look an area chart's point echoes in marker colours it has none of (tests/fixtures/series_points.json).
        """
        point = self.own()
        return point.look if self.kind() == "area" and point is not None and point.marker_line is None else None

    @member
    def MarkerBackgroundColor(self) -> object:
        look = self._echoed()
        if self._marker_colour("marker_back") is not None or look is None or not (
                look.fill_rgb or look.fill_hidden or look.fill_solid):
            return self._read_colour("marker_back")
        if look.fill_hidden:
            return VBAInt(-2, "Long")
        return PointFill(PointFormat(self)).vba_get("ForeColor").vba_get("RGB")  # pyright: ignore

    @setter("MarkerBackgroundColor")
    def _marker_background_color(self, value: object) -> None:
        self._set_colour("marker_back", value)

    @member
    def MarkerForegroundColor(self) -> object:
        look = self._echoed()
        if self._marker_colour("marker_fore") is not None or look is None or not (look.line_rgb or look.line_hidden):
            return self._read_colour("marker_fore")
        return VBAInt(-2 if look.line_hidden else bgr(look.line_rgb or "000000"), "Long")

    @setter("MarkerForegroundColor")
    def _marker_foreground_color(self, value: object) -> None:
        self._set_colour("marker_fore", value)

    @member
    def MarkerBackgroundColorIndex(self) -> object:
        return self._read_index("marker_back")

    @setter("MarkerBackgroundColorIndex")
    def _marker_background_color_index(self, value: object) -> None:
        self._set_index("marker_back", value)

    @member
    def MarkerForegroundColorIndex(self) -> object:
        return self._read_index("marker_fore")

    @setter("MarkerForegroundColorIndex")
    def _marker_foreground_color_index(self, value: object) -> None:
        self._set_index("marker_fore", value)

    @member
    def Shadow(self) -> object:
        self.kind()
        return False

    @member
    def Has3DEffect(self) -> object:
        self.kind()
        return False

    @member
    def SecondaryPlot(self) -> object:
        if self.kind() in _PIES:
            return False
        raise error(1004, "Unable to get the SecondaryPlot property of the Point class")

    @member
    def IsTotal(self) -> object:
        raise error(445, "Object doesn't support this action")

    @method
    def ClearFormats(self) -> object:
        """Not what its name says: it takes the point's fill and leaves it a black line of 2 points, or on a chart of
        lines one of its series' colour, and a marker keeps what it had."""
        kind = self.kind()
        if kind not in _MARKER_KINDS:
            raise VBAUnsupportedError("ClearFormats on a point of this kind of chart is not implemented")
        if self.series.entry.marker_line is not None:
            raise VBAUnsupportedError("ClearFormats on a point of a series whose line colour reached its marker is not "
                                      "implemented")
        point = self.held()
        colour = _CLEARED.line_rgb
        if kind in _LINES:
            if point.marker_line is None:
                old = point.look or FillLine()
                point.marker_line = FillLine(line_rgb=old.line_rgb, line_width=old.line_width, line_dash=old.line_dash,
                                             line_cap=old.line_cap)
            series_look = self.series.entry.format_look
            colour = (series_look.line_rgb if series_look is not None else None) or colour
        elif kind == "area":
            # An area point's marker colours echo its line, but not this one.
            point.marker_line = FillLine()
        point.look = FillLine(line_rgb=colour, line_width=_CLEARED.line_width, line_dash=_CLEARED.line_dash)
        self.numbered(point)
        return EMPTY

    @method
    def Select(self) -> object:
        return select_element(self.series.chart, self, self.kind() != "area" and self.series.plotted())

    @method
    def Delete(self) -> object:
        """Not what its name says: it takes away the point's series."""
        self.kind()
        return self.series.vba_get("Delete", [])


class PointFormat(Created):
    vba_type_name = "ChartFormat"

    def __init__(self, point: Point) -> None:
        self.point = point
        self.series = point.series

    @member
    def Parent(self) -> object:
        return self.point

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Fill(self) -> object:
        return PointFill(self)  # pyright: ignore[reportArgumentType]

    @member
    def Line(self) -> object:
        return PointLine(self)  # pyright: ignore[reportArgumentType]


class _OfPoint:
    """What makes a series' fill or line a point's: its own look over its series', and a c:dPt to write it in."""

    owner: PointFormat

    def own_look(self) -> FillLine:
        point = self.owner.point.own()
        return (point.look if point is not None else None) or FillLine()

    def look(self) -> FillLine:
        own = self.own_look()
        series = self.owner.series.entry.format_look or FillLine()
        fill = own if (own.fill_rgb is not None or own.fill_solid or own.fill_hidden) else series
        line = own if (own.line_rgb is not None or own.line_width is not None or own.line_dash is not None
                       or own.line_full or own.line_off) else series
        return FillLine(fill_rgb=fill.fill_rgb, fill_alpha=fill.fill_alpha, fill_solid=fill.fill_solid,
                        fill_hidden=fill.fill_hidden, line_rgb=line.line_rgb, line_width=line.line_width,
                        line_dash=line.line_dash, line_cap=line.line_cap, line_full=line.line_full,
                        line_hidden=line.line_hidden, effects=line.effects, line_off=line.line_off)

    def admit(self) -> None:
        if self.owner.point.kind() not in (*FILLED, *_LINES):
            raise unmeasured("Setting the look of a point of this kind of chart")

    def change(self) -> FillLine:
        self.admit()
        point = self.owner.point.held()
        self.owner.point.numbered(point)
        if point.look is None:
            point.look = FillLine()
        return point.look

    def accent(self) -> int:
        series = self.owner.series
        identity = series.entry.series_id
        place = identity + 1 if identity >= 0 else series.position()
        if self.owner.point.kind() in _PIES:
            place = self.owner.point.index + 1
        return bgr(accent_colour(series.chart.sheet.book.stylesheet.colors.theme, place))


class PointFill(_OfPoint, SeriesFill):
    vba_type_name = "FillFormat"

    def marker_took(self, wanted: str) -> bool:
        if self.owner.point.kind() not in _LINES:
            return False
        point = self.owner.point.held()
        point.marker_back, point.marker_hidden = wanted, None
        self.owner.point.numbered(point)
        return True

    @setter("Visible")
    def _visible(self, value: object) -> None:
        if self.owner.point.kind() not in _LINES:
            SeriesFill._visible(self, value)  # pyright: ignore[reportPrivateUsage]
            return
        wanted = int(to_integer(value, "Long"))
        own = self.owner.point.own()
        if wanted == -1 and (own is None or own.marker_hidden is None):
            return
        if wanted != 0:
            raise unmeasured("Showing a marker's hidden fill, or Visible set to anything but msoTrue or msoFalse")
        point = self.owner.point.held()
        series = self.owner.series
        accent = accent_colour(series.chart.sheet.book.stylesheet.colors.theme, series.position())
        kept = point.marker_hidden or (point.marker_back if point.marker_back not in (None, "none") else None)
        point.marker_back, point.marker_hidden = "none", kept or accent
        self.owner.point.numbered(point)

    @setter("Transparency")
    def _transparency(self, value: object) -> None:
        point = self.owner.point
        if point.kind() not in _LINES:
            SeriesFill._transparency(self, value)  # pyright: ignore[reportPrivateUsage]
            return
        own = point.own()
        if own is not None and own.marker_back is not None:
            raise unmeasured("The transparency of a marker with a colour")
        # A marker with no colour written takes none.


class PointLine(_OfPoint, SeriesLine):
    vba_type_name = "LineFormat"

    def admit(self) -> None:
        if self.owner.point.kind() == "scatter":
            raise unmeasured("Setting the line of a point of a scatter chart")
        super().admit()

    def reads_hidden(self) -> bool:
        # ClearFormats gives a scatter chart's point a real line, whatever its series'.
        return not self.own_look().line_rgb and super().reads_hidden()

    def settable(self, what: str) -> None:
        self.admit()
        own = self.owner.point.own()
        if own is not None and own.marker_line is not None:
            raise unmeasured(f"The line {what} of a point after ClearFormats")
        if self.owner.point.kind() not in _LINES:
            super().settable(what)

    def change(self) -> FillLine:
        look = super().change()
        point = self.owner.point.held()
        if self.owner.point.kind() in _LINES and point.marker_fore is not None:
            point.marker_solid = point.marker_fore != "none"
            point.marker_fore = None
        return look

    @setter("Visible")
    def _visible(self, value: object) -> None:
        if self.owner.point.kind() not in _LINES:
            SeriesLine._visible(self, value)  # pyright: ignore[reportPrivateUsage]
            return
        self.admit()
        if int(to_integer(value, "Long")) != -1:
            raise unmeasured("Hiding the line of a point of a chart of lines")
