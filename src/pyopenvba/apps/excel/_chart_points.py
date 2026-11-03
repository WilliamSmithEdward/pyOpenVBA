"""Series.Points and Point, as tests/fixtures/series_points.json measured them on each kind of chart.

A point reads what its series gives until something is set of it, which the series then keeps a c:dPt for, in order of
the point's index, with a unique id from the chart's count in the order they were set. A pie's or doughnut's point
takes its own place's accent, where another chart's takes its series'. A change to the series' fill or line takes that
part from its points. Setting a point's look on a line chart moves reads onto its marker members, so that is refused,
and so are marker settings and Delete, which takes away a series."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._chart_format import FILLED, SeriesFill, SeriesLine, accent_colour, look_markup, unmeasured
from pyopenvba._xml import escape as escape_attribute, escape_text
from pyopenvba.apps.excel._chart_xml import LABEL_POSITIONS, UNIQUE_END, LabelLook, default_run
from pyopenvba.apps.excel._charts import CHART_TYPES, PointLabel, PointState, shown_type
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.apps.excel._number_format import to_file
from pyopenvba.apps.excel._styles import bgr
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, VBAInt, error, to_bool, to_integer
from pyopenvba.shapes._values import FillLine

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Series, SeriesState

_PIES = ("pie", "doughnut")
#: ClearFormats leaves a point this line: black, 2 points and solid (tests/fixtures/series_points.json).
_CLEARED = FillLine(line_rgb="000000", line_width=25400, line_dash="solid")


def _unique_tail(number: int) -> str:
    return ('<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"'
            ' xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
            f'<c16:uniqueId val="{{{number:08X}{UNIQUE_END}}}"/></c:ext>')


def points_markup(entry: SeriesState, kind: str, place: int, theme: tuple[str, ...]) -> str:
    """The series' dPt elements, by the points' index."""
    written = ""
    for index, point in sorted(entry.points.items()):
        explosion = point.explosion if kind in _PIES and point.explosion else None
        if point.look is None and point.invert is None and explosion is None:
            continue
        invert = ""
        if kind in ("column", "bar"):
            shown = entry.invert_if_negative if point.invert is None else point.invert
            invert = f'<c:invertIfNegative val="{int(shown)}"/>'
        burst = "" if explosion is None else f'<c:explosion val="{explosion}"/>'
        accent = accent_colour(theme, index + 1 if kind in _PIES else place)
        look = "" if point.look is None else look_markup(point.look, kind, accent)
        written += (f'<c:dPt><c:idx val="{index}"/>{invert}<c:bubble3D val="0"/>{burst}{look}'
                    f"<c:extLst>{_unique_tail(point.unique)}</c:extLst></c:dPt>")
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
        if point.label_unique < 0 or (point.label is None and point.label_own.empty()):
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
                    f"{_unique_tail(point.label_unique)}</c:extLst></c:dLbl>")
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
        return self.series.entry.points.get(self.index)

    def held(self) -> PointState:
        """The point's state, made if it had none."""
        self.series.chart.changed()
        return self.series.entry.points.setdefault(self.index, PointState())

    def numbered(self, point: PointState) -> None:
        """Its dPt's unique id, from the chart's count, once it has a dPt."""
        if point.unique < 0:
            state = self.series.chart.state
            point.unique, state.next_unique = state.next_unique, state.next_unique + 1

    @member
    def Parent(self) -> object:
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Name(self) -> object:
        return f"S{self.series.position()}P{self.index + 1}"

    @member
    def Format(self) -> object:
        self.kind()
        return PointFormat(self)

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
            point.label_unique, state.next_unique = state.next_unique, state.next_unique + 1

    @member
    def DataLabel(self) -> object:
        from pyopenvba.apps.excel._chart_data_labels import DataLabel

        self.kind()
        if not self.series.has_labels():
            raise error(-2147467259, "Automation error\nUnspecified error")
        return DataLabel(self.series, self.index + 1, self.series.entry.label_generation)

    def marked(self) -> bool:
        return self.kind() in ("lineMarkers", "scatter")

    @member
    def MarkerStyle(self) -> object:
        return VBAInt(-4105 if self.marked() else -4142, "Long")

    @member
    def MarkerSize(self) -> object:
        return VBAInt(9 if self.marked() else 5, "Long")

    def plain_marker(self) -> FillLine | None:
        """The look an area chart's point echoes in its marker colours (tests/fixtures/series_points.json), or None."""
        kind = self.kind()
        entry = self.series.entry
        if entry.marker_fore is not None or entry.marker_back is not None:
            raise unmeasured("The marker colour of a point of a series with marker colours")
        point = self.own()
        return point.look if kind == "area" and point is not None else None

    @member
    def MarkerBackgroundColor(self) -> object:
        look = self.plain_marker()
        if look is None or not (look.fill_rgb or look.fill_hidden or look.fill_solid):
            return VBAInt(-1, "Long")
        if look.fill_hidden:
            return VBAInt(-2, "Long")
        return PointFill(PointFormat(self)).vba_get("ForeColor").vba_get("RGB")  # pyright: ignore

    @member
    def MarkerForegroundColor(self) -> object:
        look = self.plain_marker()
        if look is None or not (look.line_rgb or look.line_hidden):
            return VBAInt(-1, "Long")
        return VBAInt(-2 if look.line_hidden else bgr(look.line_rgb or "000000"), "Long")

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
        self.kind()
        raise error(445, "Object doesn't support this action")

    @method
    def ClearFormats(self) -> object:
        """Not what its name says: it takes the point's fill and leaves it a black line of 2 points."""
        if self.kind() not in ("column", "bar"):
            raise VBAUnsupportedError("ClearFormats on a point of this kind of chart is not implemented")
        point = self.held()
        point.look = FillLine(line_rgb=_CLEARED.line_rgb, line_width=_CLEARED.line_width, line_dash=_CLEARED.line_dash)
        self.numbered(point)
        return EMPTY


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
                       or own.line_full) else series
        return FillLine(fill_rgb=fill.fill_rgb, fill_alpha=fill.fill_alpha, fill_solid=fill.fill_solid,
                        fill_hidden=fill.fill_hidden, line_rgb=line.line_rgb, line_width=line.line_width,
                        line_dash=line.line_dash, line_cap=line.line_cap, line_full=line.line_full,
                        line_hidden=line.line_hidden, effects=line.effects)

    def admit(self) -> None:
        if self.owner.point.kind() not in FILLED:
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
        place = self.owner.point.index + 1 if self.owner.point.kind() in _PIES else series.position()
        return bgr(accent_colour(series.chart.sheet.book.stylesheet.colors.theme, place))


class PointFill(_OfPoint, SeriesFill):
    vba_type_name = "FillFormat"


class PointLine(_OfPoint, SeriesLine):
    vba_type_name = "LineFormat"
