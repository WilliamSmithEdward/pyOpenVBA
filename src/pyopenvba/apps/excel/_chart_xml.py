"""Charts written as Excel writes one made through ``ChartObjects.Add``.

Such a chart carries no formatting of its own: no colours, lines or text
properties, the legend on the right and the gaps and tick marks of Excel's
classic default, and no style or colour parts beside it. Every piece below
is the markup Excel wrote for each kind, and for each edit the model makes
(tests/fixtures/chart_markup.json).

Each series keeps, beside its references, a cache of the values they read,
as Excel writes one, so a reader that does not calculate still sees the
chart's data.
"""

from __future__ import annotations

import random
import math
import uuid
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Literal

from pyopenvba._xml import escape as escape_attribute, escape_text
from pyopenvba.shapes._xlsx import SheetGrid

# Adapted from William's pyOfficeEditor at 4723383e5d3039763e64ecd2a3ee092786c91bb5.
def format_number(value: int | float) -> str:
    if not math.isfinite(value):
        raise ValueError('Non-finite chart data cannot be written')
    if isinstance(value, int) or value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


def corner_markup(which: str, grid: SheetGrid, across: float, down: float) -> str:
    column, column_offset = grid.column_at(across)
    row, row_offset = grid.row_at(down)
    return (f'<xdr:{which}><xdr:col>{column}</xdr:col><xdr:colOff>{column_offset}</xdr:colOff>'
            f'<xdr:row>{row}</xdr:row><xdr:rowOff>{row_offset}</xdr:rowOff></xdr:{which}>')

#: The kinds of chart Excel's Insert Chart makes that can be written here.
ChartKind = Literal["column", "bar", "line", "lineMarkers", "pie", "doughnut", "scatter", "area"]
CHART_KINDS: tuple[ChartKind, ...] = ("column", "bar", "line", "lineMarkers", "pie", "doughnut", "scatter", "area")

CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"


@dataclass(frozen=True)
class SeriesData:
    """One series to write: its references, and the values they read now."""

    values: str
    value_cache: Sequence[float | None] = ()
    #: The number format of the values' first cell, which the cache carries.
    value_format: str = "General"
    name: str | None = None
    name_cache: str | None = None
    categories: str | None = None
    #: The categories as they show, or as numbers when every one is a number.
    category_cache: Sequence[str | float | None] = ()
    category_format: str = "General"
    numeric_categories: bool = False
    series_id: int | None = None
    #: The first group of its unique id, when that is not its series_id.
    unique_id: int | None = None
    #: In a line or scatter chart: the marker element, none for an automatic marker, and whether the line is smoothed.
    marker: str = ""
    smooth: bool = False
    #: In a column or bar chart, InvertIfNegative; in a pie, Explosion (tests/fixtures/chart_series_more.json).
    invert: bool = False
    explosion: int = 0
    #: HasDataLabels and HasLeaderLines (tests/fixtures/chart_data_labels.json).
    labels: bool = False
    leader_lines: bool = True


#: Chart.HasDataTable, as Excel writes the table (tests/fixtures/chart_series_more.json).
DATA_TABLE = ('<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/>'
              '<c:showKeys val="1"/></c:dTable>')


# -- the pieces every kind shares ----------------------------------------

_HEAD = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">'
    '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>'
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">'
    '<c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>'
    "<c:chart>"
)

#: DisplayBlanksAs as the file spells it; Excel writes no element for any other value.
BLANKS = {1: "gap", 2: "zero", 3: "span"}

_NOT_AVAILABLE_AS_BLANK = (
    '<c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}"'
    ' xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">'
    '<c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16>'
    "</c:ext></c:extLst>"
)


def _display(blanks: int, visible_only: bool, over_maximum: bool, not_available_as_blank: bool) -> str:
    """The chart's display options: which cells it plots, and how blanks and #N/A show."""
    spelled = BLANKS.get(blanks)
    return (
        f'<c:plotVisOnly val="{int(visible_only)}"/>'
        + (f'<c:dispBlanksAs val="{spelled}"/>' if spelled else "")
        + f'<c:showDLblsOverMax val="{int(over_maximum)}"/>'
        + (_NOT_AVAILABLE_AS_BLANK if not_available_as_blank else "")
    )


_TAIL = (
    "</c:chart>"
    '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3"'
    ' footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>'
)

_SHADOW = ('<c:spPr><a:effectLst><a:outerShdw dist="35921" dir="2700000" algn="br"><a:srgbClr val="000000"/>'
           '</a:outerShdw></a:effectLst></c:spPr>')


def shadow_markup(shadow: bool | None) -> str:
    """A title's or legend's Shadow: Excel's outer shadow, an empty effect list once turned off, nothing if never
    set."""
    return '' if shadow is None else _SHADOW if shadow else '<c:spPr><a:effectLst/></c:spPr>'


def legend_markup(position: str = 'r', *, overlay: bool = False, shadow: bool | None = None) -> str:
    return (f'<c:legend><c:legendPos val="{position}"/><c:overlay val="{int(overlay)}"/>{shadow_markup(shadow)}'
            '</c:legend>')

_LABELS = (
    '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
    '<c:showPercent val="0"/><c:showBubbleSize val="0"/>{extra}</c:dLbls>'
)

_NUMBER_FORMAT = '<c:numFmt formatCode="General" sourceLinked="1"/>'


def title_markup(text: str | None) -> str:
    """A chart or axis title: Excel's automatic one with no text, or the text typed in."""
    typed = ""
    if text is not None:
        typed = (
            "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr>"
            f'<a:r><a:rPr lang="en-US"/><a:t>{escape_text(text)}</a:t></a:r></a:p></c:rich></c:tx>'
        )
    return f'<c:title>{typed}<c:overlay val="0"/></c:title>'


def _category_axis(own: int, other: int, *, position: str, formatted: bool) -> str:
    return (
        f'<c:catAx><c:axId val="{own}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="{position}"/>{_NUMBER_FORMAT if formatted else ""}<c:majorTickMark val="out"/>'
        '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="{other}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
        '<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
    )


def _value_axis(own: int, other: int, *, position: str, between: str, formatted: bool, gridlines: bool) -> str:
    return (
        f'<c:valAx><c:axId val="{own}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="{position}"/>{"<c:majorGridlines/>" if gridlines else ""}'
        f'{_NUMBER_FORMAT if formatted else ""}<c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
        f'<c:tickLblPos val="nextTo"/><c:crossAx val="{other}"/><c:crosses val="autoZero"/>'
        f'<c:crossBetween val="{between}"/></c:valAx>'
    )


# -- series ---------------------------------------------------------------


def _string_cache(texts: Sequence[str | float | None]) -> str:
    points = "".join(
        f'<c:pt idx="{index}"><c:v>{escape_text(str(text))}</c:v></c:pt>'
        for index, text in enumerate(texts)
        if text is not None and text != ""
    )
    return f'<c:strCache><c:ptCount val="{len(texts)}"/>{points}</c:strCache>'


def _number_cache(values: Sequence[str | float | None], code: str) -> str:
    points = "".join(
        f'<c:pt idx="{index}"><c:v>{format_number(value)}</c:v></c:pt>'
        for index, value in enumerate(values)
        if isinstance(value, (int, float)) and not isinstance(value, bool)
    )
    return (
        f"<c:numCache><c:formatCode>{escape_text(code)}</c:formatCode>"
        f'<c:ptCount val="{len(values)}"/>{points}</c:numCache>'
    )


def cache_markup(values: Sequence[str | float | None], *, numeric: bool, code: str = 'General') -> str:
    return _number_cache(values, code) if numeric else _string_cache(values)


def _reference(element: str, formula: str, cache: str, *, numbers: bool) -> str:
    kind = "numRef" if numbers else "strRef"
    return f"<c:{element}><c:{kind}><c:f>{escape_text(formula)}</c:f>{cache}</c:{kind}></c:{element}>"


def _series_labels(kind: ChartKind, leader_lines: bool) -> str:
    """A series' dLbls once HasDataLabels is on, as Excel writes them; a pie's or doughnut's extension is empty."""
    shown = int(leader_lines)
    extension = ("" if kind in ("pie", "doughnut")
                 else f'<c15:showLeaderLines val="{shown}"/>')
    tail = (f'<c:ext uri="{{CE6537A1-D6FC-4f65-9D91-7224C49458BB}}"'
            ' xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"'
            + (f">{extension}</c:ext>" if extension else "/>"))
    return ('<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:showLegendKey val="0"/>'
            '<c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>'
            f'<c:showBubbleSize val="0"/><c:showLeaderLines val="{shown}"/><c:extLst>{tail}</c:extLst></c:dLbls>')


def _series(kind: ChartKind, index: int, data: SeriesData, unique: str) -> str:
    identity = index if data.series_id is None else data.series_id
    name = ""
    if data.name is not None:
        name = f"<c:tx><c:strRef><c:f>{escape_text(data.name)}</c:f>{_string_cache([data.name_cache])}</c:strRef></c:tx>"
    elif data.name_cache is not None:
        name = f'<c:tx><c:v>{escape_text(data.name_cache)}</c:v></c:tx>'
    categories = ""
    if data.categories is not None:
        cache = (
            _number_cache(list(data.category_cache), data.category_format)
            if data.numeric_categories
            else _string_cache(list(data.category_cache))
        )
        element = "xVal" if kind == "scatter" else "cat"
        categories = _reference(element, data.categories, cache, numbers=data.numeric_categories)
    if data.labels:
        # A series' data labels come just before its categories, on every kind (tests/fixtures/chart_data_labels.json).
        categories = _series_labels(kind, data.leader_lines) + categories
    values = _reference(
        "yVal" if kind == "scatter" else "val",
        data.values,
        _number_cache(list(data.value_cache), data.value_format),
        numbers=True,
    )
    if kind in ("column", "bar"):
        body = f'<c:invertIfNegative val="{int(data.invert)}"/>{categories}{values}'
    elif kind in ("line", "lineMarkers"):
        body = f'{data.marker}{categories}{values}<c:smooth val="{int(data.smooth)}"/>'
    elif kind == "scatter":
        body = (
            f'<c:spPr><a:ln w="38100"><a:noFill/></a:ln></c:spPr>{data.marker}{categories}{values}'
            f'<c:smooth val="{int(data.smooth)}"/>'
        )
    elif kind == "pie" and data.explosion:
        body = f'<c:explosion val="{data.explosion}"/>{categories}{values}'
    else:  # area, pie and doughnut
        body = f"{categories}{values}"
    return (
        f'<c:ser><c:idx val="{identity}"/><c:order val="{index}"/>{name}{body}'
        '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"'
        ' xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
        f'<c16:uniqueId val="{{{identity if data.unique_id is None else data.unique_id:08X}{unique}}}"/>'
        '</c:ext></c:extLst></c:ser>'
    )


# -- the chart --------------------------------------------------------------


def chart_part(
    kind: ChartKind,
    series: Sequence[SeriesData],
    *,
    title: str = "",
    title_deleted: bool = False,
    legend: str = "",
    axis_ids: tuple[int, int] | None = None,
    unique: str | None = None,
    blanks: int = 1,
    visible_only: bool = True,
    over_maximum: bool = False,
    not_available_as_blank: bool = True,
) -> str:
    """A chart part, as Excel writes one of this kind.

    ``title`` is the chart's title element, if it has one, and
    ``title_deleted`` says ChartTitle.Delete removed it. ``axis_ids`` and ``unique``, the tail every
    series' id shares, are random, as Excel's are, unless given. The last
    four are the chart's DisplayBlanksAs, PlotVisibleOnly,
    ShowDataLabelsOverMaximum and DisplayValueNotAvailableAsBlank,
    defaulting as a chart made through ``ChartObjects.Add`` has them,
    whatever its type.
    """
    first, second = axis_ids if axis_ids is not None else _axis_ids()
    tail = unique if unique is not None else "-" + str(uuid.uuid4()).upper().split("-", 1)[1]
    written = "".join(_series(kind, index, data, tail) for index, data in enumerate(series))
    axes_ids = f'<c:axId val="{first}"/><c:axId val="{second}"/>'
    labelled = any(data.categories is not None for data in series)
    if kind in ("column", "bar"):
        direction = "col" if kind == "column" else "bar"
        plot = (
            f'<c:barChart><c:barDir val="{direction}"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            f'{written}{_LABELS.format(extra="")}<c:gapWidth val="150"/>{axes_ids}</c:barChart>'
        )
        axes = _category_axis(first, second, position="b" if kind == "column" else "l", formatted=labelled)
        axes += _value_axis(second, first, position="l" if kind == "column" else "b", between="between",
                            formatted=bool(series), gridlines=True)
    elif kind in ("line", "lineMarkers"):
        markers = '<c:marker val="1"/>' if kind == "lineMarkers" else ""
        plot = (
            f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{written}'
            f'{_LABELS.format(extra="")}{markers}<c:smooth val="0"/>{axes_ids}</c:lineChart>'
        )
        axes = _category_axis(first, second, position="b", formatted=labelled)
        axes += _value_axis(second, first, position="l", between="between", formatted=bool(series), gridlines=True)
    elif kind == "area":
        plot = (
            f'<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>{written}'
            f'{_LABELS.format(extra="")}{axes_ids}</c:areaChart>'
        )
        axes = _category_axis(first, second, position="b", formatted=labelled)
        axes += _value_axis(second, first, position="l", between="midCat", formatted=bool(series), gridlines=True)
    elif kind == "scatter":
        plot = (
            f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{written}'
            f'{_LABELS.format(extra="")}{axes_ids}</c:scatterChart>'
        )
        axes = _value_axis(first, second, position="b", between="midCat", formatted=labelled, gridlines=False)
        axes += _value_axis(second, first, position="l", between="midCat", formatted=bool(series), gridlines=True)
    else:
        # The group's leader lines go off once any series' do (tests/fixtures/chart_data_labels.json).
        leaders = f'<c:showLeaderLines val="{int(all(data.leader_lines for data in series))}"/>'
        hole = '<c:holeSize val="50"/>' if kind == "doughnut" else ""
        plot = (
            f'<c:{kind}Chart><c:varyColors val="1"/>{written}{_LABELS.format(extra=leaders)}'
            f'<c:firstSliceAng val="0"/>{hole}</c:{kind}Chart>'
        )
        axes = ""
    return (
        f'{_HEAD}{title}<c:autoTitleDeleted val="{int(title_deleted)}"/>'
        f'<c:plotArea><c:layout/>{plot}{axes}</c:plotArea>{legend}'
        + _display(blanks, visible_only, over_maximum, not_available_as_blank)
        + _TAIL
    )


def chart_frame(
    *,
    shape_id: int,
    name: str,
    relationship: str,
    left: float,
    top: float,
    width: float,
    height: float,
    grid: SheetGrid,
) -> str:
    """The anchor a chart sits in on a sheet's drawing, as Excel writes one."""
    creation = "{" + str(uuid.uuid4()).upper() + "}"
    return (
        "<xdr:twoCellAnchor>"
        f"{corner_markup('from', grid, left, top)}{corner_markup('to', grid, left + width, top + height)}"
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
        f'<xdr:cNvPr id="{shape_id}" name="{escape_attribute(name)}"><a:extLst>'
        '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">'
        f'<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{creation}"/>'
        "</a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        f' r:id="{relationship}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>'
        "</xdr:twoCellAnchor>"
    )


def _axis_ids() -> tuple[int, int]:
    """Two ids for a chart's axes, random as Excel's are."""
    first = random.randint(10**8, 2**31 - 1)
    second = random.randint(10**8, 2**31 - 1)
    while second == first:
        second = random.randint(10**8, 2**31 - 1)
    return first, second


__all__ = ["CHART_KINDS", "CT_CHART", "ChartKind", "SeriesData", "chart_frame", "chart_part"]
