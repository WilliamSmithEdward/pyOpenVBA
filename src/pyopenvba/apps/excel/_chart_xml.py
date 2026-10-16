"""Charts written as Excel's Insert Chart writes them.

Measured with ``Shapes.AddChart2`` and its default style, one chart of each
kind from the same block of cells: every part below is the markup Excel
wrote, taken apart into the pieces the kinds share. A chart part holds its
own formatting, the text, the lines and each series' colour, so it looks as
Excel's own does without the style and colour parts Excel also writes for
its chart-style gallery; those are left out, as a chart made through
``ChartObjects.Add`` has none either.

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

_TAIL = (
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="{blanks}"/><c:showDLblsOverMax val="0"/>'
    '<c:extLst><c:ext uri="{{56B9EC1D-385E-4148-901F-78D8002777C0}}"'
    ' xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">'
    '<c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16>'
    "</c:ext></c:extLst></c:chart>"
    '<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
    '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">'
    '<a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
    "<a:effectLst/></c:spPr>"
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3"'
    ' footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>'
)

_NOTHING = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>"

#: The grey Excel draws a category axis, and gridlines, in.
_LIGHT_LINE = (
    '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">'
    '<a:lumMod val="{mod}"/><a:lumOff val="{off}"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
)

_GRIDLINES = (
    "<c:majorGridlines><c:spPr>" + _LIGHT_LINE.format(mod=15000, off=85000) + "<a:effectLst/></c:spPr>"
    "</c:majorGridlines>"
)

_BODY = (
    '<a:bodyPr rot="{rot}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square"'
    ' anchor="ctr" anchorCtr="1"/><a:lstStyle/>'
)

_RUN_PROPERTIES = (
    '<a:defRPr sz="{size}" b="0" i="0" u="none" strike="noStrike" kern="1200"{spacing} baseline="0">'
    '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
    '</a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>'
)


def _text(size: int, *, rotated: bool, spaced: bool = False) -> str:
    """Text properties: the size, and the grey Excel sets text in."""
    body = _BODY.format(rot="-60000000" if rotated else "0")
    run = _RUN_PROPERTIES.format(size=size, spacing=' spc="0"' if spaced else "")
    return f'<c:txPr>{body}<a:p><a:pPr>{run}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'


def _title(text: str | None) -> str:
    """The title: Excel's automatic one, or one typed in."""
    typed = ""
    if text is not None:
        body = _BODY.format(rot="0")
        run = _RUN_PROPERTIES.format(size=1400, spacing=' spc="0"')
        typed = (
            f"<c:tx><c:rich>{body}<a:p><a:pPr>{run}</a:pPr>"
            f'<a:r><a:rPr lang="en-US"/><a:t>{escape_text(text)}</a:t></a:r></a:p></c:rich></c:tx>'
        )
    return (
        f'<c:title>{typed}<c:overlay val="0"/>{_NOTHING}{_text(1400, rotated=False, spaced=True)}</c:title>'
        '<c:autoTitleDeleted val="0"/>'
    )


_LEGEND = f'<c:legend><c:legendPos val="b"/><c:overlay val="0"/>{_NOTHING}{_text(900, rotated=False)}</c:legend>'

_LABELS = (
    '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
    '<c:showPercent val="0"/><c:showBubbleSize val="0"/>{extra}</c:dLbls>'
)


def _category_axis(own: int, other: int, *, position: str, ticks: str = "none") -> str:
    return (
        f'<c:catAx><c:axId val="{own}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="{position}"/><c:numFmt formatCode="General" sourceLinked="1"/>'
        f'<c:majorTickMark val="{ticks}"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        "<c:spPr><a:noFill/>" + _LIGHT_LINE.format(mod=15000, off=85000) + "<a:effectLst/></c:spPr>"
        f'{_text(900, rotated=True)}<c:crossAx val="{other}"/><c:crosses val="autoZero"/><c:auto val="1"/>'
        '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
    )


def _value_axis(own: int, other: int, *, position: str, between: str, formatted: bool = True, line: bool = False) -> str:
    number = '<c:numFmt formatCode="General" sourceLinked="1"/>' if formatted else ""
    shape = (
        "<c:spPr><a:noFill/>" + _LIGHT_LINE.format(mod=25000, off=75000) + "<a:effectLst/></c:spPr>"
        if line
        else _NOTHING
    )
    return (
        f'<c:valAx><c:axId val="{own}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="{position}"/>{_GRIDLINES}{number}<c:majorTickMark val="none"/>'
        f'<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>{shape}{_text(900, rotated=True)}'
        f'<c:crossAx val="{other}"/><c:crosses val="autoZero"/><c:crossBetween val="{between}"/></c:valAx>'
    )


# -- series ---------------------------------------------------------------


#: How Excel's default palette varies the six accents on each round past
#: the first, as its colour part lists the variations and as it coloured
#: twenty series, measured.
_VARIATIONS = (
    "",
    '<a:lumMod val="60000"/>',
    '<a:lumMod val="80000"/><a:lumOff val="20000"/>',
    '<a:lumMod val="80000"/>',
    '<a:lumMod val="60000"/><a:lumOff val="40000"/>',
    '<a:lumMod val="50000"/>',
    '<a:lumMod val="70000"/><a:lumOff val="30000"/>',
    '<a:lumMod val="70000"/>',
    '<a:lumMod val="50000"/><a:lumOff val="50000"/>',
)


def _accent(index: int) -> str:
    """The theme colour of the ``index``-th series: the six accents, then
    the same six in each of the palette's variations in turn."""
    variation = _VARIATIONS[index // 6 % len(_VARIATIONS)]
    accent = index % 6 + 1
    if not variation:
        return f'<a:schemeClr val="accent{accent}"/>'
    return f'<a:schemeClr val="accent{accent}">{variation}</a:schemeClr>'


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


def _series(kind: ChartKind, index: int, data: SeriesData, unique: str) -> str:
    colour = _accent(index)
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
    values = _reference(
        "yVal" if kind == "scatter" else "val",
        data.values,
        _number_cache(list(data.value_cache), data.value_format),
        numbers=True,
    )
    if kind in ("column", "bar"):
        look = f"<c:spPr><a:solidFill>{colour}</a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>"
        body = f'{look}<c:invertIfNegative val="0"/>{categories}{values}'
    elif kind == "area":
        look = f"<c:spPr><a:solidFill>{colour}</a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>"
        body = f"{look}{categories}{values}"
    elif kind in ("line", "lineMarkers"):
        look = f'<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>{colour}</a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>'
        marker = _marker(colour) if kind == "lineMarkers" else '<c:marker><c:symbol val="none"/></c:marker>'
        body = f'{look}{marker}{categories}{values}<c:smooth val="0"/>'
    elif kind == "scatter":
        look = '<c:spPr><a:ln w="38100" cap="rnd"><a:noFill/><a:round/></a:ln><a:effectLst/></c:spPr>'
        body = f'{look}{_marker(colour)}{categories}{values}<c:smooth val="0"/>'
    else:  # pie and doughnut: a colour for each slice
        slices = "".join(
            f'<c:dPt><c:idx val="{point}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>{_accent(point)}</a:solidFill>'
            '<a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:dPt>'
            for point in range(len(data.value_cache))
        )
        body = f"{slices}{categories}{values}"
    return (
        f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/>{name}{body}'
        '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"'
        ' xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
        f'<c16:uniqueId val="{{{index:08X}{unique}}}"/></c:ext></c:extLst></c:ser>'
    )


def _marker(colour: str) -> str:
    return (
        '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>'
        f'<a:solidFill>{colour}</a:solidFill><a:ln w="9525"><a:solidFill>{colour}</a:solidFill></a:ln>'
        "<a:effectLst/></c:spPr></c:marker>"
    )


# -- the chart --------------------------------------------------------------


def chart_part(
    kind: ChartKind,
    series: Sequence[SeriesData],
    *,
    title: str | None = None,
    axis_ids: tuple[int, int] | None = None,
    unique: str | None = None,
) -> str:
    """A chart part, as Excel writes one of this kind.

    ``title`` is text typed in; with none the chart shows Excel's automatic
    title. ``axis_ids`` and ``unique``, the tail every series' id shares,
    are random, as Excel's are, unless given.
    """
    first, second = axis_ids if axis_ids is not None else _axis_ids()
    tail = unique if unique is not None else "-" + str(uuid.uuid4()).upper().split("-", 1)[1]
    written = "".join(_series(kind, index, data, tail) for index, data in enumerate(series))
    axes_ids = f'<c:axId val="{first}"/><c:axId val="{second}"/>'
    if kind in ("column", "bar"):
        direction = "col" if kind == "column" else "bar"
        spacing = '<c:gapWidth val="219"/><c:overlap val="-27"/>' if kind == "column" else '<c:gapWidth val="182"/>'
        plot = (
            f'<c:barChart><c:barDir val="{direction}"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            f"{written}{_LABELS.format(extra='')}{spacing}{axes_ids}</c:barChart>"
        )
        axes = _category_axis(first, second, position="b" if kind == "column" else "l") + _value_axis(
            second, first, position="l" if kind == "column" else "b", between="between"
        )
    elif kind in ("line", "lineMarkers"):
        markers = '<c:marker val="1"/>' if kind == "lineMarkers" else ""
        plot = (
            f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{written}'
            f'{_LABELS.format(extra="")}{markers}<c:smooth val="0"/>{axes_ids}</c:lineChart>'
        )
        axes = _category_axis(first, second, position="b") + _value_axis(second, first, position="l", between="between")
    elif kind == "area":
        plot = (
            f'<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>{written}'
            f'{_LABELS.format(extra="")}{axes_ids}</c:areaChart>'
        )
        axes = _category_axis(first, second, position="b", ticks="out") + _value_axis(
            second, first, position="l", between="midCat"
        )
    elif kind == "scatter":
        plot = (
            f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{written}'
            f'{_LABELS.format(extra="")}{axes_ids}</c:scatterChart>'
        )
        numeric = bool(series) and series[0].numeric_categories
        axes = _value_axis(first, second, position="b", between="midCat", formatted=numeric, line=True) + _value_axis(
            second, first, position="l", between="midCat", line=True
        )
    else:
        leaders = '<c:showLeaderLines val="1"/>'
        hole = '<c:holeSize val="75"/>' if kind == "doughnut" else ""
        plot = (
            f'<c:{kind}Chart><c:varyColors val="1"/>{written}{_LABELS.format(extra=leaders)}'
            f'<c:firstSliceAng val="0"/>{hole}</c:{kind}Chart>'
        )
        axes = ""
    blanks = "zero" if kind == "area" else "gap"
    return (
        f"{_HEAD}{_title(title)}<c:plotArea><c:layout/>{plot}{axes}{_NOTHING}</c:plotArea>{_LEGEND}"
        + _TAIL.format(blanks=blanks)
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
