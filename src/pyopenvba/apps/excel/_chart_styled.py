"""A chart made with one of Shapes.AddChart2's styles, from 201: its elements spell out in the chart part what their
entries in the style part say, where a new chart of the old kind leaves them to the file's defaults
(scripts/measure_add_chart2.py, tests/fixtures/add_chart2.json)."""

from __future__ import annotations

import colorsys
import re
from functools import lru_cache
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._chart_style_table import FIRST_STYLE, THEME_EFFECTS, THEME_FILLS, style_part
from pyopenvba.apps.excel._chart_xml import TextLook

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState, SeriesState

_NO_FILL = '<a:noFill/>'
_NO_LINE = '<a:ln><a:noFill/></a:ln>'
_NO_EFFECTS = '<a:effectLst/>'
#: What in a style's run stands for its colour.
_RUN_FILLS = ('<a:gradFill', '<a:solidFill', '<a:noFill', '<a:pattFill')
#: The theme colours a series takes in turn, then again through the colours part's variations.
_VARIATIONS = (
    '',
    '<a:lumMod val="60000"/>',
    '<a:lumMod val="80000"/><a:lumOff val="20000"/>',
    '<a:lumMod val="80000"/>',
    '<a:lumMod val="60000"/><a:lumOff val="40000"/>',
    '<a:lumMod val="50000"/>',
    '<a:lumMod val="70000"/><a:lumOff val="30000"/>',
    '<a:lumMod val="70000"/>',
    '<a:lumMod val="50000"/><a:lumOff val="50000"/>',
)
_TURNED = 'rot="-60000000"'
_LEVEL = 'rot="0"'
_UPRIGHT = 'rot="-5400000"'
#: What ApplyLayout gives a value axis, which a style's own says instead.
_LAYOUT_NO_LINE = '<c:spPr><a:ln w="12700"><a:noFill/></a:ln></c:spPr>'
#: The attributes of a style's text, in the order the chart part spells them, and what each is where the style says
#: nothing.
_TEXT_FIELDS = (
    ('sz', ''),
    ('b', '0'),
    ('i', '0'),
    ('u', 'none'),
    ('strike', 'noStrike'),
    ('kern', ''),
    ('cap', ''),
    ('spc', ''),
    ('normalizeH', ''),
    ('baseline', '0'),
)
_SERIES_KINDS = (
    '<c:barChart>',
    '<c:lineChart>',
    '<c:pieChart>',
    '<c:doughnutChart>',
    '<c:scatterChart>',
    '<c:areaChart>',
)
#: The theme's colours as the style names them, by their places in the theme.
_THEME_PLACES = {
    'bg1': 0,
    'lt1': 0,
    'tx1': 1,
    'dk1': 1,
    'bg2': 2,
    'lt2': 2,
    'tx2': 3,
    'dk2': 3,
    'accent1': 4,
    'accent2': 5,
    'accent3': 6,
    'accent4': 7,
    'accent5': 8,
    'accent6': 9,
}


def is_styled(style: int) -> bool:
    return style >= FIRST_STYLE


@lru_cache(maxsize=8)
def _entries(style: int) -> dict[str, str]:
    part = style_part(style)
    held: dict[str, str] = {}
    for match in re.finditer(
        r'<cs:(\w+)(?: [^>]*)?/>|<cs:(\w+)(?: [^>]*)?>.*?</cs:\2>',
        part[part.index('>', part.index('<cs:chartStyle')) + 1 :],
        re.S,
    ):
        held[match.group(1) or match.group(2)] = match.group(0)
    return held


def _split(inner: str) -> tuple[str, str, str]:
    """The fill, the line and what follows the line, its effects and the rest, of a shape's properties, each empty where
    it names none."""
    inner = inner.replace('<a:effectLst/>', '')
    line = re.search(r'<a:ln\b[^>]*?/>|<a:ln\b.*?</a:ln>', inner, re.S)
    if line is None:
        found = re.search(r'<a:(?:effectLst|scene3d|sp3d)\b', inner)
        return (inner[:found.start()], '', inner[found.start():]) if found else (inner, '', '')
    return inner[:line.start()], line.group(0), inner[line.end():]


def _theme_item(entry: str, tag: str, items: tuple[str, ...]) -> str:
    """What the entry's reference to the theme's list comes to, the theme's colour left as its placeholder."""
    found = re.search(rf'<cs:{tag} idx="(\d+)"', entry)
    place = int(found.group(1)) if found else 0
    return items[place - 1] if 1 <= place <= len(items) else ''


def _shape(entry: str) -> tuple[str, str, str]:
    """What the entry says of a shape: its own properties first, then the theme's where it refers to it."""
    found = re.search(r'<cs:spPr>(.*?)</cs:spPr>', entry, re.S)
    fill, line, after = _split(found.group(1) if found else '')
    if 'effectLst' not in after:
        after = (_theme_item(entry, 'effectRef', THEME_EFFECTS) or _NO_EFFECTS) + after
    return fill or _theme_item(entry, 'fillRef', THEME_FILLS), line, after


def _own_shape(own: str) -> tuple[str, str, str]:
    return _split(own[len('<c:spPr>') : -len('</c:spPr>')] if own else '')


def _box(style: int, name: str, own: str = '') -> str:
    """An element with an area: what was set of it, what the style says, and no fill and no line where neither says."""
    fill, line, after = _shape(_entries(style)[name])
    mine_fill, mine_line, _ = _own_shape(own)
    return f'<c:spPr>{mine_fill or fill or _NO_FILL}{mine_line or line or _NO_LINE}{after or _NO_EFFECTS}</c:spPr>'


def _stroke(style: int, name: str, own: str = '') -> str:
    """An element that is only a line."""
    _, line, after = _shape(_entries(style)[name])
    return f'<c:spPr>{_own_shape(own)[1] or line or _NO_LINE}{after or _NO_EFFECTS}</c:spPr>'


def _text_fields(style: int, name: str) -> dict[str, str]:
    found = re.search(r'<cs:defRPr\b([^>]*?)/?>', _entries(style)[name])
    return dict(re.findall(r'(\w+)="([^"]*)"', found.group(1))) if found else {}


def _text_colour(style: int, name: str) -> str:
    found = re.search(r'<cs:fontRef idx="\w+">(.*?)</cs:fontRef>', _entries(style)[name], re.S)
    return found.group(1) if found else '<a:scrgbClr r="0" g="0" b="0"/>'


def _run_children(style: int, name: str) -> str:
    """What the style's run holds besides its attributes, after its colour."""
    found = re.search(r'<cs:defRPr\b[^>]*?[^/]>(.*?)</cs:defRPr>', _entries(style)[name], re.S)
    return found.group(1) if found else ''


def _body(rotation: str) -> str:
    return (
        f'<a:bodyPr {rotation} spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr"'
        ' anchorCtr="1"/>'
    )


def _paragraph_font(style: int, name: str, font: TextLook | None) -> str:
    """The paragraph's default run, with what was set of it."""
    fields = _text_fields(style, name)
    if font is not None and font.size is not None:
        fields['sz'] = str(round(font.size * 100))
    if font is not None and font.bold is not None:
        fields['b'] = '1' if font.bold else '0'
    if font is not None and font.italic is not None:
        fields['i'] = '1' if font.italic else '0'
    spelled = ' '.join(f'{key}="{fields.get(key, plain)}"' for key, plain in _TEXT_FIELDS if fields.get(key, plain))
    colour = (
        f'<a:srgbClr val="{font.rgb}"/>' if font is not None and font.rgb is not None else _text_colour(style, name)
    )
    theme = 'mj' if '<cs:fontRef idx="major"' in _entries(style)[name] else 'mn'
    children = _run_children(style, name)
    latin = f'+{theme}-lt' if font is None or font.name is None else font.name
    return (
        f'<a:defRPr {spelled}>{"" if children.startswith(_RUN_FILLS) else f"<a:solidFill>{colour}</a:solidFill>"}'
        f'{children}'
        f'<a:latin typeface="{latin}"/>'
        f'<a:ea typeface="+{theme}-ea"/><a:cs typeface="+{theme}-cs"/></a:defRPr>'
    )


def _text(style: int, name: str, body: str, font: TextLook | None = None) -> str:
    """An element's text."""
    return (
        f'<c:txPr>{body}<a:lstStyle/><a:p><a:pPr>{_paragraph_font(style, name, font)}</a:pPr>'
        '<a:endParaRPr lang="en-US"/>'
        '</a:p></c:txPr>'
    )


def _accent(index: int) -> str:
    """The colour of the series or point at the index."""
    variation = _VARIATIONS[index // 6 % len(_VARIATIONS)]
    name = f'accent{index % 6 + 1}'
    return f'<a:schemeClr val="{name}">{variation}</a:schemeClr>' if variation else f'<a:schemeClr val="{name}"/>'


def _painted(spelled: str, colour: str) -> str:
    """The style's placeholder for the series' colour as that colour, what the style adds after what its turn does."""
    spelled = spelled.replace('<a:schemeClr val="phClr"/>', colour)
    if colour.endswith('</a:schemeClr>'):
        return spelled.replace('<a:schemeClr val="phClr">', colour[: -len('</a:schemeClr>')])
    return spelled.replace('<a:schemeClr val="phClr">', colour[:-2] + '>')


def _point_box(style: int, index: int, own: str = '') -> str:
    fill, line, after = _shape(_entries(style)['dataPoint'])
    mine_fill, mine_line, _ = _own_shape(own)
    colour = _accent(index)
    return _painted(
        f'<c:spPr>{mine_fill or fill or _NO_FILL}{mine_line or line or _NO_LINE}{after or _NO_EFFECTS}</c:spPr>',
        colour,
    )


def _line(style: int, index: int, own: str = '') -> str:
    _, line, after = _shape(_entries(style)['dataPointLine'])
    return _painted(f'<c:spPr>{_own_shape(own)[1] or line}{after or _NO_EFFECTS}</c:spPr>', _accent(index))


def _width(line: str) -> int:
    found = re.search(r'w="(\d+)"', line)
    return int(found.group(1)) if found else 0


def _scatter_line(made_scatter: bool) -> str:
    """A scatter chart's series have a line of no colour, of a width that depends on whether it was made a scatter chart
    or turned into one (tests/fixtures/add_chart2.json)."""
    width = 38100 if made_scatter else 25400
    return f'<c:spPr><a:ln w="{width}" cap="rnd"><a:noFill/><a:round/></a:ln><a:effectLst/></c:spPr>'


def _marker(style: int, index: int) -> str:
    entry = _entries(style)
    layout = dict(re.findall(r'(\w+)="([^"]*)"', entry['dataPointMarkerLayout']))
    fill, line, after = _shape(entry['dataPointMarker'])
    colour = _accent(index)
    return _painted(
        f'<c:marker><c:symbol val="{layout.get("symbol", "circle")}"/><c:size val="{layout.get("size", "5")}"/><c:spPr>'
        f'{fill or f"<a:solidFill>{colour}</a:solidFill>"}{line or _NO_LINE}{after or _NO_EFFECTS}</c:spPr></c:marker>',
        colour,
    )


def _dressed_points(series: str, style: int, index: int, slices: bool) -> str:
    """The series' points: each slice has a colour of its own, and in any chart a point given a look keeps the rest of
    its series'."""
    pattern = re.compile(r'<c:dPt><c:idx val="(\d+)"/>.*?</c:dPt>', re.S)

    def completed(found: re.Match[str]) -> str:
        place = int(found.group(1)) if slices else index
        return re.sub(
            r'<c:spPr>.*?</c:spPr>',
            lambda hit: _point_box(style, place, hit.group(0)),
            found.group(0),
            count=1,
            flags=re.S,
        )

    if not slices:
        return pattern.sub(completed, series)
    held = {int(found.group(1)): completed(found) for found in pattern.finditer(series)}
    count = re.search(r'<c:(?:val|yVal)>.*?<c:ptCount val="(\d+)"/>', series, re.S)
    for point in range(int(count.group(1)) if count else 0):
        held.setdefault(point, f'<c:dPt><c:idx val="{point}"/><c:bubble3D val="0"/>{_point_box(style, point)}</c:dPt>')
    plain = pattern.sub('', series)
    head = re.match(
        r'<c:ser><c:idx val="\d+"/><c:order val="\d+"/>(?:<c:tx>.*?</c:tx>)?(?:<c:spPr>.*?</c:spPr>)?', plain, re.S
    )
    assert head is not None
    return plain[: head.end()] + ''.join(held[point] for point in sorted(held)) + plain[head.end() :]


def _dressed_labels(series: str, style: int) -> str:
    """Data labels' text, which fits its box, and their leader lines."""
    body = (
        '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100"'
        ' tIns="19050"'
        ' rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr>'
    )

    def dressed_one(found: re.Match[str]) -> str:
        labels = found.group(0)
        if '<c:txPr>' not in labels:
            labels = labels.replace('</c:spPr>', '</c:spPr>' + _text(style, 'dataLabel', body), 1)
        return labels.replace(
            '<c15:showLeaderLines val="1"/></c:ext>',
            f'<c15:showLeaderLines val="1"/><c15:leaderLines>{_stroke(style, "leaderLine")}</c15:leaderLines></c:ext>',
            1,
        )

    return re.sub(r'<c:dLbls>.*?</c:dLbls>', dressed_one, series, flags=re.S)


def _dressed_series(xml: str, state: ChartState) -> str:
    """Each series with its look, what was set of it first."""
    style = state.chart_style
    every: list[SeriesState] = state.full_series or state.series
    found = list(re.finditer(r'<c:ser>.*?</c:ser>', xml, re.S))
    for match, entry in reversed(list(zip(found, every, strict=False))):
        series = match.group(0)
        number = re.search(r'<c:idx val="(\d+)"/>', series)
        assert number is not None
        index = int(number.group(1))
        kind = max(_SERIES_KINDS, key=xml[: match.start()].rfind)
        head = re.match(r'<c:ser><c:idx val="\d+"/><c:order val="\d+"/>(?:<c:tx>.*?</c:tx>)?', series, re.S)
        assert head is not None
        rest = series[head.end() :]
        own = re.match(r'<c:spPr>.*?</c:spPr>', rest, re.S)
        mine = own.group(0) if own else ''
        rest = rest[len(mine) :]
        if kind == '<c:scatterChart>' and entry.format_look is None:
            look = _scatter_line(state.made_scatter)
        elif kind in ('<c:lineChart>', '<c:scatterChart>'):
            look = _line(style, index, mine)
        elif kind in ('<c:pieChart>', '<c:doughnutChart>'):
            look = _box(style, 'dataPoint', mine) if mine else ''
        else:
            look = _point_box(style, index, mine)
        if kind in ('<c:lineChart>', '<c:scatterChart>') and not rest.startswith('<c:marker>'):
            look += _marker(style, index)
        series = _dressed_points(series[: head.end()] + look + rest, style, index, kind in _SERIES_KINDS[2:4])
        series = _dressed_labels(series, style)
        series = series.replace(
            '<c:trendline><c:trendlineType',
            '<c:trendline>' + _painted(_stroke(style, 'trendline'), _accent(index)) + '<c:trendlineType',
        )
        xml = xml[: match.start()] + series + xml[match.end() :]
    return xml


def _dressed_title(title: str, style: int, name: str, font: TextLook | None, rotation: str) -> str:
    """A title: what the style says of its text, in its own paragraph where it was typed, and of its box."""
    title = re.sub(
        r'<c:rich><a:bodyPr[^>]*?/><a:lstStyle/><a:p><a:pPr><a:defRPr(?:[^>]*?/>|.*?</a:defRPr>)</a:pPr>',
        lambda _: f'<c:rich>{_body(rotation)}<a:lstStyle/><a:p><a:pPr>{_paragraph_font(style, name, font)}</a:pPr>',
        title,
        count=1,
        flags=re.S,
    )
    own = re.search(r'<c:spPr>.*?</c:spPr>', title, re.S)
    mine = own.group(0) if own else ''
    if mine:
        title = title.replace(mine, '', 1)
    title = re.sub(r'<c:txPr>.*?</c:txPr>', '', title, count=1, flags=re.S)
    return title.replace(
        '</c:title>', _box(style, name, mine) + _text(style, name, _body(rotation), font) + '</c:title>', 1
    )


def _dressed_axes(xml: str, state: ChartState) -> str:
    style = state.chart_style
    kinds = iter((1, 2))

    def dressed_one(found: re.Match[str]) -> str:
        axis = found.group(0)
        kind = next(kinds)
        name = 'categoryAxis' if kind == 1 else 'valueAxis'
        if state.made_scatter and kind == 1 and '<c:xVal><c:numRef>' not in xml:
            # A new scatter chart's horizontal axis has no number format to follow until its values are numbers.
            axis = axis.replace('<c:numFmt formatCode="General" sourceLinked="1"/>', '', 1)
        axis = axis.replace(
            '<c:majorGridlines/>', f'<c:majorGridlines>{_stroke(style, "gridlineMajor")}</c:majorGridlines>', 1
        )
        axis = axis.replace(
            '<c:minorGridlines/>', f'<c:minorGridlines>{_stroke(style, "gridlineMinor")}</c:minorGridlines>', 1
        )
        axis = re.sub(
            r'<c:title>.*?</c:title>',
            lambda hit: _dressed_title(
                hit.group(0),
                style,
                'axisTitle',
                state.element_fonts.get(f'title{kind}'),
                _UPRIGHT if kind == 2 else _LEVEL,
            ),
            axis,
            count=1,
            flags=re.S,
        )
        own = re.search(r'<c:spPr>(?:(?!<c:spPr>).)*?</c:spPr>(?=(?:<c:txPr>.*?</c:txPr>)?<c:crossAx)', axis, re.S)
        mine = own.group(0) if own else ''
        if mine:
            axis = axis.replace(mine, '', 1)
        if mine == _LAYOUT_NO_LINE:
            mine = ''
        axis = re.sub(r'<c:txPr>(?:(?!<c:txPr>).)*?</c:txPr>(?=<c:crossAx)', '', axis, count=1, flags=re.S)
        font = state.element_fonts.get(f'tick{kind}')
        if '<c:delete val="1"/>' in axis:
            return axis
        look = _box(style, name, mine) + _text(style, name, _body(_TURNED), font)
        return axis.replace('<c:crossAx', look + '<c:crossAx', 1)

    return re.sub(r'<c:(valAx|catAx|dateAx)>.*?</c:\1>', dressed_one, xml, flags=re.S)


def dressed(xml: str, state: ChartState) -> str:
    """The chart part of a chart of a style from 201, its elements given what the style says."""
    style = state.chart_style
    xml = re.sub(
        r'(?<=<c:chart>)<c:title>.*?</c:title>',
        lambda hit: _dressed_title(hit.group(0), style, 'title', state.element_fonts.get('title'), _LEVEL),
        xml,
        count=1,
        flags=re.S,
    )
    xml = _dressed_series(xml, state)
    xml = _dressed_axes(xml, state)
    if '<c:ser>' not in xml:
        # With nothing drawn the leader lines are off.
        xml = xml.replace('<c:showLeaderLines val="1"/>', '<c:showLeaderLines val="0"/>')
    xml = xml.replace(
        '</c:dTable>',
        _box(style, 'dataTable')
        + _text(style, 'dataTable', _body(_LEVEL)).replace('<a:pPr>', '<a:pPr rtl="0">', 1)
        + '</c:dTable>',
        1,
    )
    plot = re.search(r'<c:spPr>(?:(?!<c:spPr>).)*?</c:spPr>(?=</c:plotArea>)', xml, re.S)
    plot_own = plot.group(0) if plot else ''
    if plot is not None:
        xml = xml[: plot.start()] + xml[plot.end() :]
    xml = xml.replace('</c:plotArea>', _box(style, 'plotArea', plot_own) + '</c:plotArea>', 1)

    def legend(found: re.Match[str]) -> str:
        whole = found.group(0)
        own = re.search(r'<c:spPr>.*?</c:spPr>(?=(?:<c:txPr>.*?</c:txPr>)?</c:legend>)', whole, re.S)
        mine = own.group(0) if own else ''
        if mine:
            whole = whole.replace(mine, '', 1)
        whole = re.sub(r'<c:txPr>(?:(?!<c:txPr>).)*?</c:txPr>(?=</c:legend>)', '', whole, flags=re.S)
        return whole.replace(
            '</c:legend>',
            _box(style, 'legend', mine)
            + _text(style, 'legend', _body(_LEVEL), state.element_fonts.get('legend'))
            + '</c:legend>',
        )

    xml = re.sub(r'<c:legend>.*?</c:legend>', legend, xml, count=1, flags=re.S)
    area = re.search(r'(?<=</c:chart>)(<c:spPr>.*?</c:spPr>)?(<c:txPr>.*?</c:txPr>)?(?=<c:printSettings>)', xml, re.S)
    assert area is not None
    text = area.group(2) or (
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    )
    made = xml[:area.start()] + _box(style, 'chartArea', area.group(1) or '') + text + xml[area.end():]
    # The elements that are no series' take the first colour for the style's placeholder.
    return _painted(made, _accent(0))


def _rgb(spelled: str, theme: tuple[str, ...]) -> str:
    """A scheme colour the style names, as the colour it comes to in the theme."""
    found = re.search(r'<a:schemeClr val="(\w+)"', spelled)
    if found is None:
        return '000000'
    held = theme[_THEME_PLACES[found.group(1)]]
    red, green, blue = (int(held[at : at + 2], 16) / 255 for at in (0, 2, 4))
    hue, light, saturation = colorsys.rgb_to_hls(red, green, blue)
    modulated = re.search(r'<a:lumMod val="(\d+)"', spelled)
    offset = re.search(r'<a:lumOff val="(\d+)"', spelled)
    light = light * (int(modulated.group(1)) / 100000 if modulated else 1) + (
        int(offset.group(1)) / 100000 if offset else 0
    )
    return ''.join(f'{round(value * 255):02X}' for value in colorsys.hls_to_rgb(hue, light, saturation))


def _font_entry(key: str) -> str:
    """The style's entry for the text of the chart element a font belongs to."""
    if key.startswith('tick'):
        return 'categoryAxis' if key.endswith('1') else 'valueAxis'
    return {'title': 'title', 'legend': 'legend', 'area': 'chartArea'}.get(
        key, 'axisTitle' if key.startswith('title') else 'dataLabel'
    )


def font_read(style: int, key: str, theme: tuple[str, ...]) -> tuple[float, bool, str]:
    """What the font of the element reads while nothing was set of it: size, bold and colour."""
    name = _font_entry(key)
    fields = _text_fields(style, name)
    return int(fields.get('sz', '1000')) / 100, fields.get('b', '0') == '1', _rgb(_text_colour(style, name), theme)


def line_read(style: int, key: str, theme: tuple[str, ...]) -> tuple[float, str] | None:
    """The weight and colour of the element's line while nothing was set of it, where the style gives it one."""
    name = {
        'area': 'chartArea',
        'plot': 'plotArea',
        'axis1': 'categoryAxis',
        'axis2': 'valueAxis',
        'grid1': 'gridlineMajor',
        'grid2': 'gridlineMajor',
    }.get(key)
    if name is None:
        return None
    line = _shape(_entries(style)[name])[1]
    return (_width(line) / 12700, _rgb(line, theme)) if _width(line) else None


def series_line_weight(style: int) -> float:
    """The weight of the line of a line chart's series."""
    return _width(_shape(_entries(style)['dataPointLine'])[1]) / 12700
