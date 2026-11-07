"""The looks of a chart read from a file, as the reads of its Format answer them.

A look is read only where its reads answer as Excel's do (scripts/measure_loaded_chart_looks.py, tests/fixtures/loaded_chart_looks.json):
colours by RGB with no alpha or mods, line width and dash, and fill or line left off. On a chart with a style part, what equals the
style's own look is the style's, read as a chart AddChart2 made with that style reads it; only what differs is the element's own. Any
other look, and any chart whose style or colours the model has no table for, leaves the element unread, and its reads refused.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState
    from pyopenvba.powerquery._opc import OpcFile
    from pyopenvba.shapes._values import FillLine

_SERIES = re.compile(r'<c:ser>.*?</c:ser>', re.DOTALL)
#: What a series holds that has its own spPr: its points, labels, trendlines and error bars.
_SERIES_PARTS = re.compile(r'<c:(dPt|dLbls|trendline|errBars)>.*?</c:\1>', re.DOTALL)
_ORDER = re.compile(r'<c:order val="(\d+)"/>')
_SPPR = re.compile(r'<c:spPr>.*?</c:spPr>', re.DOTALL)
#: What an element holds that has a spPr of its own.
_ELEMENT_PARTS = re.compile(r'<c:(title|majorGridlines|minorGridlines|legendEntry|dLbls|txPr|ser|dTable)>.*?</c:\1>', re.DOTALL)
#: The colour table every chart AddChart2 makes is given, which cycles the six accents.
_CYCLE = 10
#: What a loaded styled chart reads otherwise than a new one of its style: a plot area's style-given fill reads RGB 0 and theme -2,
#: and an axis's line reads automatic, so these stay unread on a styled chart.
_STYLED_UNREAD = ('plot', 'axis1', 'axis2')


#: The colour modifiers a colour read here may carry, which come to an RGB colour by the shapes' rule.
_MODS = frozenset({'lumMod', 'lumOff'})


def _resolved(rgb: str | None, mods: tuple[tuple[str, int], ...]) -> tuple[str | None, tuple[tuple[str, int], ...]]:
    """An RGB colour under lumMod and lumOff as the colour it comes to: shade and tint in linear light, lumMod and lumOff on HSL
    lightness, as a shape's (tests/fixtures/loaded_chart_looks.json: E97132 under 60000 and 40000 reads 8694514)."""
    from pyopenvba.apps.excel._shape_format import painted

    if rgb is None or not mods or not {name for name, _ in mods} <= _MODS:
        return rgb, mods
    bgr = painted(rgb, mods)
    return f'{bgr & 0xFF:02X}{(bgr >> 8) & 0xFF:02X}{(bgr >> 16) & 0xFF:02X}', ()


def _plain(look: FillLine) -> bool:
    """Whether a look holds only what a chart's reads answer: RGB colours, a fill's alpha as its transparency, no theme colour."""
    themed = (look.fill_scheme, look.line_scheme, look.line_alpha, look.fill_pattern, look.extra, look.chart_effect)
    return all(item is None for item in themed) and not look.fill_mods and not look.line_mods and not look.effects


def _style(state: ChartState, package: OpcFile, part: str) -> int | None:
    """The style a chart's style part names, where its colours are the cycle and the model has the style's table."""
    from pyopenvba.apps.excel._chart_styled import _entries, is_styled  # pyright: ignore[reportPrivateUsage]
    from pyopenvba.apps.excel._charts_file import REL_CHART_STYLE, rels_part

    rels = package.read(rels_part(part)).decode('utf-8')
    found: dict[str, str] = {}
    for target, kind in re.findall(r'Target="([^"]+)"[^>]*Type="[^"]*/(chartStyle|chartColorStyle)"', rels):
        found[kind] = target
    for kind, target in re.findall(r'Type="[^"]*/(chartStyle|chartColorStyle)"[^>]*Target="([^"]+)"', rels):
        found[kind] = target
    if REL_CHART_STYLE not in rels or set(found) != {'chartStyle', 'chartColorStyle'}:
        return None
    folder = part.rsplit('/', 1)[0]
    style = re.search(r'<cs:chartStyle\b[^>]*\bid="(\d+)"', package.read(f'{folder}/{found["chartStyle"]}').decode('utf-8'))
    colours = re.search(r'<cs:colorStyle\b[^>]*\bid="(\d+)"', package.read(f'{folder}/{found["chartColorStyle"]}').decode('utf-8'))
    if style is None or colours is None or int(colours.group(1)) != _CYCLE or not is_styled(int(style.group(1))):
        return None
    try:
        _entries(int(style.group(1)))
    except (KeyError, ValueError, VBAUnsupportedError):
        return None
    return int(style.group(1))


def _own(markup: str, expected: str) -> str:
    """What of an element's own spPr is not its style's: its fill, its line, or both, as spPr content."""
    from pyopenvba.apps.excel._chart_styled import _own_shape  # pyright: ignore[reportPrivateUsage]

    own = _SPPR.search(markup)
    mine_fill, mine_line, _ = _own_shape(own.group(0) if own else '')
    style_fill, style_line, _ = _own_shape(expected) if expected else ('', '', '')
    return (mine_fill if mine_fill != style_fill else '') + (mine_line if mine_line != style_line else '')


def _read(markup: str) -> FillLine | None:
    """A look from spPr content, where its reads answer."""
    from pyopenvba.shapes._read_look import read_look

    try:
        look = read_look('<xdr:spPr>' + markup + '</xdr:spPr>')
    except VBAUnsupportedError:
        return None
    look.fill_rgb, look.fill_mods = _resolved(look.fill_rgb, look.fill_mods)
    look.line_rgb, look.line_mods = _resolved(look.line_rgb, look.line_mods)
    return look if _plain(look) else None


def _element(state: ChartState, key: str, markup: str | None, expected: str) -> None:
    if markup is None or (state.style_part and key in _STYLED_UNREAD):
        return
    own = _own(markup, expected)
    if own == '':
        state.looks_read.add(key)
        return
    look = _read(own)
    if look is not None:
        state.element_looks[key] = look
        state.looks_read.add(key)


def read_looks(state: ChartState, package: OpcFile, part: str) -> None:
    """Each loaded series' and element's look that can be read, and so answered; the rest stays unread, and refused."""
    from pyopenvba.apps.excel._chart_styled import _accent, _box, _painted, _point_box, _stroke  # pyright: ignore[reportPrivateUsage]

    # What a series holds for its points, which is not read: these keep its points', and its labels', reads refused whatever else is
    # read (tests/fixtures/loaded_chart_labels.json).
    found = _SERIES.findall(state.original)
    if len(found) == len(state.series):
        for entry, markup in zip(state.series, found):
            entry.points_unread = '<c:dPt>' in markup or ('<c:dLbls>' in markup and not entry.labels_read)
            entry.labels_unread = '<c:dLbl>' in markup and not entry.labels_read
    else:
        for entry in state.series:
            entry.points_unread = entry.labels_unread = True
    style = None
    if state.style_part:
        style = _style(state, package, part)
        if style is None:
            return
        state.chart_style = style

    def box(name: str) -> str:
        return _painted(_box(style, name), _accent(0)) if style is not None else ''

    xml = state.original
    orders = [int(hit.group(1)) if (hit := _ORDER.search(one)) else -1 for one in found]
    if len(found) == len(state.series) and sorted(orders) == orders:
        for index, (entry, markup) in enumerate(zip(state.series, found)):
            expected = _point_box(style, index) if style is not None else ''
            own = _own(_SERIES_PARTS.sub('', markup), expected)
            if own == '':
                entry.look_read = True
            elif (look := _read(own)) is not None:
                entry.format_look, entry.look_read = look, True
    area = re.search(r'</c:chart>(<c:spPr>.*?</c:spPr>)?', xml, re.DOTALL)
    _element(state, 'area', area.group(1) or '' if area else None, box('chartArea'))
    plot = re.search(r'<c:plotArea>.*?</c:plotArea>', xml, re.DOTALL)
    if plot is not None:
        inner = re.sub(r'<c:(valAx|catAx|dateAx|serAx|\w+Chart)>.*?</c:\1>', '', plot.group(0), flags=re.DOTALL)
        pie = '<c:pieChart>' in plot.group(0) or '<c:doughnutChart>' in plot.group(0)
        _element(state, 'plot', _ELEMENT_PARTS.sub('', inner), box('plotArea3D' if pie else 'plotArea'))
    legend = re.search(r'<c:legend>.*?</c:legend>', xml, re.DOTALL)
    _element(state, 'legend', _ELEMENT_PARTS.sub('', legend.group(0)) if legend else None, box('legend'))
    title = re.search(r'(?<=<c:chart>)<c:title>.*?</c:title>', xml, re.DOTALL)
    _element(state, 'title', re.sub(r'<c:(tx|txPr)>.*?</c:\1>', '', title.group(0), flags=re.DOTALL) if title else None, box('title'))
    for entry in state.axes.values():
        if entry.group != 1:
            continue
        axis = re.search(rf'<c:(valAx|catAx|dateAx)><c:axId val="{entry.identifier}"/>.*?</c:\1>', xml, re.DOTALL)
        if axis is None:
            continue
        name = 'valueAxis' if entry.kind == 2 else 'categoryAxis'
        _element(state, f'axis{entry.tag}', _ELEMENT_PARTS.sub('', axis.group(0)), box(name))
        grid = re.search(r'<c:majorGridlines(?:/>|>.*?</c:majorGridlines>)', axis.group(0), re.DOTALL)
        stroke = _painted(_stroke(style, 'gridlineMajor'), _accent(0)) if style is not None else ''
        _element(state, f'grid{entry.tag}', grid.group(0) if grid else None, stroke)
