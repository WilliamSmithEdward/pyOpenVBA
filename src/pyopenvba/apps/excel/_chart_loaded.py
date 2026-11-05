"""Edits to a chart read from a file, made in its own markup so that what the model does not read stays.

Excel's save of a chart it wrote changes only what the edit touches (scripts/measure_loaded_chart_edits.py,
tests/fixtures/loaded_chart_edits.json), so the part plus the edit is what Excel saves. It rewrites a chart it did not
write whole, at any save, so such a chart keeps its part as it came and is not edited. An edit not built here stays
refused by Chart.admit_change.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._chart_title import TitleLook
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._chart_axes import AxisState
    from pyopenvba.apps.excel._charts import Chart

#: What Excel writes at the head of a chart part, whatever the chart holds.
_EXCELS_HEAD = re.compile(r'(?:<\?xml[^>]*\?>\s*)?<c:chartSpace\b[^>]*><c:date1904 val="[01]"/>(?:<c:lang [^>]*/>)?<c:roundedCorners val="[01]"/>')
_FLAG = re.compile(r'<c:autoTitleDeleted val="[01]"/>')
_TITLE = re.compile(r'(?<=<c:chart>)<c:title>.*?</c:title>', re.DOTALL)
_TITLE_TEXT = re.compile(r'<c:tx>.*?</c:tx>', re.DOTALL)
_PLAIN_TITLE = '<c:title><c:overlay val="0"/></c:title>'


def admit(chart: Chart) -> None:
    """Whether this chart's part can be edited at all: not on a protected sheet, and only as Excel wrote it."""
    from pyopenvba.apps.excel._protection import guard_drawing

    chart.admit_read()
    guard_drawing(chart.sheet, 'editing a Chart')
    if _EXCELS_HEAD.match(chart.state.original) is None or _FLAG.search(chart.state.original) is None:
        raise VBAUnsupportedError('Editing an imported chart Excel did not write is not implemented')


def _edited(chart: Chart, xml: str) -> None:
    chart.state.original, chart.state.dirty = xml, True
    chart.sheet.drawing_changed()


def take_title_away(chart: Chart) -> None:
    """HasTitle = False and ChartTitle.Delete: the element goes, kept for the title coming back, and the flag says so. With no
    title nothing changes."""
    admit(chart)
    state = chart.state
    if state.title is None:
        if state.style_part:
            raise VBAUnsupportedError('Taking the title from an imported styled chart that has none is not implemented')
        return
    found = _TITLE.search(state.original)
    if found is None:
        raise VBAUnsupportedError('Taking away this imported chart title is not implemented')
    state.title_taken = found.group(0)
    xml = state.original[:found.start()] + state.original[found.end():]
    _edited(chart, _FLAG.sub('<c:autoTitleDeleted val="1"/>', xml, count=1))
    state.title, state.title_typed, state.title_look, state.title_deleted = None, False, TitleLook(), True
    state.element_looks.pop('title', None)
    state.element_fonts.pop('title', None)


def give_title(chart: Chart) -> None:
    """HasTitle = True: an automatic title, and one taken away comes back with its look but without its text."""
    admit(chart)
    state = chart.state
    if state.title is not None:
        return
    if not state.title_taken and state.style_part:
        raise VBAUnsupportedError('Giving a title to an imported styled chart that has none is not implemented')
    title = _TITLE_TEXT.sub('', state.title_taken, count=1) if state.title_taken else _PLAIN_TITLE
    flag = _FLAG.search(state.original)
    assert flag is not None
    xml = state.original[:flag.start()] + title + '<c:autoTitleDeleted val="0"/>' + state.original[flag.end():]
    _edited(chart, xml)
    state.title, state.title_typed, state.title_deleted = '', False, False


_LEGEND = re.compile(r'<c:legend>.*?</c:legend>', re.DOTALL)
_LEGEND_LAYOUT = re.compile(r'<c:layout(?:/>|>.*?</c:layout>)', re.DOTALL)
_LEGEND_PLACE = re.compile(r'<c:legendPos val="\w+"/>')


def _legend_placed(legend: str, position: str) -> str:
    """The legend at one of its places, which ends any layout of its own."""
    return _LEGEND_PLACE.sub(f'<c:legendPos val="{position}"/>', _LEGEND_LAYOUT.sub('', legend, count=1), count=1)


def set_legend(chart: Chart, wanted: bool) -> None:
    """HasLegend: the element goes, and comes back on the right with its look but no layout."""
    admit(chart)
    state = chart.state
    if wanted == state.legend:
        return
    if wanted:
        if not state.legend_taken:
            raise VBAUnsupportedError('Giving a legend to an imported chart that has none is not implemented')
        at = state.original.index('</c:plotArea>') + len('</c:plotArea>')
        _edited(chart, state.original[:at] + _legend_placed(state.legend_taken, 'r') + state.original[at:])
        state.legend_position, state.legend_manual = 'r', False
    else:
        found = _LEGEND.search(state.original)
        if found is None:
            raise VBAUnsupportedError('Taking away this imported chart legend is not implemented')
        state.legend_taken = found.group(0)
        _edited(chart, state.original[:found.start()] + state.original[found.end():])
        state.element_looks.pop('legend', None)
        state.element_fonts.pop('legend', None)
    state.legend = wanted


def place_legend(chart: Chart, position: str) -> None:
    """Legend.Position."""
    admit(chart)
    state = chart.state
    found = _LEGEND.search(state.original)
    if found is None:
        raise VBAUnsupportedError('Placing this imported chart legend is not implemented')
    _edited(chart, state.original[:found.start()] + _legend_placed(found.group(0), position) + state.original[found.end():])
    state.legend_position, state.legend_manual = position, False


def _axis_span(chart: Chart, entry: AxisState) -> tuple[int, int]:
    found = re.search(rf'<c:(valAx|catAx)><c:axId val="{entry.identifier}"/>.*?</c:\1>', chart.state.original, re.DOTALL)
    if found is None:
        raise VBAUnsupportedError('Editing this imported chart axis is not implemented')
    return found.span()


def rescale(chart: Chart, entry: AxisState) -> None:
    """An axis's fixed ends in its c:scaling, after its orientation: maximum, then minimum."""
    from pyopenvba.apps.excel._chart_xml import format_number

    admit(chart)
    start, end = _axis_span(chart, entry)
    axis = chart.state.original[start:end]
    without = re.sub(r'<c:(?:max|min) val="[^"]*"/>', '', axis)
    ends = ''.join(f'<c:{tag} val="{format_number(value)}"/>' for tag, value in (('max', entry.maximum), ('min', entry.minimum)) if value is not None)
    _edited(chart, chart.state.original[:start] + without.replace('</c:scaling>', ends + '</c:scaling>', 1) + chart.state.original[end:])


def set_gridlines(chart: Chart, entry: AxisState, major: bool, wanted: bool) -> None:
    """HasMajorGridlines and HasMinorGridlines: taken away, and back as they were. New ones on a styled chart take the style's
    look, which is not read here, so they stay refused."""
    admit(chart)
    state = chart.state
    start, end = _axis_span(chart, entry)
    axis = state.original[start:end]
    tag = 'c:majorGridlines' if major else 'c:minorGridlines'
    found = re.search(rf'<{tag}(?:/>|>.*?</{tag}>)', axis, re.DOTALL)
    key = (entry.identifier, major)
    if wanted == (found is not None):
        return
    if found is not None:
        state.gridlines_taken[key] = found.group(0)
        axis = axis[:found.start()] + axis[found.end():]
    else:
        element = state.gridlines_taken.get(key)
        if element is None and state.style_part:
            raise VBAUnsupportedError('New gridlines on an imported styled chart are not implemented')
        # After the axis's place, and minor after major.
        place = None if major else re.search(r'<c:majorGridlines(?:/>|>.*?</c:majorGridlines>)', axis, re.DOTALL)
        place = place or re.search(r'<c:axPos val="\w+"/>', axis)
        if place is None:
            raise VBAUnsupportedError('Gridlines on this imported chart axis are not implemented')
        axis = axis[:place.end()] + (element or f'<{tag}/>') + axis[place.end():]
    _edited(chart, state.original[:start] + axis + state.original[end:])
