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
    from pyopenvba.apps.excel._chart_styled import is_styled

    admit(chart)
    state = chart.state
    if state.title is None:
        if is_styled(state.chart_style):
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
    from pyopenvba.apps.excel._chart_styled import is_styled

    admit(chart)
    state = chart.state
    if state.title is not None:
        return
    if not state.title_taken and is_styled(state.chart_style):
        raise VBAUnsupportedError('Giving a title to an imported styled chart that has none is not implemented')
    title = _TITLE_TEXT.sub('', state.title_taken, count=1) if state.title_taken else _PLAIN_TITLE
    flag = _FLAG.search(state.original)
    assert flag is not None
    xml = state.original[:flag.start()] + title + '<c:autoTitleDeleted val="0"/>' + state.original[flag.end():]
    _edited(chart, xml)
    state.title, state.title_typed, state.title_deleted = '', False, False
