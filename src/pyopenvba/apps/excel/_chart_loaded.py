"""Edits to a chart read from a file, made in its own markup so that what the model does not read stays.

Excel's save of a chart it wrote changes only what the edit touches (scripts/measure_loaded_chart_edits.py,
tests/fixtures/loaded_chart_edits.json), so the part plus the edit is what Excel saves. It rewrites a chart it did not
write whole, at any save, so such a chart keeps its part as it came and is not edited. An edit not built here stays
refused by Chart.admit_change.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._xml import escape_text
from pyopenvba.apps.excel._chart_title import TitleLook
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._chart_axes import AxisState
    from pyopenvba.apps.excel._charts import Chart, ChartState, SeriesState

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
    if entry.group != 1:
        # Only a primary axis's edits were measured (tests/fixtures/loaded_chart_edits.json).
        raise VBAUnsupportedError('Editing a secondary axis of an imported chart is not implemented')
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


_SERIES = re.compile(r'<c:ser>.*?</c:ser>', re.DOTALL)
_ORDER = re.compile(r'<c:order val="(\d+)"/>')
_SERIES_TEXT = re.compile(r'<c:tx>.*?</c:tx>', re.DOTALL)


def _series_span(chart: Chart, entry: SeriesState) -> tuple[int, int]:
    """Where a series is in the part: the model holds them by c:order."""
    state = chart.state
    if state.out_of_sight or len(state.full_series) not in (0, len(state.series)):
        raise VBAUnsupportedError('Editing a series of an imported chart with series out of sight is not implemented')
    return _span_in(state.original, state, entry)


def _span_in(xml: str, state: ChartState, entry: SeriesState) -> tuple[int, int]:
    """Where a series is in a chart part's markup, by its c:order."""
    def order(hit: re.Match[str]) -> int:
        number = _ORDER.search(hit.group(0))
        return int(number.group(1)) if number else -1

    found = sorted(_SERIES.finditer(xml), key=order)
    place = next((at for at, one in enumerate(state.series) if one is entry), None)
    if place is None or len(found) != len(state.series):
        raise VBAUnsupportedError('Finding this series in an imported chart is not implemented')
    return found[place].span()


def name_series(chart: Chart, entry: SeriesState, reference: str, literal: str, shown: str) -> None:
    """Series.Name: its own text, or a cell's with that cell's text cached, in place of what c:tx held. Once the plot area
    is written whole at save, that writes it (tests/fixtures/loaded_chart_structure_edits.json)."""
    admit(chart)
    if chart.state.plot_pending:
        return
    start, end = _series_span(chart, entry)
    series = chart.state.original[start:end]
    if _SERIES_TEXT.search(series) is None:
        raise VBAUnsupportedError('Naming a series of an imported chart that has no name is not implemented')
    if reference:
        inner = (f'<c:strRef><c:f>{escape_text(reference)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{escape_text(shown)}</c:v>'
                 '</c:pt></c:strCache></c:strRef>')
    else:
        inner = f'<c:v>{escape_text(literal)}</c:v>'
    named = _SERIES_TEXT.sub(lambda _: f'<c:tx>{inner}</c:tx>', series, count=1)
    _edited(chart, chart.state.original[:start] + named + chart.state.original[end:])


_RICH = re.compile(r'<c:tx><c:rich>(?P<head>.*?<a:lstStyle/>)(?P<body>.*?)</c:rich></c:tx>', re.DOTALL)
_PARAGRAPH = re.compile(r'<a:p>.*?</a:p>', re.DOTALL)
_RUN_TEXT = re.compile(r'<a:t>.*?</a:t>', re.DOTALL)


def _plain_paragraph(line: str) -> str:
    return f'<a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>{escape_text(line)}</a:t></a:r></a:p>'


def retitled(title: str, wanted: str) -> str:
    """A title element with this text: a carriage return starts a paragraph and a line feed is kept as a line break, written
    CRLF. Text of its own keeps its first paragraph's look; a title with none gains plain paragraphs (tests/fixtures/
    loaded_chart_edits.json)."""
    if '\r\n' in wanted:
        raise VBAUnsupportedError('A title with a CRLF in an imported chart is not implemented')
    lines = [line.replace('\n', '\r\n') for line in wanted.split('\r')]
    found = _RICH.search(title)
    if found is not None:
        paragraphs = _PARAGRAPH.findall(found.group('body'))
        if len(paragraphs) != 1 or len(_RUN_TEXT.findall(paragraphs[0])) != 1:
            raise VBAUnsupportedError('Retyping a title of several runs in an imported chart is not implemented')
        first = _RUN_TEXT.sub(lambda _: f'<a:t>{escape_text(lines[0])}</a:t>', paragraphs[0], count=1)
        body = first + ''.join(_plain_paragraph(line) for line in lines[1:])
        return title[:found.start('body')] + body + title[found.end('body'):]
    if '<c:tx>' in title or '<c:txPr>' in title:
        raise VBAUnsupportedError('Typing this imported chart title is not implemented')
    rich = '<c:tx><c:rich><a:bodyPr/><a:lstStyle/>' + ''.join(_plain_paragraph(line) for line in lines) + '</c:rich></c:tx>'
    return title.replace('<c:title>', '<c:title>' + rich, 1)


def type_title(chart: Chart, wanted: str) -> None:
    """ChartTitle.Text and Caption. Empty text takes the title away, and then raises."""
    from pyopenvba.interpreter._values import error

    admit(chart)
    if wanted == '':
        take_title_away(chart)
        raise error(-2147024809, 'The parameter is incorrect.')
    found = _TITLE.search(chart.state.original)
    if found is None:
        raise VBAUnsupportedError('Typing this imported chart title is not implemented')
    _edited(chart, chart.state.original[:found.start()] + retitled(found.group(0), wanted) + chart.state.original[found.end():])
    state = chart.state
    state.title, state.title_typed, state.title_look.formula = wanted, True, ''


def link_title(chart: Chart, formula: str, shown: str) -> None:
    """ChartTitle.Formula naming a cell: the title's text becomes that reference, its text cached."""
    admit(chart)
    found = _TITLE.search(chart.state.original)
    if found is None:
        raise VBAUnsupportedError('Linking this imported chart title is not implemented')
    linked = (f'<c:tx><c:strRef><c:f>{escape_text(formula)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0">'
              f'<c:v>{escape_text(shown)}</c:v></c:pt></c:strCache></c:strRef></c:tx>')
    title = found.group(0)
    text = re.search(r'<c:tx>.*?</c:tx>', title, re.DOTALL)
    title = title[:text.start()] + linked + title[text.end():] if text else title.replace('<c:title>', '<c:title>' + linked, 1)
    _edited(chart, chart.state.original[:found.start()] + title + chart.state.original[found.end():])
    chart.state.title_typed, chart.state.title_look.formula = False, formula


_AXIS_TITLE = re.compile(r'<c:title>.*?</c:title>', re.DOTALL)


def set_axis_title(chart: Chart, entry: AxisState, wanted: bool) -> None:
    """An axis's HasTitle: the title after the axis's gridlines, back without its text once taken away. A new one on a styled
    chart takes the style's look, which is not read here, so it stays refused."""
    admit(chart)
    state = chart.state
    start, end = _axis_span(chart, entry)
    axis = state.original[start:end]
    found = _AXIS_TITLE.search(axis)
    if wanted == (found is not None):
        return
    if found is not None:
        state.axis_titles_taken[entry.identifier] = found.group(0)
        axis = axis[:found.start()] + axis[found.end():]
    else:
        taken = state.axis_titles_taken.get(entry.identifier)
        if taken is None and state.style_part:
            raise VBAUnsupportedError('A new axis title on an imported styled chart is not implemented')
        title = _TITLE_TEXT.sub('', taken, count=1) if taken else _PLAIN_TITLE
        place = (re.search(r'<c:minorGridlines(?:/>|>.*?</c:minorGridlines>)', axis, re.DOTALL)
                 or re.search(r'<c:majorGridlines(?:/>|>.*?</c:majorGridlines>)', axis, re.DOTALL)
                 or re.search(r'<c:axPos val="\w+"/>', axis))
        if place is None:
            raise VBAUnsupportedError('A title on this imported chart axis is not implemented')
        axis = axis[:place.end()] + title + axis[place.end():]
    _edited(chart, state.original[:start] + axis + state.original[end:])


def type_axis_title(chart: Chart, entry: AxisState, wanted: str) -> None:
    """AxisTitle.Text, as a chart title's."""
    admit(chart)
    if wanted == '':
        raise VBAUnsupportedError('Typing an empty title on an imported chart axis is not implemented')
    start, end = _axis_span(chart, entry)
    axis = chart.state.original[start:end]
    found = _AXIS_TITLE.search(axis)
    if found is None:
        raise VBAUnsupportedError('Typing this imported chart axis title is not implemented')
    axis = axis[:found.start()] + retitled(found.group(0), wanted) + axis[found.end():]
    _edited(chart, chart.state.original[:start] + axis + chart.state.original[end:])


#: A series' own parts a loaded chart rewrites at save, by name: its spPr, its points' dPt elements and its dLbls.
LOOK, POINTS, LABELS = 'look', 'points', 'labels'
_PARTS = (LOOK, POINTS, LABELS)
_HEAD = re.compile(r'<c:ser><c:idx val="\d+"/><c:order val="\d+"/>(?:<c:tx>.*?</c:tx>)?', re.DOTALL)
_OWN_LOOK = re.compile(r'<c:spPr>.*?</c:spPr>', re.DOTALL)
_FOUND = {POINTS: re.compile(r'(?:<c:dPt>.*?</c:dPt>)+', re.DOTALL), LABELS: re.compile(r'<c:dLbls>.*?</c:dLbls>', re.DOTALL)}
#: What follows each in its c:ser, whichever the series has first.
_AFTER = {POINTS: re.compile(r'<c:(?:dLbls|trendline|errBars|cat|val|xVal|yVal)>'),
          LABELS: re.compile(r'<c:(?:trendline|errBars|cat|val|xVal|yVal)>')}
_IDENTIFIED = re.compile(r'<c:(dPt|dLbl)><c:idx val="(\d+)"/>.*?</c:\1>', re.DOTALL)
_UNIQUE = re.compile(r'(<c16:uniqueId val="\{)([0-9A-F]{8})([^}"]*)(\}"/>)')


def _part_span(series: str, part: str) -> tuple[int, int]:
    """Where a part is in a series' markup, or where it would go."""
    if part == LOOK:
        head = _HEAD.match(series)
        if head is None:
            raise VBAUnsupportedError('Finding the look of this imported chart series is not implemented')
        found = _OWN_LOOK.match(series, head.end())
        return head.end(), found.end() if found else head.end()
    found = _FOUND[part].search(series)
    if found is not None:
        return found.span()
    place = _AFTER[part].search(series)
    if place is None:
        raise VBAUnsupportedError('Placing this part of an imported chart series is not implemented')
    return place.start(), place.start()


def _part_now(chart: Chart, entry: SeriesState, part: str) -> str:
    """The part as the model writes it."""
    from pyopenvba.apps.excel._chart_xml import CHART_KINDS
    from pyopenvba.apps.excel._charts import CHART_TYPES, series_type
    from pyopenvba.apps.excel._charts_file import series_parts

    if CHART_TYPES.get(series_type(chart.state, entry), '') not in CHART_KINDS:
        raise VBAUnsupportedError('Editing a series of this kind of imported chart is not implemented')
    if chart.state.style_part:
        # A chart with a style gives what it writes the style's look, which the whole chart's writer does at its end
        # (tests/fixtures/loaded_chart_styled_edits.json).
        from pyopenvba.apps.excel._charts_file import markup

        if chart.sheet.chart_sheet is not None:
            raise VBAUnsupportedError('Editing a series of an imported chart sheet with a style is not implemented')
        written = markup(chart.sheet, chart.shape.shape_id)
        start, end = _span_in(written, chart.state, entry)
        series = written[start:end]
        at, to = _part_span(series, part)
        return series[at:to]
    return series_parts(chart.sheet, chart, entry)[_PARTS.index(part)]


def _without_tails(xml: str) -> str:
    """Unique ids by their number alone: the rest is the part's, or Excel's at random for a label it adds."""
    from pyopenvba.apps.excel._chart_xml import UNIQUE_END

    return _UNIQUE.sub(lambda hit: hit.group(1) + hit.group(2) + hit.group(4), xml.replace(UNIQUE_END, ''))


_SERIES_LOOKS = frozenset(f'Series.Format.{member}' for member in (
    'Fill.ForeColor.RGB', 'Fill.Visible', 'Fill.Transparency', 'Fill.Solid', 'Line.ForeColor.RGB', 'Line.Weight', 'Line.Visible',
    'Line.DashStyle'))
#: The edits measured on loaded charts, by the part they change (tests/fixtures/loaded_chart_label_edits.json,
#: tests/fixtures/loaded_chart_look_edits.json).
EDITS: dict[str, frozenset[str]] = {
    LABELS: frozenset({'Series.HasDataLabels', 'Point.HasDataLabel', 'DataLabels.ShowValue', 'DataLabels.ShowCategoryName',
                       'DataLabels.ShowSeriesName', 'DataLabels.Position', 'DataLabels.NumberFormat', 'DataLabels.Separator',
                       'DataLabel.Text', 'DataLabel.ShowCategoryName', 'DataLabel.Position', 'DataLabel.Font.size',
                       'DataLabel.Font.bold'}),
    LOOK: _SERIES_LOOKS,
    # A series' look takes what it sets from its points.
    POINTS: _SERIES_LOOKS | {'Point.Format.Fill.ForeColor.RGB', 'Point.Format.Line.ForeColor.RGB'},
}


def admit_part(chart: Chart, entry: SeriesState, part: str, member: str) -> None:
    """Whether a series' part can change in this chart's part: by a member measured, and only where the model writes it as
    the file holds it, so that writing it again changes only what the edit did. A chart with a style is refused: Excel gives
    what it adds there the style's look (tests/fixtures/loaded_chart_label_edits.json)."""
    admit(chart)
    if member not in EDITS[part]:
        raise VBAUnsupportedError(f'Changing {member or "this"} on an imported chart is not implemented')
    if part == LABELS:
        from pyopenvba.apps.excel._chart_xml import LabelLook

        if member.startswith(('DataLabel.', 'Point.')) and entry.data_labels and entry.label_look != LabelLook():
            # Excel gives such a point its series' place and the value alone, which the model does not.
            raise VBAUnsupportedError("Changing a point's label on an imported chart whose series' labels have a look is not implemented")
    if member.startswith(('Point.', 'DataLabel.')) and any(one.trendlines for one in chart.state.series):
        # Excel numbered the points and labels it added from 1 on the one book measured with a trendline, from 0 on the others
        # (tests/fixtures/loaded_chart_label_edits.json); the rule is not known.
        raise VBAUnsupportedError("Changing a point of an imported chart with a trendline is not implemented")
    if part == POINTS and member.startswith('Point.') and entry.format_look is not None:
        # Excel gives such a point the series' look beside what is set, its transparency too, which the model does not.
        raise VBAUnsupportedError("Changing a point's look on an imported chart whose series has a look is not implemented")
    state = chart.state
    if state.whole_pending or any(pending is entry and named == part for pending, named in state.parts_pending):
        # The whole part, where it is pending, writes the series' part with it.
        return
    start, end = _series_span(chart, entry)
    series = state.original[start:end]
    at, to = _part_span(series, part)
    if _without_tails(series[at:to]) != _without_tails(_part_now(chart, entry, part)):
        raise VBAUnsupportedError('Editing this series of an imported chart is not implemented')
    if part == LABELS:
        entry.labels_read = True
    state.parts_pending.append((entry, part))


def write_parts(chart: Chart) -> None:
    """The changed parts in place of what each held. A point's dPt or dLbl keeps the unique id the file gave it; those
    added share a new end (tests/fixtures/loaded_chart_label_edits.json)."""
    import uuid

    state = chart.state
    # One new end for everything the save adds, as Excel's save gives them.
    tail = '-' + str(uuid.uuid4()).upper().split('-', 1)[1]
    if state.plot_pending or state.whole_pending:
        # The plot area holds every series, so it is written whole and the parts with it; the whole part, where it is pending.
        _write_plot(chart, tail)
        state.parts_pending, state.plot_pending, state.plot_dropped, state.whole_pending = [], False, set(), False
        return
    for entry, part in state.parts_pending:
        start, end = _series_span(chart, entry)
        series = state.original[start:end]
        at, to = _part_span(series, part)
        kept = {(hit.group(1), hit.group(2)): unique.group(0) for hit in _IDENTIFIED.finditer(series[at:to])
                if (unique := _UNIQUE.search(hit.group(0))) is not None}

        def identified(hit: re.Match[str]) -> str:
            element = hit.group(0)
            unique = _UNIQUE.search(element)
            if unique is None:
                return element
            old = kept.get((hit.group(1), hit.group(2)))
            if old is not None and _without_tails(old) == _without_tails(unique.group(0)):
                return element[:unique.start()] + old + element[unique.end():]
            return element[:unique.start()] + unique.group(1) + unique.group(2) + tail + unique.group(4) + element[unique.end():]

        made = _IDENTIFIED.sub(identified, _part_now(chart, entry, part))
        state.original = state.original[:start] + series[:at] + made + series[to:] + state.original[end:]
    state.parts_pending = []


_PLOT = re.compile(r'<c:plotArea>.*?</c:plotArea>', re.DOTALL)
_SER = re.compile(r'<c:ser><c:idx val="(\d+)"/>.*?</c:ser>', re.DOTALL)
#: The edits measured that change a loaded chart's plot area as a whole: its series, their data and its chart type
#: (tests/fixtures/loaded_chart_structure_edits.json).
PLOT_EDITS = frozenset({
    'SeriesCollection.NewSeries', 'Series.Delete', 'Chart.SetSourceData', 'Chart.PlotBy', 'Series.Values', 'Series.XValues',
    'Series.Formula', 'Chart.ChartType', 'Series.ChartType'})


def _plot_now(chart: Chart) -> str:
    """The plot area as the model writes it."""
    from pyopenvba.apps.excel._charts_file import markup

    if chart.sheet.chart_sheet is not None:
        raise VBAUnsupportedError('Changing the series of an imported chart sheet is not implemented')
    found = _PLOT.search(markup(chart.sheet, chart.shape.shape_id))
    if found is None:
        raise VBAUnsupportedError('Writing the plot area of this imported chart is not implemented')
    return found.group(0)


def admit_plot(chart: Chart, member: str) -> None:
    """Whether a change to the chart's series or type can be made in its part: by a member measured, and only where the
    model writes its plot area as the file holds it, so that writing it again changes only what the edit did. Excel
    changes nothing outside it for these (tests/fixtures/loaded_chart_structure_edits.json)."""
    admit(chart)
    if member not in PLOT_EDITS:
        raise VBAUnsupportedError(f'Changing {member or "this"} on an imported chart is not implemented')
    state = chart.state
    if state.plot_pending or state.whole_pending:
        return
    found = _PLOT.search(state.original)
    if found is None or _without_tails(found.group(0)) != _without_tails(_plot_now(chart)):
        raise VBAUnsupportedError('Changing the series of this imported chart is not implemented')
    state.plot_pending = True


def _ids(plot: str) -> dict[tuple[str, ...], str]:
    """The unique ids a plot area holds, by the series' idx, the element and its idx, and the id's number."""
    found: dict[tuple[str, ...], str] = {}
    for series in _SER.finditer(plot):
        body = series.group(0)
        for element in _IDENTIFIED.finditer(body):
            if (unique := _UNIQUE.search(element.group(0))) is not None:
                found[(series.group(1), element.group(1), element.group(2), unique.group(2))] = unique.group(0)
        own = list(_UNIQUE.finditer(_IDENTIFIED.sub('', body)))
        if own:
            found[(series.group(1), own[-1].group(2))] = own[-1].group(0)
    return found


def _write_plot(chart: Chart, tail: str) -> None:
    """The plot area as the model writes it, in place of what the file held. A series, point or label the file had keeps
    its unique id; what the edit adds takes the save's new end, as Excel's does."""
    state = chart.state
    found = _PLOT.search(state.original)
    assert found is not None
    dropped = {str(identity) for identity in state.plot_dropped}
    kept = {key: unique for key, unique in _ids(found.group(0)).items() if key[0] not in dropped}

    def fresh(unique: re.Match[str], key: tuple[str, ...]) -> str:
        return kept.get(key) or unique.group(1) + unique.group(2) + tail + unique.group(4)

    def series(hit: re.Match[str]) -> str:
        identity = hit.group(1)

        def element(inner: re.Match[str]) -> str:
            text = inner.group(0)
            unique = _UNIQUE.search(text)
            if unique is None:
                return text
            key = (identity, inner.group(1), inner.group(2), unique.group(2))
            return text[:unique.start()] + fresh(unique, key) + text[unique.end():]

        body = _IDENTIFIED.sub(element, hit.group(0))
        # The series' own id is the one after its points and labels.
        last = _UNIQUE.search(body, body.rfind('<c16:uniqueId')) if '<c16:uniqueId' in body else None
        if last is None or '</c:dPt>' in body[last.start():] or '</c:dLbl>' in body[last.start():]:
            return body
        return body[:last.start()] + fresh(last, (identity, last.group(2))) + body[last.end():]

    if state.whole_pending:
        # The part as the model writes it, its plot area's ids kept as below.
        from pyopenvba.apps.excel._charts_file import markup

        written = markup(chart.sheet, chart.shape.shape_id)
        plot = _PLOT.search(written)
        assert plot is not None
        state.original = written[:plot.start()] + _SER.sub(series, plot.group(0)) + written[plot.end():]
        return
    made = _SER.sub(series, _plot_now(chart))
    state.original = state.original[:found.start()] + made + state.original[found.end():]


#: The edits measured that a loaded chart with a style takes as a new chart of its style does, its part written whole at save:
#: Excel's saves of these are its saves of the new chart (tests/fixtures/loaded_chart_styled_edits.json).
WHOLE_EDITS = frozenset({'ChartTitle.Text', 'Axis.HasTitle', 'AxisTitle.Text', 'Axis.HasMajorGridlines', 'Axis.HasMinorGridlines',
                         'area.Font.size', 'title.Font.bold'})


def admit_whole(chart: Chart, member: str) -> bool:
    """Whether this member's edit, refused in place, can be made as on a new chart of the chart's style and the part written
    whole at save: on a chart with a style, by a member measured so, and only where the model writes the whole part as the
    file holds it, so that writing it again changes only what the edit did."""
    from pyopenvba.apps.excel._charts_file import markup

    state = chart.state
    if not state.original or not state.style_part or member not in WHOLE_EDITS or chart.sheet.chart_sheet is not None:
        return False
    if state.whole_pending:
        return True
    if state.parts_pending or state.plot_pending:
        return False
    if _without_tails(state.original) != _without_tails(markup(chart.sheet, chart.shape.shape_id)):
        return False
    state.whole_pending = True
    return True
