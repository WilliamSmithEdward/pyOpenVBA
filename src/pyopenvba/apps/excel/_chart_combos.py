"""Combo charts: a series' own ChartType and AxisGroup, and the groups and secondary axes they bring, as live Excel has them
(scripts/measure_combo_charts.py, tests/fixtures/combo_charts.json).

A chart's series fall into groups by axis group and kind: columns, lines and areas here, in the order areas, columns, lines
and within each kind primary before secondary, as the file writes them and ChartGroups numbers them. A chart with more than one group reads as
xlCombination. Stacking belongs to a group, so a series that stacks stacks its group; markers belong to a line series.

A series that moves takes its group's gap and overlap along: the group it lands in, new or not, takes them over. A new series
joins the first primary group. Moving the last primary series off puts every series back on the primary axes. Bars, scatter,
pie, doughnut and radar beside other kinds, or moved to the secondary axes, force axes of their own and are refused.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any, cast

from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState, SeriesState

#: A type's kind and stacking; line types differ by markers, which each series has for itself.
FAMILIES: dict[int, tuple[str, str]] = {51: ('column', ''), 52: ('column', 'stacked'), 53: ('column', 'percentStacked'),
                                        4: ('line', ''), 65: ('line', ''), 63: ('line', 'stacked'), 66: ('line', 'stacked'),
                                        64: ('line', 'percentStacked'), 67: ('line', 'percentStacked'),
                                        1: ('area', ''), 76: ('area', 'stacked'), 77: ('area', 'percentStacked')}
_TYPES: dict[tuple[str, str, bool], int] = {('column', '', False): 51, ('column', 'stacked', False): 52, ('column', 'percentStacked', False): 53,
                                            ('line', '', False): 4, ('line', 'stacked', False): 63, ('line', 'percentStacked', False): 64,
                                            ('line', '', True): 65, ('line', 'stacked', True): 66, ('line', 'percentStacked', True): 67,
                                            ('area', '', False): 1, ('area', 'stacked', False): 76, ('area', 'percentStacked', False): 77}
_RANK = {'area': 0, 'column': 1, 'line': 2}
#: What each group keeps for itself once a chart has several; the first two travel with a moving series.
_DEFAULTS: dict[str, object] = {'drop_lines': False, 'hi_lo_lines': False, 'up_down_bars': False, 'series_lines': False, 'vary': None}
#: Every XlChartType; anything else Series.ChartType refuses as Excel does.
KNOWN_TYPES = {-4169, -4151, -4120, -4102, -4101, -4100, -4098, 1, 4, 5, 15, *range(51, 113), *range(118, 124)}
#: Types Excel sets beside others by giving them axes of their own, which are not built.
SERIES_ALONGSIDE = {57, 58, 59, 5, 68, 69, 71, -4120, 80, -4151, 81, 82, -4169, 72, 73, 74, 75, 88, 89, 90, 91}


@dataclass(frozen=True)
class Group:
    axis: int
    family: str
    grouping: str
    members: tuple[SeriesState, ...]

    @property
    def key(self) -> tuple[int, str]:
        return self.axis, self.family


def _marker_none() -> int:
    from pyopenvba.apps.excel._charts import MARKER_NONE

    return MARKER_NONE


def _base(state: ChartState) -> int:
    from pyopenvba.apps.excel._charts import shown_type

    return shown_type(state)


def is_combo(state: ChartState) -> bool:
    return any(item.own_type is not None or item.axis_group != 1 for item in state.series)


def series_type(state: ChartState, item: SeriesState) -> int:
    """The type a series reads: its own in a combo chart, with its markers, and otherwise the chart's."""
    if item.own_type is None:
        return _base(state)
    family, grouping = FAMILIES[item.own_type]
    return _TYPES[family, grouping, family == 'line' and item.marker != _marker_none()]


def groups(state: ChartState) -> list[Group]:
    """The chart's groups, in the file's order."""
    found: dict[tuple[int, str], list[SeriesState]] = {}
    grouped: dict[tuple[int, str], str] = {}
    for item in state.series:
        family, grouping = FAMILIES.get(series_type(state, item), ('', ''))
        found.setdefault((item.axis_group, family), []).append(item)
        grouped[item.axis_group, family] = grouping
    return [Group(key[0], key[1], grouped[key], tuple(found[key])) for key in sorted(found, key=lambda one: (_RANK.get(one[1], 9), one[0]))]


def plot_order(state: ChartState, item: SeriesState) -> int:
    """A series' place in its group, or its place among all of them once ChartType has put them in one."""
    if state.plot_sequence is not None and not is_combo(state):
        return next(index for index, entry in enumerate(state.plot_sequence, start=1) if entry is item)
    for group in groups(state):
        for index, entry in enumerate(group.members, start=1):
            if entry is item:
                return index
    raise ValueError('The series is not in the chart')


def group_value(state: ChartState, key: tuple[int, str], name: str) -> object:
    own = state.group_settings.get(key, {})
    if name in own:
        return own[name]
    return _DEFAULTS[name] if name in _DEFAULTS else getattr(state, name)


def set_group_value(state: ChartState, key: tuple[int, str], name: str, value: object) -> None:
    state.group_settings.setdefault(key, {})[name] = value


def _carry(state: ChartState, source: tuple[int, str], target: tuple[int, str]) -> None:
    """What a moving series' group gives the group it lands in: its gap, and only from bars an overlap."""
    set_group_value(state, target, 'gap_width', group_value(state, source, 'gap_width'))
    set_group_value(state, target, 'overlap', group_value(state, source, 'overlap') if source[1] == 'column' else 0)


def _admit_base(state: ChartState) -> None:
    if _base(state) not in FAMILIES and not is_combo(state):
        raise VBAUnsupportedError('A series of its own on a chart of this type is not implemented')
    if state.full_series or state.out_of_sight:
        raise VBAUnsupportedError('A series of its own on a chart with series filtered or out of sight is not implemented')
    if not is_combo(state) and (state.drop_lines or state.hi_lo_lines or state.up_down_bars or state.series_lines or state.vary is not None):
        raise VBAUnsupportedError('A series of its own on a chart whose group has lines or varied colours is not implemented')
    if state.layout_chart_labels or state.element_labels is not None:
        raise VBAUnsupportedError('A series of its own on a chart with a layout or labels of its own is not implemented')


def _enter(state: ChartState) -> None:
    """Every series says its type for itself once any differs."""
    for item in state.series:
        if item.own_type is None:
            item.own_type = series_type(state, item)


def settle(state: ChartState) -> None:
    """Back to a chart of one type where one group is left; the secondary axes gone with their last series."""
    if not any(item.axis_group == 2 for item in state.series):
        state.secondary_axes = {}
    if not is_combo(state):
        return
    found = groups(state)
    if not found:
        return
    first = found[0]
    if len(found) > 1:
        state.chart_type = series_type(state, first.members[0])
        return
    type_of = series_type(state, first.members[0])
    for name in ('gap_width', 'overlap', 'drop_lines', 'hi_lo_lines', 'up_down_bars', 'series_lines', 'vary'):
        setattr(state, name, group_value(state, first.key, name))
    state.group_settings = {}
    state.secondary_axes = {}
    for item in state.series:
        item.own_type, item.axis_group = None, 1
    state.chart_type = type_of


def set_series_type(state: ChartState, item: SeriesState, wanted: int) -> None:
    """A series given a type of the kinds that sit together, stacking its group if the type stacks."""
    _admit_base(state)
    family, grouping = FAMILIES[wanted]
    before = series_type(state, item)
    source = (item.axis_group, FAMILIES[before][0])
    _enter(state)
    target = (item.axis_group, family)
    if source != target:
        _carry(state, source, target)
    item.own_type = wanted
    if family == 'line':
        from pyopenvba.apps.excel._charts import MARKER_AUTOMATIC

        item.marker = _marker_none() if wanted in (4, 63, 64) else MARKER_AUTOMATIC
    changed = False
    for other in state.series:
        if other is not item and other.axis_group == item.axis_group and other.own_type is not None and FAMILIES[other.own_type][0] == family:
            changed = changed or FAMILIES[other.own_type][1] != grouping
            other.own_type = _TYPES[family, grouping, family == 'line' and other.marker != _marker_none()]
    if family == 'column' and (changed or FAMILIES[before] != (family, grouping)) and source[1] == 'column' and bool(grouping) != bool(FAMILIES[before][1]):
        # Stacking a column group overlaps its bars fully, and unstacking it parts them, as for a chart of one type.
        set_group_value(state, target, 'overlap', 100 if grouping else 0)
    settle(state)


def set_axis_group(state: ChartState, item: SeriesState, wanted: int) -> None:
    """A series on the other axes. Anything but 1 or 2 is ignored, and the last on the primary axes takes the others there."""
    from pyopenvba.apps.excel._chart_axes import secondary_states

    if wanted not in (1, 2) or wanted == item.axis_group:
        return
    if wanted == 2 or is_combo(state):
        _admit_base(state)
    family = FAMILIES[series_type(state, item)][0]
    grouping = FAMILIES[series_type(state, item)][1]
    _enter(state)
    for other in state.series:
        if other is not item and other.axis_group == wanted and other.own_type is not None and FAMILIES[other.own_type] != (family, grouping) and FAMILIES[other.own_type][0] == family:
            raise VBAUnsupportedError('A series moving to a group that stacks differently is not implemented')
    if len(state.series) > 1 and wanted == 2 and all(other.axis_group == 2 for other in state.series if other is not item):
        for other in state.series:
            other.axis_group = 1
        settle(state)
        return
    if len(state.series) == 1 and wanted == 2:
        settle(state)
        return
    _carry(state, (item.axis_group, family), (wanted, family))
    item.axis_group = wanted
    if wanted == 2:
        secondary_states(state)
    settle(state)


def flatten(state: ChartState) -> None:
    """Chart.ChartType over a combo chart: one group, each keeping its number, in the order the groups had them."""
    found = groups(state)
    if len(found) < 2:
        return
    state.plot_sequence = [item for group in found for item in group.members]
    # The one group takes over from the last, as any group a series lands in does.
    last = found[-1].key
    state.gap_width = cast(int, group_value(state, last, 'gap_width'))
    state.overlap = cast(int, group_value(state, last, 'overlap')) if last[1] == 'column' else 0
    state.group_settings = {}
    state.secondary_axes = {}
    for item in state.series:
        item.own_type, item.axis_group = None, 1


def joined_type(state: ChartState) -> tuple[int | None, str]:
    """What a new series is in a combo chart: of the first primary group."""
    if not is_combo(state):
        return None, ''
    first = next(group for group in groups(state) if group.axis == 1)
    return _TYPES[first.family, first.grouping, False], first.family


def carried_over(state: ChartState, old: list[tuple[int | None, int, int]], new: list[SeriesState]) -> None:
    """SetSourceData keeps each place's type and axes; a series past those is refused, its group not having been measured."""
    for item, (own, axis, marker) in zip(new, old, strict=False):
        item.own_type, item.axis_group = own, axis
        if own is not None:
            item.marker = marker
    settle(state)


def regrouped(xml: str, state: ChartState, write_group: Callable[[str, str, str, int, int], str], grouped: Callable[[str, ChartState, str], str],
              value_format: str = 'General') -> str:
    """The part with its series in their groups and the secondary axes after the primary. The rest of the writer sees one group
    with every series in order; this comes last, so that none of it has to know."""
    from pyopenvba.apps.excel._chart_axes_file import generated
    from pyopenvba.apps.excel._chart_xml import secondary_axes_markup

    written = re.findall(r'<c:ser>.*?</c:ser>', xml, re.DOTALL)
    by_item = {id(item): text for item, text in zip(state.series, written, strict=True)}
    start = xml.index('<c:layout/>', xml.index('<c:plotArea>')) + len('<c:layout/>')
    axes = re.search(r'<c:(?:catAx|dateAx|valAx)>', xml)
    stop = axes.start() if axes is not None else xml.index('</c:plotArea>')
    if not is_combo(state):
        sequence = state.plot_sequence or []
        plot = xml[start:stop]
        first = plot.index('<c:ser>')
        last = plot.rindex('</c:ser>') + len('</c:ser>')
        return xml[:start] + plot[:first] + ''.join(by_item[id(item)] for item in sequence) + plot[last:] + xml[stop:]
    second = _secondary_ids(state) if state.secondary_axes else state.axis_ids
    plot = ''
    for group in groups(state):
        kind = 'lineMarkers' if group.family == 'line' else group.family
        ids = state.axis_ids if group.axis == 1 else second
        members = [by_item[id(item)] for item in group.members]
        if group.family == 'area':
            # An area with nothing to draw has no line (as a series whose values are lost, tests/fixtures/chart_lost_cells.json).
            members = [one if '<c:val>' in one and '<c:val><c:numRef><c:f></c:f>' not in one else re.sub(r'(<c:order val="[0-9]+"/>(?:<c:tx>.*?</c:tx>)?)', r'\1<c:spPr><a:ln w="25400"><a:noFill/></a:ln></c:spPr>', one, count=1, flags=re.DOTALL) for one in members]
        text = write_group(kind, group.grouping, ''.join(members), *ids)
        own: Any = {name: group_value(state, group.key, name) for name in ('gap_width', 'overlap', 'drop_lines', 'hi_lo_lines', 'up_down_bars', 'series_lines', 'vary')}
        plot += grouped(text, replace(state, chart_type=_TYPES[group.family, group.grouping, False], **own), kind)
    xml = xml[:start] + plot + xml[stop:]
    # Areas alone sit their axes on the categories; beside anything else the categories have room (case 11 of the fixture).
    if not all(group.family == 'area' for group in groups(state) if group.axis == 1):
        xml = xml.replace('<c:crossBetween val="midCat"/>', '<c:crossBetween val="between"/>', 1)
    if not state.secondary_axes:
        return xml
    value, category = state.secondary_axes[2], state.secondary_axes[1]
    between = 'midCat' if all(group.family == 'area' for group in groups(state) if group.axis == 2) else 'between'
    extra = generated(secondary_axes_markup(second[1], second[0], at_top=category.present and value.far_side, between=between), replace(state, axes=state.secondary_axes), lambda _formula: '')
    if value_format != 'General':
        extra = extra.replace('<c:numFmt formatCode="General" sourceLinked="1"/>', f'<c:numFmt formatCode="{_attribute(value_format)}" sourceLinked="1"/>', 1)
    if not category.present and value.far_side:
        extra = re.sub(r'(<c:catAx>.*?)<c:crosses(?:At)? val="[^"]*"/>', r'\1', extra, count=1, flags=re.DOTALL)
    ends = [xml.rfind(f'</c:{tag}>') + len(f'</c:{tag}>') for tag in ('catAx', 'dateAx', 'valAx') if f'</c:{tag}>' in xml]
    return xml[:max(ends)] + extra + xml[max(ends):]


def _attribute(text: str) -> str:
    from pyopenvba._xml import escape

    return escape(text)


def _secondary_ids(state: ChartState) -> tuple[int, int]:
    return state.secondary_axes[1].identifier, state.secondary_axes[2].identifier


def own_type_of(element: str, grouping: str, marked: bool) -> int:
    """The type a group's element in a file gives a series."""
    family = {'barChart': 'column', 'lineChart': 'line', 'areaChart': 'area'}[element]
    stacking = '' if grouping in ('clustered', 'standard') else grouping
    return _TYPES[family, stacking, family == 'line' and marked]


def family_of_type(wanted: int) -> str:
    return FAMILIES[wanted][0]


__all__ = ['FAMILIES', 'Group', 'KNOWN_TYPES', 'SERIES_ALONGSIDE', 'carried_over', 'flatten', 'group_value', 'groups', 'is_combo', 'joined_type', 'plot_order', 'regrouped', 'series_type', 'set_axis_group', 'set_group_value', 'set_series_type', 'settle']
