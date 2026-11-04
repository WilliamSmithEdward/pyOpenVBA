"""Chart.SetElement, measured on eight kinds of chart (tests/fixtures/chart_set_element*.json).

Each element does what the members it names do, so this is a table of those members. An element a chart has no such
thing to apply to fails, and so does every element of the secondary axes, which the model does not have."""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, cast

from pyopenvba.apps.excel._chart_format import unmeasured
from pyopenvba.apps.excel._chart_xml import LabelLook
from pyopenvba.interpreter._values import MISSING, VBAInt, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Chart

_CATEGORY, _VALUE = 1, 2
_NONE = -4142
_UPWARD = -4171
_VERTICAL = -4166
_LOGARITHMIC = -4133
_THOUSANDS, _MILLIONS, _BILLIONS = -3, -6, -9
_X, _Y, _BOTH = -4168, 1, 1
_SCATTER = -4169
_PERCENT, _ST_DEV, _ST_ERROR = 2, -4155, 4
_LEGEND_PLACES = {'Right': -4152, 'Left': -4131, 'Top': -4160, 'Bottom': -4107}
_LABEL_PLACES = {'Center': -4108, 'InsideEnd': 3, 'InsideBase': 4, 'OutSideEnd': 2, 'BestFit': 5, 'Bottom': 1,
                 'Left': -4131, 'Right': -4152, 'Top': 0}

Step = Callable[['Chart'], None]


def _object(found: object) -> Any:
    """A member's object, which these steps set members of by name."""
    return cast(Any, found)


def _axis(chart: Chart, kind: int) -> Any:
    return _object(chart.vba_call('Axes', [VBAInt(kind, 'Long')]))


def _on_axis(kind: int, *settings: tuple[str, object]) -> Step:
    def step(chart: Chart) -> None:
        found = _axis(chart, kind)
        for name, value in settings:
            found.vba_set(name, value)
    return step


def _shown(kind: int, wanted: bool) -> Step:
    def step(chart: Chart) -> None:
        chart.vba_set('HasAxis', wanted, [VBAInt(kind, 'Long'), MISSING])
    return step


def _axis_title(kind: int, wanted: bool, orientation: int | None = None) -> Step:
    def step(chart: Chart) -> None:
        found = _axis(chart, kind)
        found.vba_set('HasTitle', wanted)
        if orientation is not None:
            title = found.vba_get('AxisTitle')
            title.vba_set('Orientation', VBAInt(orientation, 'Long'))
    return step


def _title(chart: Chart, overlaid: bool) -> None:
    chart.vba_set('HasTitle', True)
    _object(chart.vba_get('ChartTitle')).vba_set('IncludeInLayout', not overlaid)


def _no_title(chart: Chart) -> None:
    if chart.state.title is not None:
        chart.vba_set('HasTitle', False)
        chart.state.title_deleted = True


def _legend(place: str, overlaid: bool) -> Step:
    def step(chart: Chart) -> None:
        chart.vba_set('HasLegend', True)
        legend = _object(chart.vba_get('Legend'))
        legend.vba_set('Position', VBAInt(_LEGEND_PLACES[place], 'Long'))
        legend.vba_set('IncludeInLayout', not overlaid)
    return step


def _each_series(chart: Chart) -> list[Any]:
    return list(_object(chart.vba_call('FullSeriesCollection')).vba_items())


def _labels(place: str | None) -> Step:
    def step(chart: Chart) -> None:
        chart.state.element_labels = True, None if place is None else _LABEL_PLACES[place]
        for series in _each_series(chart):
            series.vba_set('HasDataLabels', True)
            series.entry.leader_lines = True
            if place is not None:
                labels = series.vba_call('DataLabels')
                labels.vba_set('Position', VBAInt(_LABEL_PLACES[place], 'Long'))
    return step


def _no_labels(chart: Chart) -> None:
    chart.state.element_labels = False, None
    for series in _each_series(chart):
        series.vba_set('HasDataLabels', False)


def _error_bars(kind: int, amount: object = MISSING) -> Step:
    def step(chart: Chart) -> None:
        # A scatter chart takes them both ways, X first in the file, where Series.ErrorBar keeps Y first.
        directions = (_X, _Y) if chart.state.chart_type == _SCATTER else (_Y,)
        for series in _each_series(chart):
            for direction in directions:
                series.vba_call('ErrorBar', [VBAInt(direction, 'Long'), VBAInt(_BOTH, 'Long'), VBAInt(kind, 'Long'),
                                             amount])
            series.entry.error_bars = sorted(series.entry.error_bars, key=lambda bar: bar.direction != 'x')
    return step


def _no_error_bars(chart: Chart) -> None:
    for series in _each_series(chart):
        series.vba_set('HasErrorBars', False)


def _on_group(**settings: bool) -> Step:
    def step(chart: Chart) -> None:
        group = _object(chart.vba_call('ChartGroups', [VBAInt(1, 'Long')]))
        for name, value in settings.items():
            group.vba_set(name, value)
    return step


def _data_table(chart: Chart, keys: bool) -> None:
    chart.vba_set('HasDataTable', True)
    chart.state.data_table_keys = keys


def _nothing(chart: Chart) -> None:
    del chart


def _refused(what: str) -> Step:
    def step(chart: Chart) -> None:
        del chart
        raise unmeasured(what)
    return step


def _grid(kind: int, major: bool, minor: bool) -> Step:
    return _on_axis(kind, ('HasMajorGridlines', major), ('HasMinorGridlines', minor))


#: What every kind measured takes, by the name after msoElement.
_NAMES: dict[str, Step] = {
    'ChartTitleNone': _no_title,
    'ChartTitleAboveChart': lambda chart: _title(chart, False),
    'ChartTitleCenteredOverlay': lambda chart: _title(chart, True),
    'LegendNone': lambda chart: chart.vba_set('HasLegend', False),
    'LegendRight': _legend('Right', False), 'LegendTop': _legend('Top', False), 'LegendLeft': _legend('Left', False),
    'LegendBottom': _legend('Bottom', False), 'LegendRightOverlay': _legend('Right', True),
    'LegendLeftOverlay': _legend('Left', True),
    'DataLabelNone': _no_labels, 'DataLabelShow': _labels(None),
    'DataLabelCallout': _refused('A data label callout'),
}
#: What a chart without a pie takes.
_WITH_AXES: dict[str, Step] = {
    'ErrorBarNone': _no_error_bars, 'ErrorBarPercentage': _error_bars(_PERCENT, VBAInt(5, 'Long')),
    'ErrorBarStandardDeviation': _error_bars(_ST_DEV, VBAInt(1, 'Long')),
    'ErrorBarStandardError': _error_bars(_ST_ERROR),
    'PlotAreaShow': _nothing, 'PlotAreaNone': _refused('Taking away the plot area'),
    'PrimaryCategoryAxisNone': _shown(_CATEGORY, False), 'PrimaryCategoryAxisShow': _shown(_CATEGORY, True),
    'PrimaryValueAxisNone': _shown(_VALUE, False), 'PrimaryValueAxisShow': _shown(_VALUE, True),
    'PrimaryCategoryAxisReverse': _on_axis(_CATEGORY, ('ReversePlotOrder', True)),
    'PrimaryCategoryAxisWithoutLabels': _on_axis(_CATEGORY, ('MajorTickMark', VBAInt(_NONE, 'Long')),
                                                 ('TickLabelPosition', VBAInt(_NONE, 'Long'))),
    'PrimaryValueAxisLogScale': _on_axis(_VALUE, ('ScaleType', VBAInt(_LOGARITHMIC, 'Long'))),
    'PrimaryValueAxisThousands': _on_axis(_VALUE, ('DisplayUnit', VBAInt(_THOUSANDS, 'Long'))),
    'PrimaryValueAxisMillions': _on_axis(_VALUE, ('DisplayUnit', VBAInt(_MILLIONS, 'Long'))),
    'PrimaryValueAxisBillions': _on_axis(_VALUE, ('DisplayUnit', VBAInt(_BILLIONS, 'Long'))),
    'PrimaryCategoryAxisTitleNone': _axis_title(_CATEGORY, False),
    'PrimaryCategoryAxisTitleAdjacentToAxis': _axis_title(_CATEGORY, True),
    'PrimaryCategoryAxisTitleBelowAxis': _axis_title(_CATEGORY, True),
    'PrimaryCategoryAxisTitleHorizontal': _axis_title(_CATEGORY, True, 0),
    'PrimaryCategoryAxisTitleRotated': _axis_title(_CATEGORY, True, _UPWARD),
    'PrimaryCategoryAxisTitleVertical': _axis_title(_CATEGORY, True, _VERTICAL),
    'PrimaryValueAxisTitleNone': _axis_title(_VALUE, False),
    # One number in Office's enumeration with the name after it, so it takes the title away.
    'PrimaryValueAxisTitleAdjacentToAxis': _axis_title(_VALUE, False),
    'PrimaryValueAxisTitleBelowAxis': _axis_title(_VALUE, True),
    'PrimaryValueAxisTitleHorizontal': _axis_title(_VALUE, True, 0),
    'PrimaryValueAxisTitleRotated': _axis_title(_VALUE, True, _UPWARD),
    'PrimaryValueAxisTitleVertical': _axis_title(_VALUE, True, _VERTICAL),
    'PrimaryCategoryGridLinesNone': _grid(_CATEGORY, False, False),
    'PrimaryCategoryGridLinesMajor': _grid(_CATEGORY, True, False),
    'PrimaryCategoryGridLinesMinor': _grid(_CATEGORY, False, True),
    'PrimaryCategoryGridLinesMinorMajor': _grid(_CATEGORY, True, True),
    'PrimaryValueGridLinesNone': _grid(_VALUE, False, False),
    'PrimaryValueGridLinesMajor': _grid(_VALUE, True, False),
    'PrimaryValueGridLinesMinor': _grid(_VALUE, False, True),
    'PrimaryValueGridLinesMinorMajor': _grid(_VALUE, True, True),
}
_LINES = ('line', 'lineMarkers')
#: What takes only some kinds of chart: its step, and those kinds.
_ONLY: dict[str, tuple[Step, tuple[str, ...]]] = {
    'DataLabelCenter': (_labels('Center'), ('column', 'bar', *_LINES, 'scatter', 'pie')),
    'DataLabelInsideEnd': (_labels('InsideEnd'), ('column', 'bar', 'pie')),
    'DataLabelOutSideEnd': (_labels('OutSideEnd'), ('column', 'bar', 'pie')),
    'DataLabelInsideBase': (_labels('InsideBase'), ('column', 'bar')),
    'DataLabelBestFit': (_labels('BestFit'), ('pie',)),
    'DataLabelBottom': (_labels('Bottom'), (*_LINES, 'scatter')),
    'DataLabelLeft': (_labels('Left'), (*_LINES, 'scatter')),
    'DataLabelRight': (_labels('Right'), (*_LINES, 'scatter')),
    'DataLabelTop': (_labels('Top'), (*_LINES, 'scatter')),
    'LineDropLine': (_on_group(HasDropLines=True), (*_LINES, 'area')),
    'LineNone': (_on_group(HasDropLines=False, HasHiLoLines=False), (*_LINES, 'area')),
    'LineHiLoLine': (_on_group(HasHiLoLines=True), _LINES),
    'UpDownBarsShow': (_on_group(HasUpDownBars=True), _LINES),
    'UpDownBarsNone': (_on_group(HasUpDownBars=False), _LINES),
    'TrendlineNone': (_nothing, ('column', 'bar', *_LINES, 'scatter')),
    'DataTableNone': (lambda chart: chart.vba_set('HasDataTable', False), ('column', 'bar', *_LINES, 'area')),
    'DataTableShow': (lambda chart: _data_table(chart, False), ('column', 'bar', *_LINES, 'area')),
    'DataTableWithLegendKeys': (lambda chart: _data_table(chart, True), ('column', 'bar', *_LINES, 'area')),
    'PrimaryCategoryAxisLogScale': (_on_axis(_CATEGORY, ('ScaleType', VBAInt(_LOGARITHMIC, 'Long'))), ('scatter',)),
    'PrimaryCategoryAxisThousands': (_on_axis(_CATEGORY, ('DisplayUnit', VBAInt(_THOUSANDS, 'Long'))), ('scatter',)),
    'PrimaryCategoryAxisMillions': (_on_axis(_CATEGORY, ('DisplayUnit', VBAInt(_MILLIONS, 'Long'))), ('scatter',)),
    'PrimaryCategoryAxisBillions': (_on_axis(_CATEGORY, ('DisplayUnit', VBAInt(_BILLIONS, 'Long'))), ('scatter',)),
}
_NO_AXES = ('pie', 'doughnut')
_KINDS = ('column', 'bar', *_LINES, 'scatter', 'area', *_NO_AXES)
#: What raises 1004 on these kinds, where the others have the usual failure.
_ANOTHER_FAILURE: dict[str, tuple[str, ...]] = {
    'PrimaryCategoryAxisLogScale': ('column', 'bar', *_LINES, 'area'),
    'TrendlineAddExponential': ('column', 'bar', *_LINES, 'scatter'),
    'TrendlineAddLinear': ('column', 'bar', *_LINES, 'scatter'),
    'TrendlineAddLinearForecast': ('column', 'bar', *_LINES, 'scatter'),
    'TrendlineAddTwoPeriodMovingAverage': ('column', 'bar', *_LINES, 'scatter'),
}


def _step(name: str, kind: str) -> Step | None:
    if name in _ONLY:
        step, kinds = _ONLY[name]
        return step if kind in kinds else None
    if name in _NAMES:
        return _NAMES[name]
    return None if kind in _NO_AXES else _WITH_AXES.get(name)


def _secondary_element(chart: Chart, name: str) -> object:
    """The elements of a combo chart's secondary axes that were measured (tests/fixtures/combo_charts.json)."""
    from pyopenvba.apps.excel._chart_axes import axis
    from pyopenvba.interpreter._values import VBAInt

    two = VBAInt(2, 'Long')
    if name == 'SecondaryValueAxisTitleAdjacentToAxis':
        axis(chart, VBAInt(2, 'Long'), two).vba_set('HasTitle', True)
    elif name == 'SecondaryCategoryAxisShow':
        chart.vba_set('HasAxis', True, [VBAInt(1, 'Long'), two])
    elif name == 'SecondaryValueAxisNone':
        chart.vba_set('HasAxis', False, [VBAInt(2, 'Long'), two])
    else:
        raise unmeasured('This SetElement on secondary axes')
    return True


def set_element(chart: Chart, value: object) -> object:
    """Chart.SetElement."""
    from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type
    from pyopenvba.interpreter._constants_data import OFFICE_CONSTANTS

    wanted = int(to_integer(value, 'Long'))
    chart.admit_change()
    kind = CHART_TYPES.get(shown_type(chart.state), '')
    if kind not in _KINDS:
        raise unmeasured('SetElement on this kind of chart')
    name = next((key[len('msoElement'):] for key, number in OFFICE_CONSTANTS.items()
                 if number == wanted and key.startswith('msoElement')), '')
    if name.startswith('Secondary') and chart.state.secondary_axes:
        return _secondary_element(chart, name)
    step = _step(name, kind)
    if step is None:
        if kind in _ANOTHER_FAILURE.get(name, ()):
            raise error(1004, 'Application-defined or object-defined error')
        raise error(-2147467259, 'Automation error\nUnspecified error')
    step(chart)
    return True


#: ChartWizard's Gallery as a chart type: xlBar, xlColumn, xlLine and xlPie (tests/fixtures/chart_layouts.json).
_GALLERIES = {2: 57, 3: 51, 4: 65, 5: 5}
_UNKNOWN_GALLERY = 999


def chart_wizard(chart: Chart, source: object, gallery: object, shape: object, plot_by: object, category_labels: object,
                 series_labels: object, has_legend: object, title: object, category_title: object,
                 value_title: object) -> object:
    """Chart.ChartWizard: what the members it names do. A gallery also takes away the value axis' major gridlines, and
    a line gallery sets the axis to cross at the categories."""
    chart.admit_change()
    if category_labels is not MISSING or series_labels is not MISSING:
        raise unmeasured('ChartWizard with CategoryLabels or SeriesLabels')
    if shape is not MISSING and int(to_integer(shape, 'Long')) != 1:
        raise unmeasured('ChartWizard with a Format but 1')
    if source is not MISSING:
        # With no PlotBy it keeps the chart's, where SetSourceData would work one out from the range.
        chart.vba_call('SetSourceData', [source, chart.vba_get('PlotBy') if plot_by is MISSING else plot_by])
    elif plot_by is not MISSING and int(to_integer(plot_by, 'Long')) != chart.vba_get('PlotBy'):
        raise unmeasured('ChartWizard changing PlotBy with no Source')
    if gallery is not MISSING:
        wanted = int(to_integer(gallery, 'Long'))
        if wanted in _GALLERIES:
            chart.vba_set('ChartType', VBAInt(_GALLERIES[wanted], 'Long'))
            if wanted != 5:
                _axis(chart, _VALUE).vba_set('HasMajorGridlines', False)
            if wanted == 4:
                _axis(chart, _CATEGORY).vba_set('AxisBetweenCategories', False)
        elif wanted != _UNKNOWN_GALLERY:
            raise unmeasured('This ChartWizard gallery')
    if has_legend is not MISSING:
        chart.vba_set('HasLegend', has_legend)
    if title is not MISSING:
        chart.vba_set('HasTitle', True)
        _object(chart.vba_get('ChartTitle')).vba_set('Text', title)
    for kind, text in ((_CATEGORY, category_title), (_VALUE, value_title)):
        if text is not MISSING:
            found = _axis(chart, kind)
            found.vba_set('HasTitle', True)
            found.vba_get('AxisTitle').vba_set('Text', text)
    return True

_LAST_TICK = 3
_FRESH_OVERLAP = {'line': 100, 'lineMarkers': 100, 'scatter': 100, 'area': 100}


def _layout_axes(chart: Chart, spec: dict[str, object], kind: str) -> None:
    for number, name in ((_VALUE, 'value'), (_CATEGORY, 'category')):
        found = _axis(chart, number)
        found.vba_set('HasMajorGridlines', spec.get(f'{name}_major', number == _VALUE))
        found.vba_set('HasMinorGridlines', spec.get(f'{name}_minor', False))
        found.vba_set('HasTitle', spec.get(f'{name}_title', False))
        found.vba_set('MajorTickMark', VBAInt(int(str(spec.get(f'{name}_tick', _LAST_TICK))), 'Long'))
    group = _object(chart.vba_call('ChartGroups', [VBAInt(1, 'Long')]))
    chart.state.gap_width = int(str(spec.get('gap', 150)))
    overlap = int(str(spec.get('overlap', _FRESH_OVERLAP.get(kind, 0))))
    # A chart without bars keeps it to read back alone.
    chart.state.overlap = overlap
    chart.state.layout_overlap = None if kind not in _FRESH_OVERLAP else overlap
    chart.vba_set('HasDataTable', spec.get('data_table', False))
    chart.state.data_table_keys = True
    chart.state.layout_value_look = str(spec.get('value_axis_look', ''))
    chart.state.layout_category_look = str(spec.get('category_axis_look', ''))
    if kind in ('line', 'lineMarkers', 'area'):
        group.vba_set('HasDropLines', spec.get('drop_lines', False))
    if kind in ('line', 'lineMarkers'):
        group.vba_set('HasHiLoLines', spec.get('hi_lo_lines', False))
        group.vba_set('HasUpDownBars', spec.get('up_down_bars', False))
    for number, name in ((_CATEGORY, 'category'), (_VALUE, 'value')):
        chart.vba_set('HasAxis', spec.get(f'{name}_axis', True), [VBAInt(number, 'Long'), MISSING])


def _layout_labels(chart: Chart, spec: dict[str, object]) -> None:
    from pyopenvba.apps.excel._charts import PointLabel

    every = _each_series(chart)
    labelled = every if spec.get('first_labels') else every[-1:] if spec.get('last_labels') else []
    place = spec.get('label_place')
    chart.state.element_labels = None
    chart.state.layout_chart_labels = str(spec.get('chart_labels', ''))
    for series in every:
        # Taking away labels and giving them leaves the leader lines as they were.
        entry = series.entry
        leader_lines = entry.leader_lines
        for point in entry.points.values():
            point.label, point.label_own = None, PointLabel()
        if entry.data_labels:
            series.vba_set('HasDataLabels', False)
        entry.label_look = LabelLook()
        entry.leader_lines = leader_lines
    for series in labelled:
        leader_lines = series.entry.leader_lines
        series.vba_set('HasDataLabels', True)
        series.entry.leader_lines = leader_lines
        labels = _object(series.vba_call('DataLabels'))
        for member_name, field_name in (('ShowCategoryName', 'show_category'), ('ShowSeriesName', 'show_series'),
                                        ('ShowPercentage', 'show_percent'), ('ShowValue', 'show_value')):
            if field_name in spec:
                labels.vba_set(member_name, spec[field_name])
        if place is not None:
            labels.vba_set('Position', VBAInt(int(str(place)), 'Long'))
    if spec.get('point_label') and not spec.get('first_labels'):
        for series in every:
            points = _object(series.vba_call('Points'))
            last = _object(points.vba_items()[-1])
            # A scatter chart names the point's category, where the others name the series.
            named = 'ShowSeriesName' if spec.get('point_series') else 'ShowCategoryName'
            last.vba_call('ApplyDataLabels', [], {named: True})
            if 'point_place' in spec:
                last.vba_get('DataLabel').vba_set('Position', VBAInt(int(str(spec['point_place'])), 'Long'))


def apply_layout(chart: Chart, layout: object, chart_type: object) -> object:
    """Chart.ApplyLayout (tests/fixtures/apply_layout.json). A layout gives the chart all that it says, what it does not
    going back to what a new chart has. The tick label position, the style, the series' looks and a typed title stay."""
    from pyopenvba.apps.excel._chart_layouts import LAYOUTS
    from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type

    wanted = int(to_integer(layout, 'Long'))
    chart.admit_change()
    kind = CHART_TYPES.get(shown_type(chart.state), '')
    if kind not in LAYOUTS:
        raise unmeasured('ApplyLayout on this kind of chart')
    if wanted not in LAYOUTS[kind]:
        raise error(5, 'Invalid procedure call or argument')
    if chart_type is not MISSING:
        raise unmeasured('ApplyLayout with a ChartType')
    spec = LAYOUTS[kind][wanted]
    if spec.get('trendlines'):
        raise unmeasured('A layout with trendlines')
    if spec.get('title'):
        chart.vba_set('HasTitle', True)
    else:
        _no_title(chart)
    chart.vba_set('HasLegend', spec.get('legend', True))
    if spec.get('legend', True):
        legend = _object(chart.vba_get('Legend'))
        legend.vba_set('Position', VBAInt(int(str(spec.get('legend_place', -4152))), 'Long'))
    if kind not in _NO_AXES:
        _layout_axes(chart, spec, kind)
    _layout_labels(chart, spec)
    return True
