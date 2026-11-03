"""Chart.SetElement, measured on eight kinds of chart (tests/fixtures/chart_set_element*.json).

Each element does what the members it names do, so this is a table of those members. An element a chart has no such
thing to apply to fails, and so does every element of the secondary axes, which the model does not have."""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, cast

from pyopenvba.apps.excel._chart_format import unmeasured
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
    return list(_object(chart.vba_call('SeriesCollection')).vba_items())


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
    step = _step(name, kind)
    if step is None:
        if kind in _ANOTHER_FAILURE.get(name, ()):
            raise error(1004, 'Application-defined or object-defined error')
        raise error(-2147467259, 'Automation error\nUnspecified error')
    step(chart)
    return True
