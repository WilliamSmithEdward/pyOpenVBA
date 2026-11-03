"""Chart.SetElement, measured on column, line-with-markers and pie charts (tests/fixtures/chart_set_element.json).

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
_Y, _BOTH = 1, 1
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
        for series in _each_series(chart):
            series.vba_call('ErrorBar', [VBAInt(_Y, 'Long'), VBAInt(_BOTH, 'Long'), VBAInt(kind, 'Long'), amount])
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
    'DataLabelNone': _no_labels, 'DataLabelShow': _labels(None), 'DataLabelCenter': _labels('Center'),
    'DataLabelCallout': _refused('A data label callout'),
}
#: What a chart without a pie takes.
_WITH_AXES: dict[str, Step] = {
    'DataTableNone': lambda chart: chart.vba_set('HasDataTable', False),
    'DataTableShow': lambda chart: _data_table(chart, False),
    'DataTableWithLegendKeys': lambda chart: _data_table(chart, True),
    'ErrorBarNone': _no_error_bars, 'ErrorBarPercentage': _error_bars(_PERCENT, VBAInt(5, 'Long')),
    'ErrorBarStandardDeviation': _error_bars(_ST_DEV, VBAInt(1, 'Long')),
    'ErrorBarStandardError': _error_bars(_ST_ERROR),
    'PlotAreaShow': _nothing, 'PlotAreaNone': _refused('Taking away the plot area'), 'TrendlineNone': _nothing,
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
_BY_KIND: dict[str, dict[str, Step]] = {
    'column': {'DataLabelInsideEnd': _labels('InsideEnd'), 'DataLabelInsideBase': _labels('InsideBase'),
               'DataLabelOutSideEnd': _labels('OutSideEnd')},
    'lineMarkers': {'DataLabelBottom': _labels('Bottom'), 'DataLabelLeft': _labels('Left'),
                    'DataLabelRight': _labels('Right'), 'DataLabelTop': _labels('Top'),
                    'LineDropLine': _on_group(HasDropLines=True), 'LineHiLoLine': _on_group(HasHiLoLines=True),
                    'LineNone': _on_group(HasDropLines=False, HasHiLoLines=False),
                    'UpDownBarsShow': _on_group(HasUpDownBars=True), 'UpDownBarsNone': _on_group(HasUpDownBars=False)},
    'pie': {'DataLabelInsideEnd': _labels('InsideEnd'), 'DataLabelOutSideEnd': _labels('OutSideEnd'),
            'DataLabelBestFit': _labels('BestFit')},
}
#: These raise 1004 on a chart with axes, and the usual failure on a pie.
_ANOTHER_FAILURE = ('PrimaryCategoryAxisLogScale', 'TrendlineAddExponential', 'TrendlineAddLinear',
                    'TrendlineAddLinearForecast', 'TrendlineAddTwoPeriodMovingAverage')


def set_element(chart: Chart, value: object) -> object:
    """Chart.SetElement."""
    from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type
    from pyopenvba.interpreter._constants_data import OFFICE_CONSTANTS

    wanted = int(to_integer(value, 'Long'))
    chart.admit_change()
    kind = CHART_TYPES.get(shown_type(chart.state), '')
    if kind not in _BY_KIND:
        raise unmeasured('SetElement on this kind of chart')
    name = next((key[len('msoElement'):] for key, number in OFFICE_CONSTANTS.items()
                 if number == wanted and key.startswith('msoElement')), '')
    step = _NAMES.get(name) or _BY_KIND[kind].get(name) or (None if kind == 'pie' else _WITH_AXES.get(name))
    if step is None:
        if name in _ANOTHER_FAILURE and kind != 'pie':
            raise error(1004, 'Application-defined or object-defined error')
        raise error(-2147467259, 'Automation error\nUnspecified error')
    step(chart)
    return True
