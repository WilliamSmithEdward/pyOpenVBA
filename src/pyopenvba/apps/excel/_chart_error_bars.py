"""A chart series' error bars: Series.ErrorBar, Series.HasErrorBars and the ErrorBars object, as Excel sets, reads and
writes them (tests/fixtures/chart_error_bars.json).

A series has a Y bar and, in a scatter chart, an X bar. The file writes each as an errBars element after the
series' data labels and before its categories, and names the direction only in line and scatter charts.
"""
from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_integer, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Series

#: ErrorBar's Direction: xlX and xlY.
X_DIRECTION, Y_DIRECTION = -4168, 1
#: ErrorBar's Include, as the file spells it; xlErrorBarIncludeNone (-4142) takes the bars away.
INCLUDES = {1: 'both', 2: 'plus', 3: 'minus'}
INCLUDE_NONE = -4142
#: ErrorBar's Type, as the file spells it: fixed value, percent, standard deviation, standard error and custom.
KINDS = {1: 'fixedVal', 2: 'percentage', -4155: 'stdDev', 4: 'stdErr', -4114: 'cust'}
#: The kinds of chart whose file names a bar's direction.
DIRECTED = ('line', 'lineMarkers', 'scatter')
#: The chart kinds the model writes error bars for.
BARRED = ('column', 'bar', 'line', 'lineMarkers', 'scatter', 'pie', 'doughnut')


@dataclass
class ErrorBarState:
    """One of a series' error bars."""

    direction: str = 'y'
    include: str = 'both'
    kind: str = 'fixedVal'
    amount: float | None = None
    #: A custom bar's plus and minus values: a reference, or the values given.
    plus: tuple[str, list[object]] = ('', [])
    minus: tuple[str, list[object]] = ('', [])
    no_cap: bool = False


def _default_amount(series: Series) -> float:
    """The fixed value HasErrorBars gives a new bar: 1 on a chart with no value axis to measure by, scatter charts
    among them; otherwise the value axis' major unit (tests/fixtures/chart_error_bars.json)."""
    from pyopenvba.apps.excel._chart_axes import states
    from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type

    kind = CHART_TYPES.get(shown_type(series.chart.state), '')
    if kind in ('scatter', 'pie', 'doughnut'):
        return 1.0
    major = states(series.chart)[2].major
    if major is None:
        raise VBAUnsupportedError('HasErrorBars with an automatic value axis scale requires native layout')
    return major


def _kind_of(series: Series) -> str:
    from pyopenvba.apps.excel._charts import CHART_TYPES, shown_type

    kind = CHART_TYPES.get(shown_type(series.chart.state), '')
    if kind not in BARRED:
        raise VBAUnsupportedError(f'Error bars on a {kind or "this"} chart are not implemented')
    return kind


def has_error_bars(series: Series) -> bool:
    series.position()
    return bool(series.entry.error_bars)


def set_has_error_bars(series: Series, wanted: bool) -> None:
    """HasErrorBars: on, a bar both ways at a fixed value, two in a scatter chart; off, none at all."""
    series.position()
    series.chart.admit_change()
    if not wanted:
        series.chart.changed()
        series.entry.error_bars = []
        return
    if series.entry.error_bars:
        return
    kind = _kind_of(series)
    amount = _default_amount(series)
    series.chart.changed()
    series.entry.error_bars = [ErrorBarState(direction=direction, amount=amount)
                               for direction in (('y', 'x') if kind == 'scatter' else ('y',))]


def error_bar(series: Series, direction: object, include: object, kind: object, amount: object,
              minus_values: object) -> None:
    """Series.ErrorBar: the series' bar in one direction, made anew; an X bar in a chart that has none is its Y bar.
    A value the method does not take raises 1004 (tests/fixtures/chart_error_bars.json)."""
    from pyopenvba.apps.excel._charts import series_input

    series.position()
    series.chart.admit_change()
    if MISSING in (direction, include, kind) or NULL in (direction, include, kind):
        raise error(1004, 'Unable to get the ErrorBar property of the Series class')
    chart_kind = _kind_of(series)
    way = 'x' if int(to_integer(direction, 'Long')) == X_DIRECTION and chart_kind == 'scatter' else 'y'
    included = int(to_integer(include, 'Long'))
    typed = int(to_integer(kind, 'Long'))
    if included != INCLUDE_NONE and included not in INCLUDES or typed not in KINDS:
        raise error(1004, 'Unable to get the ErrorBar property of the Series class')
    made = ErrorBarState(direction=way, include=INCLUDES.get(included, 'both'), kind=KINDS[typed])
    if made.kind == 'cust':
        if amount is MISSING or minus_values is MISSING or made.include != 'both':
            raise VBAUnsupportedError('Custom error bars without both plus and minus values are not implemented')
        made.plus = _custom(amount, series_input, series)
        made.minus = _custom(minus_values, series_input, series)
    elif made.kind != 'stdErr':
        if amount is MISSING:
            raise error(1004, 'Unable to get the ErrorBar property of the Series class')
        made.amount = float(to_number(amount))
        if made.amount < 0:
            raise error(1004, 'Unable to get the ErrorBar property of the Series class')
    kept = [bar for bar in series.entry.error_bars if bar.direction != way]
    series.chart.changed()
    series.entry.error_bars = sorted(kept + ([made] if included != INCLUDE_NONE else []),
                                     key=lambda bar: bar.direction != 'y')


def _custom(value: object, series_input: Callable[..., tuple[str, list[object]]],
            series: Series) -> tuple[str, list[object]]:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.interpreter._values import VBAArray

    if isinstance(value, (Range, VBAArray)):
        return series_input(value, sheet=series.chart.sheet)
    return '', [float(to_number(value))]


def markup(bars: Sequence[ErrorBarState], kind: str, values: Callable[[str, tuple[str, list[object]]], str]) -> str:
    """The series' errBars elements, Y first. ``values`` writes a custom bar's plus or minus element from its
    reference, or from the values given where it has none."""
    written: list[str] = []
    for bar in bars:
        if kind != 'scatter' and bar.direction != 'y':
            continue
        direction = f'<c:errDir val="{bar.direction}"/>' if kind in DIRECTED else ''
        body = (f'{direction}<c:errBarType val="{bar.include}"/><c:errValType val="{bar.kind}"/>'
                f'<c:noEndCap val="{int(bar.no_cap)}"/>')
        if bar.kind == 'cust':
            body += values('plus', bar.plus) + values('minus', bar.minus)
        elif bar.kind != 'stdErr' and bar.amount is not None:
            body += f'<c:val val="{_number_text(bar.amount)}"/>'
        written.append(f'<c:errBars>{body}</c:errBars>')
    return ''.join(written)


def _number_text(value: float) -> str:
    from pyopenvba.apps.excel._chart_xml import format_number

    return format_number(value)


class ErrorBars(Created):
    """A series' error bars, both ways at once."""

    vba_type_name = 'ErrorBars'

    def __init__(self, series: Series) -> None:
        self.series = series

    def _bars(self) -> list[ErrorBarState]:
        self.series.position()
        if not self.series.entry.error_bars:
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        return self.series.entry.error_bars

    @member
    def Parent(self) -> object:
        self._bars()
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Name(self) -> object:
        """"S", the series' place, and "E" (tests/fixtures/chart_error_bars.json)."""
        self._bars()
        return f'S{self.series.position()}E'

    @member
    def EndStyle(self) -> object:
        # xlCap 1, xlNoCap 2.
        return VBAInt(2 if self._bars()[0].no_cap else 1, 'Long')

    @setter('EndStyle')
    def _end_style(self, value: object) -> None:
        bars = self._bars()
        wanted = int(to_integer(value, 'Long'))
        self.series.chart.admit_change()
        if wanted not in (1, 2):
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        self.series.chart.changed()
        self.series.entry.error_bars = [replace(bar, no_cap=wanted == 2) for bar in bars]

    @method
    def Delete(self) -> object:
        self._bars()
        self.series.chart.admit_change()
        self.series.chart.changed()
        self.series.entry.error_bars = []
        return EMPTY


__all__ = ['ErrorBarState', 'ErrorBars', 'error_bar', 'has_error_bars', 'markup', 'set_has_error_bars']
