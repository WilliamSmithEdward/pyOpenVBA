"""A chart series' trendlines: Series.Trendlines, Trendlines.Add and the Trendline object, as Excel sets, reads and
writes them (tests/fixtures/chart_trendlines.json).

The file writes each trendline as a trendline element after the series' data labels and before its error bars.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._xml import escape_text
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.shapes._values import FillLine
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Series

#: XlTrendlineType, as the file spells it.
TYPES = {-4132: 'linear', 5: 'exp', -4133: 'log', 4: 'power', 3: 'poly', 6: 'movingAvg'}
#: How a trendline's automatic name starts, before the series' name in brackets.
LABELS = {'linear': 'Linear', 'exp': 'Expon.', 'log': 'Log.', 'power': 'Power', 'poly': 'Poly.'}
#: The kinds of chart that raise an automation error for a trendline, as measured.
REFUSED = (5, 52)


@dataclass
class TrendlineState:
    """One of a series' trendlines."""

    kind: str = 'linear'
    order: int = 2
    period: int = 2
    forward: float = 0.0
    backward: float = 0.0
    #: None while Excel works the intercept out.
    intercept: float | None = None
    #: None while the name is automatic.
    name: str | None = None
    equation: bool = False
    r_squared: bool = False
    #: What Border has set of its line (tests/fixtures/chart_borders.json).
    look: FillLine = field(default_factory=FillLine)


def _unable(what: str) -> None:
    raise error(1004, f'Unable to set the {what} property of the Trendline class')


def _admitted(series: Series) -> None:
    """A pie or stacked column chart takes no trendline; other kinds Excel refuses alike are not measured."""
    from pyopenvba.apps.excel._charts import CHART_TYPES, STACKED, shown_type

    shown = shown_type(series.chart.state)
    if shown in REFUSED:
        raise error(-2147467259, 'Automation error\nUnspecified error ')
    if shown in STACKED or CHART_TYPES.get(shown) == 'doughnut':
        raise VBAUnsupportedError('Trendlines on this kind of chart are not implemented')


def _checked_order(value: object) -> int:
    wanted = int(to_integer(value, 'Long'))
    if not 2 <= wanted <= 6:
        _unable('Order')
    return wanted


def _checked_period(series: Series, value: object) -> int:
    wanted = int(to_integer(value, 'Long'))
    if not 2 <= wanted < len(series.values()):
        _unable('Period')
    return wanted


def _checked_forward(value: object) -> float:
    wanted = float(to_number(value))
    if wanted < 0:
        raise error(-2147467259, 'Automation error\nUnspecified error ')
    return wanted


def markup(lines: list[TrendlineState]) -> str:
    from pyopenvba.apps.excel._chart_line_format import shape_properties

    written: list[str] = []
    for line in lines:
        body = f'<c:name>{escape_text(line.name)}</c:name>' if line.name is not None else ''
        body += shape_properties(line.look, 'rnd', False)
        body += f'<c:trendlineType val="{line.kind}"/>'
        if line.kind == 'poly':
            body += f'<c:order val="{line.order}"/>'
        if line.kind == 'movingAvg':
            body += f'<c:period val="{line.period}"/>'
        if line.forward:
            body += f'<c:forward val="{_number_text(line.forward)}"/>'
        if line.backward:
            body += f'<c:backward val="{_number_text(line.backward)}"/>'
        if line.intercept is not None:
            body += f'<c:intercept val="{_number_text(line.intercept)}"/>'
        body += f'<c:dispRSqr val="{int(line.r_squared)}"/><c:dispEq val="{int(line.equation)}"/>'
        if line.equation or line.r_squared:
            body += '<c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/></c:trendlineLbl>'
        written.append(f'<c:trendline>{body}</c:trendline>')
    return ''.join(written)


def _number_text(value: float) -> str:
    from pyopenvba.apps.excel._chart_xml import format_number

    return format_number(value)


class Trendlines(VBACollection, ItemDefault, Created):
    """A series' trendlines, in the order they were added."""

    vba_type_name = 'Trendlines'

    def __init__(self, series: Series) -> None:
        self.series = series

    def vba_items(self) -> list[object]:
        self.series.position()
        return [Trendline(self.series, line) for line in self.series.entry.trendlines]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        wanted = int(to_integer(index, 'Long'))
        if not 1 <= wanted <= len(items):
            raise error(1004, 'Unable to get the Trendlines property of the Series class')
        return items[wanted - 1]

    @member
    def Parent(self) -> object:
        self.series.position()
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @method
    def Add(self, Type: object = MISSING, Order: object = MISSING, Period: object = MISSING,  # noqa: N803
            Forward: object = MISSING, Backward: object = MISSING, Intercept: object = MISSING,
            DisplayEquation: object = MISSING, DisplayRSquared: object = MISSING, Name: object = MISSING) -> object:
        """A new trendline, linear unless Type says. Order counts only for a polynomial and Period only for a moving
        average; Backward is ignored (tests/fixtures/chart_trendlines.json)."""
        series = self.series
        series.position()
        series.chart.admit_change()
        _admitted(series)
        typed = -4132 if Type is MISSING else int(to_integer(Type, 'Long'))
        if typed not in TYPES:
            raise error(1004, 'Unable to add a trendline')
        line = TrendlineState(kind=TYPES[typed])
        if line.kind == 'poly' and Order is not MISSING:
            line.order = _checked_order(Order)
        if line.kind == 'movingAvg':
            line.period = _checked_period(series, 2 if Period is MISSING else Period)
        if Forward is not MISSING:
            line.forward = _checked_forward(Forward)
        if Intercept is not MISSING:
            line.intercept = float(to_number(Intercept))
        line.equation = DisplayEquation is not MISSING and to_bool(DisplayEquation)
        line.r_squared = DisplayRSquared is not MISSING and to_bool(DisplayRSquared)
        if Name is not MISSING:
            line.name = to_text(Name)
        series.chart.changed()
        series.entry.trendlines.append(line)
        return Trendline(series, line)


class _TrendlineLine:
    """The line of a trendline, for its Border."""

    def __init__(self, owner: Trendline) -> None:
        self.owner = owner

    def look(self) -> FillLine:
        self.owner.vba_get('Index', [])
        return self.owner.line.look

    def change(self) -> FillLine:
        self.owner.series.chart.changed()
        return self.look()


class TrendlineLabel(Created):
    """A trendline's equation and R-squared label, read while no window has drawn it (tests/fixtures/
    trendline_members.json). Changing it is not implemented."""

    vba_type_name = 'DataLabel'

    def __init__(self, owner: Trendline) -> None:
        self.owner = owner

    @member
    def Text(self) -> object:
        return ''

    @member
    def NumberFormat(self) -> object:
        return 'General'

    @member
    def Name(self) -> object:
        return f'Text S{self.owner.series.position()}T{self.owner.vba_get("Index", [])}'

    @method
    def Delete(self) -> object:
        self.owner.vba_set('DisplayEquation', False)
        self.owner.vba_set('DisplayRSquared', False)
        return EMPTY


class Trendline(Created):
    vba_type_name = 'Trendline'

    def __init__(self, series: Series, line: TrendlineState) -> None:
        self.series, self.line = series, line

    @method
    def GetProperty(self, IdsProperty: object = MISSING) -> object:
        raise error(-2147220480, "Method 'GetProperty' of object 'Trendline' failed")

    @method
    def SetProperty(self, IdsProperty: object = MISSING, Value: object = MISSING) -> object:
        raise error(-2147220480, "Method 'SetProperty' of object 'Trendline' failed")

    @method
    def ClearFormats(self) -> object:
        self._index()
        self.series.chart.changed()
        self.line.look = FillLine()
        return EMPTY

    @member
    def Format(self) -> object:
        from pyopenvba.apps.excel._chart_line_format import LineChartFormat

        self._index()
        return LineChartFormat(self, _TrendlineLine(self))

    @member
    def Border(self) -> object:
        from pyopenvba.apps.excel._chart_borders import TRENDLINE, ChartBorder

        self._index()
        return ChartBorder(_TrendlineLine(self), TRENDLINE, self, self.series.chart.sheet.book.stylesheet.colors)

    def _index(self) -> int:
        self.series.position()
        for index, line in enumerate(self.series.entry.trendlines):
            if line is self.line:
                return index + 1
        raise error(-2147467259, 'Automation error\nUnspecified error ')

    def _change(self) -> None:
        self._index()
        self.series.chart.admit_change()

    @method
    def Select(self) -> object:
        from pyopenvba.apps.excel._charts import select_element

        return select_element(self.series.chart, self, True)

    @member
    def Parent(self) -> object:
        self._index()
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Index(self) -> object:
        return VBAInt(self._index(), 'Long')

    @member
    def Type(self) -> object:
        self._index()
        return VBAInt(next(number for number, kind in TYPES.items() if kind == self.line.kind), 'Long')

    @setter('Type')
    def _type(self, value: object) -> None:
        self._change()
        wanted = int(to_integer(value, 'Long'))
        if wanted not in TYPES:
            _unable('Type')
        self.series.chart.changed()
        self.line.kind = TYPES[wanted]

    @member
    def Order(self) -> object:
        self._index()
        return VBAInt(self.line.order, 'Long')

    @setter('Order')
    def _order(self, value: object) -> None:
        # Only a polynomial keeps an order; any other trendline takes it silently (tests/fixtures/chart_trendlines.json).
        self._change()
        if self.line.kind != 'poly':
            return
        wanted = _checked_order(value)
        self.series.chart.changed()
        self.line.order = wanted

    @member
    def Period(self) -> object:
        self._index()
        return VBAInt(self.line.period, 'Long')

    @setter('Period')
    def _period(self, value: object) -> None:
        self._change()
        if self.line.kind != 'movingAvg':
            _unable('Period')
        wanted = _checked_period(self.series, value)
        self.series.chart.changed()
        self.line.period = wanted

    @member
    def Forward(self) -> object:
        self._index()
        return self.line.forward

    @setter('Forward')
    def _forward(self, value: object) -> None:
        self._change()
        wanted = _checked_forward(value)
        self.series.chart.changed()
        self.line.forward = wanted

    @member
    def Forward2(self) -> object:
        return self.vba_get('Forward', [])

    @setter('Forward2')
    def _forward2(self, value: object) -> None:
        self.vba_set('Forward', value)

    @member
    def Backward(self) -> object:
        self._index()
        return self.line.backward

    @setter('Backward')
    def _backward(self, value: object) -> None:
        self._change()
        wanted = _checked_forward(value)
        # Only a chart with a value axis both ways runs back more than half a category
        # (tests/fixtures/trendline_members.json).
        if wanted > 0.5 and self.series.chart.vba_get('ChartType', []) != -4169:
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        self.series.chart.changed()
        self.line.backward = wanted

    @member
    def Backward2(self) -> object:
        return self.vba_get('Backward', [])

    @setter('Backward2')
    def _backward2(self, value: object) -> None:
        self.vba_set('Backward', value)

    @member
    def DataLabel(self) -> object:
        self._index()
        return TrendlineLabel(self)

    @member
    def Intercept(self) -> object:
        self._index()
        if self.line.intercept is None:
            raise VBAUnsupportedError("An automatic trendline intercept is not implemented")
        return self.line.intercept

    @member
    def InterceptIsAuto(self) -> object:
        self._index()
        return self.line.intercept is None

    @member
    def Name(self) -> object:
        """The name given, or the type's label and the series' name in brackets (tests/fixtures/chart_trendlines.json)."""
        self._index()
        if self.line.name is not None:
            return self.line.name
        label = (f'{self.line.period} per. Mov. Avg.' if self.line.kind == 'movingAvg'
                 else LABELS[self.line.kind])
        return f'{label} ({to_text(self.series.Name())})'

    @setter('Name')
    def _name(self, value: object) -> None:
        self._change()
        self.series.chart.changed()
        self.line.name = to_text(value)

    @member
    def NameIsAuto(self) -> object:
        self._index()
        return self.line.name is None

    @setter('NameIsAuto')
    def _name_is_auto(self, value: object) -> None:
        self._change()
        if not to_bool(value):
            raise VBAUnsupportedError('Turning off an automatic trendline name is not implemented')
        self.series.chart.changed()
        self.line.name = None

    @member
    def DisplayEquation(self) -> object:
        self._index()
        return self.line.equation

    @setter('DisplayEquation')
    def _display_equation(self, value: object) -> None:
        self._change()
        self.series.chart.changed()
        self.line.equation = to_bool(value)

    @member
    def DisplayRSquared(self) -> object:
        self._index()
        return self.line.r_squared

    @setter('DisplayRSquared')
    def _display_r_squared(self, value: object) -> None:
        self._change()
        self.series.chart.changed()
        self.line.r_squared = to_bool(value)

    @method
    def Delete(self) -> object:
        self._change()
        self.series.chart.changed()
        self.series.entry.trendlines.remove(self.line)
        return EMPTY


__all__ = ['TrendlineState', 'Trendline', 'Trendlines', 'markup']
