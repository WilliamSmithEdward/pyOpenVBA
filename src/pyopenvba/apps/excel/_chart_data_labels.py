"""A series' DataLabels, as tests/fixtures/data_labels_object.json measured them: which parts they show, where they
stand, their number format and separator, what each kind of chart takes, and the errors a DataLabels taken before the
labels went raises."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._chart_xml import LABEL_POSITIONS, LabelLook
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import Series

#: Where each chart type puts its labels until told: outside the end of a clustered column or bar, centred on stacks,
#: areas and doughnuts, right of a line's or scatter's points, and best fit on a pie.
DEFAULT_POSITIONS = {51: 2, 57: 2, 52: -4108, 4: -4152, 65: -4152, 5: 5, -4120: -4108, -4169: -4152, 1: -4108}
#: The positions each measured chart type takes; another raises -2147467259 where the type was measured refusing one.
TAKEN = {51: {3, -4108, 4, 2}, 57: {3, -4108, 4, 2}, 4: {0, 1, -4131, -4152}, 65: {0, 1, -4131, -4152}, 5: {5, 2}}
REFUSING = {51: {0}, 57: {0}, 1: {-4108}}


class DataLabels(VBACollection, ItemDefault, Created):
    vba_type_name = 'DataLabels'

    def __init__(self, series: Series, generation: int) -> None:
        self.series, self.generation = series, generation

    def look(self, name: str = '') -> LabelLook:
        """The labels' look, or the error a DataLabels raises once its labels have gone: -2147417848 for its Parent,
        -2147467259 for Position, -2147467261 for the rest."""
        self.series.position()
        entry = self.series.entry
        if self.series.chart.state.original:
            raise VBAUnsupportedError('The data labels of an imported chart are not implemented')
        if not entry.data_labels or entry.label_generation != self.generation:
            if name == 'Parent':
                raise error(-2147417848, 'Automation error\nThe object invoked has disconnected from its clients.')
            if name == 'Position':
                raise error(-2147467259, 'Automation error\nUnspecified error')
            raise error(-2147467261, 'Automation error\nInvalid pointer')
        return entry.label_look

    def change(self) -> LabelLook:
        look = self.look()
        self.series.chart.changed()
        return look

    def chart_type(self) -> int:
        return self.series.chart.state.chart_type

    def vba_items(self) -> list[object]:
        self.look()
        return [DataLabel(self, index) for index in range(1, len(self.series.values()) + 1)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        place = int(to_integer(index, 'Long'))
        if not 1 <= place <= len(items):
            raise error(-2147024809, 'The parameter is incorrect.')
        return items[place - 1]

    @member
    def Count(self) -> object:
        return VBAInt(len(self.vba_items()), 'Long')

    @member
    def Parent(self) -> object:
        self.look('Parent')
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Name(self) -> object:
        self.look()
        return f'Text S{self.series.position()}'

    # -- the parts shown

    def flag(self, field: str) -> object:
        return getattr(self.look(), field)

    def set_flag(self, field: str, value: object) -> None:
        wanted = to_bool(value)
        setattr(self.change(), field, wanted)

    @member
    def ShowValue(self) -> object:
        return self.flag('value')

    @setter('ShowValue')
    def _show_value(self, value: object) -> None:
        self.set_flag('value', value)

    @member
    def ShowCategoryName(self) -> object:
        return self.flag('category')

    @setter('ShowCategoryName')
    def _show_category_name(self, value: object) -> None:
        self.set_flag('category', value)

    @member
    def ShowSeriesName(self) -> object:
        return self.flag('series')

    @setter('ShowSeriesName')
    def _show_series_name(self, value: object) -> None:
        self.set_flag('series', value)

    @member
    def ShowLegendKey(self) -> object:
        return self.flag('legend_key')

    @setter('ShowLegendKey')
    def _show_legend_key(self, value: object) -> None:
        self.set_flag('legend_key', value)

    def only_on(self, field: str, kinds: set[int], value: object) -> None:
        """A part only some kinds show: a clustered column or bar refuses it with -2147024809."""
        self.look()
        if to_bool(value) and self.chart_type() not in kinds:
            if self.chart_type() in (51, 57):
                raise error(-2147024809, 'The parameter is incorrect.')
            raise VBAUnsupportedError(f'DataLabels.{field} on this chart type is not implemented')
        self.set_flag(field, value)

    @member
    def ShowPercentage(self) -> object:
        return self.flag('percent')

    @setter('ShowPercentage')
    def _show_percentage(self, value: object) -> None:
        self.only_on('percent', {5}, value)

    @member
    def ShowBubbleSize(self) -> object:
        return self.flag('bubble')

    @setter('ShowBubbleSize')
    def _show_bubble_size(self, value: object) -> None:
        self.only_on('bubble', set(), value)

    # -- where they stand and how they read

    @member
    def Position(self) -> object:
        look = self.look('Position')
        if look.position is not None:
            return VBAInt(look.position, 'Long')
        if self.chart_type() not in DEFAULT_POSITIONS:
            raise VBAUnsupportedError('DataLabels.Position on this chart type is not implemented')
        return VBAInt(DEFAULT_POSITIONS[self.chart_type()], 'Long')

    @setter('Position')
    def _position(self, value: object) -> None:
        self.look()
        wanted = int(to_integer(value, 'Long'))
        kind = self.chart_type()
        if wanted in REFUSING.get(kind, set()):
            raise error(-2147467259, 'Automation error\nUnspecified error')
        if wanted not in TAKEN.get(kind, set()) or wanted not in LABEL_POSITIONS:
            raise VBAUnsupportedError('This DataLabels.Position on this chart type is not implemented')
        self.change().position = wanted

    @member
    def NumberFormat(self) -> object:
        return self.look().number_format or 'General'

    @setter('NumberFormat')
    def _number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._number_format import normalized

        self.look()
        code = normalized(to_text(value))
        look = self.change()
        look.number_format, look.linked = code, False

    @member
    def NumberFormatLocal(self) -> object:
        from pyopenvba.apps.excel._number_format import to_local

        return to_local(to_text(self.NumberFormat()), self.series.chart.sheet.book.application.locale)

    @setter('NumberFormatLocal')
    def _number_format_local(self, value: object) -> None:
        from pyopenvba.apps.excel._number_format import from_local

        self._number_format(from_local(to_text(value), self.series.chart.sheet.book.application.locale))

    @member
    def NumberFormatLinked(self) -> object:
        look = self.look()
        if not look.shown():
            # Labels showing nothing have no format to be linked.
            raise error(-2147024809, 'The parameter is incorrect.')
        return look.linked

    @setter('NumberFormatLinked')
    def _number_format_linked(self, value: object) -> None:
        # Linking drops the labels' own format; unlinking keeps it, General with none.
        wanted = to_bool(value)
        look = self.change()
        look.linked = wanted
        if wanted:
            look.number_format = None

    @member
    def Separator(self) -> object:
        separator = self.look().separator
        return VBAInt(1, 'Long') if separator is None else separator

    @setter('Separator')
    def _separator(self, value: object) -> None:
        if not isinstance(value, str):
            raise VBAUnsupportedError('DataLabels.Separator set to a constant is not implemented')
        self.change().separator = value

    # -- what new labels read

    @member
    def AutoText(self) -> object:
        self.look()
        return True

    @member
    def ShowRange(self) -> object:
        self.look()
        return False

    @member
    def Orientation(self) -> object:
        self.look()
        return VBAInt(-4128, 'Long')

    @member
    def HorizontalAlignment(self) -> object:
        self.look()
        return VBAInt(-4108, 'Long')

    @member
    def VerticalAlignment(self) -> object:
        self.look()
        return VBAInt(-4108, 'Long')

    @member
    def ReadingOrder(self) -> object:
        self.look()
        return VBAInt(-5002, 'Long')

    @member
    def Shadow(self) -> object:
        self.look()
        return False

    @method
    def Delete(self) -> object:
        """The labels gone, as HasDataLabels set False leaves them."""
        self.look()
        self.series.vba_set('HasDataLabels', False)
        return True


class DataLabel(Created):
    """One point's label, which the model reaches but does not read yet."""

    vba_type_name = 'DataLabel'

    def __init__(self, labels: DataLabels | None, index: int) -> None:
        self.labels, self.index = labels, index
