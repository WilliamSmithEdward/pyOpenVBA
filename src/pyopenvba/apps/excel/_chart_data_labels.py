"""A series' DataLabels, as tests/fixtures/data_labels_object.json measured them: which parts they show, where they
stand, their number format and separator, what each kind of chart takes, and the errors a DataLabels taken before the
labels went raises."""
from __future__ import annotations

from typing import Any, TYPE_CHECKING

from pyopenvba.apps.excel._chart_format import unmeasured
from pyopenvba.apps.excel._chart_xml import LABEL_POSITIONS, LabelLook, TextLook
from pyopenvba.apps.excel._charts import PointLabel, PointState
from pyopenvba.apps.excel._chart_borders import INVALID_CALL, UNSPECIFIED, ElementProperties
from pyopenvba.apps.excel._chart_format import EffectMembers
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.shapes._values import FillLine
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._format import vba_format
from pyopenvba.interpreter._values import MISSING, VBAInt, VBASingle, error, to_bool, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._chart_points import Point
    from pyopenvba.apps.excel._charts import Series

#: Where each chart type puts its labels until told: outside the end of a clustered column or bar, centred on stacks,
#: areas and doughnuts, right of a line's or scatter's points, and best fit on a pie.
DEFAULT_POSITIONS = {51: 2, 57: 2, 52: -4108, 4: -4152, 65: -4152, 5: 5, -4120: -4108, -4169: -4152, 1: -4108}
#: The positions each measured chart type takes; another raises -2147467259 where the type was measured refusing one.
TAKEN = {51: {3, -4108, 4, 2}, 57: {3, -4108, 4, 2}, 4: {0, 1, -4131, -4152, -4108}, 65: {0, 1, -4131, -4152, -4108},
         -4169: {0, 1, -4131, -4152, -4108}, 5: {5, 2, 3, -4108}}
REFUSING = {51: {0}, 57: {0}, 1: {-4108}}


#: What ApplyDataLabels' Type shows of itself: value, category and percent (tests/fixtures/apply_data_labels.json). A
#: bubble-size type shows the value on charts that have no bubbles.
_APPLIED = {2: (True, False, False), 4: (False, True, False), 3: (False, False, True), 5: (False, True, True),
            6: (True, False, False), -4142: (False, False, False)}
_PIE = 5


def _flag(value: object) -> bool | None:
    return None if value is MISSING else to_bool(value)


def applied_look(kind: int, label_type: object, legend_key: object, auto_text: object, series_name: object,
                 category_name: object, value: object, percentage: object, separator: object) -> LabelLook:
    """The parts ApplyDataLabels asks for: its Type's, with its named arguments over them. Percent is for a pie alone,
    so another chart's percent type is error 1004, and a percent argument does nothing there."""
    wanted = 2 if label_type is MISSING else int(to_integer(label_type, "Long"))
    if wanted not in _APPLIED:
        raise error(5, "Invalid procedure call or argument")
    if _flag(auto_text) is False:
        raise unmeasured("ApplyDataLabels with AutoText:=False")
    shows_value, shows_category, shows_percent = _APPLIED[wanted]
    if shows_percent and kind != _PIE:
        raise error(1004, "Application-defined or object-defined error")
    look = LabelLook(value=shows_value, category=shows_category, percent=shows_percent)
    for field_name, given in (("legend_key", legend_key), ("series", series_name), ("category", category_name),
                              ("value", value)):
        chosen = _flag(given)
        if chosen is not None:
            setattr(look, field_name, chosen)
    if kind == _PIE and _flag(percentage) is not None:
        look.percent = bool(_flag(percentage))
    if isinstance(separator, str):
        look.separator = separator
    return look


def apply_to_series(series: Series, look: LabelLook, leader_lines: object) -> None:
    """A series given these parts, or with none its labels gone. Its position, number format, separator and font stay,
    and its points go back to what it gives, but for a typed text, whose parts follow."""
    series.position()
    entry = series.entry
    series.chart.changed()
    if not look.shown():
        if entry.data_labels:
            entry.label_generation += 1
            entry.leader_lines = False
        entry.data_labels = False
        entry.label_look = LabelLook()
        for point in entry.points.values():
            point.label, point.label_own = None, PointLabel()
        return
    entry.data_labels = True
    kept = entry.label_look
    kept.value, kept.category, kept.series = look.value, look.category, look.series
    kept.percent, kept.legend_key, kept.bubble = look.percent, look.legend_key, False
    if look.separator is not None:
        kept.separator = look.separator
    entry.leader_lines = _flag(leader_lines) is not False
    for point in entry.points.values():
        if point.label_own.text is None:
            point.label, point.label_own = None, PointLabel()
        else:
            point.label_own.shows = {}


def apply_to_point(point: Point, look: LabelLook, leader_lines: object) -> None:
    """One point given these parts, or with none its label gone."""
    entry = point.series.entry
    own = point.held()
    chart = point.series.chart
    if own.label_unique < 0:
        own.label_unique, chart.state.next_unique = chart.state.next_unique, chart.state.next_unique + 1
    if _flag(leader_lines) is False:
        entry.leader_lines = False
    if not entry.data_labels:
        entry.label_look = LabelLook(value=False)
    if not look.shown():
        own.label, own.label_own = False, PointLabel()
        return
    if _flag(leader_lines) is not False:
        entry.leader_lines = True
    shows = {"value": look.value, "category": look.category, "series": look.series, "legend_key": look.legend_key,
             "percent": look.percent}
    series_look = entry.label_look
    if entry.data_labels and own.label_own.text is None and all(
            getattr(series_look, name) == shown for name, shown in shows.items()):
        # What its series' labels give already: nothing of its own.
        own.label, own.label_own = None, PointLabel()
        return
    own.label = True
    own.label_own.shows = shows
    if entry.data_labels:
        # It writes down what it had from its series' labels.
        own.label_own.position = series_look.position
        own.label_own.separator = series_look.separator
    if look.separator is not None:
        own.label_own.separator = look.separator


def _own_type(series: Any) -> int:
    """The type a series' labels go by: its own in a combo chart (see _chart_combos), and otherwise the chart's."""
    from pyopenvba.apps.excel._chart_combos import is_combo, series_type

    state = series.chart.state
    return series_type(state, series.entry) if is_combo(state) else state.chart_type


class DataLabels(ElementProperties, VBACollection, ItemDefault, Created):
    vba_type_name = 'DataLabels'

    def __init__(self, series: Series, generation: int) -> None:
        self.series, self.generation = series, generation

    @member
    def Format(self) -> object:
        return _LabelsFormat(self)

    @method
    def Propagate(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        if Index is MISSING:
            raise error(450, 'Wrong number of arguments or invalid property assignment')
        raise VBAUnsupportedError('DataLabels.Propagate from a point is not implemented')

    def look(self, name: str = '') -> LabelLook:
        """The labels' look, or the error a DataLabels raises once its labels have gone: -2147417848 for its Parent,
        -2147467259 for Position, -2147467261 for the rest."""
        self.series.position()
        entry = self.series.entry
        if self.series.chart.state.original:
            raise VBAUnsupportedError('The data labels of an imported chart are not implemented')
        if not self.series.has_labels() or entry.label_generation != self.generation:
            if name == 'Parent':
                raise error(-2147417848, 'Automation error\nThe object invoked has disconnected from its clients.')
            if name == 'Position':
                raise error(-2147467259, 'Automation error\nUnspecified error')
            raise error(-2147467261, 'Automation error\nInvalid pointer')
        return entry.label_look

    def change(self) -> LabelLook:
        look = self.look()
        if not self.series.entry.data_labels:
            raise unmeasured("Changing the labels of a series whose labels are on some points alone")
        self.series.chart.changed()
        return look

    def chart_type(self) -> int:
        return _own_type(self.series)

    def vba_items(self) -> list[object]:
        self.look()
        return [DataLabel(self.series, index, self.generation) for index in range(1, len(self.series.values()) + 1)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        place = int(to_integer(index, 'Long'))
        if not 1 <= place <= len(items):
            raise error(-2147024809, 'The parameter is incorrect.')
        return items[place - 1]

    @member
    def Count(self) -> object:
        return VBAInt(len(self.vba_items()), 'Long')

    @method
    def Select(self) -> object:
        from pyopenvba.apps.excel._charts import select_element

        self.look()
        return select_element(self.series.chart, self, True)

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
        self.only_on('percent', {5, -4120}, value)

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

    @member
    def Font(self) -> object:
        # A chart a file holds reads its labels' font, though not the rest of them yet.
        chart = self.series.chart
        return chart.font_of(self, "labels", labels=self.series.entry.label_look if chart.state.original else self.look())

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


class DataLabel(ElementProperties, Created):
    get_error, set_error = INVALID_CALL, UNSPECIFIED

    """One point's label, as tests/fixtures/data_label.json measured it. What is set of it stands apart from its series'
    labels, whatever is not set following them, and a label showing nothing is a deleted one.

    AutoText = False and a formula write a field with a random id, and the place and size of a label depend on the
    display, so they are refused."""

    vba_type_name = "DataLabel"

    def __init__(self, series: Series, index: int, generation: int) -> None:
        self.series, self.index, self.generation = series, index, generation

    # -- where it stands

    def point_state(self) -> PointState | None:
        return self.series.entry.points.get(self.series.point_key(self.index - 1))

    def _stale(self) -> bool:
        return self.series.entry.label_generation != self.generation

    def own(self) -> PointLabel:
        state = self.point_state()
        return state.label_own if state is not None else PointLabel()

    def held(self) -> PointLabel:
        """Its state, made if it had none, with the unique id its dLbl takes."""
        chart = self.series.chart
        chart.changed()
        state = self.series.entry.points.setdefault(self.series.point_key(self.index - 1), PointState())
        if state.drawn_at < 0:
            state.drawn_at = self.index - 1
        if state.label_unique < 0 and state.unique >= 0:
            # A point with a dPt shares its id with its dLbl (tests/fixtures/chart_copies.json).
            state.label_unique = state.unique
        elif state.label_unique < 0:
            state.label_unique, chart.state.next_unique = chart.state.next_unique, chart.state.next_unique + 1
        return state.label_own

    def _touched(self) -> None:
        """Characters reads back from now on, with nothing in the file to show for it."""
        point = self.series.entry.points.setdefault(self.series.point_key(self.index - 1), PointState())
        point.label_own.touched = True

    def shows_anything(self) -> bool:
        state = self.point_state()
        if state is not None and state.label is False:
            return False
        return any(self._flags().values())

    def _flags(self) -> dict[str, bool]:
        state = self.point_state()
        look = self.series.entry.label_look
        if state is not None and state.label is False:
            return dict.fromkeys(("value", "category", "series", "legend_key", "percent", "bubble"), False)
        base = {"value": look.value, "category": look.category, "series": look.series, "legend_key": look.legend_key,
                "percent": look.percent, "bubble": look.bubble}
        if state is not None and state.label is True and not self.series.entry.data_labels:
            base = {"value": True, "category": False, "series": False, "legend_key": False, "percent": False,
                    "bubble": False}
        return {**base, **self.own().shows}

    def _alive(self, name: str = "") -> None:
        """The errors of a label that has gone: its series' labels taken away, or this one showing nothing."""
        if self._stale():
            if name == "Position":
                raise error(-2147467259, "Automation error\nUnspecified error")
            raise error(-2147467261, "Automation error\nInvalid pointer")

    def _showing(self, name: str = "") -> None:
        self._alive(name)
        if not self.shows_anything():
            if name == "Position":
                raise error(-2147467259, "Automation error\nUnspecified error")
            raise error(-2147024809, "The parameter is incorrect.")

    def _formatted(self) -> None:
        if self._stale() or not self.shows_anything():
            raise error(1004, "Application-defined or object-defined error")

    def _flag(self, field: str) -> object:
        self._alive()
        return self._flags()[field]

    def _set_flag(self, field: str, value: object) -> None:
        self._alive()
        wanted = to_bool(value)
        if wanted == self._flags()[field] and field not in self.own().shows and self.point_state() is None:
            return
        own = self.held()
        own.shows[field] = wanted
        own.touched = True
        state = self.point_state()
        if state is not None and not wanted and not any(self._flags().values()):
            state.label = False
            state.label_own = PointLabel()

    def _point(self) -> object:
        return self.series.vba_get("Points", [VBAInt(self.index, "Long")])

    # -- identity

    @member
    def Parent(self) -> object:
        self._alive()
        return self._point()

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Name(self) -> object:
        self._alive()
        return f"Text S{self.series.position()}P{self.index}"

    # -- what it says

    def _automatic(self) -> str:
        flags = self._flags()
        parts: list[str] = []
        if flags["series"]:
            parts.append(to_text(self.series.vba_get("Name", [])))
        if flags["category"]:
            categories = self.series.categories()
            parts.append(to_text(categories[self.index - 1]))
        value = self.series.values()[self.index - 1]
        if flags["value"]:
            parts.append(self._number(value))
        if flags["percent"]:
            total = sum(float(to_number(one)) for one in self.series.values())
            parts.append(f"{round(float(to_number(value)) / total * 100)}%" if total else "0%")
        separator = self._separator()
        if separator is None and len(parts) == 2 and flags["category"] and flags["percent"] and not self.own().shows:
            # A pie's labels break the line here; a point's own do not.
            separator = "\n"
        return (", " if separator is None else separator).join(parts)

    def _number(self, value: object) -> str:
        code = self._number_format()
        return to_text(value) if code == "General" else to_text(vba_format(value, code))

    def _number_format(self) -> str:
        own = self.own().number_format
        return own if own is not None else self.series.entry.label_look.number_format or "General"

    def _linked(self) -> bool:
        own = self.own()
        return own.linked if own.linked is not None else self.series.entry.label_look.linked

    def _separator(self) -> str | None:
        own = self.own().separator
        return own if own is not None else self.series.entry.label_look.separator

    def said(self) -> str:
        own = self.own()
        return own.text if own.text is not None else self._automatic()

    @member
    def Text(self) -> object:
        self._showing()
        return self.said()

    @setter("Text")
    def set_text(self, value: object) -> None:
        self._showing()
        self._type(to_text(value))

    def _type(self, text: str) -> None:
        if self.own().plain_format is False:
            raise unmeasured("Typing a label's text once its look is set")
        own = self.held()
        own.text, own.touched = text, True

    @member
    def Caption(self) -> object:
        return self.Text()

    @setter("Caption")
    def _set_caption(self, value: object) -> None:
        self.set_text(value)

    @member
    def AutoText(self) -> object:
        self._showing()
        own = self.own()
        return own.text is None

    @setter("AutoText")
    def _set_auto_text(self, value: object) -> None:
        self._showing()
        if not to_bool(value):
            raise unmeasured("AutoText = False, which writes a field with a random id,")
        own = self.own()
        if own.text is not None:
            own.text, own.touched = None, False
            self.series.chart.changed()

    @member
    def Formula(self) -> object:
        self._showing()
        return self.said()

    @setter("Formula")
    def _set_formula(self, value: object) -> None:
        self._showing()
        text = to_text(value)
        if not text.startswith("="):
            self._type(text)
            return
        raise unmeasured("A label's formula, whose file part has a random id,")

    @member
    def FormulaLocal(self) -> object:
        return self.Formula()

    @setter("FormulaLocal")
    def _set_formula_local(self, value: object) -> None:
        self._set_formula(value)

    @member
    def FormulaR1C1(self) -> object:
        self._showing()
        return self.said()

    @setter("FormulaR1C1")
    def _set_formula_r1c1(self, value: object) -> None:
        self._set_formula(value)

    @member
    def FormulaR1C1Local(self) -> object:
        return self.FormulaR1C1()

    @setter("FormulaR1C1Local")
    def _set_formula_r1c1_local(self, value: object) -> None:
        self._set_formula_r1c1(value)

    @member
    def Characters(self) -> object:
        self._alive()
        return LabelCharacters(self)

    # -- the parts shown

    @member
    def ShowValue(self) -> object:
        return self._flag("value")

    @setter("ShowValue")
    def _show_value(self, value: object) -> None:
        self._set_flag("value", value)

    @member
    def ShowCategoryName(self) -> object:
        return self._flag("category")

    @setter("ShowCategoryName")
    def _show_category_name(self, value: object) -> None:
        self._set_flag("category", value)

    @member
    def ShowSeriesName(self) -> object:
        return self._flag("series")

    @setter("ShowSeriesName")
    def _show_series_name(self, value: object) -> None:
        self._set_flag("series", value)

    @member
    def ShowLegendKey(self) -> object:
        return self._flag("legend_key")

    @setter("ShowLegendKey")
    def _show_legend_key(self, value: object) -> None:
        self._set_flag("legend_key", value)

    @member
    def ShowPercentage(self) -> object:
        return self._flag("percent")

    @setter("ShowPercentage")
    def _show_percentage(self, value: object) -> None:
        self._only_on("percent", {5}, value)

    @member
    def ShowBubbleSize(self) -> object:
        return self._flag("bubble")

    @setter("ShowBubbleSize")
    def _show_bubble_size(self, value: object) -> None:
        self._only_on("bubble", set(), value)

    def _only_on(self, field: str, kinds: set[int], value: object) -> None:
        self._alive()
        kind = _own_type(self.series)
        if to_bool(value) and kind not in kinds:
            if kind in (5, 51, 57):
                raise error(-2147024809, "The parameter is incorrect.")
            raise VBAUnsupportedError(f"DataLabel.{field} on this chart type is not implemented")
        self._set_flag(field, value)

    @member
    def ShowRange(self) -> object:
        self._alive()
        return self.own().show_range

    @setter("ShowRange")
    def _show_range(self, value: object) -> None:
        self._alive()
        own = self.held()
        own.show_range, own.touched = to_bool(value), True

    # -- where it stands and how it reads

    @member
    def Position(self) -> object:
        self._showing("Position")
        own = self.own().position
        look = self.series.entry.label_look.position
        kind = _own_type(self.series)
        if own is None and look is None and kind not in DEFAULT_POSITIONS:
            raise VBAUnsupportedError("DataLabel.Position on this chart type is not implemented")
        return VBAInt(own if own is not None else look if look is not None else DEFAULT_POSITIONS[kind], "Long")

    @setter("Position")
    def _position(self, value: object) -> None:
        self._showing("Position")
        wanted = int(to_integer(value, "Long"))
        kind = _own_type(self.series)
        if kind not in _POINT_POSITIONS:
            raise VBAUnsupportedError("DataLabel.Position on this chart type is not implemented")
        if wanted not in _POINT_POSITIONS[kind]:
            raise error(-2147467259, "Automation error\nUnspecified error")
        own = self.held()
        own.position, own.touched = wanted, True

    @member
    def NumberFormat(self) -> object:
        self._showing()
        return self._number_format()

    @setter("NumberFormat")
    def _set_number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._number_format import normalized

        self._showing()
        try:
            code = normalized(to_text(value))
        except VBARuntimeError:
            raise error(1004, "Unable to set the NumberFormat property of the DataLabel class") from None
        own = self.held()
        own.number_format, own.linked = code, False

    @member
    def NumberFormatLocal(self) -> object:
        from pyopenvba.apps.excel._number_format import to_local

        self._showing()
        return to_local(self._number_format(), self.series.chart.sheet.book.application.locale)

    @setter("NumberFormatLocal")
    def _set_number_format_local(self, value: object) -> None:
        from pyopenvba.apps.excel._number_format import from_local

        self._set_number_format(from_local(to_text(value), self.series.chart.sheet.book.application.locale))

    @member
    def NumberFormatLinked(self) -> object:
        self._showing()
        return self._linked()

    @setter("NumberFormatLinked")
    def _set_number_format_linked(self, value: object) -> None:
        self._showing()
        wanted = to_bool(value)
        if wanted:
            own = self.own()
            if own.number_format is not None or own.linked is not None:
                own.number_format, own.linked = None, None
                self.series.chart.changed()
            return
        own = self.held()
        own.linked = False
        own.number_format = own.number_format or "General"

    @member
    def Separator(self) -> object:
        self._showing()
        separator = self._separator()
        return VBAInt(1, "Long") if separator is None else separator

    @setter("Separator")
    def _set_separator(self, value: object) -> None:
        self._showing()
        if isinstance(value, str):
            own = self.held()
            own.separator, own.touched = value, True
        else:
            self._touched()

    @member
    def Orientation(self) -> object:
        self._alive()
        return VBAInt(self.own().orientation, "Long")

    @setter("Orientation")
    def _set_orientation(self, value: object) -> None:
        self._alive()
        wanted = int(to_integer(value, "Long"))
        if wanted not in _ROTATIONS and not -90 <= wanted <= 90:
            raise error(-2147467259, "Automation error\nUnspecified error")
        self.held().orientation = wanted

    @member
    def HorizontalAlignment(self) -> object:
        self._alive()
        return VBAInt(self.own().horizontal, "Long")

    @setter("HorizontalAlignment")
    def _set_horizontal(self, value: object) -> None:
        self._alive()
        wanted = int(to_integer(value, "Long"))
        if wanted not in _ALIGNED:
            raise unmeasured("This DataLabel.HorizontalAlignment")
        self.held().horizontal = wanted

    @member
    def VerticalAlignment(self) -> object:
        self._alive()
        return VBAInt(-4108, "Long")

    @setter("VerticalAlignment")
    def _set_vertical(self, value: object) -> None:
        self._alive()
        to_integer(value, "Long")

    @member
    def ReadingOrder(self) -> object:
        self._alive()
        return VBAInt(self.own().reading, "Long")

    @setter("ReadingOrder")
    def _set_reading(self, value: object) -> None:
        self._alive()
        self.held().reading = int(to_integer(value, "Long"))

    @member
    def Shadow(self) -> object:
        self._alive()
        return self.own().shadow

    @setter("Shadow")
    def _set_shadow(self, value: object) -> None:
        self._alive()
        self.held().shadow = to_bool(value)

    # -- its look

    @member
    def Font(self) -> object:
        self._formatted()
        return self.series.chart.font_of(self, "label", labels=self)

    @property
    def font(self) -> TextLook | None:
        """What Font reads: its own, or else what its series' labels set."""
        return self.own().font or self.series.entry.label_look.font

    @font.setter
    def font(self, look: TextLook | None) -> None:
        self.held().font = look

    @member
    def Format(self) -> object:
        self._alive("Position")
        return _LabelFormat(self)

    @member
    def Border(self) -> object:
        self._formatted()
        return _Unmodelled("Border")

    @member
    def Interior(self) -> object:
        if self._stale():
            raise error(1004, "Application-defined or object-defined error")
        return _Unmodelled("Interior")

    @member
    def Fill(self) -> object:
        if self._stale():
            raise error(1004, "Application-defined or object-defined error")
        return _Unmodelled("ChartFillFormat")

    # -- what it does

    @method
    def Delete(self) -> object:
        self._showing()
        chart = self.series.chart
        chart.changed()
        state = self.series.entry.points.setdefault(self.series.point_key(self.index - 1), PointState())
        if state.label_unique < 0:
            state.label_unique, chart.state.next_unique = chart.state.next_unique, chart.state.next_unique + 1
        state.label, state.label_own = False, PointLabel()
        return True

    @method
    def Select(self) -> object:
        self._showing()
        if _own_type(self.series) == 5:
            # On a pie it brings the chart a title (tests/fixtures/data_label.json).
            raise unmeasured("Selecting a pie's label")
        from pyopenvba.apps.excel._charts import select_element

        return select_element(self.series.chart, self, True)


#: The places a point's label can take, by chart type (tests/fixtures/data_label.json).
_POINT_POSITIONS = {51: {3, -4108, 4, 2}, 5: {3, -4108, 2, 5}, 4: {-4152}, 65: {-4152}, -4169: {-4152}}
_ROTATIONS = {-4171: 90, -4170: -90, -4128: 0}
_ALIGNED = {-4131: "l", -4152: "r", -4108: "ctr"}


class LabelCharacters(Created):
    """A label's Characters, whose text it reads back only once something has been set of the label."""

    vba_type_name = "Characters"

    def __init__(self, label: DataLabel) -> None:
        self.label = label

    @member
    def Text(self) -> object:
        return self._read()

    def _read(self) -> str:
        label = self.label
        if not label.shows_anything():
            return ""
        state = label.point_state()
        live = (state is not None and state.label is True) or label.own().touched
        return label.said() if live else ""

    @setter("Text")
    def _set_text(self, value: object) -> None:
        self.label.set_text(value)

    @member
    def Count(self) -> object:
        return VBAInt(len(self._read()), "Long")


class _Unmodelled(Created):
    """Reached for what it is, and not modelled further."""

    vba_type_name = ""

    def __init__(self, type_name: str) -> None:
        self.vba_type_name = type_name


class _NoAdjustments(Created):
    vba_type_name = "Adjustments"

    @member
    def Count(self) -> object:
        return VBAInt(0, "Long")


class _LabelsFormat(EffectMembers, Created):
    """The Format of a series' labels, a fill and an outline, shared with each point's until that has its own
    (tests/fixtures/trendline_members.json)."""

    vba_type_name = "ChartFormat"

    def __init__(self, labels: DataLabels) -> None:
        self.labels = labels

    def effect_look(self) -> FillLine:
        return FillLine()

    @member
    def Adjustments(self) -> object:
        return _NoAdjustments()

    @member
    def AutoShapeType(self) -> object:
        return VBAInt(0, "Long")

    @setter("AutoShapeType")
    def _labels_shape(self, value: object) -> None:
        del value
        raise VBAUnsupportedError("The shape of a chart's data labels is not implemented")

    @member
    def Parent(self) -> object:
        return self.labels

    @member
    def Fill(self) -> object:
        return _LabelsFill(self.labels)

    @member
    def Line(self) -> object:
        return _LabelsLine(self.labels)


class _LabelsFill(Created):
    vba_type_name = "FillFormat"

    def __init__(self, labels: DataLabels) -> None:
        self.labels = labels

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.labels.look().fill_rgb is None else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        if int(to_integer(value, "Long")) != 0 or self.labels.look().fill_rgb is not None:
            raise VBAUnsupportedError("Showing the labels' fill, or hiding one with a colour, is not implemented")

    @member
    def ForeColor(self) -> object:
        return _LabelsColour(self.labels)


class _LabelsColour(Created):
    vba_type_name = "ColorFormat"

    def __init__(self, labels: DataLabels) -> None:
        self.labels = labels

    @member
    def RGB(self) -> object:
        rgb = self.labels.look().fill_rgb or "FFFFFF"
        return VBAInt(int(rgb[4:6] + rgb[2:4] + rgb[:2], 16), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        colour = int(to_integer(value, "Long")) & 0xFFFFFF
        look = self.labels.look()
        self.labels.series.chart.changed()
        look.fill_rgb = f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


class _LabelsLine(Created):
    """The labels' outline, which a colour does not bring, nor read back, and a weight does."""

    vba_type_name = "LineFormat"

    def __init__(self, labels: DataLabels) -> None:
        self.labels = labels

    @member
    def Visible(self) -> object:
        return VBAInt(0 if self.labels.look().line_width is None else -1, "Long")

    @setter("Visible")
    def _visible(self, value: object) -> None:
        if int(to_integer(value, "Long")) != 0 or self.labels.look().line_width is not None:
            raise VBAUnsupportedError("Showing the labels' outline, or hiding one with a weight, is not implemented")

    @member
    def Weight(self) -> object:
        width = self.labels.look().line_width
        if width is None:
            raise VBAUnsupportedError("The weight of labels with no outline is not implemented")
        return VBASingle(width / 12700)

    @setter("Weight")
    def _weight(self, value: object) -> None:
        look = self.labels.look()
        self.labels.series.chart.changed()
        look.line_width = round(float(to_number(value)) * 12700)

    @member
    def ForeColor(self) -> object:
        return _LabelsLineColour()


class _LabelsLineColour(Created):
    vba_type_name = "ColorFormat"

    @member
    def RGB(self) -> object:
        return VBAInt(16777215, "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        to_integer(value, "Long")


class _LabelFormat(Created):
    vba_type_name = "ChartFormat"

    def __init__(self, label: DataLabel) -> None:
        self.label = label

    @member
    def Fill(self) -> object:
        return _LabelFill(self.label)

    @member
    def Line(self) -> object:
        return _LabelLine(self.label)


class _LabelFill(Created):
    vba_type_name = "FillFormat"

    def __init__(self, label: DataLabel) -> None:
        self.label = label

    @member
    def Visible(self) -> object:
        return VBAInt(-1 if self.label.own().fill_rgb is not None else 0, "Long")

    @member
    def ForeColor(self) -> object:
        return _LabelColour(self.label)


class _LabelColour(Created):
    vba_type_name = "ColorFormat"

    def __init__(self, label: DataLabel) -> None:
        self.label = label

    @member
    def RGB(self) -> object:
        """Its own colour, or its series' labels' (tests/fixtures/trendline_members.json)."""
        rgb = self.label.own().fill_rgb or self.label.series.entry.label_look.fill_rgb or "FFFFFF"
        return VBAInt(int(rgb[4:6] + rgb[2:4] + rgb[:2], 16), "Long")

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        colour = int(to_integer(value, "Long")) & 0xFFFFFF
        self.label.held().fill_rgb = f"{colour & 0xFF:02X}{colour >> 8 & 0xFF:02X}{colour >> 16 & 0xFF:02X}"


class _LabelLine(Created):
    """A label's outline, which is not there, and colouring it does not bring it (tests/fixtures/data_label.json)."""

    vba_type_name = "LineFormat"

    def __init__(self, label: DataLabel) -> None:
        self.label = label

    @member
    def Visible(self) -> object:
        return VBAInt(0, "Long")

    @member
    def ForeColor(self) -> object:
        return _LabelLineColour()


class _LabelLineColour(Created):
    vba_type_name = "ColorFormat"

    @setter("RGB")
    def _rgb(self, value: object) -> None:
        to_integer(value, "Long")
