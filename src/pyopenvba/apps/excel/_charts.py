"""Embedded charts share their worksheet's drawing and cell state."""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from decimal import Decimal
import math
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference
from pyopenvba.apps.excel._chart_axes import AxisState
from pyopenvba.apps.excel._shapes import ShapeObject, ShapeRange, Shapes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.apps.excel._creator import Created
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, NULL, VBAArray, VBAInt, VBASingle, error, to_bool, to_integer, to_number, to_text
from pyopenvba.shapes._values import Shape, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

# Engine admission bounds, not native Excel limits.
MAX_CHART_POINTS = 1_000_000
MAX_CHART_SERIES = 1024
MAX_CHART_REFERENCE_TEXT = 32_768

CHART_TYPES = {51: 'column', 57: 'bar', 4: 'line', 65: 'lineMarkers', 5: 'pie', -4120: 'doughnut',
               -4169: 'scatter', 1: 'area'}

@dataclass
class SeriesState:
    name: str | None = None
    name_reference: str = ''
    categories: list[object] = field(default_factory=lambda: [])
    values: list[object] = field(default_factory=lambda: [])
    category_reference: str = ''
    value_reference: str = ''
    series_id: int = -1
    #: The first group of the series' unique id; -1 to take its series_id.
    unique: int = -1

@dataclass
class ChartState:
    chart_type: int = 51
    title: str | None = None
    #: False for Excel's automatic title, which the file holds with no text.
    title_typed: bool = False
    legend: bool = True
    plot_by: int = 2
    series: list[SeriesState] = field(default_factory=lambda: [])
    part: str = ''
    original: str = ''
    dirty: bool = True
    axes_readable: bool = True
    axes: dict[int, AxisState] = field(default_factory=lambda: {})
    axis_ids: tuple[int, int] = (10000001, 10000002)
    #: Where SetSourceData starts numbering unique ids, and the next one NewSeries takes.
    unique_base: int = 0
    next_unique: int = 0
    #: DisplayBlanksAs, any Long; the file holds only 1 to 3.
    blanks: int = 1
    visible_only: bool = True
    over_maximum: bool = False
    not_available_as_blank: bool = True
    #: False when an imported chart spells one of the four display options in a way the model does not read.
    display_readable: bool = True


def plan_series(sheet: Worksheet, area: Area, orientation: int) -> list[SeriesState]:
    """The series SetSourceData makes from a block of cells, by rows (1) or columns (2). A first row or column of text
    names the series or labels the categories; the corner cell takes no part in deciding either."""
    from pyopenvba.apps.excel._model import Range

    height, width = area.bottom - area.top + 1, area.right - area.left + 1

    def value(row: int, column: int) -> object:
        return Range(sheet, [Area(row, column, row, column)]).Value()

    def text_axis(values: list[object]) -> bool:
        return bool(values) and all(isinstance(item, str) or item is EMPTY for item in values) and any(
            isinstance(item, str) and bool(item) for item in values)
    series_plan: list[SeriesState] = []
    if orientation == 2:
        names = text_axis([value(area.top, col) for col in range(area.left + (width > 1), area.right + 1)])
        top = area.top + names
        labels = text_axis([value(row, area.left) for row in range(area.top + (height > 1), area.bottom + 1)])
        if width - labels > MAX_CHART_SERIES:
            raise VBAUnsupportedError('Chart source data exceeds the supported series budget')
        for column in range(area.left + labels, area.right + 1):
            series = SeriesState(series_id=len(series_plan))
            if names:
                series.name_reference = reference(sheet, Area(area.top, column, area.top, column))
            series.value_reference = reference(sheet, Area(top, column, area.bottom, column))
            if labels:
                series.category_reference = reference(sheet, Area(top, area.left, area.bottom, area.left))
            series_plan.append(series)
    else:
        names = text_axis([value(row, area.left) for row in range(area.top + (height > 1), area.bottom + 1)])
        left = area.left + names
        labels = text_axis([value(area.top, col) for col in range(area.left + (width > 1), area.right + 1)])
        if height - labels > MAX_CHART_SERIES:
            raise VBAUnsupportedError('Chart source data exceeds the supported series budget')
        for row in range(area.top + labels, area.bottom + 1):
            series = SeriesState(series_id=len(series_plan))
            if names:
                series.name_reference = reference(sheet, Area(row, area.left, row, area.left))
            series.value_reference = reference(sheet, Area(row, left, row, area.right))
            if labels:
                series.category_reference = reference(sheet, Area(area.top, left, area.top, area.right))
            series_plan.append(series)
    return series_plan


def plotted_before(sheet: Worksheet) -> int:
    """How many series ChartObjects.Add would plot, and drops: those of the selection, or of the active cell's current
    region, as SetSourceData plots them by default. Excel numbers later series' unique ids past them."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._region import current_region

    selection = sheet.selection
    if not isinstance(selection, Range):
        return 0
    area = selection.first
    if len(selection.areas) == 1 and (area.top, area.left) == (area.bottom, area.right):
        area = current_region(sheet, area.top, area.left)
        if (area.top, area.left) == (area.bottom, area.right) and Range(sheet, [area]).Value() is EMPTY:
            return 0
    height, width = area.bottom - area.top + 1, area.right - area.left + 1
    if height * width > MAX_CHART_POINTS:
        raise VBAUnsupportedError('Chart source data exceeds the supported cell budget')
    return len(plan_series(sheet, area, 1 if width >= height else 2))


def shown_type(chart: ChartState) -> int:
    """The type a chart reads and is written as: a line chart with no series is one with markers until it has one."""
    return 65 if chart.chart_type == 4 and not chart.series else chart.chart_type


def plot_by(chart: ChartState) -> int:
    """PlotBy as Excel works it out from the series: xlColumns (2) when their values are columns of the same rows from
    left to right, each name the cell above its own column and every series the same categories; xlRows (1) the same
    turned round; 0 when neither holds, as for a literal or a series moved out of order. A chart with no series reads
    xlRows, and series that fit both read the way their source data was given."""
    if not chart.series:
        return 1
    try:
        values = [parse_reference(entry.value_reference) for entry in chart.series]
        names = [parse_reference(entry.name_reference) if entry.name_reference else None for entry in chart.series]
    except ValueError:
        return 0
    literal = any(not entry.value_reference or entry.name is not None and not entry.name_reference
                  or entry.categories and not entry.category_reference for entry in chart.series)
    if (literal or any(len(areas) != 1 for areas in values) or any(name is not None and len(name) != 1 for name in names)
            or len({entry.category_reference for entry in chart.series}) != 1
            or len({name is None for name in names}) != 1):
        return 0
    blocks = [areas[0] for areas in values]
    cells = [name[0] if name is not None else None for name in names]

    def fits(columns: bool) -> bool:
        def span(area: Area) -> tuple[int, int]:
            return (area.top, area.bottom) if columns else (area.left, area.right)

        def place(area: Area) -> int:
            return area.left if columns else area.top

        if any(place(area) != (area.right if columns else area.bottom) for area in blocks):
            return False
        if len({(area.sheet, span(area)) for area in blocks}) != 1:
            return False
        places = [place(area) for area in blocks]
        if any(later <= earlier for earlier, later in zip(places, places[1:])):
            return False
        for cell, area in zip(cells, blocks):
            if cell is not None and (cell.sheet != area.sheet or (cell.top, cell.left) != (cell.bottom, cell.right)
                                     or place(cell) != place(area) or span(cell)[0] >= span(area)[0]):
                return False
        return len({span(cell) for cell in cells if cell is not None}) <= 1

    by_columns, by_rows = fits(True), fits(False)
    if by_columns and by_rows:
        return chart.plot_by
    return 2 if by_columns else 1 if by_rows else 0


def state(sheet: Worksheet, shape: Shape) -> ChartState:
    if shape.shape_id not in sheet.chart_states:
        if shape.source:
            raise VBAUnsupportedError('Chart metadata for this drawing has not been loaded')
        sheet.chart_states[shape.shape_id] = ChartState()
    return sheet.chart_states[shape.shape_id]


class ChartObjects(VBACollection, Created):
    vba_type_name = 'ChartObjects'

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return [ChartObject(self.sheet, shape) for shape in self.sheet.shapes_ if shape.kind == 'chart']

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                if isinstance(item, ChartObject) and item.shape.name.lower() == index.lower():
                    return item
            raise error(-2147024809, 'The item with the specified name was not found')
        position = int(to_number(index))
        if not 1 <= position <= len(items):
            raise error(1004, 'Unable to get the ChartObjects property of the Worksheet class')
        return items[position - 1]

    @method
    def Add(self, Left: object = MISSING, Top: object = MISSING, Width: object = MISSING,
            Height: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, 'ChartObjects.Add')
        if any(value is MISSING for value in (Left, Top, Width, Height)):
            raise error(449, 'Argument not optional')
        left, top, width, height = (float(to_number(value)) for value in (Left, Top, Width, Height))
        if not all(math.isfinite(value) for value in (left, top, width, height)):
            raise error(1004, 'Invalid chart geometry')
        if left < 0 or top < 0:
            raise VBAUnsupportedError('Chart placement centered in a native window is not implemented')
        if width < 0 or height < 0:
            width, height = 360, 216
        left, top, width, height = (float(VBASingle(points(emu(value)))) for value in (left, top, width, height))
        shape = Shapes(self.sheet).new_chart_shape(left, top, width, height)
        base = plotted_before(self.sheet)
        self.sheet.chart_states[shape.shape_id] = ChartState(unique_base=base, next_unique=base)
        return ChartObject(self.sheet, shape)

    @method
    def Delete(self) -> object:
        for item in self.vba_items():
            assert isinstance(item, ChartObject)
            item.Delete()
        return EMPTY


class ChartObject(Created):
    vba_type_name = 'ChartObject'

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet, self.shape = sheet, shape

    @member
    def Name(self) -> object:
        return self.shape.name

    @setter('Name')
    def _name(self, value: object) -> None:
        ShapeObject(self.sheet, self.shape).vba_set('Name', value)

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Index(self) -> object:
        return VBAInt([item for item in self.sheet.shapes_ if item.kind == 'chart'].index(self.shape) + 1, 'Long')

    @member
    def Chart(self) -> object:
        return Chart(self.sheet, self.shape)

    @member
    def ShapeRange(self) -> object:
        return ShapeRange([ShapeObject(self.sheet, self.shape)])

    @member
    def Left(self) -> object:
        return float(self.shape.left)

    @setter('Left')
    def _left(self, value: object) -> None:
        self._dimension('left', value)

    @member
    def Top(self) -> object:
        return float(self.shape.top)

    @setter('Top')
    def _top(self, value: object) -> None:
        self._dimension('top', value)

    @member
    def Width(self) -> object:
        return float(self.shape.width)

    @setter('Width')
    def _width(self, value: object) -> None:
        self._dimension('width', value)

    @member
    def Height(self) -> object:
        return float(self.shape.height)

    @setter('Height')
    def _height(self, value: object) -> None:
        self._dimension('height', value)

    def _dimension(self, name: str, value: object) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, f'ChartObject.{name}')
        number = float(to_number(value))
        if not math.isfinite(number):
            raise error(1004, 'Invalid chart geometry')
        quantized = float(VBASingle(points(emu(number))))
        setattr(self.shape, name, quantized)
        self.sheet.drawing_changed()

    @method
    def Delete(self) -> object:
        return ShapeObject(self.sheet, self.shape).Delete()


class Chart(Created):
    vba_type_name = 'Chart'

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet, self.shape = sheet, shape
        self.state = state(sheet, shape)

    def admit_read(self) -> None:
        if (not any(book is self.sheet.book for book in self.sheet.book.application.workbooks_.books)
                or not any(sheet is self.sheet for sheet in self.sheet.book.sheets_)
                or not any(shape is self.shape for shape in self.sheet.shapes_)
                or self.sheet.chart_states.get(self.shape.shape_id) is not self.state):
            raise VBAUnsupportedError('Accessing a detached chart is not implemented')

    def admit_change(self) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        self.admit_read()
        guard_drawing(self.sheet, 'editing a Chart')
        if self.state.original:
            raise VBAUnsupportedError('Editing the contents of an imported chart is not implemented')

    def changed(self) -> None:
        self.admit_change()
        self.state.dirty = True
        self.sheet.drawing_changed()

    @member
    def Parent(self) -> object:
        return ChartObject(self.sheet, self.shape)

    @member
    def ChartType(self) -> object:
        if self.state.chart_type == 0:
            raise VBAUnsupportedError('Reading this imported chart type is not implemented')
        return VBAInt(shown_type(self.state), 'Long')

    @setter('ChartType')
    def _chart_type(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in CHART_TYPES:
            raise VBAUnsupportedError(f'Chart type {wanted} is not implemented')
        from pyopenvba.apps.excel._chart_axes import transition

        self.admit_change()
        transition(self, wanted)
        self.changed()
        self.state.chart_type = wanted

    @member
    def HasTitle(self) -> object:
        return self.state.title is not None

    @setter('HasTitle')
    def _has_title(self, value: object) -> None:
        self.admit_change()
        wanted = to_bool(value)
        self.changed()
        if not wanted:
            self.state.title, self.state.title_typed = None, False
        elif self.state.title is None:
            self.state.title = 'Chart Title'

    @member
    def HasLegend(self) -> object:
        return self.state.legend

    @setter('HasLegend')
    def _legend(self, value: object) -> None:
        self.admit_change()
        wanted = to_bool(value)
        self.changed()
        self.state.legend = wanted

    @member
    def ChartTitle(self) -> object:
        if self.state.title is None:
            # Excel gives the chart its automatic title on the way to the object.
            self.changed()
            self.state.title = 'Chart Title'
        return ChartTitle(self)

    @member
    def PlotBy(self) -> object:
        self.admit_read()
        return VBAInt(plot_by(self.state), 'Long')

    def display(self) -> ChartState:
        self.admit_read()
        if not self.state.display_readable:
            raise VBAUnsupportedError("Reading this imported chart's display options is not implemented")
        return self.state

    @member
    def DisplayBlanksAs(self) -> object:
        return VBAInt(self.display().blanks, 'Long')

    @setter('DisplayBlanksAs')
    def _blanks(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        self.changed()
        self.state.blanks = wanted

    @member
    def PlotVisibleOnly(self) -> object:
        return self.display().visible_only

    @setter('PlotVisibleOnly')
    def _visible_only(self, value: object) -> None:
        wanted = to_bool(value)
        self.changed()
        self.state.visible_only = wanted

    @member
    def ShowDataLabelsOverMaximum(self) -> object:
        return self.display().over_maximum

    @setter('ShowDataLabelsOverMaximum')
    def _over_maximum(self, value: object) -> None:
        wanted = to_bool(value)
        self.changed()
        self.state.over_maximum = wanted

    @member
    def DisplayValueNotAvailableAsBlank(self) -> object:
        return self.display().not_available_as_blank

    @setter('DisplayValueNotAvailableAsBlank')
    def _not_available_as_blank(self, value: object) -> None:
        wanted = to_bool(value)
        self.changed()
        self.state.not_available_as_blank = wanted

    @method
    def SeriesCollection(self, Index: object = MISSING) -> object:
        collection = SeriesCollection(self)
        return collection if Index is MISSING else collection.vba_get('Item', [Index])

    @method
    def Axes(self, Type: object = MISSING, AxisGroup: object = MISSING) -> object:  # noqa: N803 - native named arguments
        from pyopenvba.apps.excel._chart_axes import Axes, axis

        self.admit_read()
        if Type is MISSING and AxisGroup is not MISSING:
            raise VBAUnsupportedError('Filtered chart axis collections are not implemented')
        return Axes(self) if Type is MISSING else axis(self, Type, AxisGroup)

    @member
    def HasAxis(self, Index1: object = MISSING, Index2: object = MISSING) -> object:  # noqa: N803 - native named arguments
        from pyopenvba.apps.excel._chart_axes import axis

        if Index1 is MISSING:
            raise VBAUnsupportedError('Unindexed chart axis presence is not implemented')
        found = axis(self, Index1, Index2)
        return found.entry.present

    @setter('HasAxis')
    def _axis_presence(self, Index1: object, Index2: object, value: object) -> None:  # noqa: N803 - native named arguments
        from pyopenvba.apps.excel._chart_axes import NO_AXES, axis

        found = axis(self, Index1, Index2)
        self.admit_change()
        wanted = to_bool(value)
        self.admit_change()
        found = axis(self, found.entry.kind, 1)
        self.changed()
        found.entry.present = wanted and self.state.chart_type not in NO_AXES

    @method
    def SetSourceData(self, Source: object = MISSING, PlotBy: object = MISSING) -> object:
        from pyopenvba.apps.excel._model import Range

        if Source is MISSING:
            raise error(450)
        if Source is NOTHING:
            raise VBAUnsupportedError('Chart source Nothing object lifetime behavior is not implemented')
        if not isinstance(Source, Range):
            raise error(13)
        if len(Source.areas) != 1:
            raise VBAUnsupportedError('Multi-area chart source data is not implemented')
        if Source.sheet.book is not self.sheet.book:
            raise VBAUnsupportedError('Cross-workbook chart source data is not implemented')
        area = Source.first
        height, width = area.bottom - area.top + 1, area.right - area.left + 1
        if height * width > MAX_CHART_POINTS:
            raise VBAUnsupportedError('Chart source data exceeds the supported cell budget')
        if PlotBy is NULL:
            raise error(13)
        orientation = (1 if width >= height else 2) if PlotBy is MISSING else int(to_integer(PlotBy, 'Long'))
        if orientation not in (1, 2):
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        self.admit_change()
        series_plan = plan_series(Source.sheet, area, orientation)
        self.changed()
        self.state.plot_by = orientation
        for index, series in enumerate(series_plan):
            series.unique = self.state.unique_base + index
        self.state.next_unique = self.state.unique_base + len(series_plan)
        self.state.series[:] = series_plan
        return EMPTY


class ChartTitle(Created):
    vba_type_name = 'ChartTitle'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    @member
    def Text(self) -> object:
        state = self.chart.state
        if state.title is not None and not state.title_typed and len(state.series) == 1:
            # The automatic title of a chart with one series is that series' name.
            return Series(self.chart, state.series[0]).Name()
        return state.title or ''

    @setter('Text')
    def _text(self, value: object) -> None:
        self.chart.admit_change()
        wanted = to_text(value)
        self.chart.changed()
        self.chart.state.title, self.chart.state.title_typed = wanted, True


class SeriesCollection(VBACollection, Created):
    vba_type_name = 'SeriesCollection'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def vba_items(self) -> list[object]:
        self.chart.admit_read()
        return [Series(self.chart, entry) for entry in self.chart.state.series]

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @method
    def NewSeries(self) -> object:
        self.chart.admit_change()
        used = {entry.series_id if entry.series_id >= 0 else index
                for index, entry in enumerate(self.chart.state.series)}
        series_id = 0
        while series_id in used:
            series_id += 1
        entry = SeriesState(series_id=series_id, unique=self.chart.state.next_unique)
        self.chart.changed()
        self.chart.state.next_unique += 1
        self.chart.state.series.append(entry)
        return Series(self.chart, entry)


class Series(Created):
    vba_type_name = 'Series'

    def __init__(self, chart: Chart, entry: SeriesState) -> None:
        self.chart, self.entry = chart, entry

    def position(self) -> int:
        self.chart.admit_read()
        for index, entry in enumerate(self.chart.state.series):
            if entry is self.entry:
                return index + 1
        raise error(-2147467261, 'This object is no longer valid.')

    def vba_call_by_name(self, name: str, args: Sequence[object], *, assignment: bool = False,
                         by_ref: bool = False) -> object:
        self.chart.admit_read()
        if (name.lower() in ('name', 'plotorder', 'formula', 'values', 'xvalues')
                and not any(entry is self.entry for entry in self.chart.state.series)):
            raise error(440, 'Automation error')
        return super().vba_call_by_name(name, args, assignment=assignment, by_ref=by_ref)

    @member
    def PlotOrder(self) -> object:
        return VBAInt(self.position(), 'Long')

    @setter('PlotOrder')
    def _plot_order(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        if value is NULL:
            raise error(13, 'Type mismatch')
        wanted = int(to_integer(value, 'Long'))
        current = self.position()
        self.chart.admit_change()
        if not 1 <= wanted <= len(self.chart.state.series):
            raise error(1004, 'Application-defined or object-defined error')
        if current == wanted:
            return
        entries = self.chart.state.series[:]
        entries.pop(current - 1)
        entries.insert(wanted - 1, self.entry)
        self.chart.changed()
        self.chart.state.series[:] = entries

    @method
    def Delete(self) -> object:
        position = self.position()
        self.chart.changed()
        self.chart.state.series.pop(position - 1)
        return True

    @member
    def Name(self) -> object:
        self.position()
        if self.entry.name_reference:
            values = referenced_values(self.chart.sheet, self.entry.name_reference)
            return to_text(values[0]) if values else ''
        identity = self.entry.series_id + 1 if self.entry.series_id >= 0 else self.position()
        return self.entry.name if self.entry.name is not None else f'Series{identity}'

    @setter('Name')
    def _name(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        text = to_text(value)
        if text.startswith('=') and len(text) > MAX_CHART_REFERENCE_TEXT:
            raise VBAUnsupportedError('Chart references exceed the supported text budget')
        self.position()
        self.chart.changed()
        self.entry.name_reference = text[1:] if text.startswith('=') else ''
        self.entry.name = text if not self.entry.name_reference else ''

    @member
    def Values(self) -> object:
        values = [float(value) if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) else value
                  for value in self.values()]
        return VBAArray([(1, len(values))], items=values)

    @setter('Values')
    def _values(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        reference, values = series_input(value, sheet=self.chart.sheet)
        self.position()
        self.chart.changed()
        self.entry.value_reference, self.entry.values = reference, values

    @member
    def XValues(self) -> object:
        values = [float(value) if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) else value
                  for value in self.categories()]
        return VBAArray([(1, len(values))], items=values)

    @setter('XValues')
    def _categories(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        reference, values = series_input(value, sheet=self.chart.sheet)
        self.position()
        self.chart.changed()
        self.entry.category_reference, self.entry.categories = reference, values

    def values(self) -> list[object]:
        self.position()
        return referenced_values(self.chart.sheet, self.entry.value_reference) if self.entry.value_reference else self.entry.values

    def categories(self) -> list[object]:
        self.position()
        if self.entry.category_reference:
            return referenced_values(self.chart.sheet, self.entry.category_reference)
        return self.entry.categories or [float(index + 1) for index in range(len(self.values()))]

    @member
    def Formula(self) -> object:
        order = self.position()
        name = self.entry.name_reference or (quoted(self.entry.name) if self.entry.name is not None else '')
        categories = self.entry.category_reference or literal_array(self.entry.categories)
        values = self.entry.value_reference or literal_array(self.entry.values)
        return f'=SERIES({name},{categories},{values},{order})'


def quoted(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'

def literal_array(values: list[object]) -> str:
    if not values:
        return ''
    return '{' + ','.join(quoted(value) if isinstance(value, str) else to_text(value) for value in values) + '}'

def reference(sheet: Worksheet, area: Area) -> str:
    from pyopenvba.apps.excel._model import Range

    return quote_sheet_name(sheet.name) + '!' + str(Range(sheet, [area]).Address())

def series_input(value: object, *, sheet: Worksheet) -> tuple[str, list[object]]:
    from pyopenvba.apps.excel._model import Range

    if isinstance(value, Range):
        if value.sheet.book is not sheet.book or len(value.areas) != 1:
            raise VBAUnsupportedError('Cross-workbook or multi-area chart series input is not implemented')
        if value.first.rows * value.first.columns > MAX_CHART_POINTS:
            raise VBAUnsupportedError('Chart series input exceeds the supported cell budget')
        return reference(value.sheet, value.first), []
    if isinstance(value, VBAArray):
        if len(value.items) > MAX_CHART_POINTS:
            raise VBAUnsupportedError('Chart series input exceeds the supported point budget')
        return '', value.elements()
    if isinstance(value, str) and value.startswith('='):
        if len(value) > MAX_CHART_REFERENCE_TEXT:
            raise VBAUnsupportedError('Chart references exceed the supported text budget')
        return value[1:], []
    raise error(1004, 'Series data requires a range or array')

def referenced_values(sheet: Worksheet, formula: str) -> list[object]:
    from pyopenvba.apps.excel._model import Range

    if len(formula) > MAX_CHART_REFERENCE_TEXT:
        raise VBAUnsupportedError('Chart references exceed the supported text budget')
    try:
        areas = parse_reference(formula)
    except ValueError as failure:
        raise VBAUnsupportedError(f'Chart reference expression {formula!r} is not implemented') from failure
    if sum((area.bottom - area.top + 1) * (area.right - area.left + 1) for area in areas) > MAX_CHART_POINTS:
        raise VBAUnsupportedError('Chart references exceed the supported cell budget')
    owners = [sheet.book.sheet_named(area.sheet) if area.sheet else sheet for area in areas]
    result: list[object] = []
    for area, owner in zip(areas, owners, strict=True):
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                result.append(Range(owner, [Area(row, column, row, column)]).Value2())
    return result
