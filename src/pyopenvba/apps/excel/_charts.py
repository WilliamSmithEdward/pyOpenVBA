"""Embedded charts share their worksheet's drawing and cell state."""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
import math
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference
from pyopenvba.apps.excel._shapes import ShapeObject, ShapeRange, Shapes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, VBAInt, VBASingle, error, to_bool, to_number, to_text
from pyopenvba.shapes._values import Shape, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

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

@dataclass
class ChartState:
    chart_type: int = 51
    title: str | None = None
    legend: bool = True
    plot_by: int = 2
    series: list[SeriesState] = field(default_factory=lambda: [])
    part: str = ''
    original: str = ''
    dirty: bool = True


def state(sheet: Worksheet, shape: Shape) -> ChartState:
    if shape.shape_id not in sheet.chart_states:
        if shape.source:
            raise VBAUnsupportedError('Chart metadata for this drawing has not been loaded')
        sheet.chart_states[shape.shape_id] = ChartState()
    return sheet.chart_states[shape.shape_id]


class ChartObjects(VBACollection):
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
        self.sheet.chart_states[shape.shape_id] = ChartState()
        return ChartObject(self.sheet, shape)

    @method
    def Delete(self) -> object:
        for item in self.vba_items():
            assert isinstance(item, ChartObject)
            item.Delete()
        return EMPTY


class ChartObject(VBAObject):
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


class Chart(VBAObject):
    vba_type_name = 'Chart'

    def __init__(self, sheet: Worksheet, shape: Shape) -> None:
        self.sheet, self.shape = sheet, shape
        self.state = state(sheet, shape)

    def changed(self) -> None:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, 'editing a Chart')
        if self.state.original:
            raise VBAUnsupportedError('Editing the contents of an imported chart is not implemented')
        self.state.dirty = True
        self.sheet.drawing_changed()

    @member
    def Parent(self) -> object:
        return ChartObject(self.sheet, self.shape)

    @member
    def ChartType(self) -> object:
        if self.state.chart_type == 0:
            raise VBAUnsupportedError('Reading this imported chart type is not implemented')
        return VBAInt(self.state.chart_type, 'Long')

    @setter('ChartType')
    def _chart_type(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in CHART_TYPES:
            raise VBAUnsupportedError(f'Chart type {wanted} is not implemented')
        self.changed()
        self.state.chart_type = 65 if wanted == 4 and not self.state.series else wanted

    @member
    def HasTitle(self) -> object:
        return self.state.title is not None

    @setter('HasTitle')
    def _has_title(self, value: object) -> None:
        self.changed()
        self.state.title = (self.state.title or 'Chart Title') if to_bool(value) else None

    @member
    def HasLegend(self) -> object:
        return self.state.legend

    @setter('HasLegend')
    def _legend(self, value: object) -> None:
        self.changed()
        self.state.legend = to_bool(value)

    @member
    def ChartTitle(self) -> object:
        if self.state.title is None:
            raise error(1004, 'This chart has no title')
        return ChartTitle(self)

    @member
    def PlotBy(self) -> object:
        return VBAInt(self.state.plot_by, 'Long')

    @method
    def SeriesCollection(self, Index: object = MISSING) -> object:
        collection = SeriesCollection(self)
        return collection if Index is MISSING else collection.vba_get('Item', [Index])

    @method
    def SetSourceData(self, Source: object = MISSING, PlotBy: object = MISSING) -> object:
        from pyopenvba.apps.excel._model import Range

        if not isinstance(Source, Range) or len(Source.areas) != 1:
            raise error(1004, 'SetSourceData requires a contiguous range')
        area = Source.first
        height, width = area.bottom - area.top + 1, area.right - area.left + 1
        orientation = (1 if width >= height else 2) if PlotBy is MISSING else int(to_number(PlotBy))
        if orientation not in (1, 2):
            raise error(1004, 'Invalid PlotBy value')
        def value(row: int, column: int) -> object:
            return Range(Source.sheet, [Area(row, column, row, column)]).Value()
        def text_axis(values: list[object]) -> bool:
            return bool(values) and all(isinstance(item, str) or item is EMPTY for item in values) and any(
                isinstance(item, str) and bool(item) for item in values)
        self.changed()
        self.state.plot_by = orientation
        self.state.series.clear()
        if orientation == 2:
            names = text_axis([value(area.top, col) for col in range(area.left + (width > 1), area.right + 1)])
            top = area.top + names
            labels = text_axis([value(row, area.left) for row in range(top, area.bottom + 1)])
            for column in range(area.left + labels, area.right + 1):
                series = SeriesState()
                if names:
                    series.name_reference = reference(Source.sheet, Area(area.top, column, area.top, column))
                series.value_reference = reference(Source.sheet, Area(top, column, area.bottom, column))
                if labels:
                    series.category_reference = reference(Source.sheet, Area(top, area.left, area.bottom, area.left))
                self.state.series.append(series)
        else:
            names = text_axis([value(row, area.left) for row in range(area.top + (height > 1), area.bottom + 1)])
            left = area.left + names
            labels = text_axis([value(area.top, col) for col in range(left, area.right + 1)])
            for row in range(area.top + labels, area.bottom + 1):
                series = SeriesState()
                if names:
                    series.name_reference = reference(Source.sheet, Area(row, area.left, row, area.left))
                series.value_reference = reference(Source.sheet, Area(row, left, row, area.right))
                if labels:
                    series.category_reference = reference(Source.sheet, Area(area.top, left, area.top, area.right))
                self.state.series.append(series)
        return EMPTY


class ChartTitle(VBAObject):
    vba_type_name = 'ChartTitle'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    @member
    def Text(self) -> object:
        return self.chart.state.title or ''

    @setter('Text')
    def _text(self, value: object) -> None:
        self.chart.changed()
        self.chart.state.title = to_text(value)


class SeriesCollection(VBACollection):
    vba_type_name = 'SeriesCollection'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def vba_items(self) -> list[object]:
        return [Series(self.chart, entry) for entry in self.chart.state.series]

    @method
    def NewSeries(self) -> object:
        self.chart.changed()
        entry = SeriesState()
        self.chart.state.series.append(entry)
        return Series(self.chart, entry)


class Series(VBAObject):
    vba_type_name = 'Series'

    def __init__(self, chart: Chart, entry: SeriesState) -> None:
        self.chart, self.entry = chart, entry

    def position(self) -> int:
        return next(index + 1 for index, entry in enumerate(self.chart.state.series) if entry is self.entry)

    @member
    def Name(self) -> object:
        if self.entry.name_reference:
            values = referenced_values(self.chart.sheet, self.entry.name_reference)
            return to_text(values[0]) if values else ''
        return self.entry.name if self.entry.name is not None else f'Series{self.position()}'

    @setter('Name')
    def _name(self, value: object) -> None:
        self.chart.changed()
        text = to_text(value)
        self.entry.name_reference = text[1:] if text.startswith('=') else ''
        self.entry.name = text if not self.entry.name_reference else ''

    @member
    def Values(self) -> object:
        values = [float(value) if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) else value
                  for value in self.values()]
        return VBAArray([(1, len(values))], items=values)

    @setter('Values')
    def _values(self, value: object) -> None:
        self.chart.changed()
        self.entry.value_reference, self.entry.values = series_input(value)

    @member
    def XValues(self) -> object:
        values = self.categories()
        return VBAArray([(1, len(values))], items=values)

    @setter('XValues')
    def _categories(self, value: object) -> None:
        self.chart.changed()
        self.entry.category_reference, self.entry.categories = series_input(value)

    def values(self) -> list[object]:
        return referenced_values(self.chart.sheet, self.entry.value_reference) if self.entry.value_reference else self.entry.values

    def categories(self) -> list[object]:
        if self.entry.category_reference:
            return referenced_values(self.chart.sheet, self.entry.category_reference)
        return self.entry.categories or [float(index + 1) for index in range(len(self.values()))]

    @member
    def Formula(self) -> object:
        name = self.entry.name_reference or (quoted(self.entry.name) if self.entry.name is not None else '')
        categories = self.entry.category_reference or literal_array(self.entry.categories)
        values = self.entry.value_reference or literal_array(self.entry.values)
        order = self.position()
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

def series_input(value: object) -> tuple[str, list[object]]:
    from pyopenvba.apps.excel._model import Range

    if isinstance(value, Range):
        return reference(value.sheet, value.first), []
    if isinstance(value, VBAArray):
        return '', list(value.elements())
    if isinstance(value, str) and value.startswith('='):
        return value[1:], []
    raise error(1004, 'Series data requires a range or array')

def referenced_values(sheet: Worksheet, formula: str) -> list[object]:
    from pyopenvba.apps.excel._model import Range

    try:
        areas = parse_reference(formula)
    except ValueError as failure:
        raise VBAUnsupportedError(f'Chart reference expression {formula!r} is not implemented') from failure
    result: list[object] = []
    for area in areas:
        owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                result.append(Range(owner, [Area(row, column, row, column)]).Value2())
    return result
