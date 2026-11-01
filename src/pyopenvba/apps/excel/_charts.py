"""Embedded charts share their worksheet's drawing and cell state."""
from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from decimal import Decimal
import math
import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference
from pyopenvba.apps.excel._chart_axes import AxisState
from pyopenvba.apps.excel._chart_title import TitleLook, TitleMembers
from pyopenvba.apps.excel._shapes import PLACEMENTS, ShapeObject, ShapeRange, Shapes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, NULL, VBAArray, VBADate, VBAInt, VBASingle, error, to_bool, to_integer, to_number, to_text
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
    smooth: bool = False
    marker: int = -4142
    #: MarkerSize once given; until then 9 for an automatic marker and 5 for the rest.
    marker_size: int | None = None
    invert_if_negative: bool = False
    explosion: int = 0
    #: HasDataLabels, and HasLeaderLines, which turning the labels on turns off (tests/fixtures/chart_data_labels.json).
    data_labels: bool = False
    leader_lines: bool = True
    #: MarkerForegroundColor and MarkerBackgroundColor as RRGGBB, "none", or None while automatic; "?" for an
    #: imported marker filled in a way the model does not read (tests/fixtures/chart_marker_colours.json).
    marker_fore: str | None = None
    marker_back: str | None = None

@dataclass
class ChartState:
    chart_type: int = 51
    title: str | None = None
    #: False for Excel's automatic title, which the file holds with no text.
    title_typed: bool = False
    title_look: TitleLook = field(default_factory=TitleLook)
    #: Set by ChartTitle.Delete, which Excel records as autoTitleDeleted.
    title_deleted: bool = False
    legend: bool = True
    #: The legend's Position as the file spells it, IncludeInLayout turned round, and Shadow once set.
    legend_position: str = 'r'
    legend_overlay: bool = False
    legend_shadow: bool | None = None
    #: The legend entries deleted, by series or, in a pie, point index (tests/fixtures/chart_legend_entries.json).
    legend_deleted: list[int] = field(default_factory=lambda: [])
    #: The chart group's GapWidth, Overlap, VaryByCategories once set, FirstSliceAngle and DoughnutHoleSize once set
    #: (tests/fixtures/chart_groups.json).
    gap_width: int = 150
    overlap: int = 0
    vary: bool | None = None
    first_slice: int = 0
    hole: int | None = None
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
    data_table: bool = False


def label_block(sheet: Worksheet, area: Area) -> tuple[int, int]:
    """How many of a block's top rows and left columns label the rest, as SetSourceData reads it
    (tests/fixtures/chart_multilevel_categories.json).

    A blank top left corner says so itself: its blank cells along the top row and down the left column, one where
    they run the whole way. Otherwise the left columns run to the last text or date in the bottom row, and the top
    rows to the last in the right column, whatever the cells before them hold. Either way at least one row and one
    column are left for the values."""
    from pyopenvba.apps.excel._model import Range

    height, width = area.bottom - area.top + 1, area.right - area.left + 1

    def value(row: int, column: int) -> object:
        return Range(sheet, [Area(row, column, row, column)]).Value()

    def labelling(item: object) -> bool:
        # Dates label as text does (tests/fixtures/chart_category_type.json).
        return isinstance(item, str) and bool(item) or isinstance(item, VBADate)

    def run(cells: list[object], blank: bool) -> int:
        if blank:
            count = next((index for index, item in enumerate(cells) if item is not EMPTY), len(cells))
            return 1 if count == len(cells) else count
        return max((index + 1 for index, item in enumerate(cells) if labelling(item)), default=0)

    blank = value(area.top, area.left) is EMPTY
    rows = run([value(row, area.left if blank else area.right) for row in range(area.top, area.bottom + 1)], blank)
    columns = run([value(area.top if blank else area.bottom, column)
                   for column in range(area.left, area.right + 1)], blank)
    return min(rows, height - 1), min(columns, width - 1)


def plan_series(sheet: Worksheet, area: Area, orientation: int) -> list[SeriesState]:
    """The series SetSourceData makes from a block of cells, by rows (1) or columns (2). The rows and columns
    label_block finds name the series and label the categories, more than one of either making a name or categories
    of several levels (tests/fixtures/chart_multilevel_categories.json)."""
    rows, columns = label_block(sheet, area)
    series_plan: list[SeriesState] = []
    if orientation == 2:
        top, left = area.top + rows, area.left + columns
        if area.right - left + 1 > MAX_CHART_SERIES:
            raise VBAUnsupportedError('Chart source data exceeds the supported series budget')
        for column in range(left, area.right + 1):
            series = SeriesState(series_id=len(series_plan))
            if rows:
                series.name_reference = reference(sheet, Area(area.top, column, top - 1, column))
            series.value_reference = reference(sheet, Area(top, column, area.bottom, column))
            if columns:
                series.category_reference = reference(sheet, Area(top, area.left, area.bottom, left - 1))
            series_plan.append(series)
    else:
        top, left = area.top + rows, area.left + columns
        if area.bottom - top + 1 > MAX_CHART_SERIES:
            raise VBAUnsupportedError('Chart source data exceeds the supported series budget')
        for row in range(top, area.bottom + 1):
            series = SeriesState(series_id=len(series_plan))
            if columns:
                series.name_reference = reference(sheet, Area(row, area.left, row, left - 1))
            series.value_reference = reference(sheet, Area(row, left, row, area.right))
            if rows:
                series.category_reference = reference(sheet, Area(area.top, left, top - 1, area.right))
            series_plan.append(series)
    return series_plan


def shown_cells(sheet: Worksheet, formula: str) -> list[list[str]]:
    """A reference's cells, row by row, as their formats show them whatever their columns' width: how a chart's
    cache holds a name or categories of several cells (tests/fixtures/chart_multilevel_categories.json)."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.formula._display import UndisplayableError, displayed

    area = parse_reference(formula)[0]
    owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
    grid: list[list[str]] = []
    for row in range(area.top, area.bottom + 1):
        texts: list[str] = []
        for column in range(area.left, area.right + 1):
            cell = Range(owner, [Area(row, column, row, column)])
            value = cell.Value2()
            try:
                texts.append('' if value is EMPTY else displayed(
                    value, to_text(cell.NumberFormat()), epoch_1904=owner.book.epoch_1904,
                    locale=owner.book.application.locale).text if isinstance(value, (bool, str, int, float))
                    else to_text(value))
            except UndisplayableError:
                texts.append(to_text(value))
        grid.append(texts)
    return grid


def category_levels(sheet: Worksheet, entry: SeriesState) -> list[list[str]] | None:
    """A series' categories when they take several rows or columns, one list of texts for each level, the outer
    first; None for categories of one level. Its points run along the side as long as its values, down the rows
    where both sides are (tests/fixtures/chart_multilevel_categories.json)."""
    if not entry.category_reference:
        return None
    try:
        areas = parse_reference(entry.category_reference)
    except ValueError:
        return None
    if len(areas) != 1:
        return None
    area = areas[0]
    height, width = area.bottom - area.top + 1, area.right - area.left + 1
    if height * width > MAX_CHART_POINTS:
        raise VBAUnsupportedError('Chart references exceed the supported cell budget')
    points = len(referenced_values(sheet, entry.value_reference)) if entry.value_reference else len(entry.values)
    down = height == points if (height == points) != (width == points) else height >= width
    if (width if down else height) == 1:
        return None
    grid = shown_cells(sheet, entry.category_reference)
    return [list(level) for level in zip(*grid, strict=True)] if down else grid


def joined_categories(levels: list[list[str]]) -> list[str]:
    """What XValues reads of categories of several levels: each point's texts, the outer first and blanks left out,
    an outer level's blank taking the text above it (tests/fixtures/chart_multilevel_categories.json)."""
    carried = [''] * len(levels)
    texts: list[str] = []
    for index in range(len(levels[0])):
        parts: list[str] = []
        for depth, level in enumerate(levels):
            text = level[index]
            if depth < len(levels) - 1:
                carried[depth] = text = text or carried[depth]
            if text:
                parts.append(text)
        texts.append(' '.join(parts))
    return texts


def plotted_before(sheet: Worksheet) -> int:
    """How many series ChartObjects.Add would plot, and drops: those of the selection, or of the active cell's current
    region, as SetSourceData plots them by default. Excel numbers later series' unique ids past them."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._region import current_region

    selection = sheet.selection
    if not isinstance(selection, Range):
        return 0
    area = selection.first
    region = len(selection.areas) == 1 and (area.top, area.left) == (area.bottom, area.right)
    if region:
        area = current_region(sheet, area.top, area.left)
    height, width = area.bottom - area.top + 1, area.right - area.left + 1
    if height * width > MAX_CHART_POINTS:
        raise VBAUnsupportedError('Chart source data exceeds the supported cell budget')
    if region:
        # A current region loses its rows and columns blank right across, a selected block does not
        # (tests/fixtures/chart_add_plotted.json).
        found = trimmed(sheet, area)
        if found is None:
            return 0
        area = found
    return len(plan_series(sheet, area, guessed_orientation(sheet, area)))


def trimmed(sheet: Worksheet, area: Area) -> Area | None:
    """The block without the rows and columns at its edges that are blank right across; None if nothing is left."""
    from pyopenvba.apps.excel._model import Range

    def blank(part: Area) -> bool:
        return all(Range(sheet, [Area(row, column, row, column)]).Value() is EMPTY
                   for row in range(part.top, part.bottom + 1) for column in range(part.left, part.right + 1))

    top, left, bottom, right = area.top, area.left, area.bottom, area.right
    while top <= bottom and blank(Area(top, left, top, right)):
        top += 1
    while bottom >= top and blank(Area(bottom, left, bottom, right)):
        bottom -= 1
    if top > bottom:
        return None
    while left <= right and blank(Area(top, left, bottom, left)):
        left += 1
    while right >= left and blank(Area(top, right, bottom, right)):
        right -= 1
    return Area(top, left, bottom, right)


def guessed_orientation(sheet: Worksheet, area: Area) -> int:
    """How a block is plotted when nothing says: by rows (1) unless what its labels leave is taller than wide
    (tests/fixtures/chart_add_plotted.json)."""
    rows, columns = label_block(sheet, area)
    height, width = area.bottom - area.top + 1 - rows, area.right - area.left + 1 - columns
    return 1 if width >= height else 2


def shown_type(chart: ChartState) -> int:
    """The type a chart reads and is written as. A line chart follows its first series: with markers (65) unless that
    series has none, and with markers while it has no series at all."""
    if chart.chart_type in (4, 65):
        return 4 if chart.series and chart.series[0].marker == MARKER_NONE else 65
    return chart.chart_type


#: Series.MarkerStyle and the symbol the file spells it with; automatic has no symbol of its own.
MARKER_AUTOMATIC, MARKER_NONE = -4105, -4142
#: An imported series whose marker symbol the model does not read.
MARKER_UNREAD = 0
MARKER_SYMBOLS = {MARKER_AUTOMATIC: '', MARKER_NONE: 'none', 1: 'square', 2: 'diamond', 3: 'triangle', -4168: 'x',
                  5: 'star', 8: 'circle', 9: 'plus', -4118: 'dot', -4115: 'dash', -4147: 'picture'}
#: Older marker numbers MarkerStyle still takes: 0 for none, and Excel 4's x, short bar and long bar.
MARKER_ALIASES = {0: MARKER_NONE, 4: -4168, 6: -4118, 7: -4115}
#: The marker an automatic series shows, by its idx; MarkerSize fixes it.
AUTOMATIC_MARKERS = (2, 1, 3, -4168, 5, 8, 9, -4118, -4115)


def default_marker(chart: ChartState) -> int:
    """The marker a new series takes: automatic in a chart drawn with markers, none otherwise."""
    return MARKER_AUTOMATIC if chart.chart_type in (65, -4169) else MARKER_NONE


def marker_markup(entry: SeriesState) -> str:
    """The series' marker element: none for an automatic marker, a symbol for the rest, with its size where it has
    one."""
    symbol = MARKER_SYMBOLS[entry.marker]
    look = marker_look(entry)
    if not symbol:
        return f'<c:marker>{look}</c:marker>' if look else ''
    if symbol in ('none', 'picture'):
        return f'<c:marker><c:symbol val="{symbol}"/>{look}</c:marker>'
    return f'<c:marker><c:symbol val="{symbol}"/><c:size val="{entry.marker_size or 5}"/>{look}</c:marker>'


def marker_look(entry: SeriesState) -> str:
    """The marker's spPr: its background as the fill, its foreground as the line, as Excel writes them
    (tests/fixtures/chart_marker_colours.json)."""
    fill = {None: '', 'none': '<a:noFill/>'}.get(
        entry.marker_back, f'<a:solidFill><a:srgbClr val="{entry.marker_back}"/></a:solidFill>')
    line = {None: '', 'none': '<a:ln w="12700"><a:noFill/></a:ln>'}.get(
        entry.marker_fore,
        f'<a:ln><a:solidFill><a:srgbClr val="{entry.marker_fore}"/></a:solidFill><a:prstDash val="solid"/></a:ln>')
    return f'<c:spPr>{fill}{line}</c:spPr>' if fill or line else ''


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


class ChartObjects(VBACollection, ItemDefault, Created):
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

    # -- every chart at once (tests/fixtures/chart_object_more.json)

    def _shared(self, name: str) -> object:
        values = [item.vba_get(name) for item in self.vba_items() if isinstance(item, ChartObject)]
        if not values or any(value != values[0] for value in values[1:]):
            raise VBAUnsupportedError(f'ChartObjects.{name} over charts that differ, or none, is not implemented')
        return values[0]

    def _spread(self, name: str, value: object) -> None:
        for item in self.vba_items():
            assert isinstance(item, ChartObject)
            item.vba_set(name, value)

    @member
    def Left(self) -> object:
        return self._shared('Left')

    @setter('Left')
    def _left(self, value: object) -> None:
        self._spread('Left', value)

    @member
    def Top(self) -> object:
        return self._shared('Top')

    @setter('Top')
    def _top(self, value: object) -> None:
        self._spread('Top', value)

    @member
    def Width(self) -> object:
        return self._shared('Width')

    @setter('Width')
    def _width(self, value: object) -> None:
        self._spread('Width', value)

    @member
    def Height(self) -> object:
        return self._shared('Height')

    @setter('Height')
    def _height(self, value: object) -> None:
        self._spread('Height', value)

    @member
    def Visible(self) -> object:
        return self._shared('Visible')

    @setter('Visible')
    def _visible(self, value: object) -> None:
        self._spread('Visible', value)

    # Excel's ChartObjects takes none of these three settings: 445, the object does not support this action.

    @setter('Locked')
    def _locked(self, value: object) -> None:
        raise error(445, 'Object doesn\'t support this action')

    @setter('PrintObject')
    def _print_object(self, value: object) -> None:
        raise error(445, 'Object doesn\'t support this action')

    @setter('Placement')
    def _placement(self, value: object) -> None:
        raise error(445, 'Object doesn\'t support this action')

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

    # -- tests/fixtures/chart_object_more.json

    def _changed(self, name: str) -> Shape:
        from pyopenvba.apps.excel._protection import guard_drawing

        guard_drawing(self.sheet, f'ChartObject.{name}')
        self.sheet.drawing_changed()
        return self.shape

    @member
    def Visible(self) -> object:
        return not self.shape.hidden

    @setter('Visible')
    def _visible(self, value: object) -> None:
        self._changed('Visible').hidden = not to_bool(value)

    @member
    def Locked(self) -> object:
        return self.shape.locked

    @setter('Locked')
    def _locked(self, value: object) -> None:
        self._changed('Locked').locked = to_bool(value)

    @member
    def PrintObject(self) -> object:
        return self.shape.prints

    @setter('PrintObject')
    def _print_object(self, value: object) -> None:
        self._changed('PrintObject').prints = to_bool(value)

    @member
    def Placement(self) -> object:
        return VBAInt(PLACEMENTS.index(self.shape.edit_as) + 1, 'Long')

    @setter('Placement')
    def _placement(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in (1, 2, 3):
            raise error(-2147467259, 'Unspecified error')
        self._changed('Placement').edit_as = PLACEMENTS[wanted - 1]

    @method
    def BringToFront(self) -> object:
        """The chart on top of the sheet's shapes; Excel answers True."""
        ShapeObject(self.sheet, self.shape).ZOrder(0)
        return True

    @method
    def SendToBack(self) -> object:
        """The chart behind the sheet's shapes; unlike BringToFront, Excel answers nothing."""
        ShapeObject(self.sheet, self.shape).ZOrder(1)
        return EMPTY


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
        # A line chart turns every marker off; one with markers, or a scatter chart, turns those that were off back on.
        for entry in self.state.series:
            if wanted == 4:
                entry.marker = MARKER_NONE
            elif wanted in (65, -4169) and entry.marker == MARKER_NONE:
                entry.marker = MARKER_AUTOMATIC

    @member
    def HasTitle(self) -> object:
        return self.state.title is not None

    @setter('HasTitle')
    def _has_title(self, value: object) -> None:
        self.admit_change()
        wanted = to_bool(value)
        self.changed()
        self.state.title_deleted = False
        if not wanted:
            self.state.title, self.state.title_typed, self.state.title_look = None, False, TitleLook()
        elif self.state.title is None:
            self.state.title = 'Chart Title'

    @member
    def HasLegend(self) -> object:
        return self.state.legend

    @method
    def ChartGroups(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        groups = ChartGroups(self)
        return groups if Index is MISSING else groups.vba_lookup(Index, groups.vba_items())

    @member
    def HasDataTable(self) -> object:
        """Whether the chart shows a data table, on any kind, a pie's included (tests/fixtures/chart_series_more.json)."""
        self.admit_read()
        return self.state.data_table

    @setter('HasDataTable')
    def _has_data_table(self, value: object) -> None:
        self.admit_change()
        wanted = to_bool(value)
        self.changed()
        self.state.data_table = wanted

    @member
    def Name(self) -> object:
        """The sheet's name and the chart object's, as Excel answers it (tests/fixtures/chart_series_more.json)."""
        return f'{self.sheet.name} {self.shape.name}'

    @setter('Name')
    def _chart_name(self, value: object) -> None:
        # An embedded chart's own name cannot be set: Excel raises 7.
        raise error(7, 'Out of memory')

    # An embedded chart's hidden members (tests/fixtures/hidden_members.json): _CodeName raises -2147467259, and
    # _Evaluate takes no name, 450; what it answers given none is not measured.

    @member(name='_CodeName')
    def hidden_code_name(self) -> object:
        raise error(-2147467259, 'Automation error\nUnspecified error')

    @member(name='_Evaluate')
    def hidden_evaluate(self, *arguments: object) -> object:
        if arguments:
            raise error(450, 'Wrong number of arguments or invalid property assignment')
        raise VBAUnsupportedError("An embedded chart's _Evaluate read without a name is not implemented")

    @setter('HasLegend')
    def _legend(self, value: object) -> None:
        self.admit_change()
        wanted = to_bool(value)
        self.changed()
        if not wanted:
            # A legend taken away comes back as new.
            self.state.legend_position, self.state.legend_overlay, self.state.legend_shadow = 'r', False, None
            self.state.legend_deleted = []
        self.state.legend = wanted

    @member
    def Legend(self) -> object:
        self.admit_read()
        if not self.state.legend:
            raise error(1004, 'Unable to get the Legend property of the Chart class')
        return Legend(self)

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
        orientation = (guessed_orientation(Source.sheet, area) if PlotBy is MISSING
                       else int(to_integer(PlotBy, 'Long')))
        if orientation not in (1, 2):
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        self.admit_change()
        series_plan = plan_series(Source.sheet, area, orientation)
        self.changed()
        self.state.plot_by = orientation
        for index, series in enumerate(series_plan):
            series.unique = self.state.unique_base + index
            series.marker = default_marker(self.state)
        self.state.next_unique = self.state.unique_base + len(series_plan)
        self.state.series[:] = series_plan
        return EMPTY


class ChartTitle(TitleMembers):
    vba_type_name = 'ChartTitle'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def title_look(self) -> TitleLook:
        self.chart.admit_read()
        return self.chart.state.title_look

    def title_restyle(self) -> TitleLook:
        self.chart.admit_change()
        self.chart.changed()
        return self.chart.state.title_look

    def title_sheet(self) -> Worksheet:
        self.chart.admit_change()
        return self.chart.sheet

    def title_link(self, formula: str) -> None:
        self.chart.changed()
        self.chart.state.title_typed = False
        self.chart.state.title_look.formula = formula

    @member
    def Text(self) -> object:
        state = self.chart.state
        if state.title_look.formula:
            return linked_title(self.chart.sheet, state.title_look.formula)
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
        self.chart.state.title_look.formula = ''

    @method
    def Delete(self) -> object:
        self.chart.admit_change()
        self.chart.changed()
        state = self.chart.state
        state.title, state.title_typed, state.title_look, state.title_deleted = None, False, TitleLook(), True
        return EMPTY


def linked_title(sheet: Worksheet, formula: str) -> str:
    values = referenced_values(sheet, formula)
    return to_text(values[0]) if values and values[0] is not EMPTY else ''


#: Legend.Position and how the file spells it; xlLegendPositionCustom needs native layout, and Excel refuses it.
LEGEND_POSITIONS = {-4107: 'b', -4160: 't', -4131: 'l', -4152: 'r', 2: 'tr'}


class Legend(Created):
    vba_type_name = 'Legend'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def shown(self) -> ChartState:
        self.chart.admit_read()
        if not self.chart.state.legend:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        return self.chart.state

    def change(self) -> ChartState:
        state = self.shown()
        self.chart.changed()
        return state

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Position(self) -> object:
        spelled = self.shown().legend_position
        return VBAInt(next(number for number, name in LEGEND_POSITIONS.items() if name == spelled), 'Long')

    @setter('Position')
    def _position(self, value: object) -> None:
        wanted = LEGEND_POSITIONS.get(int(to_integer(value, 'Long')))
        if wanted is None:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        self.change().legend_position = wanted

    @member
    def IncludeInLayout(self) -> object:
        return not self.shown().legend_overlay

    @setter('IncludeInLayout')
    def _include_in_layout(self, value: object) -> None:
        wanted = to_bool(value)
        self.change().legend_overlay = not wanted

    @member
    def Shadow(self) -> object:
        return bool(self.shown().legend_shadow)

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        wanted = to_bool(value)
        self.change().legend_shadow = wanted

    @method
    def Delete(self) -> object:
        self.shown()
        self.chart.vba_set('HasLegend', False)
        return EMPTY

    @method
    def Clear(self) -> object:
        """Takes the legend away, as Delete does (tests/fixtures/chart_legend_entries.json)."""
        return self.Delete()

    @method
    def LegendEntries(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        self.shown()
        entries = LegendEntries(self)
        return entries if Index is MISSING else entries.vba_lookup(Index, entries.vba_items())


class LegendEntries(VBACollection, Created):
    """The legend's entries still shown: one per series, or in a pie or doughnut one per point
    (tests/fixtures/chart_legend_entries.json)."""

    vba_type_name = 'LegendEntries'

    def __init__(self, legend: Legend) -> None:
        self.legend = legend

    def vba_items(self) -> list[object]:
        state = self.legend.shown()
        if CHART_TYPES.get(shown_type(state)) in ('pie', 'doughnut'):
            first = state.series[0] if state.series else None
            total = len(Series(self.legend.chart, first).values()) if first is not None else 0
        else:
            total = len(state.series)
        return [LegendEntry(self.legend, index) for index in range(total) if index not in state.legend_deleted]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_number(index))
        if not 1 <= position <= len(items):
            raise error(-2147467259, 'Automation error\nUnspecified error')
        return items[position - 1]


class LegendEntry(Created):
    vba_type_name = 'LegendEntry'

    def __init__(self, legend: Legend, index: int) -> None:
        self.legend, self.index = legend, index

    @member
    def Parent(self) -> object:
        return self.legend

    @method
    def Delete(self) -> object:
        """The entry gone from the legend, answering nothing; deleting the last one takes the legend away and raises
        1004."""
        state = self.legend.change()
        if len(LegendEntries(self.legend).vba_items()) <= 1:
            self.legend.chart.vba_set('HasLegend', False)
            raise error(1004, 'Delete method of LegendEntry class failed')
        state.legend_deleted.append(self.index)
        return EMPTY


class ChartGroups(VBACollection, Created):
    """A chart's groups: the one group a chart of a single kind has (tests/fixtures/chart_groups.json)."""

    vba_type_name = 'ChartGroups'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def vba_items(self) -> list[object]:
        self.chart.admit_read()
        return [ChartGroup(self.chart)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if int(to_number(index)) != 1:
            raise error(1004, 'Unable to get the ChartGroups property of the Chart class')
        return items[0]


class ChartGroup(Created):
    """The settings a chart's series share as a group (tests/fixtures/chart_groups.json). A kind that draws one of
    them not keeps what is set and writes nothing."""

    vba_type_name = 'ChartGroup'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def _kind(self) -> str:
        self.chart.admit_read()
        return CHART_TYPES.get(shown_type(self.chart.state), '')

    def _set(self, field_name: str, value: object, lowest: int, highest: int) -> None:
        wanted = int(to_integer(value, 'Long'))
        self.chart.admit_change()
        if not lowest <= wanted <= highest:
            raise error(1004, 'Unable to set the property of the ChartGroup class')
        self.chart.changed()
        setattr(self.chart.state, field_name, wanted)

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def GapWidth(self) -> object:
        self._kind()
        return VBAInt(self.chart.state.gap_width, 'Long')

    @setter('GapWidth')
    def _gap_width(self, value: object) -> None:
        self._set('gap_width', value, 0, 500)

    @member
    def Overlap(self) -> object:
        # Lines, areas and scatter charts answer 100, the series drawn over each other.
        return VBAInt(100 if self._kind() in ('line', 'lineMarkers', 'area', 'scatter') else self.chart.state.overlap,
                      'Long')

    @setter('Overlap')
    def _overlap(self, value: object) -> None:
        if self._kind() not in ('column', 'bar'):
            raise VBAUnsupportedError('Overlap on a chart without bars is not implemented')
        self._set('overlap', value, -100, 100)

    @member
    def VaryByCategories(self) -> object:
        vary = self.chart.state.vary
        return self._kind() in ('pie', 'doughnut') if vary is None else vary

    @setter('VaryByCategories')
    def _vary_by_categories(self, value: object) -> None:
        wanted = to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        self.chart.state.vary = wanted

    @member
    def FirstSliceAngle(self) -> object:
        self._kind()
        return VBAInt(self.chart.state.first_slice, 'Long')

    @setter('FirstSliceAngle')
    def _first_slice_angle(self, value: object) -> None:
        self._set('first_slice', value, 0, 360)

    @member
    def DoughnutHoleSize(self) -> object:
        hole = self.chart.state.hole
        return VBAInt((50 if self._kind() == 'doughnut' else 0) if hole is None else hole, 'Long')

    @setter('DoughnutHoleSize')
    def _doughnut_hole_size(self, value: object) -> None:
        self._set('hole', value, 10, 90)


class SeriesCollection(VBACollection, ItemDefault, Created):
    vba_type_name = 'SeriesCollection'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def vba_items(self) -> list[object]:
        self.chart.admit_read()
        return [Series(self.chart, entry) for entry in self.chart.state.series]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if 1 <= position <= len(items):
            return items[position - 1]
        # Excel answers a series it does not have with 1004, not a subscript error.
        raise error(1004, 'Invalid Parameter')

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
        entry = SeriesState(series_id=series_id, unique=self.chart.state.next_unique,
                            marker=default_marker(self.chart.state))
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
            if len(values) > 1:
                # A name of several cells reads their texts with a space between (tests/fixtures/
                # chart_multilevel_categories.json).
                return ' '.join(text for row in shown_cells(self.chart.sheet, self.entry.name_reference)
                                for text in row if text)
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

    @setter('Values', objects=True)
    def _values(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        reference, values = series_input(value, sheet=self.chart.sheet)
        self.position()
        self.chart.changed()
        self.entry.value_reference, self.entry.values = reference, values

    @member
    def XValues(self) -> object:
        values = self.categories()
        if self.entry.category_reference and category_levels(self.chart.sheet, self.entry) is None:
            # Categories with text among them read every cell as it shows, a blank as ""; numbers with a blank among
            # them cannot be read at all (tests/fixtures/chart_multilevel_categories.json).
            if any(isinstance(value, str) for value in values):
                values = [text for row in shown_cells(self.chart.sheet, self.entry.category_reference)
                          for text in row]
            elif any(value is EMPTY for value in values):
                raise error(-2147467259, 'Automation error\nUnspecified error ')
        values = [float(value) if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) else value
                  for value in values]
        return VBAArray([(1, len(values))], items=values)

    @setter('XValues', objects=True)
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
            levels = category_levels(self.chart.sheet, self.entry)
            if levels is not None:
                return list(joined_categories(levels))
            return referenced_values(self.chart.sheet, self.entry.category_reference)
        return self.entry.categories or [float(index + 1) for index in range(len(self.values()))]

    @member
    def Formula(self) -> object:
        order = self.position()
        name = self.entry.name_reference or (quoted(self.entry.name) if self.entry.name is not None else '')
        categories = self.entry.category_reference or literal_array(self.entry.categories)
        values = self.entry.value_reference or literal_array(self.entry.values)
        return f'=SERIES({name},{categories},{values},{order})'

    @member
    def FormulaR1C1(self) -> object:
        return r1c1_series(str(self.Formula()))

    @member
    def FormulaLocal(self) -> object:
        return self.Formula()

    @member
    def FormulaR1C1Local(self) -> object:
        return self.FormulaR1C1()

    def lines(self) -> SeriesState:
        """The series, where its chart draws lines or markers; Excel raises 1004 for the other kinds."""
        self.position()
        if CHART_TYPES.get(shown_type(self.chart.state)) not in ('line', 'lineMarkers', 'scatter'):
            raise error(1004, 'Unable to get the property of the Series class')
        if self.entry.marker == MARKER_UNREAD:
            raise VBAUnsupportedError("Reading this imported series' marker is not implemented")
        return self.entry

    # -- tests/fixtures/chart_series_more.json

    @member
    def InvertIfNegative(self) -> object:
        self.position()
        return self.entry.invert_if_negative

    @setter('InvertIfNegative')
    def _invert_if_negative(self, value: object) -> None:
        self.position()
        wanted = to_bool(value)
        # A chart of another kind takes the setting without a word and writes nothing.
        if CHART_TYPES.get(shown_type(self.chart.state)) in ('column', 'bar'):
            self.chart.changed()
            self.entry.invert_if_negative = wanted

    @member
    def HasDataLabels(self) -> object:
        self.position()
        return self.entry.data_labels

    @setter('HasDataLabels')
    def _has_data_labels(self, value: object) -> None:
        self.position()
        wanted = to_bool(value)
        self.chart.changed()
        self.entry.data_labels = wanted
        # Excel turns the series' leader lines off whichever way the labels go.
        self.entry.leader_lines = False

    @member
    def HasLeaderLines(self) -> object:
        self.position()
        return self.entry.leader_lines

    @setter('HasLeaderLines')
    def _has_leader_lines(self, value: object) -> None:
        self.position()
        wanted = to_bool(value)
        self.chart.changed()
        self.entry.leader_lines = wanted

    @member
    def Explosion(self) -> object:
        self.position()
        return VBAInt(self.entry.explosion, 'Long')

    @setter('Explosion')
    def _explosion(self, value: object) -> None:
        self.position()
        wanted = int(to_integer(value, 'Long'))
        # Only a pie takes it, a doughnut no more than a column chart.
        if CHART_TYPES.get(shown_type(self.chart.state)) != 'pie' or wanted < 0:
            raise error(1004, 'Unable to set the Explosion property of the Series class')
        self.chart.changed()
        self.entry.explosion = wanted

    @member
    def Smooth(self) -> object:
        return self.lines().smooth

    @setter('Smooth')
    def _smooth(self, value: object) -> None:
        entry = self.lines()
        wanted = to_bool(value)
        self.chart.changed()
        entry.smooth = wanted

    @member
    def MarkerStyle(self) -> object:
        return VBAInt(self.lines().marker, 'Long')

    @setter('MarkerStyle')
    def _marker_style(self, value: object) -> None:
        entry = self.lines()
        wanted = int(to_integer(value, 'Long'))
        wanted = MARKER_ALIASES.get(wanted, wanted)
        if wanted not in MARKER_SYMBOLS:
            raise error(1004, 'Unable to set the MarkerStyle property of the Series class')
        self.chart.changed()
        entry.marker = wanted

    # -- the marker's colours (tests/fixtures/chart_marker_colours.json)

    def _colour(self, field_name: str) -> str | None:
        found: str | None = getattr(self.lines(), field_name)
        if found == '?':
            raise VBAUnsupportedError("Reading this imported marker's colour is not implemented")
        return found

    def _recoloured(self, field_name: str, wanted: str | None) -> None:
        entry = self.lines()
        self.chart.changed()
        setattr(entry, field_name, wanted)
        if wanted is not None and entry.marker == MARKER_NONE:
            # A series drawn without markers takes automatic ones once given a marker colour.
            entry.marker = MARKER_AUTOMATIC

    def _read_colour(self, field_name: str) -> object:
        found = self._colour(field_name)
        if found is None or found == 'none':
            return VBAInt(-1 if found is None else -2, 'Long')
        return VBAInt(int(found[4:6] + found[2:4] + found[0:2], 16), 'Long')

    def _set_colour(self, field_name: str, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted == -1:
            self._recoloured(field_name, None)
            return
        if wanted < 0:
            raise VBAUnsupportedError(f'Marker colour {wanted} is not implemented')
        # Excel keeps the low 24 bits: 16777216 is black.
        blue, green, red = (wanted >> 16) & 255, (wanted >> 8) & 255, wanted & 255
        self._recoloured(field_name, f'{red:02X}{green:02X}{blue:02X}')

    def _read_index(self, field_name: str) -> object:
        found = self._colour(field_name)
        if found is None or found == 'none':
            # An automatic marker colour answers slot 2.
            return VBAInt(2 if found is None else -4142, 'Long')
        return VBAInt(self.chart.sheet.book.stylesheet.colors.nearest(found), 'Long')

    def _set_index(self, field_name: str, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted in (-4105, -4142):
            self._recoloured(field_name, None if wanted == -4105 else 'none')
            return
        palette = self.chart.sheet.book.stylesheet.colors.indexed
        if wanted == 57:
            # One past the 56 slots, Excel writes black: the system foreground.
            self._recoloured(field_name, '000000')
            return
        if not 1 <= wanted <= 56 or wanted + 7 >= len(palette):
            raise VBAUnsupportedError(f'Marker colour index {wanted} is not implemented')
        self._recoloured(field_name, palette[wanted + 7])

    @member
    def MarkerForegroundColor(self) -> object:
        return self._read_colour('marker_fore')

    @setter('MarkerForegroundColor')
    def _marker_foreground_color(self, value: object) -> None:
        self._set_colour('marker_fore', value)

    @member
    def MarkerBackgroundColor(self) -> object:
        return self._read_colour('marker_back')

    @setter('MarkerBackgroundColor')
    def _marker_background_color(self, value: object) -> None:
        self._set_colour('marker_back', value)

    @member
    def MarkerForegroundColorIndex(self) -> object:
        return self._read_index('marker_fore')

    @setter('MarkerForegroundColorIndex')
    def _marker_foreground_color_index(self, value: object) -> None:
        self._set_index('marker_fore', value)

    @member
    def MarkerBackgroundColorIndex(self) -> object:
        return self._read_index('marker_back')

    @setter('MarkerBackgroundColorIndex')
    def _marker_background_color_index(self, value: object) -> None:
        self._set_index('marker_back', value)

    @member
    def MarkerSize(self) -> object:
        entry = self.lines()
        size = entry.marker_size or (9 if entry.marker == MARKER_AUTOMATIC else 5)
        return VBAInt(size, 'Long')

    @setter('MarkerSize')
    def _marker_size(self, value: object) -> None:
        entry = self.lines()
        wanted = int(to_integer(value, 'Long'))
        if not 2 <= wanted <= 72:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        identity = entry.series_id if entry.series_id >= 0 else self.position() - 1
        if entry.marker == MARKER_AUTOMATIC and identity >= len(AUTOMATIC_MARKERS):
            raise VBAUnsupportedError('The automatic marker of a series past the ninth is not implemented')
        self.chart.changed()
        entry.marker_size = wanted
        if entry.marker == MARKER_AUTOMATIC:
            entry.marker = AUTOMATIC_MARKERS[identity]


def r1c1_series(formula: str) -> str:
    """A SERIES formula with its absolute A1 references spelled R1C1, as FormulaR1C1 answers."""
    from pyopenvba._a1 import column_number

    def cell(found: re.Match[str]) -> str:
        if not (found.group(1) and found.group(3)):
            raise VBAUnsupportedError('Relative references in a series formula are not implemented')
        return f'R{found.group(4)}C{column_number(found.group(2))}'

    parts = re.split(r'("(?:[^"]|"")*"|\{[^}]*\})', formula)
    return ''.join(part if part.startswith(('"', '{')) else re.sub(
        r"(?<=[!:,(])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])", cell, part) for part in parts)


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
