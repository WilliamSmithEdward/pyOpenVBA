"""Embedded charts share their worksheet's drawing and cell state."""
from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from decimal import Decimal
import math
import re
from typing import TYPE_CHECKING, Any, cast

from pyopenvba._a1 import Area, parse_reference
from pyopenvba.apps.excel import _events
from pyopenvba.apps.excel._chart_axes import AxisState
from pyopenvba.apps.excel._chart_error_bars import ErrorBars, ErrorBarState, error_bar, has_error_bars, set_has_error_bars
from pyopenvba.apps.excel._chart_trendlines import Trendlines, TrendlineState
from pyopenvba.apps.excel._chart_title import TitleLook, TitleMembers
from pyopenvba.apps.excel._chart_xml import LabelLook, TextLook
from pyopenvba.apps.excel._shapes import PLACEMENTS, ShapeObject, ShapeRange, Shapes
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.apps.excel._chart_borders import INVALID_CALL, NOT_A_PROPERTY, UNSPECIFIED, ElementProperties, PictureMembers
from pyopenvba._xml import escape_text
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, NULL, VBAArray, VBADate, VBAErrorValue, VBAInt, VBASingle, error, to_bool, to_integer, to_number, to_text
from pyopenvba.shapes._values import FillLine, Shape, emu, points

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

# Engine admission bounds, not native Excel limits.
MAX_CHART_POINTS = 1_000_000
MAX_CHART_SERIES = 1024
MAX_CHART_REFERENCE_TEXT = 32_768

CHART_TYPES = {51: 'column', 57: 'bar', 4: 'line', 65: 'lineMarkers', 5: 'pie', -4120: 'doughnut',
               -4169: 'scatter', 1: 'area', 52: 'column', 53: 'column', 58: 'bar', 59: 'bar', 63: 'line', 64: 'line',
               66: 'lineMarkers', 67: 'lineMarkers', 76: 'area', 77: 'area'}
#: The stacked types, and how the file groups their series (tests/fixtures/chart_stacked.json).
STACKED = {52: 'stacked', 53: 'percentStacked', 58: 'stacked', 59: 'percentStacked', 63: 'stacked',
           64: 'percentStacked', 66: 'stacked', 67: 'percentStacked', 76: 'stacked', 77: 'percentStacked'}
#: Series.Type for each kind: the gallery constants xlArea, xlBar, xlColumn, xlLine, xlPie, xlDoughnut and xlXYScatter
#: (tests/fixtures/chart_series_reads.json).
SERIES_KINDS = {'column': 3, 'bar': 2, 'line': 4, 'lineMarkers': 4, 'pie': 5, 'doughnut': -4120, 'scatter': -4169,
                'area': 1}

@dataclass
class PointLabel:
    """What was set of one point's label, each left as it is while its series' labels give it."""

    shows: dict[str, bool] = field(default_factory=lambda: {})
    show_range: bool = False
    text: str | None = None
    number_format: str | None = None
    linked: bool | None = None
    position: int | None = None
    separator: str | None = None
    orientation: int = -4128
    horizontal: int = -4108
    reading: int = -5002
    shadow: bool = False
    font: TextLook | None = None
    fill_rgb: str | None = None
    #: Whether anything has made its Characters read back.
    touched: bool = False

    def has_look(self) -> bool:
        return (self.number_format is not None or self.orientation != -4128 or self.horizontal != -4108
                or self.reading != -5002 or self.shadow or self.font is not None or self.fill_rgb is not None)

    @property
    def plain_format(self) -> bool:
        return not self.has_look()

    def empty(self) -> bool:
        return not (self.shows or self.show_range or self.text is not None
                    or self.position is not None or self.separator is not None or self.has_look())


@dataclass
class PointState:
    """What was set of one point, each None while its series gives it, and the unique ids of its dPt and dLbl."""
    #: Its place among the cells shown when it was made; moved from there it loses its unique ids.
    drawn_at: int = field(default=-1, compare=False)

    look: FillLine | None = None
    invert: bool | None = None
    explosion: int | None = None
    label: bool | None = None
    label_own: PointLabel = field(default_factory=lambda: PointLabel())
    #: Its marker, each None while its series gives it (tests/fixtures/point_markers.json).
    marker: int | None = None
    marker_size: int | None = None
    marker_fore: str | None = None
    marker_back: str | None = None
    #: The colour Format.Fill.Visible = msoFalse keeps in the file for a marker's hidden fill.
    marker_hidden: str | None = None
    #: A line's look set over a foreground colour keeps that colour's solid dash in the marker.
    marker_solid: bool = False
    #: What ClearFormats left the marker of the point's line, once it has taken the line itself.
    marker_line: FillLine | None = None
    unique: int = -1
    label_unique: int = -1

    def has_marker(self) -> bool:
        return not (self.marker is None and self.marker_size is None and self.marker_fore is None
                    and self.marker_back is None)


@dataclass
class SeriesState:
    name: str | None = None
    name_reference: str = ''
    categories: list[object] = field(default_factory=lambda: [])
    values: list[object] = field(default_factory=lambda: [])
    category_reference: str = ''
    value_reference: str = ''
    #: On a scatter chart, whether ClearFormats has given it the line it otherwise goes without
    #: (tests/fixtures/chart_borders.json).
    scatter_line: bool = False
    #: Whether categories whose cells were taken away held numbers, which the file's empty reference still says.
    lost_numbers: bool = False
    series_id: int = -1
    #: The first group of the series' unique id; -1 to take its series_id.
    unique: int = -1
    smooth: bool = False
    marker: int = -4142
    #: MarkerSize once given; until then 9 for an automatic marker and 5 for the rest.
    marker_size: int | None = None
    invert_if_negative: bool = False
    #: Shadow: None until set, which a shadow turned off differs from in the file.
    shadow: bool | None = None
    explosion: int = 0
    #: HasDataLabels, and HasLeaderLines, which turning the labels on turns off (tests/fixtures/chart_data_labels.json).
    data_labels: bool = False
    leader_lines: bool = True
    #: What the labels show (tests/fixtures/data_labels_object.json), and a count of the times they went, which
    #: leaves a DataLabels taken before stale.
    label_look: LabelLook = field(default_factory=lambda: LabelLook())
    label_generation: int = 0
    #: What Series.Format set, None until then (tests/fixtures/series_format.json).
    format_look: FillLine | None = None
    #: What was set of its points, by their index from 0 (tests/fixtures/series_points.json).
    points: dict[int, PointState] = field(default_factory=lambda: {})
    #: MarkerForegroundColor and MarkerBackgroundColor as RRGGBB, "none", or None while automatic; "?" for an
    #: imported marker filled in a way the model does not read (tests/fixtures/chart_marker_colours.json).
    marker_fore: str | None = None
    marker_back: str | None = None
    #: The colour its line gave its marker, "" for a line of no colour; None when it gave none.
    marker_line: str | None = None
    #: Its Y error bar and, in a scatter chart, its X one (tests/fixtures/chart_error_bars.json).
    error_bars: list[ErrorBarState] = field(default_factory=lambda: [])
    #: Its trendlines, in the order they were added (tests/fixtures/chart_trendlines.json).
    trendlines: list[TrendlineState] = field(default_factory=lambda: [])

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
    #: What ApplyLayout gave a chart without bars to read back.
    layout_overlap: int | None = None
    #: The chart's own c:dLbls and its value axis' c:spPr, as a layout writes them.
    layout_chart_labels: str = ''
    layout_value_look: str = ''
    layout_category_look: str = ''
    vary: bool | None = None
    first_slice: int = 0
    hole: int | None = None
    #: HasDropLines, HasHiLoLines, HasUpDownBars and HasSeriesLines, which only some kinds draw
    #: (tests/fixtures/chart_group_lines.json).
    #: What a group has for a bubble or pie-of-pie chart, which the model does not draw.
    bubble_scale: int = 100
    size_represents: int = 1
    negative_bubbles: bool = False
    shading: bool = True
    split_type: int = 1
    split_value: int | None = 1
    second_plot_size: int = 75
    drop_lines: bool = False
    hi_lo_lines: bool = False
    up_down_bars: bool = False
    series_lines: bool = False
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
    #: ChartStyle, from 1 to 48, or from 201 for one of AddChart2's styles.
    chart_style: int = 2
    #: Whether AddChart2 made it a scatter chart, which differs from another chart turned into one.
    made_scatter: bool = False
    #: The data labels its style gives its series: their value, category, name and percentage, and their place.
    style_labels: list[Any] | None = None
    #: Every series in order, while any is filtered out of ``series``, which keeps those drawn.
    full_series: list[SeriesState] = field(default_factory=lambda: [])
    #: Those whose cells are all hidden, who leave SeriesCollection and FullSeriesCollection until they are shown
    #: again and stay in the file (tests/fixtures/chart_hidden_cells.json); kept in full_series meanwhile.
    out_of_sight: list[SeriesState] = field(default_factory=lambda: [])
    #: ProtectData, ProtectFormatting and ProtectSelection once set, and what PageSetup was told.
    protection: dict[str, bool] = field(default_factory=lambda: {})
    page: dict[str, object] = field(default_factory=lambda: {})
    data_table: bool = False
    #: A data table's HasBorderHorizontal, HasBorderVertical and HasBorderOutline.
    data_table_borders: tuple[bool, bool, bool] = (True, True, True)
    #: Whether its data table shows legend keys, as it does until SetElement says otherwise.
    data_table_keys: bool = True
    #: What SetElement last said of the chart's labels: whether they show, and where. The chart's own c:dLbls holds
    #: it.
    element_labels: tuple[bool, int | None] | None = None
    #: The chart object's RoundedCorners, and its Shadow: None until set, which a shadow turned off differs from in
    #: the file (tests/fixtures/chart_object_reads.json).
    rounded: bool = False
    shadow: bool | None = None
    #: What Format set on the chart's other elements, by the element: area, plot, legend, title, axis and its kind,
    #: grid and its axis's kind, and title and its axis's kind (tests/fixtures/chart_element_format.json).
    element_looks: dict[str, FillLine] = field(default_factory=lambda: {})
    #: What Font set on the chart's text, by the element: area, legend, title, title and its axis's kind, and tick and
    #: its axis's kind (tests/fixtures/chart_fonts.json).
    element_fonts: dict[str, TextLook] = field(default_factory=lambda: {})

    def has_automatic_title(self) -> bool:
        """Whether it has an automatic title to give: it draws one series with a name, or is a pie whose first has one
        (tests/fixtures/chart_set_element.json, chart_auto_titles.json)."""
        first = self.series[0] if self.series else None
        if first is None or not (first.name_reference or first.name) or self.names_lost():
            return False
        return len(self.series) == 1 or CHART_TYPES.get(shown_type(self)) == 'pie'

    def names_lost(self) -> bool:
        """Whether every series' name has lost its cell, which leaves them all named by number."""
        every = self.full_series or self.series
        return bool(every) and all(lost(entry.name_reference) for entry in every)


def label_block(sheet: Worksheet, area: Area, at: Squeezed | None = None) -> tuple[int, int]:
    """How many of a block's top rows and left columns label the rest, as SetSourceData reads it
    (tests/fixtures/chart_multilevel_categories.json).

    A blank top left corner says so itself: its blank cells along the top row and down the left column, one where
    they run the whole way. Otherwise the left columns run to the last text or date in the bottom row, and the top
    rows to the last in the right column, whatever the cells before them hold. Either way at least one row and one
    column are left for the values."""
    from pyopenvba.apps.excel._model import Range

    height, width = area.bottom - area.top + 1, area.right - area.left + 1

    def value(row: int, column: int) -> object:
        if at is not None:
            row, column = at.rows[row - 1], at.columns[column - 1]
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


@dataclass
class Squeezed:
    """Several areas read as one block with the rows and columns between them taken out: a block from row and column 1
    stands for these rows and columns (tests/fixtures/range_tracking.json)."""

    rows: list[int]
    columns: list[int]

    def block(self) -> Area:
        return Area(1, 1, len(self.rows), len(self.columns))

    def real(self, area: Area) -> Area:
        return Area(self.rows[area.top - 1], self.columns[area.left - 1], self.rows[area.bottom - 1],
                    self.columns[area.right - 1])


def squeezed(sheet: Worksheet, areas: list[Area]) -> Squeezed:
    """Areas of the same rows, or of the same columns, as one block. Excel's series from any others follow no rule found."""
    used = sheet.used_bounds() or (1, 1, 1, 1)
    first = areas[0]
    bottom, right = min(first.bottom, used[2]), min(first.right, used[3])
    if all((area.top, area.bottom) == (first.top, first.bottom) for area in areas):
        columns = sorted(column for area in areas for column in range(area.left, min(area.right, used[3]) + 1))
        made = Squeezed(list(range(first.top, bottom + 1)), columns)
        alike = len(set(columns)) == len(columns)
    elif all((area.left, area.right) == (first.left, first.right) for area in areas):
        rows = sorted(row for area in areas for row in range(area.top, min(area.bottom, used[2]) + 1))
        made = Squeezed(rows, list(range(first.left, right + 1)))
        alike = len(set(rows)) == len(rows)
    else:
        alike, made = False, Squeezed([], [])
    if not alike:
        raise VBAUnsupportedError('Chart source data of areas that differ in their rows or columns, or overlap, is not '
                                  'implemented')
    return made


def plan_series(sheet: Worksheet, area: Area, orientation: int, at: Squeezed | None = None) -> list[SeriesState]:
    """The series SetSourceData makes from a block of cells, by rows (1) or columns (2). The rows and columns
    label_block finds name the series and label the categories, more than one of either making a name or categories
    of several levels (tests/fixtures/chart_multilevel_categories.json)."""
    rows, columns = label_block(sheet, area, at)

    def named(block: Area) -> str:
        return reference(sheet, at.real(block) if at is not None else block)

    series_plan: list[SeriesState] = []
    if orientation == 2:
        top, left = area.top + rows, area.left + columns
        if area.right - left + 1 > MAX_CHART_SERIES:
            raise VBAUnsupportedError('Chart source data exceeds the supported series budget')
        for column in range(left, area.right + 1):
            series = SeriesState(series_id=len(series_plan))
            if rows:
                series.name_reference = named(Area(area.top, column, top - 1, column))
            series.value_reference = named(Area(top, column, area.bottom, column))
            if columns:
                series.category_reference = named(Area(top, area.left, area.bottom, left - 1))
            series_plan.append(series)
    else:
        top, left = area.top + rows, area.left + columns
        if area.bottom - top + 1 > MAX_CHART_SERIES:
            raise VBAUnsupportedError('Chart source data exceeds the supported series budget')
        for row in range(top, area.bottom + 1):
            series = SeriesState(series_id=len(series_plan))
            if columns:
                series.name_reference = named(Area(row, area.left, row, left - 1))
            series.value_reference = named(Area(row, left, row, area.right))
            if rows:
                series.category_reference = named(Area(area.top, left, top - 1, area.right))
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


def lost(reference: str) -> bool:
    """Whether the cells a chart's reference read have been taken away (tests/fixtures/chart_lost_cells.json)."""
    return reference.endswith('#REF!')


def any_shown(sheet: Worksheet, formula: str) -> bool:
    """Whether a cell of a reference is not hidden; a reference the model cannot follow counts as shown."""
    try:
        areas = parse_reference(formula)
        for area in areas:
            owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
            if (any(not owner.dims.row_hidden(row) for row in range(area.top, area.bottom + 1))
                    and any(not owner.dims.column_hidden(column) for column in range(area.left, area.right + 1))):
                return True
    except (ValueError, KeyError, VBAUnsupportedError, VBARuntimeError):
        return True
    return False


def settle_hidden(sheet: Worksheet, state: ChartState) -> None:
    """The series out of sight as the cells stand now, with PlotVisibleOnly on."""
    every = state.full_series or state.series
    gone = [entry for entry in every if state.visible_only and entry.value_reference
            and not lost(entry.value_reference) and not any_shown(sheet, entry.value_reference)]
    if len(gone) == len(state.out_of_sight) and all(a is b for a, b in zip(gone, state.out_of_sight, strict=True)):
        return
    filtered = [entry for entry in every if not any(entry is drawn for drawn in state.series)
                and not any(entry is hidden for hidden in state.out_of_sight)]
    away = [*gone, *filtered]
    everyone = list(every)
    state.out_of_sight = gone
    state.series[:] = [entry for entry in everyone if not any(entry is other for other in away)]
    state.full_series = everyone if away else []


def categories_gone(sheet: Worksheet, state: ChartState, entry: SeriesState) -> bool:
    """Whether a series' categories have been taken away, or are all hidden with PlotVisibleOnly on."""
    return lost(entry.category_reference) or bool(
        state.visible_only and entry.category_reference and not any_shown(sheet, entry.category_reference))


def category_levels(sheet: Worksheet, entry: SeriesState) -> list[list[str]] | None:
    """A series' categories when they take several rows or columns, one list of texts for each level, the outer
    first; None for categories of one level. Its points run along the side as long as its values, down the rows
    where both sides are (tests/fixtures/chart_multilevel_categories.json)."""
    if not entry.category_reference or lost(entry.value_reference):
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
    if height == 1 or width == 1:
        return None
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
    if not isinstance(selection, Range) or len(selection.areas) != 1:
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


def guessed_orientation(sheet: Worksheet, area: Area, at: Squeezed | None = None) -> int:
    """How a block is plotted when nothing says: by rows (1) unless what its labels leave is taller than wide
    (tests/fixtures/chart_add_plotted.json)."""
    rows, columns = label_block(sheet, area, at)
    height, width = area.bottom - area.top + 1 - rows, area.right - area.left + 1 - columns
    return 1 if width >= height else 2


def shown_type(chart: ChartState) -> int:
    """The type a chart reads and is written as. A line chart follows its first series: with markers (65) unless that
    series has none, and with markers while it has no series at all."""
    for plain, marked in LINE_PAIRS:
        if chart.chart_type in (plain, marked):
            return plain if chart.series and chart.series[0].marker == MARKER_NONE else marked
    return chart.chart_type


#: Each line chart type and its type with markers: plain, stacked and 100% stacked (tests/fixtures/chart_stacked.json).
LINE_PAIRS = ((4, 65), (63, 66), (64, 67))
#: The types drawn with markers.
MARKED_TYPES = (65, 66, 67, -4169)
_DRAWS_MARKERS = ('line', 'lineMarkers', 'scatter')


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
    return MARKER_AUTOMATIC if chart.chart_type in MARKED_TYPES else MARKER_NONE


def marker_markup(entry: SeriesState) -> str:
    """The series' marker element: none for an automatic marker, a symbol for the rest, with its size where it has
    one."""
    symbol = MARKER_SYMBOLS[entry.marker]
    # A colour set on the series' line reaches its marker (tests/fixtures/series_format.json).
    stroke = {None: '', '': '<a:ln/>'}.get(
        entry.marker_line, f'<a:ln><a:solidFill><a:srgbClr val="{entry.marker_line}"/></a:solidFill></a:ln>')
    look = marker_look(entry, stroke)
    if not symbol:
        return f'<c:marker>{look}</c:marker>' if look else ''
    if symbol == 'none' and entry.marker_back is None and entry.marker_line is None:
        # A foreground colour alone stays for the session and leaves no mark on a series with no marker
        # (tests/fixtures/chart_group_looks.json).
        return '<c:marker><c:symbol val="none"/></c:marker>'
    if symbol in ('none', 'picture'):
        return f'<c:marker><c:symbol val="{symbol}"/>{look}</c:marker>'
    return f'<c:marker><c:symbol val="{symbol}"/><c:size val="{entry.marker_size or 5}"/>{look}</c:marker>'


def marker_look(entry: SeriesState | PointState, stroke: str = '', tail: str = '') -> str:
    """The marker's spPr: its background as the fill, its foreground as the line, as Excel writes them
    (tests/fixtures/chart_marker_colours.json), or the stroke given where it has no foreground."""
    fill = {None: '', 'none': '<a:noFill/>'}.get(
        entry.marker_back, f'<a:solidFill><a:srgbClr val="{entry.marker_back}"/></a:solidFill>')
    line = {None: '', 'none': '<a:ln w="12700"><a:noFill/></a:ln>'}.get(
        entry.marker_fore,
        f'<a:ln><a:solidFill><a:srgbClr val="{entry.marker_fore}"/></a:solidFill><a:prstDash val="solid"/></a:ln>')
    line = line or stroke
    return f'<c:spPr>{fill}{line}{tail}</c:spPr>' if fill or line else ''


def plot_by(chart: ChartState) -> int:
    """PlotBy as Excel works it out from the series: xlColumns (2) when their values are columns of the same rows from
    left to right, each name the cell above its own column and every series the same categories; xlRows (1) the same
    turned round; 0 when neither holds, as for a literal or a series moved out of order. A chart with no series reads
    xlRows, and series that fit both read the way their source data was given."""
    if not chart.series and not chart.out_of_sight:
        return 1
    # A series filtered out counts, though it is not drawn.
    every = chart.full_series or chart.series
    if any(lost(text) for entry in every
           for text in (entry.name_reference, entry.category_reference, entry.value_reference)):
        return 0
    try:
        values = [parse_reference(entry.value_reference) for entry in every]
        names = [parse_reference(entry.name_reference) if entry.name_reference else None for entry in every]
    except ValueError:
        return 0
    literal = any(not entry.value_reference or entry.name is not None and not entry.name_reference
                  or entry.categories and not entry.category_reference for entry in every)
    if (literal or any(len(areas) != 1 for areas in values) or any(name is not None and len(name) != 1 for name in names)
            or len({entry.category_reference for entry in every}) != 1
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
        category = every[0].category_reference
        if category and span(parse_reference(category)[0]) != span(blocks[0]):
            return False
        places = [place(area) for area in blocks]
        if any(later <= earlier for earlier, later in zip(places, places[1:])):
            return False
        for cell, area in zip(cells, blocks):
            if cell is not None and (cell.sheet != area.sheet
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


def add_chart(sheet: Worksheet, style: object, kind: object, place: tuple[object, object, object, object],
              new_layout: object) -> Shape:
    """Shapes.AddChart2 and AddChart: a new chart holding the data round its sheet's selection, laid out as a chart of the
    type and style starts (tests/fixtures/add_chart2.json)."""
    from pyopenvba.apps.excel._chart_style_table import (DEFAULT_STYLES, FIRST_STYLE, LAST_STYLE, PROVEN, REFUSED_TYPES,
                                                         extras, layout)
    from pyopenvba.apps.excel._protection import drawing_enforced

    if drawing_enforced(sheet):
        raise error(1004, 'The specified value is out of range.')
    wanted = _whole(kind, 51)
    if wanted not in DEFAULT_STYLES and wanted not in REFUSED_TYPES:
        wanted = 51
    given = _whole(style, -1)
    chosen = DEFAULT_STYLES.get(wanted, 0) if given == -1 else given
    if wanted in REFUSED_TYPES or not (1 <= chosen <= 48 or FIRST_STYLE <= chosen <= LAST_STYLE):
        raise error(1004, 'Method \'AddChart2\' of object \'Shapes\' failed')
    if wanted not in CHART_TYPES:
        raise VBAUnsupportedError(f'Chart type {wanted} is not implemented')
    settled: list[str] = []
    if FIRST_STYLE <= chosen <= LAST_STYLE:
        if chosen not in PROVEN.get(wanted, ()):
            # The rest bring more than a style part says, data labels and axis titles among them, or were not measured.
            raise VBAUnsupportedError(f'Chart style {chosen} is not implemented')
        settled = layout(chosen, wanted)
        wanted = int(settled[0])
        if wanted not in CHART_TYPES:
            raise VBAUnsupportedError(f'Chart type {wanted} is not implemented')
    left, top, width, height = (float(to_number(value)) if value is not MISSING else None for value in place)
    shape = Shapes(sheet).new_chart_shape(left if left is not None and left >= 0 else 192.0,
                                          top if top is not None and top >= 0 else 103.125,
                                          360.0 if width is None or width < 0 else width,
                                          216.0 if height is None or height < 0 else height)
    base = plotted_before(sheet)
    fresh = ChartState(unique_base=base, next_unique=base, chart_type=wanted, chart_style=chosen)
    sheet.chart_states[shape.shape_id] = fresh
    chart = Chart(sheet, shape)
    source = _picked_up(sheet)
    if settled:
        _set_out(chart, settled)
        if source is None and fresh.style_labels:
            # A chart made with nothing to plot keeps its labels' parts and loses their place.
            fresh.style_labels = [*fresh.style_labels[:4], '']
    if source is not None:
        fresh.unique_base = fresh.next_unique = 0
        chart.SetSourceData(source)
        _labelled(fresh)
        if CHART_TYPES[fresh.chart_type] == 'scatter':
            _across_first(fresh)
    new = new_layout is MISSING or to_bool(new_layout)
    if new:
        # A new layout has a title, and a legend only where there is something for one to tell apart.
        fresh.title = 'Chart Title'
        fresh.title_typed = False
        fresh.legend = len(fresh.series) > 1 or CHART_TYPES[fresh.chart_type] in ('pie', 'doughnut')
    if settled:
        said = extras(chosen, CHART_TYPES[fresh.chart_type])
        if new and not said[0]:
            fresh.title, fresh.title_deleted = None, True
        if new and not said[1]:
            fresh.legend = False
        if fresh.legend:
            fresh.legend_position = said[1] or 'b'
    _events.new_chart(sheet.book, ChartObject(sheet, shape).Chart())
    return shape


def _labelled(chart: ChartState) -> None:
    """Series take the data labels a style has: what they show, and where (tests/fixtures/chart_style_parts.json)."""
    from pyopenvba.apps.excel._chart_xml import LABEL_POSITIONS

    if chart.style_labels is None:
        return
    value, category, name, percent, place = chart.style_labels
    position = next((number for number, spelled in LABEL_POSITIONS.items() if spelled == place), None)
    for entry in chart.series:
        if entry.data_labels:
            continue
        entry.data_labels = True
        entry.label_look = LabelLook(value=value, category=category, series=name, percent=percent, position=position)


def _across_first(chart: ChartState) -> None:
    """A new scatter chart's first column of numbers, with no labels beside it, runs across
    (tests/fixtures/add_chart2.json)."""
    if len(chart.series) > 1 and not chart.series[0].category_reference:
        across, *rest = chart.series
        for index, entry in enumerate(rest):
            entry.category_reference, entry.series_id, entry.unique = across.value_reference, index, index
        chart.series[:] = rest
        chart.next_unique = len(rest)


def _whole(value: object, otherwise: int) -> int:
    """A style or type as a number; what is not one, or is left out, is the default."""
    if value is MISSING:
        return otherwise
    try:
        return int(to_number(value))
    except VBARuntimeError:
        return otherwise


def _picked_up(sheet: Worksheet) -> Range | None:
    """The cells a new chart takes for its data: those selected, or a single cell's region, on the sheet it goes on."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._region import current_region

    chosen = sheet.selection
    if not isinstance(chosen, Range):
        return None
    if len(chosen.areas) != 1:
        return chosen
    area = chosen.first
    if (area.top, area.left) != (area.bottom, area.right):
        return chosen
    region = current_region(sheet, area.top, area.left)
    return None if trimmed(sheet, region) is None else Range(sheet, [region])


def _set_out(chart: Chart, settled: list[str]) -> None:
    """A chart as its style lays it out (tests/fixtures/chart_style_parts.json)."""
    from pyopenvba.apps.excel._chart_axes import TICK_MARKS, TitleState, states
    from pyopenvba.apps.excel._chart_style_table import DOUGHNUTS, extras

    _, gap, overlap, category_tick, value_tick, _, _, category_grid, value_grid = settled
    chart.state.gap_width, chart.state.overlap = int(gap), int(overlap)
    kind = CHART_TYPES[chart.state.chart_type]
    chart.state.made_scatter = kind == 'scatter'
    if kind == 'doughnut':
        chart.state.hole, chart.state.first_slice = DOUGHNUTS[chart.state.chart_style]
    if kind == 'area' or chart.state.chart_type in (63, 64, 66, 67):
        # What stacks draws nothing as nought.
        chart.state.blanks = 2
    axes = states(chart)
    for axis, tick, grid in ((axes[1], category_tick, category_grid), (axes[2], value_tick, value_grid)):
        if tick != 'E':
            axis.major_tick = TICK_MARKS[int(tick)]
            axis.major_gridlines = grid == 'True'
    said = extras(chart.state.chart_style, kind)
    if kind not in ('pie', 'doughnut'):
        for axis, (minor_grid, minor_tick, deleted) in zip((axes[1], axes[2]), (said[2:5], said[5:8]), strict=True):
            axis.minor_gridlines, axis.minor_tick, axis.present = minor_grid, minor_tick, not deleted
    labels, chart.state.drop_lines, chart.state.hi_lo_lines, category_title, value_title = said[8:13]
    if kind not in ('pie', 'doughnut'):
        for axis, titled in ((axes[1], category_title), (axes[2], value_title)):
            axis.title = TitleState() if titled else None
    # An area or a doughnut has no place to put them but its own.
    chart.state.style_labels = ([*labels[:4], ''] if kind in ('area', 'doughnut') else labels) if labels else None


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
        made = ChartObject(self.sheet, shape)
        _events.new_chart(self.sheet.book, made.Chart())
        return made

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def ShapeRange(self) -> object:
        """The sheet's charts as one range (tests/fixtures/chart_areas.json)."""
        return ShapeRange([ShapeObject(self.sheet, shape) for shape in self.sheet.shapes_ if shape.kind == 'chart'])

    @member
    def ProtectChartObject(self) -> object:
        return False

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

    @method
    def Duplicate(self) -> object:
        for item in self.vba_items():
            cast(ChartObject, item).Duplicate()
        return self

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        items = self.vba_items()
        if len(items) != 1:
            raise VBAUnsupportedError('Selecting several chart objects together is not implemented')
        # Alone, it is the chart that is selected, not its frame.
        cast(ChartObject, items[0]).made_active('area')
        return EMPTY


def active_chart(sheet: Worksheet) -> object:
    """The sheet's active embedded chart, Nothing once a range is selected or the chart is gone
    (tests/fixtures/active_chart.json)."""
    shape = sheet.active_chart
    if shape is None or not any(one is shape for one in sheet.shapes_):
        sheet.active_chart = None
        return NOTHING
    return Chart(sheet, shape)


def selection(sheet: Worksheet) -> object:
    """What the sheet has selected: while a chart is active, its frame, chart area or plot area, else its range
    (tests/fixtures/active_chart.json, tests/fixtures/chart_areas.json)."""
    chart = active_chart(sheet)
    if not isinstance(chart, Chart):
        from pyopenvba.apps.excel._selected_shapes import answer

        drawn = answer(sheet)
        return sheet.selection if drawn is None else drawn
    if sheet.chart_part == 'area':
        return ChartArea(chart)
    if sheet.chart_part == 'plot':
        return PlotArea(chart)
    if sheet.chart_part == 'element':
        # A selected element that has been taken away leaves the chart area selected.
        return sheet.chart_element if chart.element_there() else ChartArea(chart)
    return ChartObject(sheet, chart.shape)


def select_element(chart: Chart, element: object, plotted: bool) -> object:
    """A series' or point's Select (tests/fixtures/chart_element_select.json): nothing on a sheet that is not active.
    One the chart cannot select, as an area chart's point, fails but activates the chart if it was not, with its area
    selected."""
    chart.admit_read()
    sheet = chart.sheet
    book = sheet.book
    if book.application.active_book is book and book.active_sheet is not sheet:
        return True
    framed = ChartObject(sheet, chart.shape)
    if not plotted:
        if sheet.active_chart is not chart.shape:
            framed.made_active('area')
        raise error(-2147467259, 'Automation error\nUnspecified error')
    framed.made_active('element')
    sheet.chart_element = element
    return True


#: What a chart's PageSetup answers until told otherwise, in points where it is a margin; and each margin's name in the file.
_PAGE_DEFAULTS: dict[str, object] = {'Orientation': 1, 'Zoom': 100, 'PaperSize': 1, 'LeftMargin': 50.4, 'RightMargin': 50.4, 'TopMargin': 54.0,
                                     'BottomMargin': 54.0, 'HeaderMargin': 21.6, 'FooterMargin': 21.6, 'FirstPageNumber': -4105,
                                     'BlackAndWhite': False, 'Draft': False}
_PAGE_MARGINS = {'BottomMargin': 'b', 'LeftMargin': 'l', 'RightMargin': 'r', 'TopMargin': 't', 'HeaderMargin': 'header',
                 'FooterMargin': 'footer'}
_HEADERS = {'LeftHeader': ('oddHeader', '&L'), 'CenterHeader': ('oddHeader', '&C'), 'RightHeader': ('oddHeader', '&R'),
            'LeftFooter': ('oddFooter', '&L'), 'CenterFooter': ('oddFooter', '&C'), 'RightFooter': ('oddFooter', '&R')}


def page_markup(state: ChartState) -> str:
    """The chart's printSettings: an orientation once anything of the page was read, what was set, and otherwise the
    bare elements every chart writes (tests/fixtures/chart_protection.json)."""
    page = state.page
    margins = ' '.join(f'{name}="{format(float(cast(float, page[key])) / 72, ".17g") if key in page else default}"'
                       for key, name, default in (('BottomMargin', 'b', '0.75'), ('LeftMargin', 'l', '0.7'), ('RightMargin', 'r', '0.7'),
                                                  ('TopMargin', 't', '0.75'), ('HeaderMargin', 'header', '0.3'), ('FooterMargin', 'footer', '0.3')))
    parts = {tag: ''.join(escape_text(code + to_text(page[key])) for key, (where, code) in _HEADERS.items() if where == tag and key in page)
             for tag in ('oddHeader', 'oddFooter')}
    header = ''.join(f'<c:{tag}>{text}</c:{tag}>' for tag, text in parts.items() if text)
    if not page.get('read'):
        setup = '<c:pageSetup/>'
    else:
        attributes = (f' paperSize="{page["PaperSize"]}"' if 'PaperSize' in page else '') + (
            f' firstPageNumber="{page["FirstPageNumber"]}"' if 'FirstPageNumber' in page else '') + (
            f' orientation="{"landscape" if page.get("Orientation") == 2 else "portrait"}"') + (
            ' blackAndWhite="1"' if page.get('BlackAndWhite') else '') + (' draft="1"' if page.get('Draft') else '') + (
            ' useFirstPageNumber="1"' if 'FirstPageNumber' in page else '')
        setup = f'<c:pageSetup{attributes}/>'
    return (f'<c:printSettings>{f"<c:headerFooter>{header}</c:headerFooter>" if header else "<c:headerFooter/>"}'
            f'<c:pageMargins {margins}/>{setup}</c:printSettings>')


class ChartPageSetup(Created):
    """The PageSetup of a chart on a sheet, its own and not the sheet's. Reading or setting nearly any of it writes the
    orientation into the file (tests/fixtures/chart_protection.json)."""

    vba_type_name = 'PageSetup'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def _read(self, name: str) -> object:
        self.chart.admit_read()
        self.chart.state.page['read'] = True
        return self.chart.state.page.get(name, _PAGE_DEFAULTS.get(name, ''))

    def _write(self, name: str, value: object) -> None:
        self.chart.admit_change()
        self.chart.changed()
        self.chart.state.page['read'] = True
        self.chart.state.page[name] = value

    def _whole(self, name: str, value: object) -> None:
        self._write(name, int(to_integer(value, 'Long')))

    def _switch(self, name: str, value: object) -> None:
        self._write(name, to_bool(value))

    def _points(self, name: str, value: object) -> None:
        wanted = float(to_number(value))
        if wanted < 0:
            raise VBAUnsupportedError('A negative margin on a chart is not implemented')
        self._write(name, wanted)

    def _headed(self, name: str, value: object) -> None:
        where = _HEADERS[name][0]
        if any(other != name and tag == where and other in self.chart.state.page for other, (tag, _) in _HEADERS.items()):
            raise VBAUnsupportedError('A second part of a chart\'s header or footer is not implemented')
        self._write(name, to_text(value))

    def _no_setting(self) -> object:
        self.chart.admit_read()
        self.chart.state.page['read'] = True
        raise error(1004, 'Unable to get the property of the PageSetup class')

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Orientation(self) -> object:
        return VBAInt(cast(int, self._read('Orientation')), 'Long')

    @setter('Orientation')
    def _orientation(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in (1, 2):
            raise VBAUnsupportedError('This chart orientation is not implemented')
        self._write('Orientation', wanted)

    @member
    def ChartSize(self) -> object:
        self._read('ChartSize')
        return VBAInt(1, 'Long')

    @setter('ChartSize')
    def _chart_size(self, value: object) -> None:
        # It takes any size and goes on answering the same.
        to_integer(value, 'Long')
        self.chart.admit_change()
        self.chart.changed()
        self.chart.state.page['read'] = True

    @member
    def Zoom(self) -> object:
        return VBAInt(cast(int, self._read('Zoom')), 'Long')

    @setter('Zoom')
    def _zoom(self, value: object) -> None:
        if value is False:
            self._no_setting()
        wanted = int(to_integer(value, 'Long'))
        if not 10 <= wanted <= 400:
            raise VBAUnsupportedError('This chart zoom is not implemented')
        self._write('Zoom', wanted)

    @member
    def FitToPagesWide(self) -> object:
        return self._no_setting()

    @setter('FitToPagesWide')
    def _fit_to_pages_wide(self, value: object) -> None:
        del value
        self._no_setting()

    @member
    def FitToPagesTall(self) -> object:
        return self._no_setting()

    @setter('FitToPagesTall')
    def _fit_to_pages_tall(self, value: object) -> None:
        del value
        self._no_setting()

    @member
    def PaperSize(self) -> object:
        return VBAInt(cast(int, self._read('PaperSize')), 'Long')

    @setter('PaperSize')
    def _paper_size(self, value: object) -> None:
        self._whole('PaperSize', value)

    @member
    def LeftMargin(self) -> object:
        return self._read('LeftMargin')

    @setter('LeftMargin')
    def _left_margin(self, value: object) -> None:
        self._points('LeftMargin', value)

    @member
    def RightMargin(self) -> object:
        return self._read('RightMargin')

    @setter('RightMargin')
    def _right_margin(self, value: object) -> None:
        self._points('RightMargin', value)

    @member
    def TopMargin(self) -> object:
        return self._read('TopMargin')

    @setter('TopMargin')
    def _top_margin(self, value: object) -> None:
        self._points('TopMargin', value)

    @member
    def BottomMargin(self) -> object:
        return self._read('BottomMargin')

    @setter('BottomMargin')
    def _bottom_margin(self, value: object) -> None:
        self._points('BottomMargin', value)

    @member
    def HeaderMargin(self) -> object:
        return self._read('HeaderMargin')

    @setter('HeaderMargin')
    def _header_margin(self, value: object) -> None:
        self._points('HeaderMargin', value)

    @member
    def FooterMargin(self) -> object:
        return self._read('FooterMargin')

    @setter('FooterMargin')
    def _footer_margin(self, value: object) -> None:
        self._points('FooterMargin', value)

    @member
    def CenterHorizontally(self) -> object:
        self._read('CenterHorizontally')
        return False

    @setter('CenterHorizontally')
    def _center_horizontally(self, value: object) -> None:
        # A chart's page takes it and goes on answering False.
        to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        self.chart.state.page['read'] = True

    @member
    def CenterVertically(self) -> object:
        self._read('CenterVertically')
        return False

    @member
    def LeftHeader(self) -> object:
        return self._read('LeftHeader')

    @setter('LeftHeader')
    def _left_header(self, value: object) -> None:
        self._headed('LeftHeader', value)

    @member
    def CenterHeader(self) -> object:
        return self._read('CenterHeader')

    @setter('CenterHeader')
    def _center_header(self, value: object) -> None:
        self._headed('CenterHeader', value)

    @member
    def RightHeader(self) -> object:
        return self._read('RightHeader')

    @setter('RightHeader')
    def _right_header(self, value: object) -> None:
        self._headed('RightHeader', value)

    @member
    def LeftFooter(self) -> object:
        return self._read('LeftFooter')

    @setter('LeftFooter')
    def _left_footer(self, value: object) -> None:
        self._headed('LeftFooter', value)

    @member
    def CenterFooter(self) -> object:
        return self._read('CenterFooter')

    @setter('CenterFooter')
    def _center_footer(self, value: object) -> None:
        self._headed('CenterFooter', value)

    @member
    def RightFooter(self) -> object:
        return self._read('RightFooter')

    @setter('RightFooter')
    def _right_footer(self, value: object) -> None:
        self._headed('RightFooter', value)

    @member
    def FirstPageNumber(self) -> object:
        return VBAInt(cast(int, self._read('FirstPageNumber')), 'Long')

    @setter('FirstPageNumber')
    def _first_page_number(self, value: object) -> None:
        self._whole('FirstPageNumber', value)

    @member
    def BlackAndWhite(self) -> object:
        return self._read('BlackAndWhite')

    @setter('BlackAndWhite')
    def _black_and_white(self, value: object) -> None:
        self._switch('BlackAndWhite', value)

    @member
    def Draft(self) -> object:
        return self._read('Draft')

    @setter('Draft')
    def _draft(self, value: object) -> None:
        self._switch('Draft', value)

    @member
    def PrintArea(self) -> object:
        self.chart.admit_read()
        raise error(1004, 'Unable to get the PrintArea property of the PageSetup class')

    @setter('PrintArea')
    def _print_area(self, value: object) -> None:
        del value
        self.chart.admit_read()
        raise error(1004, 'Unable to set the PrintArea property of the PageSetup class')

    @member
    def PrintTitleRows(self) -> object:
        self._read('PrintTitleRows')
        return VBAErrorValue(2042)

    @setter('PrintTitleRows')
    def _print_title_rows(self, value: object) -> None:
        del value
        self._no_setting()

    @member
    def PrintGridlines(self) -> object:
        self._read('PrintGridlines')
        return VBAErrorValue(2042)

    @member
    def Order(self) -> object:
        self._read('Order')
        return VBAErrorValue(2042)

    @member
    def PrintNotes(self) -> object:
        self._read('PrintNotes')
        return VBAErrorValue(2042)


class ChartDataTable(Created):
    """The table of values under a chart (tests/fixtures/chart_members_more.json)."""

    vba_type_name = 'DataTable'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Font(self) -> object:
        return self.chart.font_of(self, 'dataTable')

    @member
    def ShowLegendKey(self) -> object:
        self.chart.admit_read()
        return self.chart.state.data_table_keys

    @setter('ShowLegendKey')
    def _show_legend_key(self, value: object) -> None:
        wanted = to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        self.chart.state.data_table_keys = wanted

    def _border(self, place: int) -> object:
        self.chart.admit_read()
        return self.chart.state.data_table_borders[place]

    def _set_border(self, place: int, value: object) -> None:
        wanted = to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        borders = list(self.chart.state.data_table_borders)
        borders[place] = wanted
        self.chart.state.data_table_borders = (borders[0], borders[1], borders[2])

    @member
    def HasBorderHorizontal(self) -> object:
        return self._border(0)

    @setter('HasBorderHorizontal')
    def _has_border_horizontal(self, value: object) -> None:
        self._set_border(0, value)

    @member
    def HasBorderVertical(self) -> object:
        return self._border(1)

    @setter('HasBorderVertical')
    def _has_border_vertical(self, value: object) -> None:
        self._set_border(1, value)

    @member
    def HasBorderOutline(self) -> object:
        return self._border(2)

    @setter('HasBorderOutline')
    def _has_border_outline(self, value: object) -> None:
        self._set_border(2, value)

    @method
    def Select(self) -> object:
        self.chart.admit_read()
        raise error(438, "Object doesn't support this property or method")

    @method
    def Delete(self) -> object:
        self.chart.vba_set('HasDataTable', False)
        return EMPTY


class ChartArea(Created):
    """The whole of an embedded chart: its frame's place, 5 points smaller each way, and its look
    (tests/fixtures/chart_areas.json)."""

    vba_type_name = 'ChartArea'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    @member
    def Parent(self) -> object:
        self.chart.admit_read()
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Name(self) -> object:
        self.chart.admit_read()
        return 'Chart'

    @member
    def Left(self) -> object:
        return ChartObject(self.chart.sheet, self.chart.shape).vba_get('Left', [])

    @member
    def Top(self) -> object:
        return ChartObject(self.chart.sheet, self.chart.shape).vba_get('Top', [])

    def inside(self, size: float) -> float:
        self.chart.admit_read()
        self.chart.sheet.dims.settle_shapes()
        return points(emu(size)) - 5

    @member
    def Width(self) -> object:
        return self.inside(self.chart.shape.width)

    @member
    def Height(self) -> object:
        return self.inside(self.chart.shape.height)

    @member
    def Shadow(self) -> object:
        self.chart.admit_read()
        return bool(self.chart.state.shadow)

    @member
    def RoundedCorners(self) -> object:
        self.chart.admit_read()
        return self.chart.state.rounded

    @member
    def Format(self) -> object:
        return self.chart.format_of(self, 'area')

    @member
    def Border(self) -> object:
        return self.chart.border_of(self, 'area')

    @member
    def Font(self) -> object:
        return self.chart.font_of(self, 'area')

    # -- tests/fixtures/chart_clears.json

    @method
    def Copy(self) -> object:
        """Selects the chart, and puts nothing that Worksheet.Paste will take (tests/fixtures/chart_copies.json)."""
        self.chart.admit_read()
        ChartObject(self.chart.sheet, self.chart.shape).made_active('area')
        self.chart.sheet.book.application.copied_chart = None
        return EMPTY

    @method
    def ClearFormats(self) -> object:
        self.chart.changed()
        self.chart.state.element_looks.pop('area', None)
        return EMPTY

    @method
    def ClearContents(self) -> object:
        """Takes away the series, leaving the chart's looks, its title and its legend."""
        self.chart.changed()
        state = self.chart.state
        state.series[:] = []
        state.full_series, state.out_of_sight = [], []
        # Series made afterwards go on numbering from those that went.
        state.unique_base = state.next_unique
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        # Selecting the chart area makes its chart the active one.
        self.chart.admit_read()
        ChartObject(self.chart.sheet, self.chart.shape).made_active('area')
        return True


class PlotArea(ElementProperties, Created):
    get_error, set_error = INVALID_CALL, UNSPECIFIED

    """Where a chart draws its series (tests/fixtures/chart_areas.json); its place, which Excel lays out, is not
    modelled."""

    vba_type_name = 'PlotArea'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    @member
    def Parent(self) -> object:
        self.chart.admit_read()
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Name(self) -> object:
        self.chart.admit_read()
        return 'Plot'

    @member
    def Position(self) -> object:
        self.chart.admit_read()
        if self.chart.state.original:
            raise VBAUnsupportedError('The plot area position of an imported chart is not implemented')
        return VBAInt(-4105, 'Long')

    @member
    def Format(self) -> object:
        return self.chart.format_of(self, 'plot')

    @method
    def ClearFormats(self) -> object:
        self.chart.changed()
        self.chart.state.element_looks.pop('plot', None)
        return EMPTY

    @member
    def Border(self) -> object:
        return self.chart.border_of(self, 'plot')

    @method
    def Select(self) -> object:
        self.chart.admit_read()
        ChartObject(self.chart.sheet, self.chart.shape).made_active('plot')
        return True


class HeldByChart(Created):
    """What an embedded chart keeps of its own, which is nothing: its ChartObjects, Shapes and Hyperlinks, ChartObjects
    answering the same for an index (tests/fixtures/chart_members.json)."""

    def __init__(self, type_name: str, chart: Chart) -> None:
        self.vba_type_name, self.chart = type_name, chart

    @member
    def Count(self) -> object:
        return VBAInt(0, 'Long')

    @member
    def Parent(self) -> object:
        """The chart, though Excel fails on an embedded chart's ChartObjects (tests/fixtures/object_trio.json)."""
        if self.vba_type_name == 'ChartObjects':
            raise error(-2147417848, 'Automation error')
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application


class ChartSurface(Created):
    """The walls and floor of a chart. Only a 3-D chart, which the model does not make, has them to name
    (tests/fixtures/chart_members.json)."""

    def __init__(self, chart: Chart, type_name: str) -> None:
        self.chart, self.vba_type_name = chart, type_name

    @member
    def Name(self) -> object:
        self.chart.admit_read()
        raise error(1004, f'Unable to get the Name property of the {self.vba_type_name} class')


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
        if not any(shape is self.shape for shape in self.sheet.shapes_):
            # A deleted chart object has no chart to give (tests/fixtures/active_chart.json).
            raise error(1004, 'Unable to get the Chart property of the ChartObject class')
        return Chart(self.sheet, self.shape)

    @member
    def ShapeRange(self) -> object:
        return ShapeRange([ShapeObject(self.sheet, self.shape)])

    @method
    def Activate(self) -> object:
        self.made_active()
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        # Selecting the chart's frame makes it the active chart as Activate does (tests/fixtures/active_chart.json).
        if Replace is not MISSING:
            raise VBAUnsupportedError('ChartObject.Select with Replace is not implemented')
        self.made_active()
        return EMPTY

    def made_active(self, part: str = '') -> None:
        book = self.sheet.book
        if book.application.active_book is not book or book.active_sheet is not self.sheet:
            raise VBAUnsupportedError('Activating a chart on a sheet that is not active is not implemented')
        if not any(shape is self.shape for shape in self.sheet.shapes_):
            raise VBAUnsupportedError('Activating a deleted chart is not implemented')
        self.sheet.active_chart, self.sheet.chart_part = self.shape, part

    @member
    def ZOrder(self) -> object:
        # The chart's place among all the sheet's shapes (tests/fixtures/missing_reads.json).
        return ShapeObject(self.sheet, self.shape).vba_get('ZOrderPosition', [])

    @member
    def ProtectChartObject(self) -> object:
        return False

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def TopLeftCell(self) -> object:
        return ShapeObject(self.sheet, self.shape).vba_get('TopLeftCell', [])

    @member
    def BottomRightCell(self) -> object:
        return ShapeObject(self.sheet, self.shape).vba_get('BottomRightCell', [])

    @member
    def RoundedCorners(self) -> object:
        return Chart(self.sheet, self.shape).state.rounded

    @setter('RoundedCorners')
    def _rounded_corners(self, value: object) -> None:
        self._look('rounded', value)

    @member
    def Shadow(self) -> object:
        return bool(Chart(self.sheet, self.shape).state.shadow)

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        self._look('shadow', value)

    def _look(self, field_name: str, value: object) -> None:
        """RoundedCorners or Shadow, which the chart part holds (tests/fixtures/chart_object_reads.json)."""
        chart = Chart(self.sheet, self.shape)
        wanted = to_bool(value)
        chart.admit_change()
        chart.changed()
        setattr(chart.state, field_name, wanted)

    @member
    def Left(self) -> object:
        self.sheet.dims.settle_shapes()
        return float(self.shape.left)

    @setter('Left')
    def _left(self, value: object) -> None:
        self._dimension('left', value)

    @member
    def Top(self) -> object:
        self.sheet.dims.settle_shapes()
        return float(self.shape.top)

    @setter('Top')
    def _top(self, value: object) -> None:
        self._dimension('top', value)

    @member
    def Width(self) -> object:
        self.sheet.dims.settle_shapes()
        return float(self.shape.width)

    @setter('Width')
    def _width(self, value: object) -> None:
        self._dimension('width', value)

    @member
    def Height(self) -> object:
        self.sheet.dims.settle_shapes()
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

    # -- tests/fixtures/chart_copies.json

    @method
    def Duplicate(self) -> object:
        """The copy, as a Shape, as Excel answers."""
        return ShapeObject(self.sheet, self.shape).Duplicate()

    @method
    def Copy(self) -> object:
        return ShapeObject(self.sheet, self.shape).Copy()

    @method
    def Cut(self) -> object:
        return ShapeObject(self.sheet, self.shape).Cut()

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


class Chart(ElementProperties, Created):
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
        settle_hidden(self.sheet, self.state)

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

    def nothing_out(self, what: str) -> None:
        """As nothing_filtered, for what series with their cells all hidden make no odd answers of
        (tests/fixtures/chart_hidden_cells.json)."""
        state = self.state
        away = [*state.series, *state.out_of_sight]
        if any(not any(entry is other for other in away) for entry in state.full_series):
            self.nothing_filtered(what)

    def nothing_filtered(self, what: str) -> None:
        """Refuses what gives odd answers around a series filtered out (tests/fixtures/series_filter.json)."""
        if self.state.full_series:
            raise VBAUnsupportedError(f'{what} on a chart with a series filtered out is not implemented')

    def keeps_look(self, key: str, wanted: bool = False) -> None:
        """Refuses taking away an element with a look: what it comes back with was not measured."""
        if not wanted and (key in self.state.element_looks or key in self.state.element_fonts):
            raise VBAUnsupportedError('Taking away a chart element with a look of its own is not implemented')

    @method
    def Activate(self) -> object:
        return self.not_a_sheet('Activate')

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        return self.not_a_sheet('Select')

    def not_a_sheet(self, verb: str) -> object:
        # An embedded chart that is not active cannot be activated or selected as a chart sheet would be
        # (tests/fixtures/active_chart.json).
        self.admit_read()
        if self.sheet.active_chart is self.shape:
            raise VBAUnsupportedError(f'Chart.{verb} on the active embedded chart is not implemented')
        raise error(1004, f'{verb} method of Chart class failed')

    @member
    def Parent(self) -> object:
        return ChartObject(self.sheet, self.shape)

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    # -- what an embedded chart answers for what only a chart sheet or a pivot chart has, which the model holds
    # neither of (tests/fixtures/missing_reads.json)

    def sheet_only(self, name: str) -> object:
        self.admit_read()
        raise error(1004, f'Unable to get the {name} property of the Chart class')

    @member
    def Index(self) -> object:
        return self.sheet_only('Index')

    @member
    def CodeName(self) -> object:
        return self.sheet_only('CodeName')

    @member
    def Visible(self) -> object:
        return self.sheet_only('Visible')

    @member
    def Tab(self) -> object:
        return self.sheet_only('Tab')

    @member
    def Next(self) -> object:
        return self.sheet_only('Next')

    @member
    def Previous(self) -> object:
        return self.sheet_only('Previous')

    @member
    def ProtectContents(self) -> object:
        return self.sheet_only('ProtectContents')

    @member
    def ProtectDrawingObjects(self) -> object:
        return self.sheet_only('ProtectDrawingObjects')

    @member
    def ProtectionMode(self) -> object:
        return self.sheet_only('ProtectionMode')

    @member
    def ProtectData(self) -> object:
        self.admit_read()
        return self.state.protection.get('data', False)

    @setter('ProtectData')
    def _protectdata(self, value: object) -> None:
        wanted = to_bool(value)
        self.admit_change()
        self.changed()
        self.state.protection['data'] = wanted

    @member
    def ProtectFormatting(self) -> object:
        self.admit_read()
        return self.state.protection.get('formatting', False)

    @setter('ProtectFormatting')
    def _protectformatting(self, value: object) -> None:
        wanted = to_bool(value)
        self.admit_change()
        self.changed()
        self.state.protection['formatting'] = wanted

    @member
    def ProtectSelection(self) -> object:
        self.admit_read()
        return self.state.protection.get('selection', False)

    @setter('ProtectSelection')
    def _protectselection(self, value: object) -> None:
        wanted = to_bool(value)
        self.admit_change()
        self.changed()
        self.state.protection['selection'] = wanted

    @member
    def ProtectGoalSeek(self) -> object:
        """Takes True and goes on answering False (tests/fixtures/chart_protection.json)."""
        self.admit_read()
        return False

    @setter('ProtectGoalSeek')
    def _protect_goal_seek(self, value: object) -> None:
        to_bool(value)
        self.admit_read()

    @method
    def Protect(self, Password: object = MISSING, DrawingObjects: object = MISSING, Contents: object = MISSING,  # noqa: N803 - native named arguments
                Scenarios: object = MISSING, UserInterfaceOnly: object = MISSING) -> object:
        """Refused on a chart on a sheet (tests/fixtures/chart_protection.json)."""
        del Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly
        self.admit_read()
        raise error(1004, 'Protect method of Chart class failed')

    @method
    def Unprotect(self, Password: object = MISSING) -> object:  # noqa: N803 - native named argument
        del Password
        self.admit_read()
        raise error(1004, 'Unprotect method of Chart class failed')

    @member
    def PageSetup(self) -> object:
        self.admit_read()
        return ChartPageSetup(self)

    @member
    def PivotLayout(self) -> object:
        self.admit_read()
        return NOTHING

    # -- what a 2-D chart the model made reads for the 3-D view, its style and its label levels; a pie or doughnut is
    # seen from higher and unturned (tests/fixtures/chart_type_reads.json)

    def made(self) -> ChartState:
        self.admit_read()
        if self.state.original:
            raise VBAUnsupportedError("Reading this setting of an imported chart is not implemented")
        return self.state

    @member
    def Elevation(self) -> object:
        return VBAInt(75 if self.made().chart_type in (5, -4120) else 15, 'Long')

    @member
    def Rotation(self) -> object:
        return VBAInt(0 if self.made().chart_type in (5, -4120) else 20, 'Long')

    @member
    def Perspective(self) -> object:
        self.made()
        return VBAInt(30, 'Long')

    @member
    def RightAngleAxes(self) -> object:
        self.made()
        return False

    @member
    def DepthPercent(self) -> object:
        self.made()
        return VBAInt(100, 'Long')

    @member
    def HeightPercent(self) -> object:
        self.made()
        return VBAInt(100, 'Long')

    @member
    def GapDepth(self) -> object:
        self.made()
        return VBAInt(150, 'Long')

    @member
    def AutoScaling(self) -> object:
        self.made()
        return True

    @member
    def BarShape(self) -> object:
        self.made()
        return VBAInt(0, 'Long')

    @member
    def ChartStyle(self) -> object:
        return VBAInt(self.made().chart_style, 'Long')

    @setter('ChartStyle')
    def _chart_style(self, value: object) -> None:
        from pyopenvba.apps.excel._chart_style_table import DEFAULT_STYLES, PROVEN, extras, layout
        from pyopenvba.apps.excel._chart_styled import is_styled

        wanted = int(to_integer(value, 'Long'))
        self.admit_change()
        if wanted < 1 or wanted == 49:
            raise error(5, 'Invalid procedure call or argument')
        if wanted > 48 and not (is_styled(self.state.chart_style) and wanted in PROVEN.get(self.state.chart_type, ())):
            # Turning a chart of the old kind to one of AddChart2's styles was not measured, nor the rest of those.
            raise VBAUnsupportedError('A chart style past 48 is not implemented')
        self.changed()
        self.state.chart_style = wanted
        kind = self.state.chart_type
        if wanted > 48 and (wanted not in DEFAULT_STYLES.values() or wanted == DEFAULT_STYLES.get(kind)):
            # A style is a layout too, but a style that is another chart type's default brings only its looks
            # (tests/fixtures/add_chart2.json, three cases).
            _set_out(self, layout(wanted, kind))
            _labelled(self.state)
            place = extras(wanted, CHART_TYPES[kind])[1]
            if self.state.legend and place:
                self.state.legend_position = place

    @member
    def ChartColor(self) -> object:
        # The eight colour sets repeat through the 48 styles (tests/fixtures/chart_layouts.json); the styles from 201
        # share one colours part (tests/fixtures/add_chart2.json).
        style = self.made().chart_style
        return VBAInt(10 if style > 48 else (style - 1) % 8 + 1, 'Long')

    @setter('ChartColor')
    def _chart_color(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        self.admit_change()
        if not 1 <= wanted <= 26:
            raise error(5, 'Invalid procedure call or argument')
        raise VBAUnsupportedError('ChartColor, which brings parts of its own, is not implemented')

    # -- tests/fixtures/chart_members.json

    @method
    def Delete(self) -> object:
        self.admit_read()
        raise error(1004, 'Delete method of Chart class failed')

    @method
    def Move(self, Before: object = MISSING, After: object = MISSING) -> object:  # noqa: N803 - native named args
        self.admit_read()
        raise error(1004, 'Move method of Chart class failed')

    @method
    def Copy(self, Before: object = MISSING, After: object = MISSING) -> object:  # noqa: N803 - native named args
        self.admit_read()
        raise error(1004, 'Copy method of Chart class failed')

    @method
    def Refresh(self) -> object:
        self.admit_read()
        return EMPTY

    @method
    def Calculate(self) -> object:
        raise error(438, "Object doesn't support this property or method")

    @method
    def Resize(self, *_: object) -> object:
        raise error(438, "Object doesn't support this property or method")

    @method
    def ClearToMatchColorStyle(self) -> object:
        self.admit_read()
        return EMPTY

    @member
    def ChartObjects(self, Index: object = MISSING) -> object:  # noqa: N803 - native named arg
        self.admit_read()
        return HeldByChart('ChartObjects', self)

    @member
    def Shapes(self) -> object:
        self.admit_read()
        return HeldByChart('Shapes', self)

    @member
    def Hyperlinks(self) -> object:
        self.admit_read()
        return HeldByChart('Hyperlinks', self)

    @member
    def OLEObjects(self, Index: object = MISSING) -> object:  # noqa: N803 - native named arg
        self.admit_read()
        raise error(1004, 'OLEObjects method of Chart class failed')

    @method
    def Evaluate(self, Name: object) -> object:  # noqa: N803 - native named arg
        """A chart's Evaluate answers the name of an element, not a formula."""
        self.admit_read()
        if to_text(Name).casefold() == 'legend' and self.state.legend:
            return Legend(self)
        raise error(-2147467259, 'Automation error\nUnspecified error ')

    @method
    def Location(self, Where: object, Name: object = MISSING) -> object:  # noqa: N803 - native named args
        self.admit_read()
        where = to_integer(Where, 'Long')
        if where not in (1, 2, 3):
            raise error(13, 'Type mismatch')
        if where != 2:
            raise VBAUnsupportedError('Moving a chart to a chart sheet is not implemented')
        if Name is MISSING:
            raise error(5, 'Invalid procedure call or argument')
        wanted = to_text(Name)
        if not any(sheet.name.casefold() == wanted.casefold() for sheet in self.sheet.book.sheets_):
            raise error(1004, 'Location method of Chart class failed')
        if wanted.casefold() != self.sheet.name.casefold():
            raise VBAUnsupportedError('Moving a chart to another sheet is not implemented')
        return self

    @member
    def Walls(self) -> object:
        return ChartSurface(self, 'Walls')

    @member
    def BackWall(self) -> object:
        return ChartSurface(self, 'Walls')

    @member
    def SideWall(self) -> object:
        return ChartSurface(self, 'Walls')

    @member
    def Floor(self) -> object:
        return ChartSurface(self, 'Floor')

    @method
    def ClearToMatchStyle(self) -> object:
        """Takes away every look set on the chart, leaving what it has and its style."""
        self.admit_change()
        state = self.state
        if any(entry.data_labels for entry in state.series):
            raise VBAUnsupportedError('ClearToMatchStyle on a chart with data labels is not implemented')
        self.changed()
        state.element_looks.clear()
        state.element_fonts.clear()
        state.title_look = TitleLook()
        for entry in state.series:
            entry.format_look = None
            entry.points.clear()
            for line in entry.trendlines:
                line.look = FillLine()
            for bar in entry.error_bars:
                bar.look = FillLine()
        return True

    @method
    def ApplyLayout(self, Layout: object, ChartType: object = MISSING) -> object:  # noqa: N803 - native named arguments
        from pyopenvba.apps.excel._chart_elements import apply_layout

        return apply_layout(self, Layout, ChartType)

    @method
    def ChartWizard(self, Source: object = MISSING, Gallery: object = MISSING, Format: object = MISSING,  # noqa: N803
                    PlotBy: object = MISSING, CategoryLabels: object = MISSING,  # noqa: N803
                    SeriesLabels: object = MISSING, HasLegend: object = MISSING,  # noqa: N803
                    Title: object = MISSING, CategoryTitle: object = MISSING, ValueTitle: object = MISSING,  # noqa: N803
                    ExtraTitle: object = MISSING) -> object:  # noqa: N803
        from pyopenvba.apps.excel._chart_elements import chart_wizard

        del ExtraTitle
        return chart_wizard(self, Source, Gallery, Format, PlotBy, CategoryLabels, SeriesLabels, HasLegend, Title,
                            CategoryTitle, ValueTitle)

    def single_levels(self, what: str) -> None:
        state = self.made()
        for entry in state.series:
            if category_levels(self.sheet, entry) or (
                    entry.name_reference and len(referenced_values(self.sheet, entry.name_reference)) > 1):
                raise VBAUnsupportedError(f'{what} over labels of several levels is not implemented')

    @member
    def CategoryLabelLevel(self) -> object:
        self.single_levels('CategoryLabelLevel')
        return VBAInt(0, 'Long')

    @member
    def SeriesNameLevel(self) -> object:
        self.single_levels('SeriesNameLevel')
        return VBAInt(0, 'Long')

    @member
    def ChartArea(self) -> object:
        self.admit_read()
        return ChartArea(self)

    @member
    def PlotArea(self) -> object:
        self.admit_read()
        return PlotArea(self)

    def border_of(self, element: object, key: str) -> object:
        """The older way to the element's line (tests/fixtures/chart_borders.json)."""
        from pyopenvba.apps.excel._chart_borders import ChartBorder, family_of
        from pyopenvba.apps.excel._chart_format import ElementFormat, ElementLine

        self.admit_read()
        line = ElementLine(ElementFormat(element, self, key))
        return ChartBorder(line, family_of(key), element, self.sheet.book.stylesheet.colors)

    def format_of(self, element: object, key: str) -> object:
        from pyopenvba.apps.excel._chart_format import ElementFormat

        self.admit_read()
        return ElementFormat(element, self, key)

    def font_of(self, element: object, key: str, *, size: float = 10.0, bold: bool = False,
                labels: Any = None) -> object:
        """The Font of the element's text, whose defaults are given."""
        from pyopenvba.apps.excel._chart_fonts import ChartFont

        self.admit_read()
        return ChartFont(element, self, key, size, bold, labels)

    @member
    def ShowAllFieldButtons(self) -> object:
        self.admit_read()
        return False

    @member
    def ShowExpandCollapseEntireFieldButtons(self) -> object:
        self.admit_read()
        return False

    @member
    def ShowAxisFieldButtons(self) -> object:
        return self.sheet_only('ShowAxisFieldButtons')

    @member
    def ShowLegendFieldButtons(self) -> object:
        return self.sheet_only('ShowLegendFieldButtons')

    @member
    def ShowReportFilterFieldButtons(self) -> object:
        return self.sheet_only('ShowReportFilterFieldButtons')

    @member
    def ShowValueFieldButtons(self) -> object:
        return self.sheet_only('ShowValueFieldButtons')

    @member
    def ChartType(self) -> object:
        if self.state.chart_type == 0:
            raise VBAUnsupportedError('Reading this imported chart type is not implemented')
        if self.state.made_scatter and not self.state.series and self.state.chart_type == -4169:
            # One of AddChart2's scatter charts reads as with lines until it has a series
            # (tests/fixtures/add_chart2.json).
            return VBAInt(74, 'Long')
        return VBAInt(shown_type(self.state), 'Long')

    @setter('ChartType')
    def _chart_type(self, value: object) -> None:
        wanted = int(to_number(value))
        if wanted not in CHART_TYPES:
            raise VBAUnsupportedError(f'Chart type {wanted} is not implemented')
        if self.state.element_looks or self.state.element_fonts or any(
                one.format_look is not None or one.points or one.label_look.font is not None
                for one in self.state.series):
            raise VBAUnsupportedError('Changing the type of a chart with a look of its own is not implemented')
        from pyopenvba.apps.excel._chart_axes import transition

        self.admit_change()
        transition(self, wanted)
        self.changed()
        if CHART_TYPES[wanted] in ('column', 'bar') and (wanted in STACKED) != (self.state.chart_type in STACKED):
            # Stacking a column or bar chart overlaps its bars fully, and unstacking it parts them
            # (tests/fixtures/chart_stacked.json).
            self.state.overlap = 100 if wanted in STACKED else 0
        if self.state.style_labels:
            # A style's labels keep what they show and lose their place (tests/fixtures/add_chart2.json).
            self.state.style_labels = [*self.state.style_labels[:4], '']
            for entry in self.state.series:
                entry.label_look.position = None
        if (self.state.chart_style > 48 and wanted not in STACKED and self.state.chart_type not in STACKED
                and {CHART_TYPES[wanted], CHART_TYPES.get(self.state.chart_type)} == {'column', 'bar'}):
            # A style's overlap does not survive its bars turning (tests/fixtures/add_chart2.json).
            self.state.overlap = 0
        self.state.chart_type = wanted
        if self.sheet.active_chart is self.shape:
            # A new type for the active chart selects its chart area; a title does not
            # (tests/fixtures/active_chart.json, tests/fixtures/chart_areas.json).
            self.sheet.chart_part = 'area'
        # A line chart turns every marker off; one with markers, or a scatter chart, turns those that were off back on.
        for entry in self.state.series:
            if wanted in (plain for plain, _ in LINE_PAIRS):
                entry.marker = MARKER_NONE
            elif wanted in MARKED_TYPES and entry.marker == MARKER_NONE:
                entry.marker = MARKER_AUTOMATIC

    @member
    def HasTitle(self) -> object:
        return self.state.title is not None

    @setter('HasTitle')
    def _has_title(self, value: object) -> None:
        from pyopenvba.apps.excel._chart_styled import is_styled

        self.admit_change()
        wanted = to_bool(value)
        self.keeps_look('title', wanted)
        self.changed()
        # A styled chart's title is taken away as deleting it does (tests/fixtures/add_chart2.json).
        self.state.title_deleted = not wanted and is_styled(self.state.chart_style)
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
    def DataTable(self) -> object:
        self.admit_read()
        if not self.state.data_table:
            raise error(1004, 'Unable to get the DataTable property of the Chart class')
        return ChartDataTable(self)

    @member
    def HasDataTable(self) -> object:
        """Whether the chart shows a data table (tests/fixtures/chart_series_more.json)."""
        self.admit_read()
        # A pie, a doughnut or a scatter chart keeps the table in its file and answers that it has none
        # (tests/fixtures/chart_members_more.json).
        return self.state.data_table and CHART_TYPES.get(shown_type(self.state)) not in ('pie', 'doughnut', 'scatter')

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
        self.keeps_look('legend', wanted)
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
        self.admit_read()
        if self.state.title is None:
            if not self.state.has_automatic_title():
                raise error(-2147024809, 'The parameter is incorrect.')
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
    def ApplyDataLabels(self, Type: object = MISSING, LegendKey: object = MISSING, AutoText: object = MISSING,
                        HasLeaderLines: object = MISSING, ShowSeriesName: object = MISSING,
                        ShowCategoryName: object = MISSING, ShowValue: object = MISSING,
                        ShowPercentage: object = MISSING, ShowBubbleSize: object = MISSING,
                        Separator: object = MISSING) -> object:
        from pyopenvba.apps.excel._chart_data_labels import applied_look, apply_to_series

        del ShowBubbleSize
        try:
            look = applied_look(self.state.chart_type, Type, LegendKey, AutoText, ShowSeriesName,
                                ShowCategoryName, ShowValue, ShowPercentage, Separator)
        except VBARuntimeError:
            # A chart takes a type its series refuse: no labels, and no error.
            return True
        for entry in self.state.series:
            apply_to_series(Series(self, entry), look, HasLeaderLines)
        return True

    @method
    def SeriesCollection(self, Index: object = MISSING) -> object:
        collection = SeriesCollection(self)
        return collection if Index is MISSING else collection.vba_get('Item', [Index])

    @method
    def SetElement(self, Element: object) -> object:  # noqa: N803 - native named argument
        from pyopenvba.apps.excel._chart_elements import set_element

        return set_element(self, Element)

    @method
    def FullSeriesCollection(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        collection = FullSeriesCollection(self)
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
        self.keeps_look(f'axis{found.entry.kind}', wanted)
        self.keeps_look(f'tick{found.entry.kind}', wanted)
        self.changed()
        found.entry.present = wanted and self.state.chart_type not in NO_AXES

    @method
    def SetSourceData(self, Source: object = MISSING, PlotBy: object = MISSING) -> object:
        from pyopenvba.apps.excel._model import Range

        if Source is MISSING:
            raise error(450)
        self.nothing_out('SetSourceData')
        if Source is NOTHING:
            raise VBAUnsupportedError('Chart source Nothing object lifetime behavior is not implemented')
        if not isinstance(Source, Range):
            raise error(13)
        if Source.sheet.book is not self.sheet.book:
            raise VBAUnsupportedError('Cross-workbook chart source data is not implemented')
        at = squeezed(Source.sheet, Source.areas) if len(Source.areas) != 1 else None
        area = at.block() if at is not None else Source.first
        height, width = area.bottom - area.top + 1, area.right - area.left + 1
        if height * width > MAX_CHART_POINTS:
            raise VBAUnsupportedError('Chart source data exceeds the supported cell budget')
        if PlotBy is NULL:
            raise error(13)
        orientation = (guessed_orientation(Source.sheet, area, at) if PlotBy is MISSING
                       else int(to_integer(PlotBy, 'Long')))
        if orientation not in (1, 2):
            raise error(-2147467259, 'Automation error\nUnspecified error ')
        self.admit_change()
        series_plan = plan_series(Source.sheet, area, orientation, at)
        self.changed()
        self.state.plot_by = orientation
        for index, series in enumerate(series_plan):
            series.unique = self.state.unique_base + index
            series.marker = default_marker(self.state)
        self.state.next_unique = self.state.unique_base + len(series_plan)
        selected = self.selected_series()
        place = -1 if selected is None else selected.position() - 1
        self.state.series[:] = series_plan
        self.state.full_series, self.state.out_of_sight = [], []
        _labelled(self.state)
        if selected is not None and place < len(series_plan):
            # What is selected stays selected, in the same place among the new series.
            selected.entry = series_plan[place]
        return EMPTY

    def element_there(self) -> bool:
        """Whether the element selected is still in the chart: its series, or for the ones that come and go, itself."""
        element = self.sheet.chart_element
        state = self.state
        if isinstance(element, Legend):
            return state.legend
        if isinstance(element, ChartTitle):
            return state.title is not None
        entry = getattr(element, 'entry', None)
        if entry is not None and hasattr(entry, 'present'):
            return bool(entry.present)
        series = element if isinstance(element, Series) else getattr(element, 'series', None)
        return series is None or self.selected_series() is not None

    def selected_series(self) -> Series | None:
        """The series selected, or the series of the point selected, of this chart."""
        sheet = self.sheet
        if sheet.active_chart is not self.shape or sheet.chart_part != 'element':
            return None
        element = sheet.chart_element
        series = element if isinstance(element, Series) else getattr(element, 'series', None)
        return series if isinstance(series, Series) and any(
            entry is series.entry for entry in self.state.series) else None


class ChartTitle(ElementProperties, TitleMembers):
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

    @method
    def Select(self) -> object:
        return select_element(self.chart, self, True)

    @member
    def Parent(self) -> object:
        self.chart.admit_read()
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def Name(self) -> object:
        # tests/fixtures/chart_object_reads.json
        self.chart.admit_read()
        return 'Title'

    @member
    def Format(self) -> object:
        return self.chart.format_of(self, 'title')

    @member
    def Border(self) -> object:
        return self.chart.border_of(self, 'title')

    @member
    def Font(self) -> object:
        return self.chart.font_of(self, 'title', size=18.0, bold=True)

    def title_link(self, formula: str) -> None:
        self.chart.changed()
        self.chart.state.title_typed = False
        self.chart.state.title_look.formula = formula

    @member
    def Text(self) -> object:
        state = self.chart.state
        if state.title_look.formula:
            return linked_title(self.chart.sheet, state.title_look.formula)
        drawn_one = len(state.series) == 1 or CHART_TYPES.get(shown_type(state)) == 'pie'
        if state.title is not None and not state.title_typed and drawn_one and state.series:
            if state.chart_style > 48 and not (state.series[0].name_reference or state.series[0].name):
                # One of AddChart2's new charts has no name to take from a series with none
                # (tests/fixtures/add_chart2.json).
                return state.title
            # The automatic title of a chart that draws one series is that series' name.
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
        self.chart.keeps_look('title')
        self.chart.changed()
        state = self.chart.state
        state.title, state.title_typed, state.title_look, state.title_deleted = None, False, TitleLook(), True
        return EMPTY


def linked_title(sheet: Worksheet, formula: str) -> str:
    values = referenced_values(sheet, formula)
    return to_text(values[0]) if values and values[0] is not EMPTY else ''


#: Legend.Position and how the file spells it; xlLegendPositionCustom needs native layout, and Excel refuses it.
LEGEND_POSITIONS = {-4107: 'b', -4160: 't', -4131: 'l', -4152: 'r', 2: 'tr'}


class Legend(ElementProperties, Created):
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

    @method
    def Select(self) -> object:
        return select_element(self.chart, self, True)

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def Name(self) -> object:
        # tests/fixtures/chart_object_reads.json
        self.shown()
        return 'Legend'

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

    @member
    def Format(self) -> object:
        self.shown()
        return self.chart.format_of(self, 'legend')

    @member
    def Border(self) -> object:
        self.shown()
        return self.chart.border_of(self, 'legend')

    @member
    def Font(self) -> object:
        self.shown()
        return self.chart.font_of(self, 'legend')

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

    @member
    def Parent(self) -> object:
        return self.legend

    @member
    def Application(self) -> object:
        return self.legend.chart.sheet.book.application

    def vba_items(self) -> list[object]:
        state = self.legend.shown()
        if CHART_TYPES.get(shown_type(state)) in ('pie', 'doughnut'):
            first = state.series[0] if state.series else None
            # Values lost from the series it shows leave the categories their entries, while another keeps its own.
            total = (0 if first is None else len(Series(self.legend.chart, first).categories())
                     if lost(first.value_reference) and not all(lost(entry.value_reference) for entry in state.series)
                     else len(Series(self.legend.chart, first).values()))
        else:
            total = len(state.series) + sum(len(entry.trendlines) for entry in state.series)
        return [LegendEntry(self.legend, index) for index in range(total) if index not in state.legend_deleted]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_number(index))
        if not 1 <= position <= len(items):
            raise error(-2147467259, 'Automation error\nUnspecified error')
        return items[position - 1]


class LegendKey(Created):
    """The key beside a legend entry, which is its series under another name: what is done to it is done to the series,
    its deletion and selection included (tests/fixtures/chart_clears.json)."""

    vba_type_name = 'LegendKey'

    def __init__(self, series: Series, entry: LegendEntry) -> None:
        self.series, self.entry = series, entry

    @member
    def Parent(self) -> object:
        return self.entry

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Format(self) -> object:
        return self.series.vba_get('Format', [])

    @member
    def MarkerStyle(self) -> object:
        return self.series.vba_get('MarkerStyle', [])

    @setter('MarkerStyle')
    def _marker_style(self, value: object) -> None:
        if CHART_TYPES.get(shown_type(self.series.chart.state)) not in _DRAWS_MARKERS:
            # The key takes a style its series refuses.
            raise VBAUnsupportedError('LegendKey.MarkerStyle on a chart that draws no markers is not implemented')
        self.series.vba_set('MarkerStyle', value)

    @member
    def MarkerSize(self) -> object:
        return self.series.vba_get('MarkerSize', [])

    @setter('MarkerSize')
    def _marker_size(self, value: object) -> None:
        self.series.vba_set('MarkerSize', value)

    @member
    def MarkerBackgroundColor(self) -> object:
        return self.series.vba_get('MarkerBackgroundColor', [])

    @setter('MarkerBackgroundColor')
    def _marker_background_color(self, value: object) -> None:
        self.series.vba_set('MarkerBackgroundColor', value)

    @member
    def MarkerBackgroundColorIndex(self) -> object:
        return self.series.vba_get('MarkerBackgroundColorIndex', [])

    @setter('MarkerBackgroundColorIndex')
    def _marker_background_color_index(self, value: object) -> None:
        self.series.vba_set('MarkerBackgroundColorIndex', value)

    @member
    def MarkerForegroundColor(self) -> object:
        return self.series.vba_get('MarkerForegroundColor', [])

    @setter('MarkerForegroundColor')
    def _marker_foreground_color(self, value: object) -> None:
        self.series.vba_set('MarkerForegroundColor', value)

    @member
    def MarkerForegroundColorIndex(self) -> object:
        return self.series.vba_get('MarkerForegroundColorIndex', [])

    @setter('MarkerForegroundColorIndex')
    def _marker_foreground_color_index(self, value: object) -> None:
        self.series.vba_set('MarkerForegroundColorIndex', value)

    @member
    def Shadow(self) -> object:
        return self.series.vba_get('Shadow', [])

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        self.series.vba_set('Shadow', value)

    @member
    def PictureType(self) -> object:
        return VBAInt(1, 'Long')

    @member
    def PictureUnit2(self) -> object:
        return VBASingle(0.0)

    @method
    def ClearFormats(self) -> object:
        return self.series.vba_call('ClearFormats', [])

    @member
    def Smooth(self) -> object:
        return self.series.vba_get('Smooth', [])

    @member
    def InvertIfNegative(self) -> object:
        return self.series.vba_get('InvertIfNegative', [])

    @method
    def Delete(self) -> object:
        return self.series.Delete()

    @method
    def Select(self) -> object:
        return self.series.Select()


class LegendEntry(Created):
    vba_type_name = 'LegendEntry'

    def __init__(self, legend: Legend, index: int) -> None:
        self.legend, self.index = legend, index

    @method
    def Select(self) -> object:
        return select_element(self.legend.chart, self, True)

    # -- tests/fixtures/chart_clears.json

    @member
    def Index(self) -> object:
        return VBAInt(self.index + 1, 'Long')

    @member
    def Format(self) -> object:
        """A fill that reads as set and reaches no file, and no line (tests/fixtures/chart_members_more.json)."""
        from pyopenvba.apps.excel._chart_format import EntryFormat

        self.legend.shown()
        return EntryFormat(self, self.legend.chart, f'legendEntry{self.index}')

    @member
    def Font(self) -> object:
        """Its own, which starts as the legend's."""
        from pyopenvba.apps.excel._chart_fonts import ChartFont

        chart = self.legend.chart
        shown = cast(ChartFont, self.legend.Font())
        return chart.font_of(self, f'legend{self.index}', size=float(cast(float, to_number(shown.vba_get('Size', [])))),
                             bold=bool(shown.vba_get('Bold', [])))

    @member
    def LegendKey(self) -> object:
        state = self.legend.shown()
        if self.index >= len(state.series) or CHART_TYPES.get(shown_type(state)) in ('pie', 'doughnut'):
            raise VBAUnsupportedError('The legend key of a trendline or of a pie slice is not implemented')
        return LegendKey(Series(self.legend.chart, state.series[self.index]), self)

    @member
    def Parent(self) -> object:
        return self.legend

    @member
    def Application(self) -> object:
        return self.legend.chart.sheet.book.application

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

    @member
    def Parent(self) -> object:
        return self.chart

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    def vba_items(self) -> list[object]:
        self.chart.admit_read()
        return [ChartGroup(self.chart)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if int(to_integer(index, 'Long')) != 1:
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
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def AxisGroup(self) -> object:
        # The one group the model draws is on the primary axes (tests/fixtures/chart_object_reads.json).
        self._kind()
        return VBAInt(1, 'Long')

    @member
    def Index(self) -> object:
        self._kind()
        return VBAInt(1, 'Long')

    @method
    def SeriesCollection(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        """The chart's series, every one of them in the group (tests/fixtures/chart_object_reads.json)."""
        self._kind()
        return self.chart.SeriesCollection(Index)

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
        kept = self.chart.state.layout_overlap
        drawn = 100 if kept is None else kept
        return VBAInt(drawn if self._kind() in ('line', 'lineMarkers', 'area', 'scatter') else self.chart.state.overlap,
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

    def _flag(self, field_name: str, value: object) -> None:
        """One of the group's lines or bars on or off; a kind that does not draw it keeps it and writes nothing
        (tests/fixtures/chart_group_lines.json)."""
        wanted = to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        setattr(self.chart.state, field_name, wanted)
        if not wanted:
            for key in _GROUP_LOOKS[field_name]:
                self.chart.state.element_looks.pop(key, None)

    # What a group has for a kind the model does not draw: kept, and written nowhere
    # (tests/fixtures/chart_group_looks.json).

    def _kept(self, field_name: str) -> object:
        self._kind()
        return getattr(self.chart.state, field_name)

    def _ranged(self, field_name: str, value: object, lowest: int, highest: int, code: int = 5) -> None:
        wanted = int(to_integer(value, 'Long'))
        self.chart.admit_change()
        if not lowest <= wanted <= highest:
            raise error(code, 'Invalid procedure call or argument' if code == 5 else 'Automation error')
        self.chart.changed()
        setattr(self.chart.state, field_name, wanted)

    def _switched(self, field_name: str, value: object) -> None:
        wanted = to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        setattr(self.chart.state, field_name, wanted)

    def _categories(self, index: object) -> object:
        self._kind()
        found = CategoryCollection(self)
        return found if index is MISSING else found.Item(index)

    @method
    def CategoryCollection(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        return self._categories(Index)

    @method
    def FullCategoryCollection(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        return self._categories(Index)

    @member
    def BubbleScale(self) -> object:
        return VBAInt(cast(int, self._kept('bubble_scale')), 'Long')

    @setter('BubbleScale')
    def _bubble_scale(self, value: object) -> None:
        self._ranged('bubble_scale', value, 0, 300)

    @member
    def SizeRepresents(self) -> object:
        return VBAInt(cast(int, self._kept('size_represents')), 'Long')

    @setter('SizeRepresents')
    def _size_represents(self, value: object) -> None:
        self._ranged('size_represents', value, 1, 2)

    @member
    def SecondPlotSize(self) -> object:
        return VBAInt(cast(int, self._kept('second_plot_size')), 'Long')

    @setter('SecondPlotSize')
    def _second_plot_size(self, value: object) -> None:
        self._ranged('second_plot_size', value, 5, 200)

    @member
    def ShowNegativeBubbles(self) -> object:
        return self._kept('negative_bubbles')

    @setter('ShowNegativeBubbles')
    def _show_negative_bubbles(self, value: object) -> None:
        self._switched('negative_bubbles', value)

    @member
    def Has3DShading(self) -> object:
        return self._kept('shading')

    @setter('Has3DShading')
    def _has_3d_shading(self, value: object) -> None:
        if not to_bool(value):
            # Excel writes a flat 3-D material on every series.
            raise VBAUnsupportedError('Has3DShading = False is not implemented')
        self._switched('shading', value)

    @member
    def SplitType(self) -> object:
        return VBAInt(cast(int, self._kept('split_type')), 'Long')

    @setter('SplitType')
    def _split_type(self, value: object) -> None:
        self._ranged('split_type', value, 1, 4, 1004)
        # A new type starts over: the first point, 1 of value, 10 percent, and with a custom split no value at all.
        self.chart.state.split_value = {1: 1, 2: 1, 3: 10}.get(self.chart.state.split_type)

    @member
    def SplitValue(self) -> object:
        value = self._kept('split_value')
        return EMPTY if value is None else VBAInt(cast(int, value), 'Long')

    @setter('SplitValue')
    def _split_value(self, value: object) -> None:
        self._ranged('split_value', value, 0, 2147483647, -2147467259)

    @member
    def HasRadarAxisLabels(self) -> object:
        self._kind()
        raise error(-2147467259, 'Automation error\nUnspecified error')

    @setter('HasRadarAxisLabels')
    def _has_radar_axis_labels(self, value: object) -> None:
        del value
        self._kind()
        raise error(-2147467259, 'Automation error\nUnspecified error')

    @member
    def RadarAxisLabels(self) -> object:
        self._kind()
        raise error(-2147467259, 'Automation error\nUnspecified error')

    @member
    def HasDropLines(self) -> object:
        self._kind()
        return self.chart.state.drop_lines

    @setter('HasDropLines')
    def _has_drop_lines(self, value: object) -> None:
        self._flag('drop_lines', value)

    @member
    def HasHiLoLines(self) -> object:
        self._kind()
        return self.chart.state.hi_lo_lines

    @setter('HasHiLoLines')
    def _has_hi_lo_lines(self, value: object) -> None:
        self._flag('hi_lo_lines', value)

    @member
    def HasUpDownBars(self) -> object:
        self._kind()
        return self.chart.state.up_down_bars

    @setter('HasUpDownBars')
    def _has_up_down_bars(self, value: object) -> None:
        self._flag('up_down_bars', value)

    @member
    def HasSeriesLines(self) -> object:
        self._kind()
        return self.chart.state.series_lines

    @setter('HasSeriesLines')
    def _has_series_lines(self, value: object) -> None:
        self._flag('series_lines', value)

    # -- the lines and bars themselves, which a group without them answers with 424
    # (tests/fixtures/chart_type_reads.json)

    @member
    def DropLines(self) -> object:
        return DropLines(self).shown()

    @member
    def HiLoLines(self) -> object:
        return HiLoLines(self).shown()

    @member
    def UpBars(self) -> object:
        return UpBars(self).shown()

    @member
    def DownBars(self) -> object:
        return DownBars(self).shown()

    @member
    def SeriesLines(self) -> object:
        return SeriesLines(self).shown()

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


#: The looks that go with a chart group's lines or bars.
_GROUP_LOOKS = {'drop_lines': ('dropLines',), 'hi_lo_lines': ('hiLowLines',), 'up_down_bars': ('upBars', 'downBars'),
                'series_lines': ('serLines',)}


class _GroupLines(Created):
    """A chart group's drop lines, high-low lines, up or down bars or series lines: there while the group's Has member
    reads True, named as Excel names them, and gone with Delete (tests/fixtures/chart_type_reads.json)."""

    #: A base the five types share, not a type of its own.
    vba_type_name = ''
    #: The ChartGroup member that says whether they are there, and their name; None where reading it raises.
    flag = ''
    name: str | None = None
    #: Their look's key, which is their element in the file.
    key = ''

    def __init__(self, group: ChartGroup) -> None:
        self.group = group

    def present(self) -> bool:
        return bool(self.group.vba_get(self.flag, []))

    def shown(self) -> object:
        if not self.present():
            raise error(424, 'Object required')
        return self

    @member
    def Parent(self) -> object:
        return self.group

    @member
    def Application(self) -> object:
        return self.group.chart.sheet.book.application

    @member
    def Name(self) -> object:
        if not self.present():
            raise error(424, 'Object required')
        if self.name is None:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        return self.name

    @method
    def Delete(self) -> object:
        if not self.present():
            raise error(424, 'Object required')
        if self.name is None:
            raise error(-2147467259, 'Automation error\nUnspecified error')
        self.group.vba_set(self.flag, False)
        return True

    # Their look, gone with them (tests/fixtures/chart_group_looks.json).

    def look(self) -> FillLine:
        chart = self.group.chart
        chart.admit_read()
        if chart.state.original:
            raise VBAUnsupportedError('The look of the lines or bars of an imported chart is not implemented')
        self.shown()
        return chart.state.element_looks.get(self.key) or FillLine()

    def change(self) -> FillLine:
        self.look()
        chart = self.group.chart
        chart.changed()
        return chart.state.element_looks.setdefault(self.key, FillLine())

    @member
    def Format(self) -> object:
        from pyopenvba.apps.excel._chart_line_format import BarChartFormat, LineChartFormat

        self.look()
        if self.key in ('upBars', 'downBars'):
            return BarChartFormat(self, self, self.group.chart, self.key)
        return LineChartFormat(self, self)

    @member
    def Border(self) -> object:
        from pyopenvba.apps.excel._chart_borders import GROUP_LINES, ChartBorder

        self.look()
        return ChartBorder(self, GROUP_LINES, self, self.group.chart.sheet.book.stylesheet.colors)

    @method
    def Select(self) -> object:
        self.shown()
        return select_element(self.group.chart, self, True)


class _UnansweredSelect(_GroupLines):
    """Lines whose Select answers nothing, and which do not have GetProperty or SetProperty at all."""

    vba_type_name = ''

    @method
    def Select(self) -> object:
        super().Select()
        return EMPTY


class DropLines(_UnansweredSelect):
    vba_type_name = 'DropLines'
    flag, name, key = 'HasDropLines', 'Dropline 1', 'dropLines'


class HiLoLines(_UnansweredSelect):
    vba_type_name = 'HiLoLines'
    flag, name, key = 'HasHiLoLines', 'HiLoline 1', 'hiLowLines'


class _Properties(ElementProperties):
    vba_type_name = ''
    get_error, set_error = INVALID_CALL, UNSPECIFIED


class UpBars(_Properties, _GroupLines):
    vba_type_name = 'UpBars'
    flag, name, key = 'HasUpDownBars', 'UpBar 1', 'upBars'


class DownBars(_Properties, _GroupLines):
    vba_type_name = 'DownBars'
    flag, name, key = 'HasUpDownBars', None, 'downBars'

    @member
    def Border(self) -> object:
        self.look()
        raise error(1004, 'Unable to get the Border property of the DownBars class')

    @method
    def Select(self) -> object:
        self.shown()
        return select_element(self.group.chart, self, False)


class SeriesLines(_Properties, _GroupLines):
    vba_type_name = 'SeriesLines'
    flag, name, key = 'HasSeriesLines', 'Seriesline 1', 'serLines'


class _NoLook:
    """A look that can be read as new and not yet be changed."""

    def look(self) -> FillLine:
        return FillLine()

    def change(self) -> FillLine:
        raise VBAUnsupportedError('Changing the look of leader lines is not implemented')


class LeaderLines(Created):
    """A series' leader lines (tests/fixtures/chart_members_more.json)."""

    vba_type_name = 'LeaderLines'

    def __init__(self, series: Series) -> None:
        self.series = series

    @member
    def Parent(self) -> object:
        return self.series

    @member
    def Application(self) -> object:
        return self.series.chart.sheet.book.application

    @member
    def Border(self) -> object:
        from pyopenvba.apps.excel._chart_borders import GROUP_LINES, ChartBorder

        return ChartBorder(_NoLook(), GROUP_LINES, self, self.series.chart.sheet.book.stylesheet.colors)

    @member
    def Format(self) -> object:
        from pyopenvba.apps.excel._chart_line_format import LineChartFormat

        return LineChartFormat(self, _NoLook())

    @method
    def Delete(self) -> object:
        self.series.vba_set('HasLeaderLines', False)
        return EMPTY


class ChartCategory(Created):
    """One category of a chart group (tests/fixtures/chart_members_more.json)."""

    vba_type_name = 'ChartCategory'

    def __init__(self, group: ChartGroup, name: str) -> None:
        self.group, self.name = group, name

    @member
    def Parent(self) -> object:
        return NOTHING

    @member
    def Application(self) -> object:
        return self.group.chart.sheet.book.application

    @member
    def Name(self) -> object:
        return self.name

    @setter('Name')
    def _name(self, value: object) -> None:
        del value
        raise error(450, 'Wrong number of arguments or invalid property assignment')

    @member
    def IsFiltered(self) -> object:
        return False

    @setter('IsFiltered')
    def _is_filtered(self, value: object) -> None:
        if to_bool(value):
            raise VBAUnsupportedError('Filtering a chart category is not implemented')


class CategoryCollection(ItemDefault, Created):
    """A chart group's categories, those of its first series. They cannot be walked with For Each: Excel raises 438."""

    vba_type_name = 'CategoryCollection'

    def __init__(self, group: ChartGroup) -> None:
        self.group = group

    def vba_iterate(self) -> Iterator[object]:
        self.names()
        raise error(438, "Object doesn't support this property or method")

    def names(self) -> list[str]:
        chart = self.group.chart
        chart.admit_read()
        if not chart.state.series:
            return []
        series = Series(chart, chart.state.series[0])
        if CHART_TYPES.get(shown_type(chart.state)) == 'scatter':
            return [str(place + 1) for place in range(len(series.values()))]
        names = [to_text(name) for name in series.categories()]
        # As many as the series has values, whatever the names given.
        return names + [''] * (len(series.values()) - len(names))

    @member
    def Parent(self) -> object:
        self.group.chart.admit_read()
        raise error(1004, 'Unable to get the Parent property of the CategoryCollection class')

    @member
    def Application(self) -> object:
        return self.group.chart.sheet.book.application

    @member
    def Count(self) -> object:
        return VBAInt(len(self.names()), 'Long')

    @method
    def Item(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        return ChartCategory(self.group, self._named(Index))

    def _named(self, index: object) -> str:
        position = int(to_integer(index, 'Long'))
        names = self.names()
        if not 1 <= position <= len(names):
            raise error(1004, 'Unable to get the Item property of the CategoryCollection class')
        return names[position - 1]


class SeriesCollection(VBACollection, ItemDefault, Created):
    vba_type_name = 'SeriesCollection'

    def __init__(self, chart: Chart) -> None:
        self.chart = chart

    def vba_items(self) -> list[object]:
        self.chart.admit_read()
        return [Series(self.chart, entry) for entry in self.chart.state.series]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                if to_text(cast(Series, item).vba_get('Name', [])).casefold() == index.casefold():
                    return item
            raise error(1004, 'Invalid Parameter')
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
    def Add(self, Source: object = MISSING, Rowcol: object = MISSING, SeriesLabels: object = MISSING,  # noqa: N803
            CategoryLabels: object = MISSING, Replace: object = MISSING) -> object:  # noqa: N803 - native named args
        from pyopenvba.apps.excel._series_add import add

        return add(self.chart, Source, Rowcol, SeriesLabels, CategoryLabels)

    @method
    def Extend(self, Source: object = MISSING, Rowcol: object = MISSING,  # noqa: N803 - native named args
               CategoryLabels: object = MISSING) -> object:  # noqa: N803
        from pyopenvba.apps.excel._series_add import extend

        return extend(self.chart, Source, Rowcol, CategoryLabels)

    @method
    def NewSeries(self) -> object:
        self.chart.admit_change()
        self.chart.nothing_filtered('NewSeries')
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
        _labelled(self.chart.state)
        return Series(self.chart, entry)


class FullSeriesCollection(SeriesCollection):
    """Every series, including those filtered out of SeriesCollection, which the model does not do
    (tests/fixtures/chart_full_series.json). It adds none."""

    vba_type_name = 'FullSeriesCollection'

    def vba_items(self) -> list[object]:
        self.chart.admit_read()
        state = self.chart.state
        return [Series(self.chart, entry) for entry in state.full_series or state.series
                if not any(entry is hidden for hidden in state.out_of_sight)]

    @method
    def NewSeries(self) -> object:
        raise error(438, "Object doesn't support this property or method")

    @method
    def Add(self, *_: object) -> object:
        raise error(438, "Object doesn't support this property or method")

    @method
    def Extend(self, *_: object) -> object:
        raise error(438, "Object doesn't support this property or method")


class Series(PictureMembers, Created):
    @method
    def GetProperty(self, Name: object) -> object:  # noqa: N803 - native named argument
        """Its Name by that name; any other name it does not know raises (tests/fixtures/chart_group_looks.json)."""
        wanted = to_text(Name)
        if wanted.lower() != 'name':
            raise error(NOT_A_PROPERTY, 'Automation error')
        return self.vba_get('Name', [])

    @method
    def SetProperty(self, Name: object, Value: object) -> object:  # noqa: N803 - native named argument
        wanted = to_text(Name)
        if wanted.lower() != 'name':
            raise error(NOT_A_PROPERTY, 'Automation error')
        self.vba_set('Name', Value)
        return EMPTY

    vba_type_name = 'Series'

    def __init__(self, chart: Chart, entry: SeriesState) -> None:
        self.chart, self.entry = chart, entry

    def position(self) -> int:
        self.chart.admit_read()
        state = self.chart.state
        for index, entry in enumerate(state.series):
            if entry is self.entry:
                return index + 1
        for index, entry in enumerate(state.full_series):
            if entry is self.entry:
                return index + 1
        raise error(-2147467261, 'This object is no longer valid.')

    def full_place(self) -> int:
        state = self.chart.state
        if state.full_series:
            self.position()
            full = [entry for entry in state.full_series
                    if entry is self.entry or not any(entry is hidden for hidden in state.out_of_sight)]
            return next(index + 1 for index, entry in enumerate(full) if entry is self.entry)
        return self.position()

    def filtered(self) -> bool:
        state = self.chart.state
        return any(entry is self.entry for entry in state.full_series) and not any(
            entry is self.entry for entry in [*state.series, *state.out_of_sight])

    def vba_call_by_name(self, name: str, args: Sequence[object], *, assignment: bool = False,
                         by_ref: bool = False) -> object:
        self.chart.admit_read()
        state = self.chart.state
        if (name.lower() in ('name', 'plotorder', 'formula', 'values', 'xvalues')
                and not any(entry is self.entry for entry in [*state.series, *state.full_series])):
            raise error(440, 'Automation error')
        return super().vba_call_by_name(name, args, assignment=assignment, by_ref=by_ref)

    @member
    def Parent(self) -> object:
        self.position()
        return ChartGroup(self.chart)

    @member
    def Application(self) -> object:
        return self.chart.sheet.book.application

    @member
    def AxisGroup(self) -> object:
        # Every series the model plots is on the primary axes (tests/fixtures/chart_series_reads.json).
        self.position()
        return VBAInt(1, 'Long')

    @member
    def ChartType(self) -> object:
        """The chart's own type, which every series of a chart of one kind shares
        (tests/fixtures/chart_series_reads.json)."""
        self.position()
        return VBAInt(shown_type(self.chart.state), 'Long')

    @member
    def Type(self) -> object:
        """The series' kind as the old gallery constants number it: xlColumn 3 for any column chart, xlLine 4 with
        markers or without (tests/fixtures/chart_series_reads.json)."""
        self.position()
        return VBAInt(SERIES_KINDS[CHART_TYPES[shown_type(self.chart.state)]], 'Long')

    @member
    def PlotColorIndex(self) -> object:
        """Which of the theme's colours an automatic series takes, from 1 (tests/fixtures/chart_series_reads.json)."""
        position = self.position()
        return VBAInt(self.entry.series_id + 1 if self.entry.series_id >= 0 else position, 'Long')

    @member
    def IsFiltered(self) -> object:
        self.position()
        return self.filtered()

    @setter('IsFiltered')
    def _is_filtered(self, value: object) -> None:
        """Moves it between the series drawn and those only FullSeriesCollection has
        (tests/fixtures/series_filter.json)."""
        self.position()
        wanted = to_bool(value)
        if wanted == self.filtered():
            return
        self.chart.admit_change()
        self.chart.changed()
        state = self.chart.state
        if not state.full_series:
            state.full_series = list(state.series)
        hidden = [entry for entry in state.full_series
                  if not any(entry is drawn for drawn in state.series) and entry is not self.entry]
        if wanted:
            hidden.append(self.entry)
        state.series[:] = [entry for entry in state.full_series if not any(entry is one for one in hidden)]
        if not hidden:
            state.full_series = []

    @member
    def BubbleSizes(self) -> object:
        # Only a bubble chart's series has sizes, and the model makes none (tests/fixtures/missing_reads.json).
        self.position()
        raise error(5, 'Invalid procedure call or argument')

    @member
    def HasErrorBars(self) -> object:
        return has_error_bars(self)

    @setter('HasErrorBars')
    def _has_error_bars(self, value: object) -> None:
        set_has_error_bars(self, to_bool(value))

    @method
    def ErrorBar(self, Direction: object = MISSING, Include: object = MISSING, Type: object = MISSING,
                 Amount: object = MISSING, MinusValues: object = MISSING) -> object:
        error_bar(self, Direction, Include, Type, Amount, MinusValues)
        return EMPTY

    @member
    def ErrorBars(self) -> object:
        self.position()
        return ErrorBars(self)

    @method
    def Trendlines(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        self.position()
        lines = Trendlines(self)
        return lines if Index is MISSING else lines.vba_lookup(Index, lines.vba_items())

    @member
    def Has3DEffect(self) -> object:
        self.position()
        return False

    @member
    def Shadow(self) -> object:
        self.position()
        return bool(self.entry.shadow)

    @setter('Shadow')
    def _shadow(self, value: object) -> None:
        self.position()
        wanted = to_bool(value)
        self.chart.admit_change()
        self.chart.changed()
        self.entry.shadow = wanted

    @member
    def LeaderLines(self) -> object:
        """There unless labels on the series turned them off (tests/fixtures/chart_members_more.json)."""
        self.position()
        if self.entry.data_labels and not self.entry.leader_lines:
            raise error(5, 'Invalid procedure call or argument')
        return LeaderLines(self)

    # What only the newer kinds of chart have (tests/fixtures/chart_members_more.json).

    @member
    def GeoMappingLevel(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('GeoMappingLevel')
    def _geomappinglevel(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def GeoProjectionType(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('GeoProjectionType')
    def _geoprojectiontype(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def ParentDataLabelOption(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('ParentDataLabelOption')
    def _parentdatalabeloption(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def QuartileCalculationInclusiveMedian(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('QuartileCalculationInclusiveMedian')
    def _quartilecalculationinclusivemedian(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def RegionLabelOption(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('RegionLabelOption')
    def _regionlabeloption(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def SeriesColorGradientStyle(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('SeriesColorGradientStyle')
    def _seriescolorgradientstyle(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def ValueSortOrder(self) -> object:
        raise error(INVALID_CALL, 'Automation error')

    @setter('ValueSortOrder')
    def _valuesortorder(self, value: object) -> None:
        del value
        raise error(INVALID_CALL, 'Automation error')

    @member
    def SeriesColorMaxGradientStop(self) -> object:
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def SeriesColorMidGradientStop(self) -> object:
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def SeriesColorMinGradientStop(self) -> object:
        raise error(UNSPECIFIED, 'Automation error')

    @member
    def BarShape(self) -> object:
        # xlBox (tests/fixtures/chart_series_reads.json).
        self.position()
        return VBAInt(0, 'Long')

    @member
    def InvertColorIndex(self) -> object:
        self.position()
        return VBAInt(2, 'Long')

    @member
    def InvertColor(self) -> object:
        # White, the fill InvertIfNegative gives a negative bar (tests/fixtures/chart_series_reads.json).
        self.position()
        return VBAInt(16777215, 'Long')

    @member
    def PlotOrder(self) -> object:
        place = self.position()
        return VBAInt(0 if self.filtered() else place, 'Long')

    @setter('PlotOrder')
    def _plot_order(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        self.chart.nothing_filtered('Setting PlotOrder')
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
        self.position()
        self.chart.changed()
        state = self.chart.state
        state.series[:] = [entry for entry in state.series if entry is not self.entry]
        state.full_series = [entry for entry in state.full_series if entry is not self.entry]
        if len(state.full_series) == len(state.series):
            state.full_series = []
        return True

    @member
    def Name(self) -> object:
        self.position()
        if self.entry.name_reference and not self.chart.state.names_lost():
            values = referenced_values(self.chart.sheet, self.entry.name_reference)
            if len(values) > 1:
                # A name of several cells reads their texts with a space between (tests/fixtures/
                # chart_multilevel_categories.json).
                return ' '.join(text for row in shown_cells(self.chart.sheet, self.entry.name_reference)
                                for text in row if text)
            return to_text(values[0]) if values else ''
        identity = self.entry.series_id + 1 if self.entry.series_id >= 0 else self.position()
        return self.entry.name if self.entry.name is not None and not self.entry.name_reference else f'Series{identity}'

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
    def Values(self, Index: object = MISSING) -> object:  # noqa: N803 - it takes none, as Excel answers
        if Index is not MISSING:
            raise error(451, 'Property let procedure not defined and property get procedure did not return an object')
        values = [0.0 if isinstance(value, str) else
                  float(value) if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) else value
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
        gone = categories_gone(self.chart.sheet, self.chart.state, self.entry)
        if gone:
            # They read as the first series' that still has its own.
            other = next((entry for entry in self.chart.state.series if entry.category_reference
                          and not categories_gone(self.chart.sheet, self.chart.state, entry)), None)
            if other is not None:
                return Series(self.chart, other).XValues()
        if (self.entry.category_reference and not gone
                and category_levels(self.chart.sheet, self.entry) is None):
            # Categories with text among them read every cell as it shows, a blank as ""; numbers with a blank among
            # them cannot be read at all (tests/fixtures/chart_multilevel_categories.json).
            if any(isinstance(value, str) for value in values):
                if CHART_TYPES.get(shown_type(self.chart.state)) == 'scatter' and not self.chart.state.out_of_sight:
                    # A scatter chart has no use for text: it reads their places
                    # (tests/fixtures/chart_hidden_cells.json).
                    values = [float(index + 1) for index in range(len(values))]
                else:
                    every = [text for row in shown_cells(self.chart.sheet, self.entry.category_reference)
                             for text in row]
                    kept = visible_places(self.chart.sheet, self.entry.category_reference,
                                          self.chart.state.visible_only)
                    values = [every[place] for place in kept]
            elif any(value is EMPTY for value in values):
                raise error(-2147467259, 'Automation error\nUnspecified error ')
        values = [float(value) if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) else value
                  for value in values]
        if lost(self.entry.value_reference):
            # Excel reads them as far as the one point left.
            values = values[:1]
        return VBAArray([(1, len(values))], items=values)

    @setter('XValues', objects=True)
    def _categories(self, value: object) -> None:
        self.position()
        self.chart.admit_change()
        reference, values = series_input(value, sheet=self.chart.sheet)
        self.position()
        self.chart.changed()
        self.entry.category_reference, self.entry.categories = reference, values

    def has_labels(self) -> bool:
        """Whether it has labels at all: its own, or a point's that was given one."""
        return self.entry.data_labels or any(point.label is not None for point in self.entry.points.values())

    def values(self) -> list[object]:
        self.position()
        if not self.entry.value_reference:
            return self.entry.values
        if lost(self.entry.value_reference):
            return [1.0]
        return referenced_values(self.chart.sheet, self.entry.value_reference, visible=self.chart.state.visible_only)

    def point_key(self, shown: int) -> int:
        """Where a point stands among the series' cells, by its place among those shown: a point's look and label stay
        with its cell (tests/fixtures/chart_hidden_cells.json)."""
        reference = self.entry.value_reference
        if not self.chart.state.visible_only or not reference or lost(reference):
            return shown
        places = visible_places(self.chart.sheet, reference, True)
        return places[shown] if shown < len(places) else shown

    def categories(self) -> list[object]:
        self.position()
        if categories_gone(self.chart.sheet, self.chart.state, self.entry):
            return ['' for _ in self.values()]
        if self.entry.category_reference:
            levels = category_levels(self.chart.sheet, self.entry)
            if levels is not None:
                return list(joined_categories(levels))
            return referenced_values(self.chart.sheet, self.entry.category_reference,
                                     visible=self.chart.state.visible_only)
        return self.entry.categories or [float(index + 1) for index in range(len(self.values()))]

    @member
    def Formula(self) -> object:
        # Its number among them all, filtered or not.
        order = self.full_place()
        if lost(self.entry.name_reference) or lost(self.entry.value_reference):
            raise VBARuntimeError(2023, 'Application-defined or object-defined error')
        name = self.entry.name_reference or (quoted(self.entry.name) if self.entry.name is not None else '')
        categories = ('' if categories_gone(self.chart.sheet, self.chart.state, self.entry)
                      else self.entry.category_reference or literal_array(self.entry.categories))
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
        if CHART_TYPES.get(shown_type(self.chart.state)) not in _DRAWS_MARKERS:
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
            # It reaches the points set apart (tests/fixtures/series_points.json).
            for point in self.entry.points.values():
                point.invert = wanted

    @member
    def HasDataLabels(self) -> object:
        # Labels that show nothing read as none (tests/fixtures/data_labels_object.json).
        self.position()
        return self.entry.data_labels and self.entry.label_look.shown()

    @setter('HasDataLabels')
    def _has_data_labels(self, value: object) -> None:
        self.position()
        wanted = to_bool(value)
        if any(point.label is not None for point in self.entry.points.values()):
            raise VBAUnsupportedError('Changing the data labels of a series with point labels is not implemented')
        if self.entry.label_look.font is not None:
            raise VBAUnsupportedError('Changing data labels that have a font is not implemented')
        self.chart.changed()
        if self.entry.data_labels and not wanted:
            self.entry.label_generation += 1
        if wanted:
            # Labels turned on show the value alone, however they were before.
            self.entry.label_look = LabelLook()
        self.entry.data_labels = wanted
        # Excel turns the series' leader lines off whichever way the labels go.
        self.entry.leader_lines = False

    @member
    def Format(self) -> object:
        from pyopenvba.apps.excel._chart_format import ChartFormat

        self.chart.admit_read()
        state = self.chart.state
        if not any(entry is self.entry for entry in [*state.series, *state.full_series]):
            # A series taken away answers this one differently from its other members.
            raise error(-2147467259, 'Automation error\nUnspecified error')
        return ChartFormat(self)

    @method
    def Points(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        from pyopenvba.apps.excel._chart_points import Points

        self.position()
        points = Points(self)
        return points if Index is MISSING else points.vba_get('Item', [Index])

    @method
    def ApplyDataLabels(self, Type: object = MISSING, LegendKey: object = MISSING, AutoText: object = MISSING,
                        HasLeaderLines: object = MISSING, ShowSeriesName: object = MISSING,
                        ShowCategoryName: object = MISSING, ShowValue: object = MISSING,
                        ShowPercentage: object = MISSING, ShowBubbleSize: object = MISSING,
                        Separator: object = MISSING) -> object:
        from pyopenvba.apps.excel._chart_data_labels import applied_look, apply_to_series

        del ShowBubbleSize
        look = applied_look(self.chart.state.chart_type, Type, LegendKey, AutoText, ShowSeriesName,
                            ShowCategoryName, ShowValue, ShowPercentage, Separator)
        apply_to_series(self, look, HasLeaderLines)
        return True

    @method
    def Select(self, Replace: object = MISSING) -> object:  # noqa: N803 - native named argument
        del Replace
        return select_element(self.chart, self, self.plotted())

    def plotted(self) -> bool:
        """Whether its chart draws it: a pie draws only its first series."""
        place = self.position()
        if self.filtered():
            return False
        return place == 1 or CHART_TYPES.get(shown_type(self.chart.state)) != 'pie'

    @method
    def DataLabels(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        from pyopenvba.apps.excel._chart_data_labels import DataLabels

        self.position()
        if not self.has_labels():
            raise error(-2147467259, 'Automation error\nUnspecified error')
        labels = DataLabels(self, self.entry.label_generation)
        return labels if Index is MISSING else labels.vba_get('Item', [Index])

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
        # It takes over from the points' own (tests/fixtures/series_points.json).
        for place, point in list(self.entry.points.items()):
            point.explosion = None
            if point.look is None and point.invert is None and point.label is None:
                del self.entry.points[place]

    @member
    def Border(self) -> object:
        from pyopenvba.apps.excel._chart_borders import ChartBorder, family_of_series
        from pyopenvba.apps.excel._chart_format import ChartFormat, SeriesLine

        self.chart.admit_read()
        line = SeriesLine(ChartFormat(self))
        return ChartBorder(line, family_of_series(line.kind(), line.accent()), self,
                           self.chart.sheet.book.stylesheet.colors)

    @method
    def ClearFormats(self) -> object:
        """Takes away the look set on the series and its points; their labels, its trendlines and what else it has
        stay (tests/fixtures/chart_borders.json)."""
        self.position()
        kind = CHART_TYPES.get(shown_type(self.chart.state), '')
        if kind == 'line' and self.entry.format_look is not None:
            raise VBAUnsupportedError('ClearFormats on a line series with a look is not implemented')
        if kind == 'area' and self.entry.points:
            raise VBAUnsupportedError('ClearFormats on an area series with a point set apart is not implemented')
        self.chart.changed()
        self.entry.format_look = None
        if kind in ('line', 'lineMarkers', 'scatter'):
            # They all come back automatic.
            self.entry.marker, self.entry.marker_size = MARKER_AUTOMATIC, None
            self.entry.marker_fore, self.entry.marker_back, self.entry.marker_line = None, None, None
            self.entry.smooth = False
            self.entry.scatter_line = kind == 'scatter'
        for place, point in list(self.entry.points.items()):
            point.look = None
            if point.invert is None and point.explosion is None and point.label is None and not point.has_marker():
                del self.entry.points[place]
        return True

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
        self.position()
        if CHART_TYPES.get(shown_type(self.chart.state)) in _DRAWS_MARKERS:
            return VBAInt(self.lines().marker, 'Long')
        return VBAInt(self.entry.marker, 'Long')

    @setter('MarkerStyle')
    def _marker_style(self, value: object) -> None:
        entry = self.lines()
        wanted = int(to_integer(value, 'Long'))
        wanted = MARKER_ALIASES.get(wanted, wanted)
        if wanted not in MARKER_SYMBOLS:
            raise error(1004, 'Unable to set the MarkerStyle property of the Series class')
        look = entry.format_look
        if look is not None and (look.line_rgb is None or look != FillLine(line_rgb=look.line_rgb)):
            # Excel copies the line's look into the marker, measured only for a colour (tests/fixtures/series_format.json).
            raise VBAUnsupportedError('A marker on a series with a look other than a line colour is not implemented')
        self.chart.changed()
        entry.marker = wanted
        if look is not None:
            entry.marker_line = look.line_rgb
        self.taken_from_points('marker')

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
        if wanted is not None and field_name == 'marker_back' and entry.marker == MARKER_NONE:
            # A series drawn without markers takes automatic ones once given a background colour.
            entry.marker = MARKER_AUTOMATIC
        self.taken_from_points(field_name)

    def taken_from_points(self, field_name: str) -> None:
        """What is set of a series' marker it sets of its points too, and a marker left the same as the series' goes."""
        entry = self.entry
        for point in entry.points.values():
            setattr(point, field_name, None)
            if point.marker == entry.marker:
                point.marker = None

    def _read_colour(self, field_name: str) -> object:
        found = self._colour(field_name)
        look = self.entry.format_look
        if found is None and field_name == 'marker_fore' and look is not None:
            # Its line's colour is its marker's too (tests/fixtures/point_line_format.json).
            found = look.line_rgb
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
        self.position()
        entry = self.entry
        drawn = CHART_TYPES.get(shown_type(self.chart.state)) in _DRAWS_MARKERS
        style = self.chart.state.chart_style
        from pyopenvba.apps.excel._chart_styled import is_styled, marker_size

        size = entry.marker_size or (marker_size(style) if is_styled(style) else
                                     9 if drawn and self.lines().marker == MARKER_AUTOMATIC else 5)
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
        self.taken_from_points('marker_size')


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

def visible_places(sheet: Worksheet, formula: str, visible: bool) -> list[int]:
    """Which of a reference's cells, by their place in it, count: those not hidden, or with ``visible`` off all."""
    places: list[int] = []
    place = 0
    for area in parse_reference(formula):
        owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                if not (visible and (owner.dims.row_hidden(row) or owner.dims.column_hidden(column))):
                    places.append(place)
                place += 1
    return places


def referenced_values(sheet: Worksheet, formula: str, *, visible: bool = False) -> list[object]:
    """What a reference holds, ``visible`` keeping only the cells not hidden."""
    from pyopenvba.apps.excel._model import Range

    if lost(formula):
        return ['#REF!']
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
            if visible and owner.dims.row_hidden(row):
                continue
            for column in range(area.left, area.right + 1):
                if not (visible and owner.dims.column_hidden(column)):
                    result.append(Range(owner, [Area(row, column, row, column)]).Value2())
    return result
