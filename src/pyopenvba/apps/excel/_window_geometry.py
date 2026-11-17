"""Where a window's grid ends and what it shows: Window.UsableWidth, UsableHeight, VisibleRange, LargeScroll and
ScrollIntoView, as Excel lays them out in screen pixels (scripts/measure_window_geometry.py,
tests/fixtures/window_geometry_*.json).

Excel lays the grid out in whole pixels of the display. A row or column at a zoom is its size at 100% times the zoom,
rounded half up to a pixel, and the headings scale the same way. UsableWidth and UsableHeight are the window's size less
its frame and bars, never below 0. The rows the grid shows are those that start inside it, a row cut off at the bottom
included and at least one, and the columns the same. With frozen panes, VisibleRange is the scrolling pane's, from its
first cell to the window's last. LargeScroll moves a page at a time: down, the last row shown becomes the first; up, the
first becomes the last; a page is a row or column at least, when
the grid shows less than one whole; kept on the sheet. ScrollIntoView takes a rectangle in screen pixels of the zoomed sheet,
measured from A1: with Start True or left out, the row and column holding its top left come first; with Start False,
the first column is the least that shows all of it in whole columns. Which first row Start False takes is not known:
of eight measured cases, five took the least that shows all of it in whole rows and three the least that shows it with a
row cut off at the bottom, and nothing in the window or the sheet told them apart (tests/fixtures/window_geometry1.json,
window_geometry2.json); the rule below is the first.

All of it depends on the display's DPI, through the pixels a point makes and the sizes of Excel's own frame, bars and
headings. DISPLAYS holds those for each display they were measured on, by the StandardHeight its rows have. The model
lays sheets out at 96 DPI (StandardHeight 15), which is not measured yet, so the members report themselves unsupported
until it is.
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS


@dataclass(frozen=True)
class Display:
    """One display's measures: pixels per point, Excel's headings at 100% in pixels, and in points what the window's
    frame, the vertical scroll bar, the row of sheet tabs and the horizontal scroll bar, the ribbon and its frame, the
    formula bar and the status bar take from the window's size."""

    pixels: float
    row_heading: int
    column_heading: int
    frame_width: float
    vertical_bar: float
    frame_height: float
    tab_row: float
    formula_bar: float
    status_bar: float
    #: The rows the row headings were measured for: more digits make them wider, which was not measured.
    heading_rows: int
    #: The zooms the rules held at. At 100% they matched every case, row numbers of one to three digits; at other
    #: zooms the headings are not their size at 100% times the zoom (wider at 50% and 60%, and wider again for row
    #: numbers of three digits), which one look does not pin (tests/fixtures/window_geometry1.json).
    zooms: frozenset[int]


#: By StandardHeight: 14.5 at 144 DPI (tests/fixtures/window_geometry_a.json and _b).
DISPLAYS: dict[float, Display] = {
    14.5: Display(pixels=2.0, row_heading=40, column_heading=29, frame_width=11.0, vertical_bar=15.5,
                  frame_height=140.0, tab_row=24.0, formula_bar=28.5, status_bar=16.5, heading_rows=999,
                  zooms=frozenset({100})),
}


@dataclass(frozen=True)
class Layout:
    """A window and the sheet it shows: its size in points, its settings, where it is scrolled to, and each row's and
    column's size in pixels at 100% on the display, 0 for one hidden."""

    width: float
    height: float
    zoom: int = 100
    headings: bool = True
    vertical_bar: bool = True
    horizontal_bar: bool = True
    tabs: bool = True
    formula_bar: bool = True
    status_bar: bool = True
    scroll: tuple[int, int] = (1, 1)
    #: Frozen rows and columns above and left of the scrolling pane.
    frozen: tuple[int, int] = (0, 0)
    row_pixels: Callable[[int], int] = lambda _row: 29
    column_pixels: Callable[[int], int] = lambda _column: 96


def zoomed(pixels: int, zoom: int) -> int:
    """A size at 100% in pixels, at a zoom: rounded half up to a whole pixel."""
    return int(pixels * zoom / 100 + 0.5)


def usable(display: Display, layout: Layout) -> tuple[float, float]:
    """UsableWidth and UsableHeight in points."""
    width = layout.width - display.frame_width - (display.vertical_bar if layout.vertical_bar else 0)
    height = (layout.height - display.frame_height - (display.tab_row if layout.tabs or layout.horizontal_bar else 0)
              - (display.formula_bar if layout.formula_bar else 0) - (display.status_bar if layout.status_bar else 0))
    return max(width, 0.0), max(height, 0.0)


class Grid:
    """The cells a window's grid shows, in pixels."""

    def __init__(self, display: Display, layout: Layout) -> None:
        self.display, self.layout = display, layout
        width, height = usable(display, layout)
        headings = layout.headings
        across = width * display.pixels - (zoomed(display.row_heading, layout.zoom) if headings else 0)
        down = height * display.pixels - (zoomed(display.column_heading, layout.zoom) if headings else 0)
        self.space = (down - sum(self.size(0, row) for row in range(1, layout.frozen[0] + 1)),
                      across - sum(self.size(1, column) for column in range(1, layout.frozen[1] + 1)))

    def size(self, axis: int, index: int) -> int:
        layout = self.layout
        return zoomed(layout.row_pixels(index) if axis == 0 else layout.column_pixels(index), layout.zoom)

    def first(self, axis: int) -> int:
        """The first row or column a scrolling pane may show."""
        return self.layout.frozen[axis] + 1

    def last(self, axis: int, start: int) -> int:
        """The last row or column shown from ``start``: each that starts inside the grid, at least one."""
        limit = MAX_ROWS if axis == 0 else MAX_COLUMNS
        index, used = start, 0
        while index < limit:
            used += self.size(axis, index)
            if used >= self.space[axis]:
                break
            index += 1
        return index

    def whole(self, axis: int, start: int) -> int:
        """The pixel, from A1, where the last row or column wholly shown from ``start`` ends."""
        limit = MAX_ROWS if axis == 0 else MAX_COLUMNS
        offset, used, index = self.offset(axis, start), 0, start
        while index <= limit and used + self.size(axis, index) <= self.space[axis]:
            used += self.size(axis, index)
            index += 1
        return offset + used

    def offset(self, axis: int, index: int) -> int:
        """Where a row or column starts, in pixels from A1."""
        return sum(self.size(axis, one) for one in range(1, index))

    def holding(self, axis: int, pixel: int) -> int:
        """The row or column holding a pixel from A1; the first for one before it."""
        limit = MAX_ROWS if axis == 0 else MAX_COLUMNS
        index, used = 1, 0
        while index < limit:
            used += self.size(axis, index)
            if pixel < used:
                break
            index += 1
        return index


def visible(display: Display, layout: Layout) -> tuple[int, int, int, int]:
    """VisibleRange as top, left, bottom, right."""
    grid = Grid(display, layout)
    top, left = layout.scroll
    return top, left, grid.last(0, top), grid.last(1, left)


def large_scroll(display: Display, layout: Layout, down: int, right: int) -> tuple[int, int]:
    """ScrollRow and ScrollColumn after LargeScroll moves ``down`` pages (up when below 0) and ``right`` pages."""
    grid = Grid(display, layout)
    places = list(layout.scroll)
    for axis, pages in ((0, down), (1, right)):
        limit = MAX_ROWS if axis == 0 else MAX_COLUMNS
        for _ in range(abs(pages)):
            here = places[axis]
            if pages > 0:
                places[axis] = min(max(grid.last(axis, here), here + 1), limit)
            else:
                start = here
                while start > grid.first(axis) and grid.last(axis, start - 1) >= here:
                    start -= 1
                places[axis] = max(min(start, here - 1), grid.first(axis))
    return places[0], places[1]


def scroll_into_view(display: Display, layout: Layout, rectangle: tuple[int, int, int, int],
                     start: bool) -> tuple[int, int]:
    """ScrollRow and ScrollColumn after ScrollIntoView of (left, top, width, height) in pixels from A1."""
    grid = Grid(display, layout)
    left, top, width, height = rectangle
    places = [0, 0]
    for axis, (near, extent) in ((0, (top, height)), (1, (left, width))):
        if start:
            place = grid.holding(axis, near)
        else:
            place = grid.first(axis)
            while grid.whole(axis, place) < near + extent:
                place += 1
        places[axis] = max(place, grid.first(axis))
    return places[0], places[1]
