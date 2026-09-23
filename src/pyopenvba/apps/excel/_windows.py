"""Windows: Application.ActiveWindow, Windows and Goto, and the view of each sheet a file keeps.

Measured in live Excel with its window shown and sized 700 by 500 points
(scripts/measure_windows.py, tests/fixtures/windows/). Hidden, Excel lays
out no panes, and a freeze it saves there has no split in it.

* Each sheet has its own view, which the window shows while the sheet is
  active: zoom, gridlines, headings, zeros and formulas, the view it is
  in, where it is scrolled to, its panes and each pane's selection. Each
  view keeps its own zoom: the page break preview opens at 60, the normal
  view's zoom is kept apart once another view is chosen, and 85.6 is 85.
* FreezePanes freezes the rows above and the columns left of the active
  cell, which has to be in view. SplitRow and SplitColumn count them;
  ScrollRow and ScrollColumn scroll the pane under or right of them, never
  above or left of it, and scroll the whole window along a way nothing is
  frozen. SplitVertical and SplitHorizontal are the frozen rows' and
  columns' pixels at the zoom, each rounded to a whole pixel, in points.
* The macro recorder's Freeze Top Row sets SplitColumn and SplitRow, then
  FreezePanes: a split frozen in place, which the file keeps as
  frozenSplit. A split that is not frozen is kept in twips of the rows,
  columns and headings it runs past, which depend on the window: the
  model answers for one and does not save one.
* Select sets the active pane's selection. Activate moves the active cell
  when the range's first cell is in the selection, else selects the
  range; on a sheet that is not the active one, both are error 1004. Goto
  activates the range's sheet and selects the range, and with Scroll
  puts its first cell at the top left of the window, or of the pane that
  scrolls.
* Windows lists the windows front to back. A workbook added or opened
  comes to the front; activating one brings it forward only while the
  screen updates, and closing the active workbook activates the window
  behind it.
* DisplayWorkbookTabs, TabRatio and the scroll bars belong to the
  workbook, and its workbookView keeps them with the active tab.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter, parse_area
from pyopenvba._xml import attributes, tag_attributes
from pyopenvba.apps.excel import _dimensions
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._r1c1 import to_a1
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    ERR_SUBSCRIPT_OUT_OF_RANGE,
    MISSING,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_number,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Workbook, Worksheet

#: XlWindowView by number, as a sheetView names it.
VIEWS: Final = {1: "normal", 2: "pageBreakPreview", 3: "pageLayout"}
#: What each view opens at until a zoom is chosen in it.
_OPENING_ZOOM: Final = {1: 100, 2: 60, 3: 100}
_FROZEN: Final = ("frozen", "frozenSplit")
#: The attributes of a sheetView and of a workbookView in the order the schema gives them.
_SHEET_VIEW_ORDER: Final = (
    "windowProtection", "showFormulas", "showGridLines", "showRowColHeaders", "showZeros", "rightToLeft",
    "tabSelected", "showRuler", "showOutlineSymbols", "defaultGridColor", "showWhiteSpace", "view", "topLeftCell",
    "colorId", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView", "zoomScalePageLayoutView",
    "workbookViewId",
)
_BOOK_VIEW_ORDER: Final = (
    "visibility", "minimized", "showHorizontalScroll", "showVerticalScroll", "showSheetTabs", "xWindow", "yWindow",
    "windowWidth", "windowHeight", "tabRatio", "firstSheet", "activeTab", "autoFilterDateGrouping", "xr2:uid",
)
_SHEET_VIEW = re.compile(r"<sheetView\b[^>]*?(?:/>|>.*?</sheetView>)", re.DOTALL)
_PANE = re.compile(r"<pane\b[^>]*/>")
_SELECTION = re.compile(r"<selection\b[^>]*/>")
#: What a sheetView holds besides its pane and selections, which the model keeps as it was.
_KEPT = re.compile(r"<pivotSelection\b[^>]*?(?:/>|>.*?</pivotSelection>)|<extLst\b.*?</extLst>", re.DOTALL)
_WORKBOOK_VIEW = re.compile(r"<workbookView\b[^>]*?/?>")


@dataclass
class SheetViewState:
    """A sheet's view in its workbook's window, as the sheet's first sheetView keeps it."""

    view: int = 1
    #: zoomScale, the zoom of the view the sheet is in.
    zoom: int = 100
    #: The normal view's zoom, the page break preview's and the page layout view's once each is kept apart; 0 before.
    zoom_normal: int = 0
    zoom_break: int = 0
    zoom_layout: int = 0
    gridlines: bool = True
    headings: bool = True
    zeros: bool = True
    formulas: bool = False
    top_left: tuple[int, int] = (1, 1)
    #: The pane's state: "" with no panes, "split", "frozen" or "frozenSplit".
    state: str = ""
    #: The rows and columns above and left of the split.
    rows: int = 0
    columns: int = 0
    #: The first cell the pane under and right of the split shows.
    pane_top_left: tuple[int, int] = (1, 1)
    active_pane: str = "topLeft"
    #: Each pane's selection, active cell and sqref, in the order the file lists them; the active pane's is the
    #: sheet's own selection.
    selections: dict[str, tuple[str, str]] = field(default_factory=lambda: {})
    #: A pane the model does not lay out, kept as the file had it: a split in twips, or panes frozen with nothing
    #: split, which a hidden Excel saves.
    split_xml: str = ""
    changed: bool = False


@dataclass
class BookViewState:
    """The workbook's window, as its workbookView keeps it."""

    tabs: bool = True
    tab_ratio: float = 0.6
    horizontal_scroll: bool = True
    vertical_scroll: bool = True
    #: What Caption was set to; the workbook's name until it is.
    caption: str | None = None
    #: WindowState, Width and Height once a macro sets them; before, they are the screen's.
    state: int | None = None
    width: float | None = None
    height: float | None = None
    #: The sheet whose tab the file has selected, so a save moves that only when another sheet is active.
    saved_active: Worksheet | None = None
    changed: bool = False
    window: Window | None = None


# --- the file ------------------------------------------------------------------------------------


def read_view(sheet: Worksheet, xml: str) -> None:
    """A sheet's view and its selection, from the first sheetView of its part."""
    found = _SHEET_VIEW.search(xml)
    if found is None:
        return
    element = found.group(0)
    head = tag_attributes(element)
    view = sheet.view
    view.view = next((number for number, name in VIEWS.items() if name == head.get("view", "normal")), 1)
    view.zoom = _number(head.get("zoomScale"), 100)
    view.zoom_normal = _number(head.get("zoomScaleNormal"), 0)
    view.zoom_break = _number(head.get("zoomScaleSheetLayoutView"), 0)
    view.zoom_layout = _number(head.get("zoomScalePageLayoutView"), 0)
    view.formulas = _flag(head.get("showFormulas"), False)
    view.gridlines = _flag(head.get("showGridLines"), True)
    view.headings = _flag(head.get("showRowColHeaders"), True)
    view.zeros = _flag(head.get("showZeros"), True)
    view.top_left = _cell(head.get("topLeftCell", "")) or (1, 1)
    pane = _PANE.search(element)
    if pane is not None:
        split = attributes(pane.group(0))
        view.state = split.get("state", "split")
        view.active_pane = split.get("activePane", "topLeft")
        rows, columns = int(float(split.get("ySplit", "0") or 0)), int(float(split.get("xSplit", "0") or 0))
        if view.state in _FROZEN and (rows or columns):
            view.rows, view.columns = rows, columns
            corner = _cell(split.get("topLeftCell", ""))
            view.pane_top_left = corner or (view.top_left[0] + rows, view.top_left[1] + columns)
        else:
            # A split kept in twips, or panes frozen with nothing split, as a hidden Excel saves them.
            view.split_xml = pane.group(0)
    view.selections = {}
    for selection in _SELECTION.findall(element):
        chosen = attributes(selection)
        view.selections[chosen.get("pane", "topLeft")] = (chosen.get("activeCell") or "A1",
                                                          chosen.get("sqref") or "A1")
    cell, sqref = view.selections.get(view.active_pane, ("A1", "A1"))
    if (cell, sqref) != ("A1", "A1"):
        try:
            sheet.selection_range = Range(sheet, [parse_area(part, sheet=sheet.name) for part in sqref.split()])
            sheet.active_cell_range = Range(sheet, [parse_area(cell, sheet=sheet.name)])
        except ValueError:
            sheet.selection_range = sheet.active_cell_range = None


def _number(text: str | None, default: int) -> int:
    try:
        return int(text) if text else default
    except ValueError:
        return default


def _flag(text: str | None, default: bool) -> bool:
    return default if text is None else text in ("1", "true")


def _cell(text: str) -> tuple[int, int] | None:
    try:
        area = parse_area(text)
    except ValueError:
        return None
    return area.top, area.left


def _spelled_cell(row: int, column: int) -> str:
    return f"{column_letter(column)}{row}"


def _selected(row: int, column: int) -> tuple[str, str]:
    """A pane's selection of one cell, as its active cell and its sqref."""
    cell = _spelled_cell(row, column)
    return cell, cell


def _spelled(area: Area) -> str:
    """An area as a selection's sqref spells it: whole rows and columns in full, C1:C1048576."""
    first = _spelled_cell(area.top, area.left)
    bottom, right = min(area.bottom, MAX_ROWS), min(area.right, MAX_COLUMNS)
    return first if (area.top, area.left) == (bottom, right) else f"{first}:{_spelled_cell(bottom, right)}"


def tabs_moved(book: Workbook) -> bool:
    """Whether another sheet is active than the one whose tab the file has selected."""
    return book.active_sheet is not book.view.saved_active


def with_view(sheet: Worksheet, xml: str, *, tabs: bool) -> str:
    """A worksheet part with its first sheetView as the model has the view: all of it when the view changed, else
    only its tab's selection when ``tabs`` says the active sheet moved."""
    view = sheet.view
    if not (view.changed or tabs):
        return xml
    active = sheet.book.active_sheet is sheet
    found = _SHEET_VIEW.search(xml)
    if found is None:
        at = re.search(r"<sheetFormatPr\b|<cols\b|<sheetData\b", xml)
        if not view.changed or at is None:
            return xml
        markup = "<sheetViews>" + _rewritten(sheet, '<sheetView workbookViewId="0"/>', active) + "</sheetViews>"
        return xml[:at.start()] + markup + xml[at.start():]
    element = found.group(0)
    if view.changed:
        replacement = _rewritten(sheet, element, active)
    else:
        opening, inside = _opening(element)
        replacement = _closed(_with_attribute(opening, "tabSelected", "1" if active else "", _SHEET_VIEW_ORDER),
                              inside, "sheetView")
    return xml[:found.start()] + replacement + xml[found.end():]


def _opening(element: str) -> tuple[str, str]:
    """An element's opening tag without its closing mark, and what it holds."""
    end = element.index(">")
    if element[end - 1] == "/":
        return element[:end - 1].rstrip(), ""
    return element[:end], element[end + 1:element.rindex("</")]


def _closed(opening: str, inside: str, tag: str) -> str:
    return f"{opening}>{inside}</{tag}>" if inside else f"{opening}/>"


def _with_attribute(opening: str, name: str, value: str, order: tuple[str, ...]) -> str:
    """An opening tag with one attribute set, in the place the schema gives it, or taken out when ``value`` is ""."""
    opening = re.sub(rf'\s{re.escape(name)}="[^"]*"', "", opening)
    if not value:
        return opening
    later = set(order[order.index(name) + 1:])
    for found in re.finditer(r'\s([\w:]+)="', opening):
        if found.group(1) in later:
            return f'{opening[:found.start()]} {name}="{value}"{opening[found.start():]}'
    return f'{opening} {name}="{value}"'


def _rewritten(sheet: Worksheet, element: str, active: bool) -> str:
    """A sheetView as the model has the view, keeping the attributes and children it does not model."""
    view = sheet.view
    opening, inside = _opening(element)
    zooms = (("zoomScale", "" if view.zoom == 100 else str(view.zoom)),
             ("zoomScaleNormal", str(view.zoom_normal) if view.zoom_normal else ""),
             ("zoomScaleSheetLayoutView", str(view.zoom_break) if view.zoom_break else ""),
             ("zoomScalePageLayoutView", str(view.zoom_layout) if view.zoom_layout else ""))
    for name, value in (("showFormulas", "1" if view.formulas else ""),
                        ("showGridLines", "" if view.gridlines else "0"),
                        ("showRowColHeaders", "" if view.headings else "0"),
                        ("showZeros", "" if view.zeros else "0"),
                        ("tabSelected", "1" if active else ""),
                        ("view", "" if view.view == 1 else VIEWS[view.view]),
                        ("topLeftCell", "" if view.top_left == (1, 1) else _spelled_cell(*view.top_left)),
                        *zooms):
        opening = _with_attribute(opening, name, value, _SHEET_VIEW_ORDER)
    kept = "".join(found.group(0) for found in _KEPT.finditer(inside))
    return _closed(opening, _pane_xml(view) + _selections_xml(sheet) + kept, "sheetView")


def _pane_xml(view: SheetViewState) -> str:
    if view.split_xml or not view.state:
        return view.split_xml
    if view.state == "split":
        raise VBAUnsupportedError("saving a window split without being frozen is not implemented: Excel keeps the "
                                  "split in twips of the rows, columns and headings the window shows")
    split = (f'xSplit="{view.columns}" ' if view.columns else "") + (f'ySplit="{view.rows}" ' if view.rows else "")
    return (f'<pane {split}topLeftCell="{_spelled_cell(*view.pane_top_left)}" activePane="{view.active_pane}" '
            f'state="{view.state}"/>')


def _selections_xml(sheet: Worksheet) -> str:
    """Each pane's selection, the active pane's the sheet's own; an attribute at its default, A1, is left out."""
    view = sheet.view
    chosen = dict(view.selections)
    chosen.setdefault(view.active_pane, ("A1", "A1"))
    out: list[str] = []
    for pane, (cell, sqref) in chosen.items():
        if pane == view.active_pane:
            cell, sqref = _sheet_selection(sheet)
        parts = [] if pane == "topLeft" else [f'pane="{pane}"']
        parts += [f'activeCell="{cell}"'] if cell != "A1" else []
        parts += [f'sqref="{sqref}"'] if sqref != "A1" else []
        if parts:
            out.append(f"<selection {' '.join(parts)}/>")
    return "".join(out)


def _sheet_selection(sheet: Worksheet) -> tuple[str, str]:
    active = sheet.active_cell_range
    selection = sheet.selection_range
    cell = _spelled_cell(active.first.top, active.first.left) if active is not None else "A1"
    return cell, " ".join(_spelled(area) for area in selection.areas) if selection is not None else "A1"


def read_book_view(book: Workbook, xml: str) -> None:
    """The workbook's window from its first workbookView, and the sheet its active tab makes active."""
    found = _WORKBOOK_VIEW.search(xml)
    if found is not None:
        head = attributes(found.group(0))
        view = book.view
        view.tabs = _flag(head.get("showSheetTabs"), True)
        view.horizontal_scroll = _flag(head.get("showHorizontalScroll"), True)
        view.vertical_scroll = _flag(head.get("showVerticalScroll"), True)
        view.tab_ratio = _number(head.get("tabRatio"), 600) / 1000
        if book.sheets_:
            book.active_sheet_index = min(_number(head.get("activeTab"), 0), len(book.sheets_) - 1)
    book.view.saved_active = book.active_sheet


def with_book_view(book: Workbook, xml: str) -> str:
    """A workbook part whose first workbookView has the model's tabs, scroll bars, tab ratio and active tab."""
    found = _WORKBOOK_VIEW.search(xml)
    if found is None:
        return xml
    tag = found.group(0)
    closing = "/>" if tag.endswith("/>") else ">"
    opening = tag[:-len(closing)].rstrip()
    view = book.view
    ratio = round(view.tab_ratio * 1000)
    active = book.sheets_.index(book.active_sheet) if book.active_sheet is not None else 0
    for name, value in (("showHorizontalScroll", "" if view.horizontal_scroll else "0"),
                        ("showVerticalScroll", "" if view.vertical_scroll else "0"),
                        ("showSheetTabs", "" if view.tabs else "0"),
                        ("tabRatio", "" if ratio == 600 else str(ratio)),
                        ("activeTab", str(active) if active else "")):
        opening = _with_attribute(opening, name, value, _BOOK_VIEW_ORDER)
    return xml[:found.start()] + opening + closing + xml[found.end():]


def saved(book: Workbook) -> None:
    """The file now has every view as the model has it."""
    for sheet in book.sheets_:
        sheet.view.changed = False
    book.view.changed = False
    book.view.saved_active = book.active_sheet


def copied(source: Worksheet, copy: Worksheet) -> None:
    """A sheet's copy has its view and its selection, measured."""
    copy.view = replace(source.view, selections=dict(source.view.selections), changed=True)
    for name in ("selection_range", "active_cell_range"):
        chosen = getattr(source, name)
        if chosen is not None:
            setattr(copy, name, Range(copy, [Area(a.top, a.left, a.bottom, a.right, copy.name) for a in chosen.areas]))


# --- rows and columns inserted and deleted -----------------------------------------------------------


def _frozen_after(view: SheetViewState, *, rows: bool, start: int, count: int, delete: bool) -> int | None:
    """How many rows or columns are frozen after whole ones are inserted or deleted among them, never fewer than
    one; None when the edit is at or past the split, which leaves the view alone, measured."""
    split = view.rows if rows else view.columns
    if view.split_xml or not view.state or not split:
        return None
    top = view.top_left[0 if rows else 1]
    if start >= top + split:
        return None
    if start < top or (delete and start + count > top + split):
        raise VBAUnsupportedError("inserting or deleting rows or columns that run past the frozen ones is not "
                                  "implemented")
    return max(1, split - count) if delete else split + count


def edit_admitted(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> None:
    """Refuse an edit whose effect on the frozen rows or columns is not modelled, before anything moves."""
    _frozen_after(sheet.view, rows=rows, start=start, count=count, delete=delete)


def edited(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> None:
    """The frozen rows or columns grow or shrink with an edit among them, and the pane past them moves as far; the
    selections and the window's top left stay where they were, measured."""
    view = sheet.view
    after = _frozen_after(view, rows=rows, start=start, count=count, delete=delete)
    if after is None:
        return
    axis = 0 if rows else 1
    moved = after - (view.rows if rows else view.columns)
    if rows:
        view.rows = after
    else:
        view.columns = after
    corner = list(view.pane_top_left)
    corner[axis] += moved
    view.pane_top_left = (corner[0], corner[1])
    view.changed = True


# --- selecting ---------------------------------------------------------------------------------------


def check_shown(target: Range, verb: str) -> None:
    """Select and Activate work only on the active sheet of the active workbook: error 1004 elsewhere, measured."""
    book = target.sheet.book
    if book.application.active_book is not book or book.active_sheet is not target.sheet:
        raise error(1004, f"{verb} method of Range class failed")


def activate(target: Range) -> None:
    """Range.Activate: the active cell moves when the range's first cell is in the selection, else the range is
    selected."""
    check_shown(target, "Activate")
    sheet = target.sheet
    first = target.first
    selection = sheet.selection_range
    areas = selection.areas if selection is not None else [Area(1, 1, 1, 1)]
    if not any(area.contains(first.top, first.left) for area in areas):
        target.select_as_excel_does()
        return
    sheet.active_cell_range = Range(sheet, [Area(first.top, first.left, first.top, first.left, first.sheet)])
    sheet.view.changed = True


def goto(application: Application, reference: object, scroll: object) -> None:
    """Application.Goto: the range's workbook and sheet activated and the range selected, scrolled to with Scroll."""
    if reference is MISSING:
        raise VBAUnsupportedError("Goto without a reference, which goes back to where the last Goto started, is not "
                                  "implemented")
    if isinstance(reference, Range):
        target = reference
    else:
        # A text reference is in R1C1 notation, or a name.
        found = application.vba_get("Range", [to_a1("=" + to_text(reference), 1, 1)[1:]])
        if not isinstance(found, Range):
            raise error(1004, "Reference is not valid.")
        target = found
    sheet = target.sheet
    scrolled = scroll is not MISSING and to_bool(scroll)
    if scrolled:
        _check_scrollable(sheet.view)
    if application.active_book is not sheet.book or sheet.book.active_sheet is not sheet:
        sheet.Activate()
    target.select_as_excel_does()
    if scrolled:
        _scroll_to(sheet.view, 0, target.first.top)
        _scroll_to(sheet.view, 1, target.first.left)


def _check_scrollable(view: SheetViewState) -> None:
    if view.state == "split" or view.split_xml:
        raise VBAUnsupportedError("scrolling a window split without being frozen, or one whose panes the model does "
                                  "not lay out, is not implemented")


def _scroll_to(view: SheetViewState, axis: int, number: int) -> None:
    """Scroll the window along one axis, 0 down and 1 across: the pane past a frozen split, never above or left
    of it; the whole window where nothing is split that way."""
    _check_scrollable(view)
    split = view.rows if axis == 0 else view.columns
    if view.state in _FROZEN and split:
        corner = list(view.pane_top_left)
        corner[axis] = max(number, view.top_left[axis] + split)
        view.pane_top_left = (corner[0], corner[1])
    else:
        corner = list(view.top_left)
        corner[axis] = number
        view.top_left = (corner[0], corner[1])
    view.changed = True


# --- VBA -------------------------------------------------------------------------------------


def front_to_back(application: Application) -> list[Workbook]:
    """The open workbooks by their windows, front to back."""
    books = application.workbooks_.books
    known = [book for book in application.window_order if book in books]
    return known + [book for book in reversed(books) if book not in known]


def brought_forward(application: Application, book: Workbook) -> None:
    """A workbook activated: a new one's window comes to the front, another's only while the screen updates."""
    order = application.window_order
    if book in order and not application.screen_updating:
        return
    if book in order:
        order.remove(book)
    order.insert(0, book)


def window_of(book: Workbook) -> Window:
    """The workbook's one window, the same object each time it is asked for."""
    if book.view.window is None:
        book.view.window = Window(book)
    return book.view.window


class Windows(VBACollection, ExcelObject):
    """Application.Windows, front to back, or Workbook.Windows, the one window a workbook has."""

    vba_type_name = "Windows"

    def __init__(self, application: Application, book: Workbook | None = None) -> None:
        self.application = application
        self.book = book

    def vba_items(self) -> list[object]:
        books = [self.book] if self.book is not None else front_to_back(self.application)
        return [window_of(book) for book in books]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for window in items:
                if isinstance(window, Window) and window.caption().casefold() == index.casefold():
                    return window
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        return super().vba_lookup(index, items)

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application


def pane_count(view: SheetViewState) -> int:
    if not view.state:
        return 1
    if view.split_xml:
        split = attributes(view.split_xml)
        both = float(split.get("xSplit", "0") or 0) and float(split.get("ySplit", "0") or 0)
        return 4 if both else 2
    return 4 if view.rows and view.columns else 2


class Panes(VBACollection, ExcelObject):
    """A window's panes: one, two when it is split one way, four when both."""

    vba_type_name = "Panes"

    def __init__(self, window: Window) -> None:
        self.window = window

    def vba_items(self) -> list[object]:
        raise VBAUnsupportedError("a window's panes, each scrolled on its own, are not implemented; Panes.Count is")

    @member
    def Count(self) -> object:
        return VBAInt(pane_count(self.window.sheet().view), "Long")


class Window(ExcelObject):
    """A workbook's window, which shows its active sheet."""

    vba_type_name = "Window"

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def sheet(self) -> Worksheet:
        sheet = self.book.active_sheet
        if sheet is None:
            raise error(1004, "the window shows no sheet")
        return sheet

    def view(self) -> SheetViewState:
        return self.sheet().view

    def caption(self) -> str:
        caption = self.book.view.caption
        return self.book.name if caption is None else caption

    def _changing(self) -> SheetViewState:
        """The view a setter changes; panes the model does not lay out are kept as the file had them."""
        view = self.view()
        if view.split_xml:
            raise VBAUnsupportedError("changing the view of a window whose panes the model does not lay out, a split "
                                      "Excel kept in twips of what the window shows, is not implemented")
        view.changed = True
        return view

    def _book_changing(self) -> BookViewState:
        self.book.view.changed = True
        return self.book.view

    # -- the sheet it shows

    @member
    def ActiveSheet(self) -> object:
        return self.sheet()

    @member
    def ActiveCell(self) -> object:
        return self.sheet().active_cell

    @member
    def Selection(self) -> object:
        return self.sheet().selection

    @member
    def RangeSelection(self) -> object:
        return self.sheet().selection

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Index(self) -> object:
        return VBAInt(front_to_back(self.book.application).index(self.book) + 1, "Long")

    @method
    def Activate(self) -> object:
        self.book.application.activate_book(self.book)
        return EMPTY

    # -- the window itself

    @member
    def Caption(self) -> object:
        return self.caption()

    @setter("Caption")
    def _set_caption(self, value: object) -> None:
        self.book.view.caption = to_text(value)

    @member
    def Visible(self) -> object:
        return True

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        raise VBAUnsupportedError("hiding a workbook's window is not implemented")

    @member
    def WindowState(self) -> object:
        state = self.book.view.state
        if state is None:
            raise VBAUnsupportedError("a window's state before a macro sets it is the screen's, which the model does "
                                      "not have")
        return VBAInt(state, "Long")

    @setter("WindowState")
    def _set_window_state(self, value: object) -> None:
        self.book.view.state = int(to_integer(value, "Long"))

    @member
    def Width(self) -> object:
        return self._size(self.book.view.width)

    @setter("Width")
    def _set_width(self, value: object) -> None:
        self.book.view.width = float(to_number(value))

    @member
    def Height(self) -> object:
        return self._size(self.book.view.height)

    @setter("Height")
    def _set_height(self, value: object) -> None:
        self.book.view.height = float(to_number(value))

    @staticmethod
    def _size(points: float | None) -> object:
        if points is None:
            raise VBAUnsupportedError("a window's size before a macro sets it is the screen's, which the model does "
                                      "not have")
        return points

    # -- the view of the sheet

    @member
    def Zoom(self) -> object:
        return VBAInt(self.view().zoom, "Long")

    @setter("Zoom")
    def _set_zoom(self, value: object) -> None:
        if value is True:
            raise VBAUnsupportedError("zooming to fit the selection, which depends on the window's size, is not "
                                      "implemented")
        zoom = int(float(to_number(value)))
        if not 10 <= zoom <= 400:
            raise error(1004, "Unable to set the Zoom property of the Window class")
        view = self._changing()
        view.zoom = zoom
        if view.view == 1:
            view.zoom_normal = zoom
        elif view.view == 2:
            view.zoom_break = zoom
        else:
            view.zoom_layout = zoom

    @member
    def View(self) -> object:
        return VBAInt(self.view().view, "Long")

    @setter("View")
    def _set_view(self, value: object) -> None:
        number = int(to_integer(value, "Long"))
        if number not in VIEWS:
            raise error(1004, "Unable to set the View property of the Window class")
        view = self.view()
        if number == view.view:
            return
        if view.state:
            raise VBAUnsupportedError("changing the view of a window with panes is not implemented")
        view = self._changing()
        if view.view == 1:
            view.zoom_normal = view.zoom_normal or view.zoom
        kept = {1: view.zoom_normal, 2: view.zoom_break, 3: view.zoom_layout}[number]
        view.zoom = kept or _OPENING_ZOOM[number]
        view.view = number

    def _flag(self, name: str, value: object) -> None:
        setattr(self._changing(), name, to_bool(value))

    @member
    def DisplayGridlines(self) -> object:
        return self.view().gridlines

    @setter("DisplayGridlines")
    def _set_gridlines(self, value: object) -> None:
        self._flag("gridlines", value)

    @member
    def DisplayHeadings(self) -> object:
        return self.view().headings

    @setter("DisplayHeadings")
    def _set_headings(self, value: object) -> None:
        self._flag("headings", value)

    @member
    def DisplayZeros(self) -> object:
        return self.view().zeros

    @setter("DisplayZeros")
    def _set_zeros(self, value: object) -> None:
        self._flag("zeros", value)

    @member
    def DisplayFormulas(self) -> object:
        return self.view().formulas

    @setter("DisplayFormulas")
    def _set_formulas(self, value: object) -> None:
        self._flag("formulas", value)

    # -- scrolling

    def _scrolled(self, axis: int) -> int:
        view = self.view()
        split = view.rows if axis == 0 else view.columns
        return (view.pane_top_left if view.state in _FROZEN and split else view.top_left)[axis]

    def _scroll(self, axis: int, value: object) -> None:
        number = int(to_integer(value, "Long"))
        if not 1 <= number <= (MAX_ROWS if axis == 0 else MAX_COLUMNS):
            name = "ScrollRow" if axis == 0 else "ScrollColumn"
            raise error(1004, f"Unable to set the {name} property of the Window class")
        _scroll_to(self._changing(), axis, number)

    @member
    def ScrollRow(self) -> object:
        return VBAInt(self._scrolled(0), "Long")

    @setter("ScrollRow")
    def _set_scroll_row(self, value: object) -> None:
        self._scroll(0, value)

    @member
    def ScrollColumn(self) -> object:
        return VBAInt(self._scrolled(1), "Long")

    @setter("ScrollColumn")
    def _set_scroll_column(self, value: object) -> None:
        self._scroll(1, value)

    # -- panes

    @member
    def FreezePanes(self) -> object:
        return self.view().state in _FROZEN

    @setter("FreezePanes")
    def _set_freeze_panes(self, value: object) -> None:
        view = self.view()
        if to_bool(value):
            if view.state in _FROZEN:
                return
            if view.state == "split":
                self._changing()
                _freeze(view, "frozenSplit")
                return
            active = self.sheet().active_cell_range
            row, column = (active.first.top, active.first.left) if active is not None else (1, 1)
            rows, columns = row - view.top_left[0], column - view.top_left[1]
            if rows < 0 or columns < 0 or not (rows or columns):
                raise VBAUnsupportedError("freezing panes with the active cell out of view or at the window's top "
                                          "left is not implemented: Excel then freezes the middle of a window whose "
                                          "size the model does not have")
            self._changing()
            view.rows, view.columns = rows, columns
            view.selections = {}
            if columns:
                view.selections["topRight"] = _selected(view.top_left[0], column)
            if rows:
                view.selections["bottomLeft"] = _selected(row, view.top_left[1])
            _freeze(view, "frozen")
        elif view.state == "frozen":
            self._changing()
            _unsplit(view)
        elif view.state == "frozenSplit":
            self._changing().state = "split"

    @member
    def Split(self) -> object:
        return bool(self.view().state)

    @setter("Split")
    def _set_split(self, value: object) -> None:
        view = self.view()
        if to_bool(value):
            if not view.state:
                raise VBAUnsupportedError("splitting a window at the active cell is not implemented: Excel keeps the "
                                          "split in twips of what the window shows")
            return
        if not view.state:
            return
        self._changing()
        if view.state == "frozen":
            view.state = "split"
        else:
            _unsplit(view)

    def _split_count(self, axis: int) -> int:
        view = self.view()
        if view.split_xml:
            raise VBAUnsupportedError("panes the model does not lay out, a split Excel kept in twips, are not counted "
                                      "in rows and columns")
        return view.rows if axis == 0 else view.columns

    def _set_split_count(self, axis: int, value: object) -> None:
        count = int(to_integer(value, "Long"))
        if count < 0:
            name = "SplitRow" if axis == 0 else "SplitColumn"
            raise error(1004, f"Unable to set the {name} property of the Window class")
        view = self.view()
        if not view.state and not count:
            return
        self._changing()
        if view.state in _FROZEN:
            view.state = "split"
        elif not view.state:
            view.state, view.active_pane, view.selections = "split", "topLeft", {}
        top, left = view.top_left
        if axis == 0:
            view.rows = count
        else:
            view.columns = count
        # Each pane keeps the selection it opened with, in the order the panes were made.
        for pane, there, cell in (("bottomLeft", view.rows, (top + view.rows, left)),
                                  ("topRight", view.columns, (top, left + view.columns)),
                                  ("bottomRight", view.rows and view.columns, (top + view.rows, left + view.columns))):
            if not there:
                view.selections.pop(pane, None)
            elif pane not in view.selections:
                view.selections[pane] = _selected(*cell)
        view.pane_top_left = (top + view.rows, left + view.columns)
        if not (view.rows or view.columns):
            _unsplit(view)

    @member
    def SplitRow(self) -> object:
        return VBAInt(self._split_count(0), "Long")

    @setter("SplitRow")
    def _set_split_row(self, value: object) -> None:
        self._set_split_count(0, value)

    @member
    def SplitColumn(self) -> object:
        return VBAInt(self._split_count(1), "Long")

    @setter("SplitColumn")
    def _set_split_column(self, value: object) -> None:
        self._set_split_count(1, value)

    def _split_points(self, axis: int) -> float:
        """How far from the top or the left the split runs, in points: its rows' or columns' pixels at the zoom."""
        sheet = self.sheet()
        view = sheet.view
        start, count = view.top_left[axis], self._split_count(axis)
        size = sheet.row_height_points if axis == 0 else sheet.column_width_points
        pixels = sum(int(round(size(one) / _dimensions.PIXEL) * view.zoom / 100 + 0.5)
                     for one in range(start, start + count))
        return pixels * _dimensions.PIXEL

    @member
    def SplitVertical(self) -> object:
        return self._split_points(0)

    @member
    def SplitHorizontal(self) -> object:
        return self._split_points(1)

    @member
    def Panes(self, Index: object = MISSING) -> object:
        panes = Panes(self)
        return panes if Index is MISSING else panes.vba_get("Item", [Index])

    # -- the workbook's window

    @member
    def DisplayWorkbookTabs(self) -> object:
        return self.book.view.tabs

    @setter("DisplayWorkbookTabs")
    def _set_tabs(self, value: object) -> None:
        self._book_changing().tabs = to_bool(value)

    @member
    def TabRatio(self) -> object:
        return self.book.view.tab_ratio

    @setter("TabRatio")
    def _set_tab_ratio(self, value: object) -> None:
        ratio = float(to_number(value))
        if not 0 <= ratio <= 1:
            raise error(1004, "Unable to set the TabRatio property of the Window class")
        self._book_changing().tab_ratio = ratio

    @member
    def DisplayHorizontalScrollBar(self) -> object:
        return self.book.view.horizontal_scroll

    @setter("DisplayHorizontalScrollBar")
    def _set_horizontal_scroll(self, value: object) -> None:
        self._book_changing().horizontal_scroll = to_bool(value)

    @member
    def DisplayVerticalScrollBar(self) -> object:
        return self.book.view.vertical_scroll

    @setter("DisplayVerticalScrollBar")
    def _set_vertical_scroll(self, value: object) -> None:
        self._book_changing().vertical_scroll = to_bool(value)


def _freeze(view: SheetViewState, state: str) -> None:
    """Freeze the split in place: the pane past it becomes the active one, and the top left pane's selection is
    the window's top left cell."""
    top, left = view.top_left
    view.selections = {"topLeft": _selected(top, left), **view.selections}
    view.state = state
    view.active_pane = "bottomRight" if view.rows and view.columns else "bottomLeft" if view.rows else "topRight"
    view.pane_top_left = (top + view.rows, left + view.columns)


def _unsplit(view: SheetViewState) -> None:
    view.state, view.rows, view.columns, view.active_pane, view.selections = "", 0, 0, "topLeft", {}
