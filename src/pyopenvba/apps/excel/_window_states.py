"""What each of a workbook's windows has of its own (Workbook.NewWindow), measured on a shown Excel (scripts/measure_new_window.py,
tests/fixtures/new_window.json).

A window has its own active sheet, its own selection and view of each sheet, and its own tabs and scroll bars. The workbook and its sheets hold
those of the window in front; each window behind keeps its in a KeptView, swapped in while it is asked for anything. So code that knows of one
window goes on reading the workbook and its sheets.

- A new window comes to the front, showing the sheet of the one in front. Its selection and every setting of its view start afresh, nothing copied.
- Windows are numbered from 1 in the order they were made, and closing one closes the numbers up. With more than one, a caption not set by a macro
  is the workbook's name, two spaces, a dash, two spaces and a number; a new window's number there counts the windows with such a caption.
- A window can be closed while another is left, and the workbook stays. Closing the last closes the workbook.
- A save writes a workbookView for each window, and in each sheet a sheetView for each. Every window's active tab is the front window's.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import ChartSheet, Range, Workbook, Worksheet
    from pyopenvba.apps.excel._windows import SheetViewState

#: A workbook's window settings, as a window of its own keeps them. The rest of its workbookView belongs to the workbook.
BOOK_FIELDS = ('tabs', 'first_sheet', 'tab_ratio', 'horizontal_scroll', 'vertical_scroll', 'caption', 'on_window', 'state', 'left', 'top', 'width', 'height')

SheetState = tuple['SheetViewState', 'Range | None', 'Range | None']


@dataclass
class KeptView:
    """What a window behind the front one keeps."""

    book: dict[str, object]
    sheet: Worksheet | None
    chart: ChartSheet | None = None
    selected: list[Worksheet] = field(default_factory=lambda: [])
    sheets: dict[int, SheetState] = field(default_factory=lambda: {})


@dataclass
class WindowPlace:
    """One of a workbook's windows, whichever object a macro holds for it."""

    number: int = 1
    suffix: int = 1
    #: None while the workbook and its sheets hold this window's state.
    kept: KeptView | None = None
    gone: bool = False


def fresh_sheet() -> SheetState:
    from pyopenvba.apps.excel._windows import SheetViewState

    return SheetViewState(), None, None


def fresh_view(book: Workbook) -> KeptView:
    """A new window's: the sheet in front, and everything else as a new window has it."""
    from pyopenvba.apps.excel._windows import BookViewState

    start = BookViewState()
    return KeptView({name: getattr(start, name) for name in BOOK_FIELDS}, book.active_sheet, book.active_chart_sheet)


def sheet_state(book: Workbook, place: WindowPlace, sheet: Worksheet) -> SheetState:
    """A sheet's view and selection in a window."""
    if place.kept is None:
        return sheet.view, sheet.selection_range, sheet.active_cell_range
    found = place.kept.sheets.get(id(sheet))
    return found or book.view.extra_views.get((place.number - 1, id(sheet))) or fresh_sheet()


def _stash(book: Workbook) -> KeptView:
    return KeptView({name: getattr(book.view, name) for name in BOOK_FIELDS}, book.active_sheet, book.active_chart_sheet, list(book.selected_sheets),
                    {id(one): (one.view, one.selection_range, one.active_cell_range) for one in book.sheets_})


def _restore(book: Workbook, place: WindowPlace) -> None:
    kept = place.kept
    assert kept is not None
    states = {id(one): sheet_state(book, place, one) for one in book.sheets_}
    for name, value in kept.book.items():
        setattr(book.view, name, value)
    if kept.sheet is not None and any(one is kept.sheet for one in book.sheets_):
        book.active_sheet_index = book.sheets_.index(kept.sheet)
    book.active_chart_sheet = kept.chart
    book.selected_sheets = [one for one in kept.selected if any(one is sheet for sheet in book.sheets_)]
    for sheet in book.sheets_:
        sheet.view, sheet.selection_range, sheet.active_cell_range = states[id(sheet)]
    place.kept = None


def bring_in(book: Workbook, place: WindowPlace) -> None:
    """Make the workbook and its sheets hold this window's state, keeping the other's."""
    live = book.view.live
    if live is place or live is None:
        return
    live.kept = _stash(book)
    _restore(book, place)
    book.view.live = place
