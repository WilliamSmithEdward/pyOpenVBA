"""Excel's object model, in memory.

The objects a macro touches -- Application, Workbook, Worksheet, Range
and the few that hang off them -- holding real state that VBA mutates
and that can be written back out to a file.

A cell holds a value, a formula and its format, a
:class:`~pyopenvba.apps.excel._styles.Style`.  Reading a cell whose
formula has not been worked out calculates it first, through
:mod:`pyopenvba.apps.excel._calc`, so a macro that writes a formula and
reads the answer gets one.  What a macro reads and sets on a format --
Font, Interior, Borders, alignment -- lives in
:mod:`pyopenvba.apps.excel._formats`.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, replace
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_reference
from pyopenvba.exceptions import VBAUnsupportedError, VBARuntimeError
from pyopenvba.formula._parse import shift_text
from pyopenvba.formula._r1c1 import from_a1, to_a1
from pyopenvba.apps.excel._find import FindState
from pyopenvba.apps.excel import _dimensions, _merges, _names
from pyopenvba.apps.excel._styles import applying
from pyopenvba.shapes._values import Shape as ShapeState
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    ERR_APPLICATION_DEFINED,
    ERR_SUBSCRIPT_OUT_OF_RANGE,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBAErrorValue,
    VBADate,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_text,
    type_name,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._autofilter import SheetFilter
    from pyopenvba.apps.excel._protection import BookProtection, Gate, SheetProtection
    from pyopenvba.apps.excel._clipboard import Clip
    from pyopenvba.apps.excel._sort import SortState
    from pyopenvba.apps.excel._styles import Style, Stylesheet
    from pyopenvba.apps.excel._typing import Typed
    from pyopenvba.interpreter._runtime import Interpreter

_LIBRARY = "excel"


class ExcelObject(VBAObject):
    """Every Excel class, so the member inventory knows where to look."""

    vba_library = _LIBRARY
    invalidated = False

    def _check_alive(self) -> None:
        if any(getattr(owner, "invalidated", False) for owner in
               (self, getattr(self, "sheet", None), getattr(self, "entry", None))):
            raise error(424, "Object required")

    def vba_member(self, name: str) -> MemberSpec | None:
        self._check_alive()
        return super().vba_member(name)

    def vba_default_member(self) -> MemberSpec | None:
        self._check_alive()
        return super().vba_default_member()

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        spec = self.vba_member(name)
        self.guard_set(spec.name if spec is not None else name)
        super().vba_set(name, value, args, named, by_ref=by_ref)

    def guard_set(self, member: str) -> None:
        """Refuse a property a protected sheet does not let a macro set: the classes of cells say which."""


# --- what a cell holds ---------------------------------------------------------------


@dataclass(slots=True)
class Cell:
    """One cell's state.  Absent from the sheet's map means empty."""

    value: object = EMPTY
    formula: str = ""
    #: True when the formula was written here and nothing has computed it.
    stale: bool = False
    #: The cell's format.  None is the workbook's default, its first xf.
    style: Style | None = None
    #: The xf the cell was read with, which a save keeps while the format is unchanged.
    xf: int = -1

    @property
    def number_format(self) -> str:
        return self.style.number_format if self.style is not None else "General"

    def is_blank(self) -> bool:
        return self.value is EMPTY and not self.formula and self.style is None


# --- Application ------------------------------------------------------------------------


class Application(ExcelObject):
    """The Excel instance: the root of everything a macro can reach."""

    vba_type_name = "Application"

    def __init__(self) -> None:
        self.workbooks_ = Workbooks(self)
        self.screen_updating = True
        self.display_alerts = True
        self.enable_events = True
        self.calculation = -4105  # xlCalculationAutomatic
        self.status_bar: object = False
        #: What Copy or Cut last put on the clipboard (see _clipboard).
        self.clipboard: Clip | None = None
        self.user_name = "pyOpenVBA"
        self.interpreter: Interpreter | None = None
        self.active_book: Workbook | None = None
        self._this_workbook: Workbook | None = None
        self.worksheet_function = WorksheetFunction(self)
        self.find_state = FindState()

    # -- identity

    @member
    def Name(self) -> object:
        return "Microsoft Excel"

    @member
    def Version(self) -> object:
        return "16.0"

    @member
    def Build(self) -> object:
        return VBAInt(0, "Long")

    @member
    def Application(self) -> object:
        return self

    @member
    def Parent(self) -> object:
        return self

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")

    # -- state a macro sets and reads back

    @member
    def ScreenUpdating(self) -> object:
        return self.screen_updating

    @setter("ScreenUpdating")
    def _set_screen_updating(self, value: object) -> None:
        self.screen_updating = to_bool(value)

    @member
    def DisplayAlerts(self) -> object:
        return self.display_alerts

    @setter("DisplayAlerts")
    def _set_display_alerts(self, value: object) -> None:
        self.display_alerts = to_bool(value)

    def vba_get(self, name: str, args: Any = (), named: Any = None) -> object:
        """A member, or a worksheet function called the late-bound way, which hands an error back as a value."""
        if self.vba_member(name) is None:
            from pyopenvba.interpreter._inventory import member_exists

            if member_exists("WorksheetFunction", name, "excel"):
                from pyopenvba.apps.excel._worksheet_functions import call, engine_name

                if engine_name(name) is None:
                    raise VBAUnsupportedError(f"Application.{name} is a worksheet function that pyOpenVBA does not "
                                              f"implement")
                return call(self, name, args, named, raising=False)
        return super().vba_get(name, args, named)

    @member
    def EnableEvents(self) -> object:
        return self.enable_events

    @setter("EnableEvents")
    def _set_enable_events(self, value: object) -> None:
        self.enable_events = to_bool(value)

    @member
    def Calculation(self) -> object:
        return VBAInt(self.calculation, "Long")

    @setter("Calculation")
    def _set_calculation(self, value: object) -> None:
        self.calculation = int(to_integer(value, "Long"))

    @member
    def StatusBar(self) -> object:
        return self.status_bar

    @setter("StatusBar")
    def _set_status_bar(self, value: object) -> None:
        self.status_bar = value

    @member
    def CutCopyMode(self) -> object:
        """0 with nothing on the clipboard, 1 after Copy and 2 after Cut."""
        from pyopenvba.apps.excel._clipboard import mode

        return mode(self.clipboard)

    @setter("CutCopyMode")
    def _set_cut_copy_mode(self, value: object) -> None:
        if not to_bool(value):
            self.clipboard = None

    @member
    def UserName(self) -> object:
        return self.user_name

    @setter("UserName")
    def _set_user_name(self, value: object) -> None:
        self.user_name = to_text(value)

    @member
    def Visible(self) -> object:
        return False

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        return None

    # -- the object graph

    @member
    def Workbooks(self, Index: object = MISSING) -> object:
        return self.workbooks_ if Index is MISSING else self.workbooks_.vba_get("Item", [Index])

    @member
    def ActiveWorkbook(self) -> object:
        return self.active_book if self.active_book is not None else NOTHING

    @member
    def ThisWorkbook(self) -> object:
        book = self._this_workbook or self.active_book
        return book if book is not None else NOTHING

    @member
    def ActiveSheet(self) -> object:
        book = self.active_book
        return book.active_sheet if book is not None else NOTHING

    @member
    def ActiveCell(self) -> object:
        sheet = self.active_book.active_sheet if self.active_book is not None else None
        return sheet.active_cell if sheet is not None else NOTHING

    @member
    def Selection(self) -> object:
        sheet = self.active_book.active_sheet if self.active_book is not None else None
        return sheet.selection if sheet is not None else NOTHING

    @member
    def Sheets(self, Index: object = MISSING) -> object:
        return self._require_book().vba_get("Sheets", [] if Index is MISSING else [Index])

    @member
    def Worksheets(self, Index: object = MISSING) -> object:
        return self._require_book().vba_get("Worksheets", [] if Index is MISSING else [Index])

    @member
    def Names(self, Index: object = MISSING) -> object:
        return self._require_book().vba_get("Names", [] if Index is MISSING else [Index])

    @member(default=True)
    def Range(self, Cell1: object = MISSING, Cell2: object = MISSING) -> object:
        sheet = self._require_sheet()
        return sheet.vba_get("Range", [Cell1, Cell2])

    @member
    def Cells(self, RowIndex: object = MISSING, ColumnIndex: object = MISSING) -> object:
        return self._require_sheet().vba_get("Cells", [RowIndex, ColumnIndex])

    @member
    def Rows(self, Index: object = MISSING) -> object:
        return self._require_sheet().vba_get("Rows", [Index])

    @member
    def Columns(self, Index: object = MISSING) -> object:
        return self._require_sheet().vba_get("Columns", [Index])

    @member
    def WorksheetFunction(self) -> object:
        return self.worksheet_function

    @member
    def Path(self) -> object:
        return str(Path.cwd())

    # -- methods

    @method
    def Evaluate(self, Name: object = MISSING) -> object:
        return self.evaluate_text(to_text(Name))

    @method
    def Run(self, Macro: object = MISSING, *args: object) -> object:
        if self.interpreter is None:
            raise error(ERR_APPLICATION_DEFINED, "no VBA project is loaded")
        return self.interpreter.run(to_text(Macro), list(args))

    @method
    def Intersect(self, Arg1: object = MISSING, Arg2: object = MISSING, *rest: object) -> object:
        ranges = [one for one in (Arg1, Arg2, *rest) if isinstance(one, Range)]
        if len(ranges) < 2:
            raise error(ERR_APPLICATION_DEFINED, "Intersect needs two ranges")
        sheet = ranges[0].sheet
        top = max(one.first.top for one in ranges)
        left = max(one.first.left for one in ranges)
        bottom = min(one.first.bottom for one in ranges)
        right = min(one.first.right for one in ranges)
        if top > bottom or left > right:
            return NOTHING
        return Range(sheet, [Area(top, left, bottom, right, sheet.name)])

    @method
    def Union(self, Arg1: object = MISSING, Arg2: object = MISSING, *rest: object) -> object:
        ranges = [one for one in (Arg1, Arg2, *rest) if isinstance(one, Range)]
        if not ranges:
            raise error(ERR_APPLICATION_DEFINED, "Union needs a range")
        areas: list[Area] = []
        for one in ranges:
            areas.extend(one.areas)
        return Range(ranges[0].sheet, areas)

    @method
    def Calculate(self) -> object:
        for book in self.workbooks_.books:
            book.calculator.calculate_all()
        return EMPTY

    @method
    def Quit(self) -> object:
        self.workbooks_.books.clear()
        self.active_book = None
        return EMPTY

    @method
    def Wait(self, Time: object = MISSING) -> object:
        return True

    @method
    def Volatile(self, Volatile: object = MISSING) -> object:
        return EMPTY

    # -- Python side

    def evaluate_text(self, text: str) -> object:
        """``[A1+1]`` and Evaluate("A1+1"), worked out on the active sheet as Excel's Evaluate works them."""
        from pyopenvba.apps.excel._evaluate import evaluated

        return evaluated(self._require_sheet(), text)

    def _require_book(self) -> Workbook:
        if self.active_book is None:
            raise error(ERR_APPLICATION_DEFINED, "no workbook is open")
        return self.active_book

    def _require_sheet(self) -> Worksheet:
        sheet = self._require_book().active_sheet
        if sheet is None:
            raise error(ERR_APPLICATION_DEFINED, "no sheet is active")
        return sheet

    def activate_book(self, book: Workbook) -> None:
        self.active_book = book
        if self._this_workbook is None:
            self._this_workbook = book

    def describe(self, indent: str = "") -> str:
        lines = [f"{indent}Excel {self.vba_get('Version')} ({len(self.workbooks_.books)} workbook(s))"]
        for book in self.workbooks_.books:
            lines.append(book.describe(indent + "  "))
        return "\n".join(lines)


# --- workbooks -----------------------------------------------------------------------------


class Workbooks(VBACollection, ExcelObject):
    """The open workbooks."""

    vba_type_name = "Workbooks"

    def __init__(self, application: Application) -> None:
        self.application = application
        self.books: list[Workbook] = []
        self._next_number = 1

    def next_name(self) -> str:
        while True:
            name = f"Book{self._next_number}"
            self._next_number += 1
            if not any(book.name.casefold() == name.casefold() for book in self.books):
                return name

    def vba_items(self) -> list[object]:
        return list(self.books)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for book in self.books:
                if book.name.lower() == index.lower():
                    return book
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        return super().vba_lookup(index, items)

    @method
    def Add(self, Template: object = MISSING) -> object:
        if Template is not MISSING and Template != -4167:
            raise VBAUnsupportedError("Only worksheet workbooks are supported by Workbooks.Add")
        book = Workbook(self.application, self.next_name())
        book.add_sheet("Sheet1")
        self.books.append(book)
        self.application.activate_book(book)
        return book

    @method
    def Open(self, Filename: object = MISSING) -> object:
        from pyopenvba.apps.excel._io import load_workbook

        if Filename is MISSING:
            raise error(449)
        path = Path(to_text(Filename)).resolve()
        existing = next((book for book in self.books if book.path and (Path(book.path) / book.name).resolve() == path), None)
        if existing is not None:
            self.application.activate_book(existing)
            return existing
        if any(book.name.casefold() == path.name.casefold() for book in self.books):
            raise error(1004, "A workbook with that name is already open")
        book = load_workbook(self.application, path)
        self.books.append(book)
        self.application.activate_book(book)
        return book

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application

    def remove(self, book: Workbook) -> None:
        if book in self.books:
            self.books.remove(book)
        if self.application.active_book is book:
            self.application.active_book = self.books[-1] if self.books else None


class Workbook(ExcelObject):
    """One workbook: its sheets, its names, and where it came from."""

    vba_type_name = "Workbook"

    def __init__(self, application: Application, name: str, path: str = "") -> None:
        self.application = application
        self.name = name
        self.path = path
        self.saved = True
        self.sheets_: list[Worksheet] = []
        self.names_ = Names(self)
        self.queries_ = Queries(self)
        #: How the workbook first spelled each name its formulas use but it does not define, by lower case;
        #: None until a formula is written, when the formulas it has teach it.
        self.name_spellings: dict[str, str] | None = None
        self.active_sheet_index = 0
        #: The package this was loaded from, kept so a save can put back
        #: every part pyOpenVBA does not model.
        self.package: Any = None
        self._calculator: Any = None
        #: Where each query's rows go, read from the package on demand.
        self._load_targets: Any = None
        self._stylesheet: Stylesheet | None = None
        #: What Workbook.Protect set, or None while the structure is not protected (see _protection).
        self.protection: BookProtection | None = None
        #: True once a macro protects or unprotects the workbook, so a save writes workbookProtection again.
        self.protection_changed = False

    @property
    def calculator(self) -> Any:
        """The workbook's formulas, and what each one comes to."""
        if self._calculator is None:
            from pyopenvba.apps.excel._calc import Calculator

            self._calculator = Calculator(self)
        return self._calculator

    @property
    def stylesheet(self) -> Stylesheet:
        """The workbook's cell formats: its own, or those of the template a new workbook starts from."""
        if self._stylesheet is None:
            from pyopenvba.apps.excel._io import read_stylesheet

            self._stylesheet = read_stylesheet(self.package)
        return self._stylesheet

    # -- identity

    @member
    def Name(self) -> object:
        return self.name

    @member
    def FullName(self) -> object:
        return str(Path(self.path) / self.name) if self.path else self.name

    @member
    def Path(self) -> object:
        return self.path

    @member
    def Saved(self) -> object:
        return self.saved

    @setter("Saved")
    def _set_saved(self, value: object) -> None:
        self.saved = to_bool(value)

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application

    # -- contents

    @member
    def Sheets(self, Index: object = MISSING) -> object:
        sheets = Sheets(self)
        return sheets if Index is MISSING else sheets.vba_get("Item", [Index])

    @member
    def Worksheets(self, Index: object = MISSING) -> object:
        return self.Sheets(Index)

    @member
    def ActiveSheet(self) -> object:
        sheet = self.active_sheet
        return sheet if sheet is not None else NOTHING

    @member
    def Names(self, Index: object = MISSING) -> object:
        return self.names_ if Index is MISSING else self.names_.vba_get("Item", [Index])

    @member
    def Queries(self, Index: object = MISSING) -> object:
        return self.queries_ if Index is MISSING else self.queries_.vba_get("Item", [Index])

    @member
    def VBProject(self) -> object:
        raise VBAUnsupportedError(
            "Workbook.VBProject edits the project from inside itself; use pyopenvba.ExcelFile instead"
        )

    # -- methods

    @method
    def Activate(self) -> object:
        self.application.activate_book(self)
        return EMPTY

    # -- protection (see _protection)

    @method
    def Protect(self, Password: object = MISSING, Structure: object = MISSING, Windows: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import protect_book

        protect_book(self, Password, Structure)
        return EMPTY

    @method
    def Unprotect(self, Password: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import unprotect_book

        unprotect_book(self, Password)
        return EMPTY

    @member
    def ProtectStructure(self) -> object:
        return self.protection is not None

    @member
    def ProtectWindows(self) -> object:
        """False: Excel takes Protect's Windows and protects nothing with it, measured."""
        return False

    @method
    def Close(self, SaveChanges: object = MISSING, Filename: object = MISSING, RouteWorkbook: object = MISSING) -> object:
        if SaveChanges is not MISSING and to_bool(SaveChanges):
            if Filename is not MISSING:
                self.SaveAs(Filename)
            else:
                self.Save()
        self.application.workbooks_.remove(self)
        return EMPTY

    @method
    def Save(self) -> object:
        from pyopenvba.apps.excel._io import save_workbook

        if not self.path:
            raise error(ERR_APPLICATION_DEFINED, "this workbook has never been saved, so Save has no path")
        save_workbook(self, Path(self.path) / self.name)
        self.saved = True
        return EMPTY

    @method
    def SaveAs(self, Filename: object = MISSING) -> object:
        from pyopenvba.apps.excel._io import save_workbook

        if Filename is MISSING:
            raise error(449)
        target = Path(to_text(Filename))
        save_workbook(self, target)
        self.path = str(target.parent)
        self.name = target.name
        self.saved = True
        return EMPTY

    @method
    def Calculate(self) -> object:
        self.calculator.calculate_all()
        return EMPTY

    @method
    def RefreshAll(self) -> object:
        """Evaluate every query and land each one on its sheet."""
        from pyopenvba.apps.excel._refresh import refresh

        for entry in list(self.queries_.entries):
            refresh(self, entry.name)
        self.calculator.calculate_all()
        return EMPTY

    # -- Python side

    @property
    def active_sheet(self) -> Worksheet | None:
        if not self.sheets_:
            return None
        return self.sheets_[min(self.active_sheet_index, len(self.sheets_) - 1)]

    def activate_sheet(self, sheet: Worksheet) -> None:
        self.active_sheet_index = self.sheets_.index(sheet)

    def add_sheet(self, name: str = "", *, at: int | None = None) -> Worksheet:
        """A new worksheet, named the way Excel names one when asked."""
        chosen = name or self._next_sheet_name()
        if any(sheet.name.lower() == chosen.lower() for sheet in self.sheets_):
            raise error(1004, f"a sheet called {chosen} is already there")
        sheet = Worksheet(self, chosen)
        if at is None:
            self.sheets_.append(sheet)
        else:
            self.sheets_.insert(at, sheet)
        self.saved = False
        return sheet

    def _next_sheet_name(self) -> str:
        taken = {sheet.name.lower() for sheet in self.sheets_}
        number = len(self.sheets_) + 1
        while f"sheet{number}" in taken:
            number += 1
        return f"Sheet{number}"

    def sheet_named(self, name: str) -> Worksheet:
        for sheet in self.sheets_:
            if sheet.name.lower() == name.lower():
                return sheet
        raise error(ERR_SUBSCRIPT_OUT_OF_RANGE, f"there is no sheet called {name}")

    def describe(self, indent: str = "") -> str:
        lines = [f"{indent}Workbook {self.name!r}{'' if self.saved else ' (unsaved changes)'}"]
        for sheet in self.sheets_:
            lines.append(sheet.describe(indent + "  "))
        if self.names_.entries:
            lines.append(f"{indent}  Names: " + ", ".join(one.name for one in self.names_.entries))
        if self.queries_.entries:
            lines.append(f"{indent}  Queries: " + ", ".join(one.name for one in self.queries_.entries))
        return "\n".join(lines)


class Sheets(VBACollection, ExcelObject):
    """A workbook's sheets, in tab order."""

    vba_type_name = "Sheets"

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        return list(self.book.sheets_)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            return self.book.sheet_named(index)
        return super().vba_lookup(index, items)

    @method
    def Add(
        self,
        Before: object = MISSING,
        After: object = MISSING,
        Count: object = MISSING,
        Type: object = MISSING,
    ) -> object:
        from pyopenvba.apps.excel._protection import ADDING, refuse_structure

        refuse_structure(self.book, ADDING)
        at: int | None = None
        if isinstance(Before, Worksheet):
            at = self.book.sheets_.index(Before)
        elif isinstance(After, Worksheet):
            at = self.book.sheets_.index(After) + 1
        elif self.book.active_sheet is not None:
            at = self.book.sheets_.index(self.book.active_sheet)
        sheet = self.book.add_sheet(at=at)
        self.book.activate_sheet(sheet)
        return sheet

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Parent(self) -> object:
        return self.book


class Worksheet(ExcelObject):
    """One worksheet and its cells."""

    vba_type_name = "Worksheet"

    def __init__(self, book: Workbook, name: str) -> None:
        self.book = book
        self.name = name
        self.cells_: dict[tuple[int, int], Cell] = {}
        self.merged_areas: list[Area] = []
        self.merges_dirty = False
        self.visible = -1  # xlSheetVisible
        #: Row heights, column widths and what is hidden.
        self.dims = _dimensions.SheetDimensions(self)
        self.selection_range: Range | None = None
        #: The settings Worksheet.Sort holds, made the first time it is asked for.
        self.sort_state: SortState | None = None
        #: The sheet's AutoFilter: its range and the criteria on its columns.
        self.auto_filter: SheetFilter | None = None
        #: True once the filter is made, changed or removed, so a save writes it again.
        self.filter_changed = False
        #: What Protect set, or None while the sheet is not protected (see _protection).
        self.protection: SheetProtection | None = None
        #: The Allow options the last Protect gave, which Unprotect leaves in place.
        self.protection_allows: frozenset[str] = frozenset()
        #: True once a macro protects or unprotects the sheet, so a save writes its sheetProtection again.
        self.protection_changed = False
        #: EnableSelection: xlNoRestrictions, 0, until a macro sets it.
        self.enable_selection = 0
        #: The write in progress on the protected sheet, which passes over its locked cells.
        self.write_gate: Gate | None = None
        self.active_cell_range: Range | None = None
        #: What the sheet's XML part was called in the file it came from.
        self.part_name = ""
        #: False until a cell changes.  A sheet nobody wrote to is
        #: carried through a save exactly as it arrived.
        self.dirty = False
        #: The shapes on the sheet, in the order the drawing holds them.
        self.shapes_: list[ShapeState] = []
        #: How many shapes have been added here, which is where the next
        #: one's number comes from.  It does not go back down when one
        #: is deleted, and it belongs to the sheet rather than to the
        #: workbook: both measured in live Excel.
        self.shape_count = 0
        #: The drawing part they came from, and its markup, so one that
        #: nobody touched is written back exactly as it arrived.
        self.drawing_part = ""
        self.drawing_xml = ""
        self.drawing_dirty = False

    def touched(self) -> None:
        self.dirty = True
        self.book.saved = False

    def declares_revisions(self) -> bool:
        """Whether the sheet's part declares the revision namespace, where an autoFilter's xr:uid lives."""
        import re

        package = self.book.package
        if package is None or not self.part_name or not package.has(self.part_name):
            return True
        root = re.search(r"<worksheet\b[^>]*>", package.read(self.part_name).decode("utf-8", errors="replace"))
        return root is not None and "xmlns:xr=" in root.group(0)

    def drawing_changed(self) -> None:
        """A shape was added, moved, renamed or deleted."""
        self.drawing_dirty = True
        self.book.saved = False

    def column_width_points(self, column: int) -> float:
        """How wide a column is in points, which a shape's cell needs: none when hidden."""
        return self.dims.column_shown_pixels(column) * _dimensions.PIXEL

    def row_height_points(self, row: int) -> float:
        """How tall a row is in points: none when hidden."""
        return self.dims.row_shown_pixels(row) * _dimensions.PIXEL

    def drawing_grid(self) -> Any:
        """The sheet's columns and rows, as the drawing layer wants them.

        The model counts columns and rows from one; a drawing counts from
        zero and works in points.
        """
        from pyopenvba.shapes._xlsx import SheetGrid

        dims = self.dims
        return SheetGrid(
            column_widths={column - 1: self.column_width_points(column) for column in dims.columns},
            row_heights={row - 1: self.row_height_points(row) for row in dims.rows},
            default_column=dims.standard_width() * _dimensions.PIXEL,
            default_row=dims.default_row_pixels() * _dimensions.PIXEL,
        )

    def cell_changed(self, row: int, column: int) -> None:
        """One cell's contents changed: tell the calculator what to redo."""
        self.touched()
        self.dims.fonts_changed(row)
        calculator = self.book.calculator
        cell = self.cells_.get((row, column))
        if cell is not None:
            calculator.remember(self.name, row, column, cell)
        else:
            calculator.compiled.pop((self.name.lower(), row, column), None)
        calculator.wrote(self.name, row, column)
        if cell is None or not cell.formula:
            from pyopenvba.apps.excel._controls import cell_changed

            cell_changed(self, row, column, cell.value if cell is not None else EMPTY)

    def shape_changed(self) -> None:
        """Rows or columns moved, so every formula has to be read again."""
        self.touched()
        self.book.calculator.rebuild()

    # -- identity

    @member
    def Name(self) -> object:
        return self.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse_structure

        refuse_structure(self.book, APPLICATION_DEFINED)
        wanted = to_text(value)
        if any(sheet is not self and sheet.name.lower() == wanted.lower() for sheet in self.book.sheets_):
            raise error(1004, f"a sheet called {wanted} is already there")
        self.name = wanted
        self.book.saved = False

    @member
    def Index(self) -> object:
        return VBAInt(self.book.sheets_.index(self) + 1, "Long")

    @member
    def Next(self) -> object:
        """The sheet after this one, or Nothing after the last."""
        index = self.book.sheets_.index(self)
        return self.book.sheets_[index + 1] if index + 1 < len(self.book.sheets_) else NOTHING

    @member
    def Previous(self) -> object:
        """The sheet before this one, or Nothing before the first."""
        index = self.book.sheets_.index(self)
        return self.book.sheets_[index - 1] if index > 0 else NOTHING

    @member
    def Visible(self) -> object:
        return VBAInt(self.visible, "Long")

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import HIDING, refuse_structure

        refuse_structure(self.book, HIDING)
        self.visible = -1 if value is True else (0 if value is False else int(to_integer(value, "Long")))

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Type(self) -> object:
        return VBAInt(-4167, "Long")  # xlWorksheet

    # -- the grid

    @member
    def Range(self, Cell1: object = MISSING, Cell2: object = MISSING) -> object:
        if Cell1 is MISSING:
            raise error(449)
        first = self._as_area(Cell1)
        # A name or a qualified reference can point at another sheet,
        # and the Range then belongs to that sheet, not to this one.
        owner = self.book.sheet_named(first[0].sheet) if first[0].sheet else self
        if Cell2 is MISSING:
            return Range(owner, first)
        second = self._as_area(Cell2)
        top = min(first[0].top, second[0].top)
        left = min(first[0].left, second[0].left)
        bottom = max(first[-1].bottom, second[-1].bottom)
        right = max(first[-1].right, second[-1].right)
        return Range(owner, [Area(top, left, bottom, right, owner.name)])

    def _as_area(self, value: object) -> list[Area]:
        if isinstance(value, Range):
            return list(value.areas)
        text = to_text(value)
        named = self.book.names_.find(text, scope=self)
        if named is not None:
            return list(named.refers_to_range().areas)
        try:
            return parse_reference(text, sheet=self.name)
        except ValueError:
            raise error(1004, f"{text!r} is not a reference this sheet knows") from None

    @member
    def Shapes(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._shapes import Shapes as ShapesCollection

        shapes = ShapesCollection(self)
        return shapes if Index is MISSING else shapes.vba_get("Item", [Index])

    @member
    def DrawingObjects(self, Index: object = MISSING) -> object:
        """The older name for the same shapes."""
        return self.Shapes(Index)

    @member
    def Cells(self, RowIndex: object = MISSING, ColumnIndex: object = MISSING) -> object:
        whole = Range(self, [Area(1, 1, MAX_ROWS, MAX_COLUMNS, self.name)])
        if RowIndex is MISSING:
            return whole
        return whole.vba_get("Item", [RowIndex, ColumnIndex])

    @member
    def Rows(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            return Range(self, [Area(1, 1, MAX_ROWS, MAX_COLUMNS, self.name)], whole="rows")
        if isinstance(Index, str):
            # Rows("3") is the whole row, the way Rows("3:5") is.
            wanted = Index if ":" in Index else f"{Index}:{Index}"
            return Range(self, parse_reference(wanted, sheet=self.name), whole="rows")
        row = int(to_integer(Index, "Long"))
        return Range(self, [Area(row, 1, row, MAX_COLUMNS, self.name)], whole="rows")

    @member
    def Columns(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            return Range(self, [Area(1, 1, MAX_ROWS, MAX_COLUMNS, self.name)], whole="columns")
        if isinstance(Index, str):
            # Columns("C") is the whole column, the way Columns("C:D") is.
            wanted = Index if ":" in Index else f"{Index}:{Index}"
            return Range(self, parse_reference(wanted, sheet=self.name), whole="columns")
        column = int(to_integer(Index, "Long"))
        return Range(self, [Area(1, column, MAX_ROWS, column, self.name)], whole="columns")

    @member
    def UsedRange(self) -> object:
        bounds = self.used_bounds()
        if bounds is None:
            return Range(self, [Area(1, 1, 1, 1, self.name)])
        top, left, bottom, right = bounds
        return Range(self, [Area(top, left, bottom, right, self.name)])

    @member
    def StandardHeight(self) -> object:
        from pyopenvba.apps.excel._dimensions import read_standard_height

        return read_standard_height(self)

    @setter("StandardHeight")
    def _set_standard_height(self, value: object) -> None:
        raise error(1004, "StandardHeight is worked out from the Normal style's font")

    @member
    def StandardWidth(self) -> object:
        from pyopenvba.apps.excel._dimensions import read_standard_width

        return read_standard_width(self)

    @setter("StandardWidth")
    def _set_standard_width(self, value: object) -> None:
        from pyopenvba.apps.excel._dimensions import write_standard_width

        write_standard_width(self, value)

    @member
    def Names(self, Index: object = MISSING) -> object:
        names = Names(self.book, self)
        return names if Index is MISSING else names.vba_get("Item", [Index])

    @member
    def Sort(self) -> object:
        """The sheet's sort settings, as a recorded macro sets and applies them."""
        from pyopenvba.apps.excel._sort import SortObject

        return SortObject(self)

    @member
    def AutoFilter(self) -> object:
        from pyopenvba.apps.excel._autofilter import sheet_autofilter

        return sheet_autofilter(self)

    @member
    def AutoFilterMode(self) -> object:
        return self.auto_filter is not None

    @setter("AutoFilterMode")
    def _set_auto_filter_mode(self, value: object) -> None:
        from pyopenvba.apps.excel._autofilter import set_mode

        set_mode(self, value)

    @member
    def FilterMode(self) -> object:
        return self.auto_filter is not None and bool(self.auto_filter.fields)

    @method
    def ShowAllData(self) -> object:
        from pyopenvba.apps.excel._autofilter import show_all_data
        from pyopenvba.apps.excel._protection import refuse

        refuse(self, "ShowAllData method of Worksheet class failed")
        show_all_data(self)
        return EMPTY

    # -- methods

    @method
    def Activate(self) -> object:
        self.book.activate_sheet(self)
        self.book.application.activate_book(self.book)
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:
        return self.Activate()

    @method
    def Paste(self, Destination: object = MISSING, Link: object = MISSING) -> object:
        """What Copy or Cut put on the clipboard, at Destination or the selection."""
        from pyopenvba.apps.excel._clipboard import paste

        return paste(self, Destination, Link)

    @method
    def Delete(self) -> object:
        from pyopenvba.apps.excel._protection import DELETING, refuse_structure

        refuse_structure(self.book, DELETING)
        self.book.sheets_.remove(self)
        self.book.saved = False
        return EMPTY

    @method
    def Copy(self, Before: object = MISSING, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import COPYING
        from pyopenvba.apps.excel._sheet_copy import copy_sheet

        self._guard_structure(Before, After, COPYING, "Copying")
        return copy_sheet(self, Before, After)

    @method
    def Move(self, Before: object = MISSING, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import MOVING
        from pyopenvba.apps.excel._sheet_move import move_sheet

        self._guard_structure(Before, After, MOVING, "Moving")
        return move_sheet(self, Before, After)

    def _guard_structure(self, before: object, after: object, message: str, what: str) -> None:
        """Within one protected workbook, error 1004 with ``message``; to or from another while one is protected,
        which was not measured, a report."""
        from pyopenvba.apps.excel._protection import refuse_structure

        anchor = before if isinstance(before, Worksheet) else after if isinstance(after, Worksheet) else None
        if anchor is not None and anchor.book is self.book:
            refuse_structure(self.book, message)
        elif self.book.protection is not None or (anchor is not None and anchor.book.protection is not None):
            raise VBAUnsupportedError(f"{what} a sheet between workbooks, one of them protected, is not implemented")

    @method
    def Calculate(self) -> object:
        self.book.calculator.calculate_all()
        return EMPTY

    @method
    def Evaluate(self, Name: object = MISSING) -> object:
        from pyopenvba.apps.excel._evaluate import evaluated

        return evaluated(self, to_text(Name))

    # -- protection (see _protection)

    @method
    def Protect(self, Password: object = MISSING, DrawingObjects: object = MISSING, Contents: object = MISSING,
                Scenarios: object = MISSING, UserInterfaceOnly: object = MISSING,
                AllowFormattingCells: object = MISSING, AllowFormattingColumns: object = MISSING,
                AllowFormattingRows: object = MISSING, AllowInsertingColumns: object = MISSING,
                AllowInsertingRows: object = MISSING, AllowInsertingHyperlinks: object = MISSING,
                AllowDeletingColumns: object = MISSING, AllowDeletingRows: object = MISSING,
                AllowSorting: object = MISSING, AllowFiltering: object = MISSING,
                AllowUsingPivotTables: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import protect

        options = {"DrawingObjects": DrawingObjects, "Contents": Contents, "Scenarios": Scenarios,
                   "UserInterfaceOnly": UserInterfaceOnly, "AllowFormattingCells": AllowFormattingCells,
                   "AllowFormattingColumns": AllowFormattingColumns, "AllowFormattingRows": AllowFormattingRows,
                   "AllowInsertingColumns": AllowInsertingColumns, "AllowInsertingRows": AllowInsertingRows,
                   "AllowInsertingHyperlinks": AllowInsertingHyperlinks, "AllowDeletingColumns": AllowDeletingColumns,
                   "AllowDeletingRows": AllowDeletingRows, "AllowSorting": AllowSorting,
                   "AllowFiltering": AllowFiltering, "AllowUsingPivotTables": AllowUsingPivotTables}
        protect(self, Password, options)
        return EMPTY

    @method
    def Unprotect(self, Password: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import unprotect

        unprotect(self, Password)
        return EMPTY

    @member
    def ProtectContents(self) -> object:
        return self.protection is not None and self.protection.contents

    @member
    def ProtectDrawingObjects(self) -> object:
        return self.protection is not None and self.protection.drawing_objects

    @member
    def ProtectScenarios(self) -> object:
        return self.protection is not None and self.protection.scenarios

    @member
    def ProtectionMode(self) -> object:
        """True while the sheet is protected for the user interface only, which leaves macros free."""
        return self.protection is not None and self.protection.user_interface_only

    @member
    def Protection(self) -> object:
        from pyopenvba.apps.excel._protection import Protection

        return Protection(self)

    @member
    def EnableSelection(self) -> object:
        return VBAInt(self.enable_selection, "Long")

    @setter("EnableSelection")
    def _set_enable_selection(self, value: object) -> None:
        self.enable_selection = int(to_integer(value, "Long"))

    # -- Python side

    @property
    def selection(self) -> object:
        return self.selection_range if self.selection_range is not None else self.vba_get("Range", ["A1"])

    @property
    def active_cell(self) -> object:
        return self.active_cell_range if self.active_cell_range is not None else self.vba_get("Range", ["A1"])

    def cell(self, row: int, column: int, *, create: bool = False) -> Cell | None:
        """The cell at a position; a new one starts in the format its row or column gives it."""
        found = self.cells_.get((row, column))
        if found is None and create:
            inherited = self.inherited_style(row, column)
            found = Cell(style=None if inherited == self.book.stylesheet.default else inherited)
            self.cells_[(row, column)] = found
        return found

    def inherited_style(self, row: int, column: int) -> Style:
        """The format a position shows with no cell there: its row's, else its column's, else the default."""
        record = self.dims.rows.get(row)
        if record is not None and record.style is not None:
            return record.style
        return self.dims.column_style(column) or self.book.stylesheet.default

    def style_at(self, row: int, column: int) -> Style:
        """The format a cell shows: its own, or the one its row or column gives it."""
        found = self.cells_.get((row, column))
        if found is not None:
            return found.style or self.book.stylesheet.default
        return self.inherited_style(row, column)

    def holds(self, row: int, column: int, cell: Cell) -> bool:
        """Whether a cell says anything: a value, a formula, or a format apart from the one it would inherit.

        Excel keeps a cell with nothing in it only while its format differs
        from its row's or column's, and lets it go the moment it does not.
        """
        if cell.value is not EMPTY or cell.formula:
            return True
        return (cell.style or self.book.stylesheet.default) != self.inherited_style(row, column)

    def settle(self, row: int, column: int) -> None:
        """Drop the cell at a position if it no longer says anything."""
        found = self.cells_.get((row, column))
        if found is not None and not self.holds(row, column, found):
            del self.cells_[(row, column)]

    def restyle(self, row: int, column: int, style: Style) -> None:
        """Give a position a new format, keeping a cell there only while it has to."""
        found = self.cells_.get((row, column))
        if found is None:
            if style == self.inherited_style(row, column):
                return
            found = self.cells_[(row, column)] = Cell()
        found.style, found.xf = (None if style == self.book.stylesheet.default else style), -1
        self.settle(row, column)
        self.touched()
        self.dims.fonts_changed(row)

    def set_number_format(self, row: int, column: int, code: str) -> None:
        self.restyle(row, column, applying(self.style_at(row, column), "number_format", number_format=code))

    def used_bounds(self) -> tuple[int, int, int, int] | None:
        """The sheet's used block, as UsedRange and a file's dimension give it.

        A row with a height, a hidden flag or a format of its own counts,
        as its cells do, and so does one a border draws taller; a column
        counts only while it is hidden keeping a width, or while its format
        is apart from the sheet's. What is
        missing on one side defaults to row 1 or column A. Excel's own
        block only grows while a workbook is open, and shrinks when
        something reads UsedRange; the model's is always the content's.
        """
        live = [(row, column) for (row, column), cell in self.cells_.items() if self.holds(row, column, cell)]
        for area in self.merged_areas:
            live.extend(((area.top, area.left), (area.bottom, area.right)))
        rows = [row for row, _ in live]
        rows.extend(self.dims.record_rows())
        rows.extend(self.dims.shaped_rows())
        columns = [column for _, column in live]
        columns.extend(self.dims.used_columns())
        if not rows and not columns:
            return None
        rows, columns = rows or [1], columns or [1]
        return min(rows), min(columns), max(rows), max(columns)

    def describe(self, indent: str = "") -> str:
        bounds = self.used_bounds()
        where = "empty" if bounds is None else Area(*bounds, self.name).address(absolute=False)
        head = f"{indent}Worksheet {self.name!r} ({where})"
        if bounds is None:
            return head
        lines = [head]
        top, left, bottom, right = bounds
        for row in range(top, min(bottom, top + 19) + 1):
            pieces: list[str] = []
            for column in range(left, min(right, left + 9) + 1):
                cell = self.cells_.get((row, column))
                pieces.append(_cell_text(cell))
            lines.append(f"{indent}  {row:>5} | " + " | ".join(piece.ljust(12)[:12] for piece in pieces))
        if bottom > top + 19:
            lines.append(f"{indent}  ... {bottom - top - 19} more rows")
        return "\n".join(lines)


def _cell_text(cell: Cell | None) -> str:
    if cell is None:
        return ""
    if cell.formula and cell.stale:
        return f"{cell.formula} (?)"
    if cell.formula:
        return f"{_short_value(cell.value)} [{cell.formula}]"
    return _short_value(cell.value)


def _short_value(value: object) -> str:
    from pyopenvba.formula._values import ExcelError

    if isinstance(value, ExcelError):
        return value.name
    if value is EMPTY:
        return ""
    if isinstance(value, str):
        return value
    try:
        return to_text(value)
    except Exception:  # noqa: BLE001 - the dump must never fail
        return type_name(value)


# --- Range ---------------------------------------------------------------------------------


class Range(ExcelObject):
    """A rectangle of cells, or several of them."""

    vba_type_name = "Range"

    def __init__(self, sheet: Worksheet, areas: list[Area], *, whole: str = "") -> None:
        self.sheet = sheet
        self.areas = areas or [Area(1, 1, 1, 1, sheet.name)]
        #: "rows" or "columns" for a range made as whole rows or columns.
        #: Every row and every column is the same block of cells, but
        #: Rows("1:1048576").Hidden hides rows and Columns("A:XFD").Hidden
        #: columns, where Cells.Hidden is an error.
        self.whole = whole

    def guard_set(self, member: str) -> None:
        from pyopenvba.apps.excel._protection import check_range_set

        check_range_set(self, member)

    @property
    def first(self) -> Area:
        return self.areas[0]

    @property
    def single(self) -> bool:
        return len(self.areas) == 1 and self.first.rows == 1 and self.first.columns == 1

    def positions(self) -> Iterator[tuple[int, int]]:
        for area in self.areas:
            for row in range(area.top, min(area.bottom, MAX_ROWS) + 1):
                for column in range(area.left, min(area.right, MAX_COLUMNS) + 1):
                    yield row, column

    def bounded(self) -> Area:
        """The first area, clipped to what the sheet actually uses.

        A whole-column reference is 1048576 cells tall; walking it would
        take a minute and mean nothing, so anything that iterates asks
        for this instead.
        """
        area = self.first
        used = self.sheet.used_bounds()
        if used is None:
            return Area(area.top, area.left, min(area.bottom, area.top), min(area.right, area.left), area.sheet)
        return Area(
            max(area.top, 1),
            max(area.left, 1),
            min(area.bottom, max(used[2], area.top)),
            min(area.right, max(used[3], area.left)),
            area.sheet,
        )

    # -- value

    @member(default=True)
    def Value(self, RangeValueDataType: object = MISSING) -> object:
        return self._values(raw=False)

    @setter("Value")
    def _set_value(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import writing

        with writing(self.sheet):
            self._write(value)

    @member
    def Value2(self) -> object:
        """Value without the Date and Currency a cell's format makes of its number."""
        return self._values(raw=True)

    @setter("Value2")
    def _set_value2(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import writing

        with writing(self.sheet):
            self._write(value, raw=True)

    def _values(self, *, raw: bool) -> object:
        if self.single:
            return self._read(self.first.top, self.first.left, raw=raw)
        area = self.bounded()
        items: list[object] = []
        # VBA lays a two-dimensional array out column by column.
        for column in range(area.left, area.right + 1):
            for row in range(area.top, area.bottom + 1):
                items.append(self._read(row, column, raw=raw))
        return VBAArray([(1, area.rows), (1, area.columns)], items=items)

    @member
    def PrefixCharacter(self) -> object:
        """The apostrophe a cell was typed with, while it holds text; a format keeps it, so a number hides it."""
        cell = self.sheet.cell(self.first.top, self.first.left)
        if cell is None or cell.formula or not isinstance(cell.value, str):
            return ""
        return "'" if cell.style is not None and cell.style.quote_prefix else ""

    @member
    def Text(self) -> object:
        cell = self.sheet.cell(self.first.top, self.first.left)
        if cell is None:
            return ""
        # Reading what a cell shows calculates it first, as looking at
        # one in Excel does; read raw, since a Date too large to hold still shows.
        self._read(self.first.top, self.first.left, raw=True)
        dims = self.sheet.dims
        return _display_text(cell, _dimensions.characters_read(dims.column_shown_pixels(self.first.left)))

    @member
    def Formula(self) -> object:
        return self._read_formulas(r1c1=False)

    @member
    def Name(self) -> object:
        for entry in self.sheet.book.names_.entries:
            named = DefinedName(entry)
            try:
                target = named.refers_to_range()
            except (VBARuntimeError, VBAUnsupportedError, ValueError):
                continue
            if target.sheet is self.sheet and [(a.top, a.left, a.bottom, a.right) for a in target.areas] == [(a.top, a.left, a.bottom, a.right) for a in self.areas]:
                return named
        raise error(1004, "This range has no defined name")

    @setter("Name")
    def _set_range_name(self, value: object) -> None:
        self.sheet.book.names_.Add(Name=value, RefersTo=self)

    @setter("Formula")
    def _set_formula(self, value: object) -> None:
        """A formula written to a block moves with each cell.

        Excel anchors what was written at the top left and shifts every
        reference that is not held by a dollar sign, which is why
        Range("D2:D6").Formula = "=B2*C2" leaves =B6*C6 in D6. What is not
        a formula is typed as Value types it, and a Text cell keeps even a
        formula as the text it is.
        """
        from pyopenvba.apps.excel._protection import writing

        with writing(self.sheet):
            self._write_formula(value)

    def _write_formula(self, value: object) -> None:
        if isinstance(value, VBAArray):
            self._write_formula_array(value, r1c1=False)
            return
        if isinstance(value, VBAObject):
            value = value.vba_value()
        if not isinstance(value, str):
            self._write(value)
            return
        anchor = self.first
        spelled: str | None = None
        placed: list[tuple[int, int]] = []
        for row, column in self.writable_positions():
            if not _merges.writable(self.sheet, row, column):
                continue
            if _is_formula(value) and not self._keeps_text(row, column):
                if spelled is None:
                    # Excel reads the formula once and writes it out again, as spelled_formula does.
                    spelled = spelled_formula(self.sheet, value)
                if self._put_formula(row, column, shift_text(spelled, row - anchor.top, column - anchor.left)):
                    placed.append((row, column))
            else:
                self._type_into(row, column, value)
        if spelled is not None:
            self._bring_format(placed, spelled, anchor.top, anchor.left)

    @member
    def HasFormula(self) -> object:
        """True when every cell holds a formula, False when none does, and Null for a mix."""
        area = self.first if self.single else self.bounded()
        found = {bool(cell is not None and cell.formula) for row in range(area.top, area.bottom + 1)
                 for column in range(area.left, area.right + 1) for cell in [self.sheet.cell(row, column)]}
        return NULL if len(found) > 1 else found.pop()

    @member
    def FormulaR1C1(self) -> object:
        return self._read_formulas(r1c1=True)

    def _read_formulas(self, *, r1c1: bool) -> object:
        def read(row: int, column: int) -> object:
            cell = self.sheet.cell(row, column)
            if cell is None:
                return ""
            if cell.formula:
                return from_a1(cell.formula, row, column) if r1c1 else cell.formula
            return _formula_text(cell.value)

        area = self.first
        if self.single:
            return read(area.top, area.left)
        return VBAArray([(1, area.rows), (1, area.columns)], items=[
            read(row, column) for column in range(area.left, area.right + 1)
            for row in range(area.top, area.bottom + 1)
        ])

    @setter("FormulaR1C1")
    def _set_formula_r1c1(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import writing

        with writing(self.sheet):
            self._write_formula_r1c1(value)

    def _write_formula_r1c1(self, value: object) -> None:
        if isinstance(value, VBAArray):
            self._write_formula_array(value, r1c1=True)
            return
        if isinstance(value, VBAObject):
            value = value.vba_value()
        if not isinstance(value, str):
            self._write(value)
            return
        placed: list[tuple[int, int]] = []
        first: tuple[str, int, int] | None = None
        for row, column in self.writable_positions():
            if not _merges.writable(self.sheet, row, column):
                continue
            if not _is_formula(value) or self._keeps_text(row, column):
                self._type_into(row, column, value)
                continue
            try:
                a1 = to_a1(value, row, column)
            except ValueError as exc:
                raise error(1004, str(exc)) from None
            formula = spelled_formula(self.sheet, a1)
            first = first or (formula, row, column)
            if self._put_formula(row, column, formula):
                placed.append((row, column))
        if first is not None:
            self._bring_format(placed, *first)

    def _write_formula_array(self, array: VBAArray, *, r1c1: bool) -> None:
        from pyopenvba.formula._values import NA

        if array.dimensions not in (1, 2):
            raise error(13, "Formula assignment requires a one- or two-dimensional array")
        if not array.size:
            return
        source_rows = array.bounds[0][1] - array.bounds[0][0] + 1 if array.dimensions == 2 else 1
        source_columns = array.bounds[-1][1] - array.bounds[-1][0] + 1
        # Validate the existing write-size limit before changing any area.
        self.writable_positions()
        for area in self._array_areas(array):
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    if not _merges.writable(self.sheet, row, column):
                        continue
                    down, across = row - area.top, column - area.left
                    source_row = 0 if source_rows == 1 else down
                    source_column = 0 if source_columns == 1 else across
                    if source_row >= source_rows or source_column >= source_columns:
                        self._put(row, column, NA)
                        continue
                    indices = [array.bounds[-1][0] + source_column]
                    if array.dimensions == 2:
                        indices.insert(0, array.bounds[0][0] + source_row)
                    item = array.get(indices)
                    if item is NULL:
                        item = EMPTY
                    if isinstance(item, VBAErrorValue):
                        from pyopenvba.apps.excel._calc import from_vba

                        item = from_vba(item)
                    if isinstance(item, str) and item.startswith("="):
                        # A1 repeats a singleton axis by shifting the formula.
                        # R1C1 also resolves at each destination, except that a
                        # one-element array behaves like a scalar assignment.
                        extra_row = down if source_rows == 1 else 0
                        extra_column = across if source_columns == 1 else 0
                        if r1c1 and array.size == 1:
                            extra_row = extra_column = 0
                        try:
                            item = (to_a1(item, row + extra_row, column + extra_column) if r1c1
                                    else shift_text(item, extra_row, extra_column))
                        except ValueError as exc:
                            raise error(1004, str(exc)) from None
                    self._put(row, column, item)

    @member
    def NumberFormat(self) -> object:
        from pyopenvba.apps.excel._formats import styles_of, uniform

        return uniform(style.number_format for style in styles_of(self))

    @setter("NumberFormat")
    def _set_number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._formats import restyle
        from pyopenvba.apps.excel._number_format import normalized

        text = normalized(to_text(value))
        restyle(self, lambda style: applying(style, "number_format", number_format=text))

    # The Local spellings are in the language of the user's settings, which the model keeps English: the same
    # text as the plain ones, as an English Excel has them (tests/fixtures/excel_model/probes.txt).

    @member
    def NumberFormatLocal(self) -> object:
        return self.NumberFormat()

    @setter("NumberFormatLocal")
    def _set_number_format_local(self, value: object) -> None:
        self._set_number_format(value)

    @member
    def FormulaLocal(self) -> object:
        return self._read_formulas(r1c1=False)

    @setter("FormulaLocal")
    def _set_formula_local(self, value: object) -> None:
        self._set_formula(value)

    @member
    def FormulaR1C1Local(self) -> object:
        return self._read_formulas(r1c1=True)

    @setter("FormulaR1C1Local")
    def _set_formula_r1c1_local(self, value: object) -> None:
        self._set_formula_r1c1(value)
        self.sheet.touched()

    # -- geometry

    @member
    def Address(
        self,
        RowAbsolute: object = MISSING,
        ColumnAbsolute: object = MISSING,
        ReferenceStyle: object = MISSING,
        External: object = MISSING,
        RelativeTo: object = MISSING,
    ) -> object:
        rows_fixed = True if RowAbsolute is MISSING else to_bool(RowAbsolute)
        columns_fixed = True if ColumnAbsolute is MISSING else to_bool(ColumnAbsolute)
        external = External is not MISSING and to_bool(External)
        book = f"[{self.sheet.book.name}]" if external else ""
        texts = [area.address(rows_fixed=rows_fixed, columns_fixed=columns_fixed, with_sheet=external)
                 for area in self.areas]
        if ReferenceStyle is not MISSING and int(to_integer(ReferenceStyle, "Long")) == -4150:  # xlR1C1
            # Relative parts count from RelativeTo, and from A1 without it, whatever is selected.
            anchor = RelativeTo.first if isinstance(RelativeTo, Range) else Area(1, 1, 1, 1)
            texts = [from_a1("=" + text, anchor.top, anchor.left)[1:] for text in texts]
        return ",".join(book + text for text in texts)

    @member
    def AddressLocal(
        self,
        RowAbsolute: object = MISSING,
        ColumnAbsolute: object = MISSING,
        ReferenceStyle: object = MISSING,
        External: object = MISSING,
        RelativeTo: object = MISSING,
    ) -> object:
        """Address in the language of the user's settings, which the model keeps English: the same text."""
        return self.Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)

    @member
    def Row(self) -> object:
        return VBAInt(self.first.top, "Long")

    @member
    def Column(self) -> object:
        return VBAInt(self.first.left, "Long")

    @member
    def Count(self) -> object:
        """The rows of whole rows, the columns of whole columns, else the cells: Rows.Count is 1048576.

        A count past a Long is error 6, as Cells.Count is.
        """
        total = self._count()
        if total > 2147483647:
            raise error(6, "Overflow")
        return VBAInt(total, "Long")

    @member
    def CountLarge(self) -> object:
        """Count as a LongLong, which holds a whole sheet: Cells.CountLarge is 17179869184."""
        return VBAInt(self._count(), "LongLong")

    def _count(self) -> int:
        if self.whole == "rows":
            return sum(area.rows for area in self.areas)
        if self.whole == "columns":
            return sum(area.columns for area in self.areas)
        return sum(area.rows * area.columns for area in self.areas)

    @member
    def Next(self) -> object:
        """The cell right of the range's top-left one; past the last column, error 1004."""
        return self._beside(1)

    @member
    def Previous(self) -> object:
        """The cell left of the range's top-left one; before column A, error 1004."""
        return self._beside(-1)

    def _beside(self, step: int) -> Range:
        from pyopenvba.apps.excel._protection import enforced

        if enforced(self.sheet) is not None:
            raise VBAUnsupportedError("Next and Previous on a protected sheet, where they move between its unlocked "
                                      "cells, are not implemented")
        area = self.first
        column = area.left + step
        if not 1 <= column <= MAX_COLUMNS:
            raise error(1004, "Application-defined or object-defined error")
        return Range(self.sheet, [Area(area.top, column, area.top, column, self.sheet.name)])

    @method
    def Calculate(self) -> object:
        """Work the range's formulas out again; Excel answers Null."""
        calculator = self.sheet.book.calculator
        for row, column in self.cell_positions():
            if self.sheet.cells_[(row, column)].formula:
                calculator.value_of(self.sheet.name, row, column, force=True)
        return NULL

    @member
    def Rows(self, Index: object = MISSING) -> object:
        area = self.first
        if Index is MISSING:
            return RowsOf(self)
        row = area.top + int(to_integer(Index, "Long")) - 1
        return Range(self.sheet, [Area(row, area.left, row, area.right, area.sheet)])

    @member
    def Columns(self, Index: object = MISSING) -> object:
        area = self.first
        if Index is MISSING:
            return ColumnsOf(self)
        column = area.left + int(to_integer(Index, "Long")) - 1
        return Range(self.sheet, [Area(area.top, column, area.bottom, column, area.sheet)])

    @member
    def Cells(self, RowIndex: object = MISSING, ColumnIndex: object = MISSING) -> object:
        if RowIndex is MISSING:
            return self
        return self.vba_get("Item", [RowIndex, ColumnIndex])

    @member
    def Item(self, RowIndex: object = MISSING, ColumnIndex: object = MISSING) -> object:
        area = self.first
        if RowIndex is MISSING:
            raise error(449)
        if ColumnIndex is MISSING or ColumnIndex is NOTHING:
            # One index walks the range left to right, then down.
            position = int(to_integer(RowIndex, "Long")) - 1
            if position < 0:
                raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
            row = area.top + position // area.columns
            column = area.left + position % area.columns
        else:
            row = area.top + int(to_integer(RowIndex, "Long")) - 1
            if isinstance(ColumnIndex, str):
                from pyopenvba._a1 import column_number

                column = area.left + column_number(ColumnIndex) - 1
            else:
                column = area.left + int(to_integer(ColumnIndex, "Long")) - 1
        if row < 1 or column < 1 or row > MAX_ROWS or column > MAX_COLUMNS:
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        return Range(self.sheet, [Area(row, column, row, column, area.sheet)])

    @member
    def Offset(self, RowOffset: object = MISSING, ColumnOffset: object = MISSING) -> object:
        down = 0 if RowOffset is MISSING else int(to_integer(RowOffset, "Long"))
        across = 0 if ColumnOffset is MISSING else int(to_integer(ColumnOffset, "Long"))
        areas: list[Area] = []
        for area in self.areas:
            top, left = area.top + down, area.left + across
            bottom, right = area.bottom + down, area.right + across
            if top < 1 or left < 1 or bottom > MAX_ROWS or right > MAX_COLUMNS:
                raise error(1004, "Offset would leave the sheet")
            areas.append(Area(top, left, bottom, right, area.sheet))
        return Range(self.sheet, areas)

    @member
    def Resize(self, RowSize: object = MISSING, ColumnSize: object = MISSING) -> object:
        area = self.first
        rows = area.rows if RowSize is MISSING else int(to_integer(RowSize, "Long"))
        columns = area.columns if ColumnSize is MISSING else int(to_integer(ColumnSize, "Long"))
        if rows < 1 or columns < 1:
            raise error(ERR_APPLICATION_DEFINED, "Resize needs at least one row and column")
        return Range(self.sheet, [Area(area.top, area.left, area.top + rows - 1, area.left + columns - 1, area.sheet)])

    @member
    def EntireRow(self) -> object:
        return Range(
            self.sheet,
            [Area(area.top, 1, area.bottom, MAX_COLUMNS, area.sheet) for area in self.areas],
            whole="rows",
        )

    @member
    def EntireColumn(self) -> object:
        return Range(
            self.sheet,
            [Area(1, area.left, MAX_ROWS, area.right, area.sheet) for area in self.areas],
            whole="columns",
        )

    @member
    def Areas(self, Index: object = MISSING) -> object:
        areas = Areas(self)
        return areas if Index is MISSING else areas.vba_get("Item", [Index])

    @member
    def Worksheet(self) -> object:
        return self.sheet

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Font(self) -> object:
        from pyopenvba.apps.excel._formats import Font

        return Font(self)

    @member
    def Interior(self) -> object:
        from pyopenvba.apps.excel._formats import Interior

        return Interior(self)

    @member
    def Borders(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._formats import Borders

        borders = Borders(self)
        return borders if Index is MISSING else borders.vba_get("Item", [Index])

    @method
    def BorderAround(self, LineStyle: object = MISSING, Weight: object = MISSING, ColorIndex: object = MISSING,
                     Color: object = MISSING, ThemeColor: object = MISSING) -> object:
        from pyopenvba.apps.excel._formats import border_around
        from pyopenvba.apps.excel._protection import refuse

        refuse(self.sheet, "BorderAround method of Range class failed", "AllowFormattingCells")
        border_around(self, LineStyle, Weight, ColorIndex, Color, ThemeColor)
        return EMPTY

    @method
    def ClearFormats(self) -> object:
        from pyopenvba.apps.excel._autofilter import header_cleared
        from pyopenvba.apps.excel._formats import clear_formats
        from pyopenvba.apps.excel._protection import whole
        from pyopenvba.apps.excel._visible import visible

        whole(self.sheet, visible(self).areas, "ClearFormats")
        clear_formats(self)
        header_cleared(visible(self))
        return EMPTY

    def _format(self, what: str) -> object:
        from pyopenvba.apps.excel._formats import read_alignment

        return read_alignment(self, what)

    def _set_format(self, what: str, value: object) -> None:
        from pyopenvba.apps.excel._formats import write_alignment

        write_alignment(self, what, value)

    @member
    def HorizontalAlignment(self) -> object:
        return self._format("HorizontalAlignment")

    @setter("HorizontalAlignment")
    def _set_horizontal_alignment(self, value: object) -> None:
        self._set_format("HorizontalAlignment", value)

    @member
    def VerticalAlignment(self) -> object:
        return self._format("VerticalAlignment")

    @setter("VerticalAlignment")
    def _set_vertical_alignment(self, value: object) -> None:
        self._set_format("VerticalAlignment", value)

    @member
    def WrapText(self) -> object:
        return self._format("WrapText")

    @setter("WrapText")
    def _set_wrap_text(self, value: object) -> None:
        self._set_format("WrapText", value)

    @member
    def ShrinkToFit(self) -> object:
        return self._format("ShrinkToFit")

    @setter("ShrinkToFit")
    def _set_shrink_to_fit(self, value: object) -> None:
        self._set_format("ShrinkToFit", value)

    @member
    def IndentLevel(self) -> object:
        return self._format("IndentLevel")

    @setter("IndentLevel")
    def _set_indent_level(self, value: object) -> None:
        self._set_format("IndentLevel", value)

    @member
    def AddIndent(self) -> object:
        return self._format("AddIndent")

    @setter("AddIndent")
    def _set_add_indent(self, value: object) -> None:
        self._set_format("AddIndent", value)

    @member
    def Orientation(self) -> object:
        return self._format("Orientation")

    @setter("Orientation")
    def _set_orientation(self, value: object) -> None:
        self._set_format("Orientation", value)

    @member
    def ReadingOrder(self) -> object:
        return self._format("ReadingOrder")

    @setter("ReadingOrder")
    def _set_reading_order(self, value: object) -> None:
        self._set_format("ReadingOrder", value)

    @member
    def Locked(self) -> object:
        return self._format("Locked")

    @setter("Locked")
    def _set_locked(self, value: object) -> None:
        self._set_format("Locked", value)

    @member
    def FormulaHidden(self) -> object:
        return self._format("FormulaHidden")

    @setter("FormulaHidden")
    def _set_formula_hidden(self, value: object) -> None:
        self._set_format("FormulaHidden", value)

    # -- sizes: pyopenvba.apps.excel._dimensions has the rules

    @member
    def ColumnWidth(self) -> object:
        return _dimensions.read_column_width(self)

    @setter("ColumnWidth")
    def _set_column_width(self, value: object) -> None:
        _dimensions.write_column_width(self, value)

    @member
    def RowHeight(self) -> object:
        return _dimensions.read_row_height(self)

    @setter("RowHeight")
    def _set_row_height(self, value: object) -> None:
        _dimensions.write_row_height(self, value)

    @member
    def Width(self) -> object:
        return _dimensions.read_width(self)

    @member
    def Height(self) -> object:
        return _dimensions.read_height(self)

    @member
    def Left(self) -> object:
        return _dimensions.read_left(self)

    @member
    def Top(self) -> object:
        return _dimensions.read_top(self)

    @member
    def Hidden(self) -> object:
        return _dimensions.read_hidden(self)

    @setter("Hidden")
    def _set_hidden(self, value: object) -> None:
        _dimensions.write_hidden(self, value)

    @member
    def UseStandardHeight(self) -> object:
        return _dimensions.read_use_standard_height(self)

    @setter("UseStandardHeight")
    def _set_use_standard_height(self, value: object) -> None:
        _dimensions.write_use_standard_height(self, value)

    @member
    def UseStandardWidth(self) -> object:
        return _dimensions.read_use_standard_width(self)

    @setter("UseStandardWidth")
    def _set_use_standard_width(self, value: object) -> None:
        _dimensions.write_use_standard_width(self, value)

    # -- methods

    @method
    def Select(self) -> object:
        self.sheet.selection_range = self
        self.sheet.active_cell_range = Range(
            self.sheet, [Area(self.first.top, self.first.left, self.first.top, self.first.left, self.first.sheet)]
        )
        self.sheet.book.activate_sheet(self.sheet)
        return EMPTY

    @method
    def Activate(self) -> object:
        return self.Select()

    @method
    def Clear(self) -> object:
        """ClearFormats and ClearContents: a position a row or column format reaches keeps a cell in the default."""
        from pyopenvba.apps.excel._autofilter import header_cleared
        from pyopenvba.apps.excel._formats import clear_formats
        from pyopenvba.apps.excel._visible import visible

        from pyopenvba.apps.excel._protection import enforced

        target = visible(self)
        _merges.validate_clear(target)
        if enforced(self.sheet) is not None:
            # On a protected sheet Clear clears the contents alone, as ClearContents would.
            return self.ClearContents()
        clear_formats(target)
        target._clear_contents()
        header_cleared(target)
        return EMPTY

    @method
    def ClearContents(self) -> object:
        from pyopenvba.apps.excel._autofilter import header_cleared
        from pyopenvba.apps.excel._protection import PROTECTED, any_locked, enforced
        from pyopenvba.apps.excel._visible import visible

        target = visible(self)
        _merges.validate_clear(target)
        if enforced(self.sheet) is not None and any_locked(self.sheet, target.areas):
            # Unlike a write, a clear that meets a locked cell leaves every cell as it was.
            raise error(1004, PROTECTED)
        target._clear_contents()
        header_cleared(target)
        return EMPTY

    def _clear_contents(self) -> None:
        """Every value and formula gone; a cell left with nothing to say goes with them."""
        for row, column in self.cell_positions():
            cell = self.sheet.cells_[(row, column)]
            cell.value = EMPTY
            cell.formula = ""
            cell.stale = False
            self.sheet.settle(row, column)
            self.sheet.cell_changed(row, column)
        self.sheet.touched()

    def cell_positions(self) -> list[tuple[int, int]]:
        """The positions of the range that hold a cell, found without visiting every position of a large range."""
        cells = self.sheet.cells_
        out: list[tuple[int, int]] = []
        for area in self.areas:
            if area.rows * area.columns <= len(cells):
                out.extend((row, column) for row in range(area.top, area.bottom + 1)
                           for column in range(area.left, area.right + 1) if (row, column) in cells)
            else:
                out.extend(sorted(position for position in cells if area.contains(*position)))
        return out

    @method
    def Delete(self, Shift: object = MISSING) -> object:
        """Delete, which pulls the cells below or to the right up or left.

        With no Shift a range taller than it is wide pulls cells left and
        any other range pulls them up; _editing has how references follow.
        On a filtered sheet, as _visible has it, a delete that would pull
        cells up takes the whole of each visible row instead, anywhere on
        the sheet, with no Shift too, and one that pulls them left takes
        the visible cells.
        """
        from pyopenvba.apps.excel._protection import check_delete
        from pyopenvba.apps.excel._visible import filtering

        check_delete(self)
        filtered = filtering(self.sheet) and not all(area.whole_columns for area in self.areas)
        if Shift is MISSING:
            up = filtered or self.first.rows <= self.first.columns
        else:
            up = int(to_integer(Shift, "Long")) == -4162  # xlUp
        if filtered:
            self._delete_filtered(up)
        elif all(area.whole_rows for area in self.areas) or all(area.whole_columns for area in self.areas):
            self._delete_lines()
        elif len(self.areas) != 1:
            raise VBAUnsupportedError("deleting cells in several areas is not implemented")
        else:
            self._delete_cells(self.first, up)
        return EMPTY

    def _delete_lines(self) -> None:
        """Whole rows or columns deleted, the last area first, so that each goes from where it was."""
        from pyopenvba.apps.excel._editing import edit

        rows = all(area.whole_rows for area in self.areas)
        spans = sorted((area.top, area.bottom) if rows else (area.left, area.right) for area in self.areas)
        if any(later[0] <= earlier[1] for earlier, later in zip(spans, spans[1:])):
            raise VBAUnsupportedError("deleting overlapping rows or columns is not implemented")
        for low, high in reversed(spans):
            area = Area(low, 1, high, MAX_COLUMNS, self.sheet.name) if rows else \
                Area(1, low, MAX_ROWS, high, self.sheet.name)
            edit(Range(self.sheet, [area], whole=self.whole), delete=True)

    def _delete_filtered(self, up: bool) -> None:
        from pyopenvba.apps.excel._visible import unmeasured, visible_areas, visible_rows

        if not all(area.whole_rows for area in self.areas):
            unmeasured(self, "Deleting cells", lines="columns")
        if up or all(area.whole_rows for area in self.areas):
            runs = visible_rows(self) or [(area.top, area.bottom) for area in self.areas]
            Range(self.sheet, [Area(top, 1, bottom, MAX_COLUMNS, self.sheet.name) for top, bottom in runs],
                  whole="rows")._delete_lines()
            return
        # Each visible block of rows gives up its cells on its own.
        for area in visible_areas(self) or self.areas:
            self._delete_cells(area, False)

    def _delete_cells(self, area: Area, up: bool) -> None:
        from pyopenvba.apps.excel._editing import shift_cells

        self._shifting_cells()
        shift_cells(self, area, delete=True, vertical=up)

    @method
    def Insert(self, Shift: object = MISSING, CopyOrigin: object = MISSING) -> object:
        """Insert, which pushes cells down or right.

        With no Shift a range taller than it is wide pushes cells right and
        any other range pushes them down; _editing has how references
        follow. On a filtered sheet, as _visible has it, whole rows go in
        as many as the range shows, at its top, and inserting cells is
        error 1004.
        """
        from pyopenvba.apps.excel._protection import check_insert
        from pyopenvba.apps.excel._visible import filtering, visible_rows

        check_insert(self)
        area = self.first
        if CopyOrigin is not MISSING and int(to_integer(CopyOrigin, "Long")) != 0:
            raise VBAUnsupportedError("Insert taking formats from the right or below (xlFormatFromRightOrBelow) "
                                      "is not implemented")
        filtered = filtering(self.sheet)
        if area.whole_rows or area.whole_columns:
            from pyopenvba.apps.excel._editing import edit

            runs = visible_rows(self) if filtered and area.whole_rows and len(self.areas) == 1 else None
            if runs is not None:
                count = sum(bottom - top + 1 for top, bottom in runs)
                area = Area(area.top, 1, area.top + count - 1, MAX_COLUMNS, area.sheet)
                edit(Range(self.sheet, [area], whole="rows"), delete=False)
                return EMPTY
            edit(self, delete=False)
            return EMPTY
        if filtered:
            raise error(1004, "Insert method of Range class failed")
        if len(self.areas) != 1:
            raise VBAUnsupportedError("inserting cells in several areas is not implemented")
        from pyopenvba.apps.excel._editing import shift_cells

        self._shifting_cells()
        # With no Shift a range taller than it is wide pushes cells right, any other range down.
        down = area.rows <= area.columns if Shift is MISSING else int(to_integer(Shift, "Long")) == -4121  # xlDown
        shift_cells(self, area, delete=False, vertical=down)
        return EMPTY

    def _shifting_cells(self) -> None:
        """Refuse to shift cells on a sheet whose rows or columns carry formats.

        Excel moves the format each position shows along with the cells,
        which the model does not work out.
        """
        dims = self.sheet.dims
        if any(record.style is not None for record in dims.rows.values()) or \
                any(record.style is not None for record in dims.columns.values()):
            raise VBAUnsupportedError("shifting cells on a sheet whose rows or columns carry formats is not "
                                      "implemented")

    @method
    def Copy(self, Destination: object = MISSING) -> object:
        from pyopenvba.apps.excel._clipboard import copy, copy_range

        return copy(self) if Destination is MISSING else copy_range(self, Destination)

    @method
    def Cut(self, Destination: object = MISSING) -> object:
        from pyopenvba.apps.excel._clipboard import cut

        return cut(self, Destination)

    @method
    def PasteSpecial(self, Paste: object = MISSING, Operation: object = MISSING, SkipBlanks: object = MISSING,
                     Transpose: object = MISSING) -> object:
        from pyopenvba.apps.excel._clipboard import paste_special

        return paste_special(self, Paste, Operation, SkipBlanks, Transpose)

    def copy_to(self, Destination: object = MISSING, *, name_conflict: str = "reuse") -> object:
        if name_conflict not in {"reuse", "rename", "error"}:
            raise ValueError("name_conflict must be reuse, rename, or error")
        if Destination is MISSING:
            from pyopenvba.apps.excel._clipboard import copy

            return copy(self)
        if not isinstance(Destination, Range):
            raise error(1004, "Copy needs a range to copy to")
        if Destination.sheet.book.application is not self.sheet.book.application:
            raise error(1004, "Ranges must belong to the same Excel application")
        if len(self.areas) != 1 or len(Destination.areas) != 1:
            raise VBAUnsupportedError("Copy with multiple source or destination areas is not implemented")
        area, target = self.first, Destination.first
        repeat = target.rows % area.rows == 0 and target.columns % area.columns == 0
        rows = target.rows if repeat else area.rows
        columns = target.columns if repeat else area.columns
        bottom, right = target.top + rows - 1, target.left + columns - 1
        if bottom > MAX_ROWS or right > MAX_COLUMNS:
            raise error(1004, "Copy would extend beyond the worksheet")
        written = Area(target.top, target.left, bottom, right)
        if any(_merges.intersects(area, one) for one in self.sheet.merged_areas) or any(
            _merges.intersects(written, one) for one in Destination.sheet.merged_areas
        ):
            raise VBAUnsupportedError("Copy involving merged cells is not implemented")
        if rows * columns > 1048576:
            raise VBAUnsupportedError("Copy destinations larger than 1048576 cells are not implemented")
        # Snapshot before clearing or writing: source and destination can overlap.
        sources = {(row - area.top, column - area.left): cell
                   for (row, column), cell in self.sheet.cells_.items() if area.contains(row, column)}
        from pyopenvba.apps.excel._row_formats import shown_formats

        whole_rows, whole_columns = area.whole_rows and written.whole_rows, area.whole_columns and written.whole_columns
        shown = [shown_formats(self.sheet, area, Destination.sheet, tile_row, tile_column, sources,
                               rows_follow=whole_rows, columns_follow=whole_columns)
                 for tile_row in range(target.top, bottom + 1, area.rows)
                 for tile_column in range(target.left, right + 1, area.columns)]
        from pyopenvba.apps.excel._copy_names import NameCopyPlan

        plan = NameCopyPlan(self.sheet, Destination.sheet, name_conflict)
        formulas = {position: plan.rewrite(cell.formula) for position, cell in sources.items()
                    if cell.formula} if Destination.sheet.book is not self.sheet.book else {}
        if plan.entries:
            Destination.sheet.book.names_.entries.extend(plan.entries)
            Destination.sheet.book.names_.changed = True
        for position in list(Destination.sheet.cells_):
            if written.contains(*position):
                del Destination.sheet.cells_[position]
        # Whole rows take their heights and formats along, and whole columns their widths and formats.
        if whole_rows:
            Destination.sheet.dims.copy_rows_from(self.sheet.dims, [
                (area.top + (row - target.top) % area.rows, row) for row in range(target.top, bottom + 1)])
        elif whole_columns:
            Destination.sheet.dims.copy_columns_from(self.sheet.dims, [
                (area.left + (column - target.left) % area.columns, column) for column in range(target.left, right + 1)])
        for tile_row in range(target.top, bottom + 1, area.rows):
            for tile_column in range(target.left, right + 1, area.columns):
                for (down, across), source in sources.items():
                    copy = Cell(**{field_.name: getattr(source, field_.name) for field_ in _CELL_FIELDS})
                    if copy.formula:
                        copy.formula = shift_text(formulas.get((down, across), copy.formula), tile_row - area.top, tile_column - area.left)
                        copy.stale = True
                        copy.value = EMPTY
                    Destination.sheet.cells_[(tile_row + down, tile_column + across)] = copy
                    Destination.sheet.settle(tile_row + down, tile_column + across)
        # A position that showed its row's or column's format keeps showing it where the destination's differs.
        for plan in shown:
            for (row, column), style in plan:
                if style != Destination.sheet.inherited_style(row, column):
                    Destination.sheet.restyle(row, column, style)
        Destination.sheet.shape_changed()
        return True

    @method
    def AutoFill(self, Destination: object = MISSING, Type: object = MISSING) -> object:
        from pyopenvba.apps.excel._autofill import autofill

        return autofill(self, Destination, Type)

    @method
    def FillDown(self) -> object:
        return self._fill(1, 0)

    @method
    def FillUp(self) -> object:
        return self._fill(-1, 0)

    @method
    def FillRight(self) -> object:
        return self._fill(0, 1)

    @method
    def FillLeft(self) -> object:
        return self._fill(0, -1)

    def _fill(self, down: int, across: int) -> object:
        """Copy each area's first row or column, in the direction of the fill, over the rest of it.

        Everything the source holds goes -- values, formulas moved with
        each row or column, formats, and blanks over what was there. An
        area one row deep takes its source from the row before it, and one
        with no row before it is error 1004, before anything changes.
        On a filtered sheet the visible cells fill, as _visible has it.
        """
        from pyopenvba.apps.excel._visible import visible_areas

        parts = visible_areas(self)
        if parts is not None:
            return self._fill_visible(parts, down, across)
        plans: list[tuple[Area, Area]] = []
        for area in self.areas:
            lines = area.rows if down else area.columns
            if lines == 1:
                # A single line is filled from its neighbour on the side the fill comes from.
                source = Area(area.top - down, area.left - across, area.bottom - down, area.right - across)
                if not (1 <= source.top and source.bottom <= MAX_ROWS and 1 <= source.left
                        and source.right <= MAX_COLUMNS):
                    raise error(1004, "There is nothing to fill from")
                plans.append((source, area))
            elif down:
                edge = area.top if down > 0 else area.bottom
                rest = (area.top + 1, area.bottom) if down > 0 else (area.top, area.bottom - 1)
                plans.append((Area(edge, area.left, edge, area.right), Area(rest[0], area.left, rest[1], area.right)))
            else:
                edge = area.left if across > 0 else area.right
                rest = (area.left + 1, area.right) if across > 0 else (area.left, area.right - 1)
                plans.append((Area(area.top, edge, area.bottom, edge), Area(area.top, rest[0], area.bottom, rest[1])))
        from pyopenvba.apps.excel._protection import whole

        whole(self.sheet, [target for _, target in plans], "Filling")
        for source, target in plans:
            Range(self.sheet, [source]).copy_to(Range(self.sheet, [target]))
        return True

    def _fill_visible(self, parts: list[Area], down: int, across: int) -> object:
        """A fill of a filtered range's visible cells.

        One visible block fills as a range of its own. Several, which
        only hidden rows part, fill across each on its own, and down or
        up from the first or last visible row into every other one.
        """
        from pyopenvba.apps.excel._protection import enforced

        if enforced(self.sheet) is not None:
            raise VBAUnsupportedError("filling a filtered range of a protected sheet is not implemented")
        if len(parts) == 1:
            return Range(self.sheet, parts)._fill(down, across)
        if any((area.left, area.right) != (parts[0].left, parts[0].right) for area in parts):
            raise VBAUnsupportedError("filling a filtered range that hidden columns cut is not implemented")
        if across:
            for area in parts:
                Range(self.sheet, [area])._fill(down, across)
            return True
        rows = [row for area in parts for row in range(area.top, area.bottom + 1)]
        source = rows[0] if down > 0 else rows[-1]
        left, right = parts[0].left, parts[0].right
        for row in rows:
            if row != source:
                Range(self.sheet, [Area(source, left, source, right)]).copy_to(
                    Range(self.sheet, [Area(row, left, row, right)]))
        return True

    @method
    def Sort(self, Key1: object = MISSING, Order1: object = MISSING, Key2: object = MISSING, Type: object = MISSING,
             Order2: object = MISSING, Key3: object = MISSING, Order3: object = MISSING, Header: object = MISSING,
             OrderCustom: object = MISSING, MatchCase: object = MISSING, Orientation: object = MISSING,
             SortMethod: object = MISSING, DataOption1: object = MISSING, DataOption2: object = MISSING,
             DataOption3: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import check_sort
        from pyopenvba.apps.excel._sort import range_sort

        check_sort(self)
        return range_sort(self, [(Key1, Order1, DataOption1), (Key2, Order2, DataOption2), (Key3, Order3, DataOption3)],
                          Header, OrderCustom, MatchCase, Orientation)

    @method
    def AutoFilter(self, Field: object = MISSING, Criteria1: object = MISSING, Operator: object = MISSING,
                   Criteria2: object = MISSING, VisibleDropDown: object = MISSING, SubField: object = MISSING) -> object:
        from pyopenvba.apps.excel._autofilter import range_autofilter
        from pyopenvba.apps.excel._protection import FILTERING, refuse

        # AllowFiltering lets the user filter, not a macro.
        refuse(self.sheet, FILTERING)
        return range_autofilter(self, Field, Criteria1, Operator, Criteria2, VisibleDropDown, SubField)

    @method
    def RemoveDuplicates(self, Columns: object = MISSING, Header: object = MISSING) -> object:
        from pyopenvba.apps.excel._duplicates import remove_duplicates
        from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse

        refuse(self.sheet, APPLICATION_DEFINED)
        return remove_duplicates(self, Columns, Header)

    @method
    def SpecialCells(self, Type: object = MISSING, Value: object = MISSING) -> object:
        from pyopenvba.apps.excel._special_cells import special_cells

        if Type is MISSING:
            raise error(449)
        return special_cells(self, Type, Value)

    @member
    def CurrentRegion(self) -> object:
        """The block around the first cell that empty rows and columns bound, as Ctrl+* selects it."""
        from pyopenvba.apps.excel._region import current_region

        return Range(self.sheet, [current_region(self.sheet, self.first.top, self.first.left)])

    @method
    def End(self, Direction: object = MISSING) -> object:
        """The cell you land on with Ctrl and an arrow key."""
        from pyopenvba.apps.excel._region import holds_content

        if Direction is MISSING:
            raise error(449)
        which = int(to_integer(Direction, "Long"))
        row, column = self.first.top, self.first.left
        steps = {-4162: (-1, 0), -4121: (1, 0), -4159: (0, -1), -4161: (0, 1)}
        if which not in steps:
            raise error(1004, "End takes xlUp, xlDown, xlToLeft or xlToRight")
        down, across = steps[which]

        def occupied(at_row: int, at_column: int) -> bool:
            # A cell with only a format is walked over like an empty one.
            return holds_content(self.sheet, at_row, at_column)

        def inside(at_row: int, at_column: int) -> bool:
            return 1 <= at_row <= MAX_ROWS and 1 <= at_column <= MAX_COLUMNS

        # The neighbour decides which of the two journeys this is: along
        # a run of filled cells to its last one, or across a gap to the
        # next filled cell, or to the edge of the sheet if there is none.
        step_row, step_column = row + down, column + across
        if inside(step_row, step_column) and occupied(step_row, step_column):
            while inside(row + down, column + across) and occupied(row + down, column + across):
                row, column = row + down, column + across
        else:
            while inside(row + down, column + across):
                row, column = row + down, column + across
                if occupied(row, column):
                    break
        return Range(self.sheet, [Area(row, column, row, column, self.first.sheet)])

    @method
    def Find(self, What: object, After: object = MISSING, LookIn: object = MISSING,
             LookAt: object = MISSING, SearchOrder: object = MISSING, SearchDirection: object = MISSING,
             MatchCase: object = MISSING, MatchByte: object = MISSING, SearchFormat: object = MISSING) -> object:
        from pyopenvba.apps.excel._find import find
        return find(self, What, After, LookIn, LookAt, SearchOrder, SearchDirection,
                    MatchCase, MatchByte, SearchFormat)

    @method
    def FindNext(self, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._find import find
        return find(self, MISSING, After, direction=1, again=True)

    @method
    def FindPrevious(self, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._find import find
        return find(self, MISSING, After, direction=2, again=True)

    @method
    def Replace(self, What: object = MISSING, Replacement: object = MISSING, LookAt: object = MISSING,
                SearchOrder: object = MISSING, MatchCase: object = MISSING, MatchByte: object = MISSING,
                SearchFormat: object = MISSING, ReplaceFormat: object = MISSING,
                FormulaVersion: object = MISSING) -> object:
        from pyopenvba.apps.excel._find import replace
        from pyopenvba.apps.excel._protection import enforced

        if enforced(self.sheet) is not None:
            # A protected sheet's Replace answers True and changes nothing, unlocked cells included.
            return True
        return replace(self, What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat,
                       ReplaceFormat, FormulaVersion)

    def replaced_in(self, row: int, column: int, text: str) -> bool:
        """Enter the text Replace left in one cell; False, changing nothing, for a formula Excel cannot read.

        A cell that shows a prefix character keeps its text as text; any
        other cell, a Text cell's included, takes a formula as a formula.
        """
        from pyopenvba.apps.excel._typing import replaced

        if not _merges.writable(self.sheet, row, column):
            return True
        cell = self.sheet.cell(row, column)
        prefixed = cell is not None and not cell.formula and isinstance(cell.value, str) \
            and cell.style is not None and cell.style.quote_prefix
        if _is_formula(text) and not prefixed:
            try:
                formula = spelled_formula(self.sheet, text)
            except VBARuntimeError:
                return False
            self._put_formula(row, column, formula)
            return True
        self._store(row, column, replaced(text, self.sheet.style_at(row, column).number_format, prefixed=prefixed))
        return True

    @method
    def AutoFit(self) -> object:
        from pyopenvba.apps.excel._protection import refuse, whole_lines

        allow = "AllowFormattingRows" if whole_lines(self) == "rows" else "AllowFormattingColumns"
        refuse(self.sheet, "AutoFit method of Range class failed", allow)
        _dimensions.autofit(self)
        return EMPTY

    @method
    def Merge(self, Across: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse

        refuse(self.sheet, APPLICATION_DEFINED)
        _merges.merge(self, False if Across is MISSING else to_bool(Across))
        return EMPTY

    @method
    def UnMerge(self) -> object:
        from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse

        refuse(self.sheet, APPLICATION_DEFINED)
        _merges.unmerge(self)
        return EMPTY

    @member
    def MergeCells(self) -> object:
        return _merges.flag(self)

    @setter("MergeCells")
    def _set_merge_cells(self, value: object) -> None:
        if to_bool(value):
            _merges.merge(self, False)
        else:
            _merges.unmerge(self)

    @member
    def MergeArea(self) -> object:
        if not self.single:
            raise error(1004, "MergeArea requires one cell")
        area = _merges.at(self.sheet, self.first.top, self.first.left)
        return Range(self.sheet, [area or self.first])

    # -- Python side

    def vba_iterate(self) -> Iterator[object]:
        self._check_alive()
        for row, column in self.positions():
            yield Range(self.sheet, [Area(row, column, row, column, self.sheet.name)])

    def _read(self, row: int, column: int, *, raw: bool = False) -> object:
        """A cell's value as VBA reads it; ``raw`` is Value2, which leaves a number a Double whatever its format."""
        from pyopenvba.apps.excel._calc import as_vba

        cell = self.sheet.cell(row, column)
        if cell is None:
            return EMPTY
        value = self.sheet.book.calculator.value_of(self.sheet.name, row, column) if cell.formula else cell.value
        return as_vba(value, None if raw else cell)

    def writable_positions(self) -> list[tuple[int, int]]:
        """Which cells a write touches: on a filtered sheet only the visible ones, as _visible has it.

        Excel fills a whole column, or formats it through the column's own
        format; the model does neither yet, so the write reports itself.
        """
        from pyopenvba.apps.excel._visible import visible_areas

        out: list[tuple[int, int]] = []
        for area in visible_areas(self) or self.areas:
            if area.rows * area.columns > 1048576:
                raise VBAUnsupportedError("writing to or formatting more than 1,048,576 cells at once is not implemented")
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    out.append((row, column))
        return out

    def _write(self, value: object, *, raw: bool = False) -> None:
        if isinstance(value, VBAArray):
            self._write_array(value, raw=raw)
            return
        if isinstance(value, VBAObject):
            value = value.vba_value()
        if isinstance(value, str) and value.startswith("="):
            self._set_formula(value)
            return
        for row, column in self.writable_positions():
            if _merges.writable(self.sheet, row, column):
                self._type_into(row, column, value, raw=raw)

    def _write_array(self, array: VBAArray, *, raw: bool = False) -> None:
        """An array written to each area -- each visible area, on a filtered sheet -- from the array's first item."""
        areas = self._array_areas(array)
        for area in areas:
            if array.dimensions == 1:
                # A flat array is one row, laid across and repeated down
                # every row of the range, which is what Excel does with it.
                items = array.elements()
                for offset, item in enumerate(items):
                    column = area.left + offset
                    if column > area.right:
                        break
                    for row in range(area.top, area.bottom + 1):
                        self._put(row, column, item, raw=raw)
                continue
            rows, columns = array.bounds[0], array.bounds[1]
            for row_index in range(rows[0], rows[1] + 1):
                for column_index in range(columns[0], columns[1] + 1):
                    row = area.top + row_index - rows[0]
                    column = area.left + column_index - columns[0]
                    if row > area.bottom or column > area.right:
                        continue
                    self._put(row, column, array.get([row_index, column_index]), raw=raw)
        self.sheet.touched()

    def _array_areas(self, array: VBAArray) -> list[Area]:
        """The areas an array is written to, each from the array's start.

        Excel reads a two-dimensional array for the second of several
        areas at another stride -- {11,12;21,22;...;61,62} written to
        A2:B2, A4:B4 and A6:B6 leaves 11 and 31 in A4:B4 -- so areas wider
        than one column report themselves; one column wide, each takes the
        array's first column.
        """
        from pyopenvba.apps.excel._visible import visible_areas

        areas = visible_areas(self) or self.areas
        if len(areas) > 1 and array.dimensions == 2 and any(area.columns > 1 for area in areas):
            raise VBAUnsupportedError("writing a two-dimensional array to several areas wider than one column is "
                                      "not implemented")
        return areas

    def _put(self, row: int, column: int, value: object, *, raw: bool = False) -> None:
        if not _merges.writable(self.sheet, row, column):
            return
        if isinstance(value, str) and _is_formula(value) and not self._keeps_text(row, column):
            formula = spelled_formula(self.sheet, value)
            if self._put_formula(row, column, formula):
                self._bring_format([(row, column)], formula, row, column)
        else:
            self._type_into(row, column, value, raw=raw)

    def _keeps_text(self, row: int, column: int) -> bool:
        """Whether a position is a Text cell, which keeps whatever string is written to it as text."""
        return self.sheet.style_at(row, column).number_format == "@"

    def _put_formula(self, row: int, column: int, formula: str) -> bool:
        """Write one formula into one cell; False when a protected sheet held the cell back."""
        gate = self.sheet.write_gate
        if gate is not None and not gate.admits(row, column):
            return False
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        cell.formula = formula
        cell.stale = True
        cell.value = EMPTY
        self.sheet.cell_changed(row, column)
        return True

    def _bring_format(self, placed: list[tuple[int, int]], formula: str, row: int, column: int) -> None:
        """Give the cells a formula was just written to the number format it brings, where they are General.

        Excel works the format out once, from the formula as written at the
        top left of the range, and gives it to every cell of the write.
        """
        from pyopenvba.apps.excel._formula_format import brought_format

        code = brought_format(self.sheet, formula, row, column) if placed else ""
        if not code:
            return
        for one_row, one_column in placed:
            style = self.sheet.style_at(one_row, one_column)
            if style.number_format in ("General", ""):
                self.sheet.restyle(one_row, one_column, applying(style, "number_format", number_format=code))

    def _type_into(self, row: int, column: int, value: object, *, raw: bool = False) -> None:
        """Write one value into one cell as Excel types it: the value, the format it brings, and a prefix."""
        from pyopenvba.apps.excel._typing import typed

        gate = self.sheet.write_gate
        if gate is not None and not gate.admits(row, column):
            return
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        style = cell.style or self.sheet.book.stylesheet.default
        self._store(row, column, typed(value, style.number_format, raw=raw))

    def _store(self, row: int, column: int, result: Typed) -> None:
        """Put what typing made of a write into one cell: its value, the format it ends with, and a prefix."""
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        style = cell.style or self.sheet.book.stylesheet.default
        cell.value = result.value
        cell.formula = ""
        cell.stale = False
        changed = style
        if result.number_format is not None:
            changed = applying(changed, "number_format", number_format=result.number_format)
        if result.prefix and not changed.quote_prefix:
            changed = replace(changed, quote_prefix=True)
        if changed != style:
            self.sheet.restyle(row, column, changed)
        self.sheet.cell_changed(row, column)

    def describe(self, indent: str = "") -> str:
        return f"{indent}Range {self.vba_get('Address')} on {self.sheet.name!r}"


_CELL_FIELDS = tuple(Cell.__dataclass_fields__.values())


class Areas(VBACollection, ExcelObject):
    """The rectangles a Union made."""

    vba_type_name = "Areas"

    def __init__(self, target: Range) -> None:
        self.target = target

    def vba_items(self) -> list[object]:
        return [Range(self.target.sheet, [area]) for area in self.target.areas]


class RowsOf(VBACollection, ExcelObject):
    """A range's rows, one Range each."""

    vba_type_name = "Range"

    def __init__(self, target: Range) -> None:
        self.target = target

    def vba_items(self) -> list[object]:
        area = self.target.bounded()
        return [
            Range(self.target.sheet, [Area(row, area.left, row, area.right, area.sheet)])
            for row in range(area.top, area.bottom + 1)
        ]

    @member
    def Count(self) -> object:
        return VBAInt(self.target.first.rows, "Long")


class ColumnsOf(VBACollection, ExcelObject):
    """A range's columns, one Range each."""

    vba_type_name = "Range"

    def __init__(self, target: Range) -> None:
        self.target = target

    def vba_items(self) -> list[object]:
        area = self.target.bounded()
        return [
            Range(self.target.sheet, [Area(area.top, column, area.bottom, column, area.sheet)])
            for column in range(area.left, area.right + 1)
        ]

    @member
    def Count(self) -> object:
        return VBAInt(self.target.first.columns, "Long")


# --- names -------------------------------------------------------------------------------------


@dataclass(slots=True)
class NameEntry:
    name: str
    refers_to: str
    book: Any
    sheet: Any = None
    #: The attribute text this name was read with, so a name nobody
    #: touched is written back with its localSheetId and hidden intact.
    attributes: str = ""
    visible: bool = True
    comment: str = ""
    invalidated: bool = False


class Names(VBACollection, ExcelObject):
    """A workbook's defined names."""

    vba_type_name = "Names"

    def __init__(self, book: Workbook, sheet: Worksheet | None = None) -> None:
        self._book = book
        self.sheet = sheet
        self._entries: list[NameEntry] = []
        #: False until something adds, changes or removes a name, so an
        #: untouched workbook's names are never rewritten.
        self.changed = False

    @property
    def book(self) -> Workbook:
        return self.sheet.book if self.sheet is not None else self._book

    @property
    def entries(self) -> list[NameEntry]:
        return self.book.names_.entries if self.sheet is not None else self._entries

    @entries.setter
    def entries(self, entries: list[NameEntry]) -> None:
        if self.sheet is not None:
            self.book.names_.entries = entries
        else:
            self._entries = entries

    def vba_items(self) -> list[object]:
        from pyopenvba._a1 import split_sheet

        self._check_alive()
        entries = [entry for entry in self.entries if self.sheet is None or split_sheet(entry.name)[0].casefold() == self.sheet.name.casefold()]
        return [DefinedName(entry) for entry in sorted(entries, key=lambda entry: (split_sheet(entry.name)[1].casefold(), entry.name.casefold()))]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            found = self.find(index)
            if found is None:
                raise error(1004, f"there is no name called {index}")
            return found
        return super().vba_lookup(index, items)

    def find(self, name: str, *, scope: Worksheet | None = None) -> DefinedName | None:
        from pyopenvba._a1 import split_sheet

        owner, bare = split_sheet(name)
        workbook_qualified = owner.casefold() == self.book.name.casefold()
        context = self.sheet or scope or self.book.active_sheet
        wanted_scope = owner or (context.name if context is not None else "")
        fallback = None
        for entry in self.entries:
            entry_scope, entry_bare = split_sheet(entry.name)
            if entry_bare.casefold() != bare.casefold():
                continue
            if workbook_qualified:
                if not entry_scope:
                    return DefinedName(entry)
                continue
            if entry_scope.casefold() == wanted_scope.casefold():
                return DefinedName(entry)
            if not entry_scope and not owner and self.sheet is None:
                fallback = DefinedName(entry)
        return fallback

    @method
    def Add(self, Name: object = MISSING, RefersTo: object = MISSING, Visible: object = MISSING,
            MacroType: object = MISSING, ShortcutKey: object = MISSING, Category: object = MISSING,
            NameLocal: object = MISSING, RefersToLocal: object = MISSING, CategoryLocal: object = MISSING,
            RefersToR1C1: object = MISSING, RefersToR1C1Local: object = MISSING) -> object:
        if Name is MISSING:
            Name = NameLocal
        if Name is MISSING:
            raise error(449)
        if any(value is not MISSING for value in (MacroType, ShortcutKey, Category, CategoryLocal)):
            raise VBAUnsupportedError("Macro name metadata is not implemented")
        wanted = _names.canonical(self.book, to_text(Name), self.sheet)
        owner = self.sheet or self.book.active_sheet
        if owner is None:
            raise error(1004, "No worksheet is available for the name")
        refers = RefersTo if RefersTo is not MISSING else RefersToLocal
        if refers is MISSING:
            rc = RefersToR1C1 if RefersToR1C1 is not MISSING else RefersToR1C1Local
            if rc is MISSING:
                raise error(449)
            refers = to_a1(to_text(rc), 1, 1)
        text = "=" + ",".join(area.address(with_sheet=True) for area in refers.areas) if isinstance(refers, Range) else _names.qualify(to_text(refers), owner.name)
        entry = next((one for one in self.entries if one.name.casefold() == wanted.casefold()), None)
        if entry is None:
            entry = NameEntry(wanted, text, self.book)
            self.entries.append(entry)
        entry.name, entry.refers_to = wanted, text
        if Visible is not MISSING:
            entry.visible = to_bool(Visible)
        _names.changed(self.book)
        return DefinedName(entry)


class DefinedName(ExcelObject):
    """One defined name."""

    vba_type_name = "Name"

    def __init__(self, entry: NameEntry) -> None:
        self.entry = entry

    @member
    def Name(self) -> object:
        return self.entry.name

    @member
    def NameLocal(self) -> object:
        return self.Name()

    @setter("NameLocal")
    def _set_name_local(self, value: object) -> None:
        self._set_name(value)

    @setter("Name")
    def _set_name(self, value: object) -> None:
        from pyopenvba._a1 import split_sheet

        book = self.entry.book
        scope, _ = split_sheet(self.entry.name)
        wanted = _names.canonical(book, to_text(value), book.sheet_named(scope) if scope else None)
        if any(one is not self.entry and one.name.casefold() == wanted.casefold() for one in book.names_.entries):
            return  # Excel silently leaves the original name on collision.
        for sheet in book.sheets_:
            for cell in sheet.cells_.values():
                if cell.formula:
                    updated = _names.renamed_formula(cell.formula, sheet, self.entry, wanted)
                    if updated != cell.formula:
                        cell.formula = updated
                        sheet.touched()
            for shape in sheet.shapes_:
                if shape.control is not None:
                    for field in ("linked_cell", "list_range"):
                        text = getattr(shape.control, field)
                        updated = _names.renamed_formula(text, sheet, self.entry, wanted) if text else text
                        if updated != text:
                            setattr(shape.control, field, updated)
                            sheet.drawing_changed()
        for entry in book.names_.entries:
            entry_scope, _ = split_sheet(entry.name)
            owner = book.sheet_named(entry_scope) if entry_scope else book.active_sheet
            entry.refers_to = _names.renamed_formula(entry.refers_to, owner, self.entry, wanted)
        self.entry.name = wanted
        _names.changed(book)

    @member
    def Visible(self) -> object:
        return self.entry.visible

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        self.entry.visible = to_bool(value)
        _names.changed(self.entry.book)

    @member
    def Comment(self) -> object:
        return self.entry.comment

    @setter("Comment")
    def _set_comment(self, value: object) -> None:
        self.entry.comment = to_text(value)
        _names.changed(self.entry.book)

    @member
    def Parent(self) -> object:
        return self.entry.book

    @member
    def RefersToR1C1(self) -> object:
        return from_a1(self.entry.refers_to, 1, 1)

    @setter("RefersToR1C1")
    def _set_refers_r1c1(self, value: object) -> None:
        self._set_refers_to(to_a1(to_text(value), 1, 1))

    @member
    def RefersTo(self) -> object:
        return self.entry.refers_to

    @member
    def RefersToLocal(self) -> object:
        return self.RefersTo()

    @setter("RefersToLocal")
    def _set_refers_local(self, value: object) -> None:
        self._set_refers_to(value)

    @member
    def RefersToR1C1Local(self) -> object:
        return self.RefersToR1C1()

    @setter("RefersToR1C1Local")
    def _set_refers_r1c1_local(self, value: object) -> None:
        self._set_refers_r1c1(value)

    @setter("RefersTo")
    def _set_refers_to(self, value: object) -> None:
        self.entry.refers_to = _names.qualify(to_text(value), self.entry.book.active_sheet.name)
        _names.changed(self.entry.book)

    @member(default=True)
    def Value(self) -> object:
        return self.entry.refers_to

    @setter("Value")
    def _set_value(self, value: object) -> None:
        self._set_refers_to(value)

    @member
    def RefersToRange(self) -> object:
        return self.refers_to_range()

    @method
    def Delete(self) -> object:
        book = self.entry.book
        book.names_.entries = [one for one in book.names_.entries if one is not self.entry]
        _names.changed(book)
        return EMPTY

    def refers_to_range(self) -> Range:
        text = self.entry.refers_to.lstrip("=")
        book = self.entry.book
        try:
            areas = parse_reference(text)
        except ValueError:
            from pyopenvba.apps.excel._control_refs import binding
            from pyopenvba._a1 import split_sheet

            scope, _ = split_sheet(self.entry.name)
            owner = book.sheet_named(scope) if scope else book.active_sheet
            try:
                _, area = binding(owner, self.entry.name)
            except (ValueError, VBAUnsupportedError):
                area = None
            if area is None:
                raise error(1004, f"{self.entry.name} does not refer to a range") from None
            areas = [area]
        sheet = book.sheet_named(areas[0].sheet) if areas[0].sheet else book.active_sheet
        if sheet is None:
            raise error(1004, f"{self.entry.name} refers to a sheet that is not there")
        return Range(sheet, areas)


# --- Power Query ----------------------------------------------------------------------------------


@dataclass(slots=True)
class QueryEntry:
    name: str
    formula: str
    description: str = ""


class Queries(VBACollection, ExcelObject):
    """A workbook's Power Query queries."""

    vba_type_name = "Queries"

    def __init__(self, book: Workbook) -> None:
        self.book = book
        self.entries: list[QueryEntry] = []

    def vba_items(self) -> list[object]:
        return [WorkbookQuery(entry, self.book) for entry in self.entries]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for entry in self.entries:
                if entry.name.lower() == index.lower():
                    return WorkbookQuery(entry, self.book)
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE, f"there is no query called {index}")
        return super().vba_lookup(index, items)

    @method
    def Add(self, Name: object = MISSING, Formula: object = MISSING, Description: object = MISSING) -> object:
        if Name is MISSING or Formula is MISSING:
            raise error(449)
        entry = QueryEntry(to_text(Name), to_text(Formula), "" if Description is MISSING else to_text(Description))
        self.entries.append(entry)
        self.book.saved = False
        return WorkbookQuery(entry, self.book)


class WorkbookQuery(ExcelObject):
    """One query: its M, and the refresh pyOpenVBA cannot yet run."""

    vba_type_name = "WorkbookQuery"

    def __init__(self, entry: QueryEntry, book: Workbook) -> None:
        self.entry = entry
        self.book = book

    @member
    def Name(self) -> object:
        return self.entry.name

    @member
    def Formula(self) -> object:
        return self.entry.formula

    @setter("Formula")
    def _set_formula(self, value: object) -> None:
        self.entry.formula = to_text(value)
        self.book.saved = False

    @member
    def Description(self) -> object:
        return self.entry.description

    @method
    def Refresh(self) -> object:
        """Evaluate the query's M and put what it answers on its sheet."""
        from pyopenvba.apps.excel._refresh import refresh

        refresh(self.book, self.entry.name)
        return EMPTY

    @method
    def Delete(self) -> object:
        self.book.queries_.entries = [one for one in self.book.queries_.entries if one is not self.entry]
        return EMPTY


# --- WorksheetFunction ------------------------------------------------------------------------------


class WorksheetFunction(ExcelObject):
    """The worksheet functions a macro calls through Application.WorksheetFunction.

    Each member is the formula engine's function of the same name, called
    as _worksheet_functions describes; a real Excel function the engine
    does not have says it is not implemented rather than guessing.
    """

    vba_type_name = "WorksheetFunction"

    def __init__(self, application: Application) -> None:
        self.application = application

    def vba_get(self, name: str, args: Any = (), named: Any = None) -> object:
        from pyopenvba.apps.excel._worksheet_functions import call, engine_name
        from pyopenvba.interpreter._inventory import member_exists

        if engine_name(name) is not None and member_exists("WorksheetFunction", name, "excel"):
            return call(self.application, name, args, named, raising=True)
        spec = self.vba_member(name)
        if spec is None:
            from pyopenvba.formula._inventory import excel_has_function

            if excel_has_function(name):
                raise VBAUnsupportedError(
                    f"WorksheetFunction.{name} is a real Excel function that pyOpenVBA "
                    f"does not implement"
                )
            raise error(438, f"WorksheetFunction has no member named {name}")
        return super().vba_get(name, args, named)


def _register_worksheet_functions() -> None:
    """A member for each WorksheetFunction method the formula engine has, so the inventory sees them."""
    from pyopenvba.apps.excel._worksheet_functions import call, engine_name
    from pyopenvba.interpreter._inventory import members_of

    def calling(name: str) -> Callable[..., object]:
        def run(self: WorksheetFunction, *args: object) -> object:
            return call(self.application, name, args, None, raising=True)

        return run

    for name in sorted(members_of("WorksheetFunction", "excel")):
        if engine_name(name) is not None:
            WorksheetFunction.vba_add_member(MemberSpec(name=name, kind="method", getter=calling(name), varargs=True))


_register_worksheet_functions()


# --- helpers ----------------------------------------------------------------------------------------


def _display_text(cell: Cell, characters: float) -> str:
    """What the cell shows: its value through its number format, in a column ``characters`` wide.

    A value its format cannot show, a date before 1900 say, fills the
    column with as many # as it is characters wide (measured at 40 and 60).
    What a ``*`` fill pads the cell with depends on its width in pixels and
    the font's, which is not modelled, and neither is a value too wide for
    its column, which Excel shows as # or with fewer digits.
    """
    from pyopenvba.formula._display import UndisplayableError, shown
    from pyopenvba.formula._values import ExcelError

    value = cell.value.serial if isinstance(cell.value, VBADate) else cell.value
    if cell.stale or value is EMPTY:
        return ""
    if isinstance(value, ExcelError):
        return value.name
    if not isinstance(value, (bool, str, int, float)):
        return to_text(value)
    try:
        text, fills = shown(value, cell.number_format)
    except UndisplayableError:
        return "#" * int(characters)
    if fills:
        raise VBAUnsupportedError("what a number format with a * fill shows depends on the cell's width in pixels, "
                                  "which pyOpenVBA does not model")
    return text


def _formula_text(value: object) -> str:
    """What Range.Formula reads for a constant: a number as Excel spells one there, not as CStr would."""
    from pyopenvba.formula._values import ExcelError, number_text

    if value is EMPTY:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, ExcelError):
        return value.name
    if isinstance(value, VBADate):
        value = value.serial
    if isinstance(value, (int, float)):
        return number_text(float(value), formula=True)
    return to_text(value)


class _BookNames:
    """The workbook a formula is written into, as spelling it asks (see pyopenvba.formula._spell)."""

    def __init__(self, sheet: Worksheet) -> None:
        self.book = sheet.book
        self.home = sheet.name

    def sheet(self, name: str) -> str | None:
        return next((one.name for one in self.book.sheets_ if one.name.lower() == name.lower()), None)

    def defined(self, name: str, sheet: str) -> str | None:
        from pyopenvba._a1 import split_sheet

        scope = next((one for one in self.book.sheets_ if one.name.lower() == sheet.lower()), None)
        found = self.book.names_.find(name, scope=scope)
        return None if found is None else split_sheet(found.entry.name)[1]

    def remembered(self, name: str) -> str:
        from pyopenvba.formula._spell import remembered_names

        spellings = self.book.name_spellings
        if spellings is None:
            # A workbook learns the names its formulas already use the first time it needs one.
            spellings = self.book.name_spellings = {}
            for sheet in self.book.sheets_:
                names = _BookNames(sheet)
                for _, cell in sorted(sheet.cells_.items()):
                    for one in remembered_names(cell.formula, names) if cell.formula else []:
                        spellings.setdefault(one.lower(), one)
        return spellings.setdefault(name.lower(), name)


def _is_formula(text: str) -> bool:
    """Whether a string written to a cell is a formula: it starts with =, and = alone is text."""
    return text.startswith("=") and text != "="


def spelled_formula(sheet: Worksheet, formula: str) -> str:
    """A formula a macro writes to ``sheet``, as Excel spells it back; error 1004 where Excel refuses it."""
    from pyopenvba.formula._parse import FormulaError
    from pyopenvba.formula._spell import UnmodelledFormulaError, spelled

    try:
        return spelled(formula, _BookNames(sheet))
    except UnmodelledFormulaError as exc:
        raise VBAUnsupportedError(str(exc)) from None
    except FormulaError:
        raise error(1004, "Application-defined or object-defined error") from None


def stored_value(value: object) -> object:
    """A value as a cell holds it when nothing types it: every number a Double, a Date the serial under it.

    What a macro writes goes through typing instead (see _typing); this is
    for values that arrive already typed, as a query's rows do.
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    return value
