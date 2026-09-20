"""Excel's object model, in memory.

The objects a macro touches -- Application, Workbook, Worksheet, Range
and the few that hang off them -- holding real state that VBA mutates
and that can be written back out to a file.

A cell holds a value, a formula, a number format and a little
formatting.  Reading a cell whose formula has not been worked out
calculates it first, through :mod:`pyopenvba.apps.excel._calc`, so a
macro that writes a formula and reads the answer gets one.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_reference
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import shift_text
from pyopenvba.shapes._values import Shape as ShapeState
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    ERR_APPLICATION_DEFINED,
    ERR_SUBSCRIPT_OUT_OF_RANGE,
    MISSING,
    NOTHING,
    VBAArray,
    VBADate,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_number,
    to_text,
    type_name,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter

_LIBRARY = "excel"


class ExcelObject(VBAObject):
    """Every Excel class, so the member inventory knows where to look."""

    vba_library = _LIBRARY


# --- what a cell holds ---------------------------------------------------------------


@dataclass(slots=True)
class Cell:
    """One cell's state.  Absent from the sheet's map means empty."""

    value: object = EMPTY
    formula: str = ""
    number_format: str = "General"
    #: True when the formula was written here and nothing has computed it.
    stale: bool = False
    font_bold: bool = False
    font_italic: bool = False
    font_name: str = "Calibri"
    font_size: float = 11.0
    font_color: int = 0
    interior_color: int | None = None
    horizontal_alignment: int = -4131

    def is_blank(self) -> bool:
        return (
            self.value is EMPTY
            and not self.formula
            and self.number_format == "General"
            and not self.font_bold
            and not self.font_italic
            and self.interior_color is None
        )


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
        self.cut_copy_mode: object = False
        self.user_name = "pyOpenVBA"
        self.interpreter: Interpreter | None = None
        self.active_book: Workbook | None = None
        self._this_workbook: Workbook | None = None
        self.worksheet_function = WorksheetFunction(self)

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
        return self.cut_copy_mode

    @setter("CutCopyMode")
    def _set_cut_copy_mode(self, value: object) -> None:
        self.cut_copy_mode = value

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
        """``[A1]`` and Evaluate("A1"): a reference, or a name."""
        sheet = self._require_sheet()
        book = sheet.book
        named = book.names_.find(text)
        if named is not None:
            return named.refers_to_range()
        try:
            areas = parse_reference(text, sheet=sheet.name)
        except ValueError:
            raise error(ERR_APPLICATION_DEFINED, f"cannot evaluate {text!r}") from None
        return Range(book.sheet_named(areas[0].sheet) if areas[0].sheet else sheet, areas)

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
        book = Workbook(self.application, f"Book{len(self.books) + 1}")
        book.add_sheet("Sheet1")
        self.books.append(book)
        self.application.activate_book(book)
        return book

    @method
    def Open(self, Filename: object = MISSING, *rest: object) -> object:
        from pyopenvba.apps.excel._io import load_workbook

        if Filename is MISSING:
            raise error(449)
        book = load_workbook(self.application, Path(to_text(Filename)))
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
            self.application.active_book = self.books[0] if self.books else None


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
        self.active_sheet_index = 0
        #: The package this was loaded from, kept so a save can put back
        #: every part pyOpenVBA does not model.
        self.package: Any = None
        self._calculator: Any = None
        #: Where each query's rows go, read from the package on demand.
        self._load_targets: Any = None

    @property
    def calculator(self) -> Any:
        """The workbook's formulas, and what each one comes to."""
        if self._calculator is None:
            from pyopenvba.apps.excel._calc import Calculator

            self._calculator = Calculator(self)
        return self._calculator

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

    @method
    def Close(self, SaveChanges: object = MISSING, Filename: object = MISSING, *rest: object) -> object:
        if SaveChanges is not MISSING and to_bool(SaveChanges):
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
    def SaveAs(self, Filename: object = MISSING, *rest: object) -> object:
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
        self.visible = -1  # xlSheetVisible
        self.column_widths: dict[int, float] = {}
        self.row_heights: dict[int, float] = {}
        self.selection_range: Range | None = None
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

    def drawing_changed(self) -> None:
        """A shape was added, moved, renamed or deleted."""
        self.drawing_dirty = True
        self.book.saved = False

    def column_width_points(self, column: int) -> float:
        """How wide a column is in points, which a shape's cell needs."""
        from pyopenvba.shapes._xlsx import DEFAULT_COLUMN_POINTS, characters_to_points

        width = self.column_widths.get(column)
        return characters_to_points(width) if width is not None else DEFAULT_COLUMN_POINTS

    def row_height_points(self, row: int) -> float:
        """How tall a row is in points."""
        from pyopenvba.shapes._xlsx import DEFAULT_ROW_POINTS

        return self.row_heights.get(row, DEFAULT_ROW_POINTS)

    def drawing_grid(self) -> Any:
        """The sheet's columns and rows, as the drawing layer wants them.

        The model counts columns and rows from one and keeps a column's
        width in characters; a drawing counts from zero and works in
        points.
        """
        from pyopenvba.shapes._xlsx import SheetGrid, characters_to_points

        return SheetGrid(
            column_widths={
                column - 1: characters_to_points(width)
                for column, width in self.column_widths.items()
            },
            row_heights={row - 1: height for row, height in self.row_heights.items()},
        )

    def cell_changed(self, row: int, column: int) -> None:
        """One cell's contents changed: tell the calculator what to redo."""
        self.touched()
        calculator = self.book.calculator
        cell = self.cells_.get((row, column))
        if cell is not None:
            calculator.remember(self.name, row, column, cell)
        else:
            calculator.compiled.pop((self.name.lower(), row, column), None)
        calculator.wrote(self.name, row, column)

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
        wanted = to_text(value)
        if any(sheet is not self and sheet.name.lower() == wanted.lower() for sheet in self.book.sheets_):
            raise error(1004, f"a sheet called {wanted} is already there")
        self.name = wanted
        self.book.saved = False

    @member
    def Index(self) -> object:
        return VBAInt(self.book.sheets_.index(self) + 1, "Long")

    @member
    def Visible(self) -> object:
        return VBAInt(self.visible, "Long")

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
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
        named = self.book.names_.find(text)
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
            return Range(self, [Area(1, 1, MAX_ROWS, MAX_COLUMNS, self.name)])
        if isinstance(Index, str):
            # Rows("3") is the whole row, the way Rows("3:5") is.
            wanted = Index if ":" in Index else f"{Index}:{Index}"
            return Range(self, parse_reference(wanted, sheet=self.name))
        row = int(to_integer(Index, "Long"))
        return Range(self, [Area(row, 1, row, MAX_COLUMNS, self.name)])

    @member
    def Columns(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            return Range(self, [Area(1, 1, MAX_ROWS, MAX_COLUMNS, self.name)])
        if isinstance(Index, str):
            # Columns("C") is the whole column, the way Columns("C:D") is.
            wanted = Index if ":" in Index else f"{Index}:{Index}"
            return Range(self, parse_reference(wanted, sheet=self.name))
        column = int(to_integer(Index, "Long"))
        return Range(self, [Area(1, column, MAX_ROWS, column, self.name)])

    @member
    def UsedRange(self) -> object:
        bounds = self.used_bounds()
        if bounds is None:
            return Range(self, [Area(1, 1, 1, 1, self.name)])
        top, left, bottom, right = bounds
        return Range(self, [Area(top, left, bottom, right, self.name)])

    @member
    def Names(self, Index: object = MISSING) -> object:
        return self.book.vba_get("Names", [] if Index is MISSING else [Index])

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
    def Delete(self) -> object:
        self.book.sheets_.remove(self)
        self.book.saved = False
        return EMPTY

    @method
    def Calculate(self) -> object:
        self.book.calculator.calculate_all()
        return EMPTY

    # -- Python side

    @property
    def selection(self) -> object:
        return self.selection_range if self.selection_range is not None else self.vba_get("Range", ["A1"])

    @property
    def active_cell(self) -> object:
        return self.active_cell_range if self.active_cell_range is not None else self.vba_get("Range", ["A1"])

    def cell(self, row: int, column: int, *, create: bool = False) -> Cell | None:
        found = self.cells_.get((row, column))
        if found is None and create:
            found = Cell()
            self.cells_[(row, column)] = found
        return found

    def used_bounds(self) -> tuple[int, int, int, int] | None:
        live = [(row, column) for (row, column), cell in self.cells_.items() if not cell.is_blank()]
        if not live:
            return None
        rows = [row for row, _ in live]
        columns = [column for _, column in live]
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

    def __init__(self, sheet: Worksheet, areas: list[Area]) -> None:
        self.sheet = sheet
        self.areas = areas or [Area(1, 1, 1, 1, sheet.name)]

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
        if self.single:
            return self._read(self.first.top, self.first.left)
        area = self.bounded()
        items: list[object] = []
        # VBA lays a two-dimensional array out column by column.
        for column in range(area.left, area.right + 1):
            for row in range(area.top, area.bottom + 1):
                items.append(self._read(row, column))
        return VBAArray([(1, area.rows), (1, area.columns)], items=items)

    @setter("Value")
    def _set_value(self, value: object) -> None:
        self._write(value)

    @member
    def Value2(self) -> object:
        return self.Value()

    @setter("Value2")
    def _set_value2(self, value: object) -> None:
        self._write(value)

    @member
    def Text(self) -> object:
        cell = self.sheet.cell(self.first.top, self.first.left)
        if cell is None:
            return ""
        # Reading what a cell shows calculates it first, as looking at
        # one in Excel does.
        self._read(self.first.top, self.first.left)
        return _display_text(cell)

    @member
    def Formula(self) -> object:
        if self.single:
            cell = self.sheet.cell(self.first.top, self.first.left)
            if cell is None:
                return ""
            return cell.formula or _formula_text(cell.value)
        area = self.bounded()
        items: list[object] = []
        for column in range(area.left, area.right + 1):
            for row in range(area.top, area.bottom + 1):
                cell = self.sheet.cell(row, column)
                items.append("" if cell is None else (cell.formula or _formula_text(cell.value)))
        return VBAArray([(1, area.rows), (1, area.columns)], items=items)

    @setter("Formula")
    def _set_formula(self, value: object) -> None:
        """A formula written to a block moves with each cell.

        Excel anchors what was written at the top left and shifts every
        reference that is not held by a dollar sign, which is why
        Range("D2:D6").Formula = "=B2*C2" leaves =B6*C6 in D6.
        """
        written = to_text(value)
        anchor = self.first
        for row, column in self.writable_positions():
            text = (
                shift_text(written, row - anchor.top, column - anchor.left)
                if written.startswith("=")
                else written
            )
            cell = self.sheet.cell(row, column, create=True)
            assert cell is not None
            if text.startswith("="):
                cell.formula = text
                cell.stale = True
                cell.value = EMPTY
            else:
                cell.formula = ""
                cell.stale = False
                cell.value = _from_text(text)
            self.sheet.cell_changed(row, column)

    @member
    def FormulaR1C1(self) -> object:
        raise VBAUnsupportedError("FormulaR1C1 is not implemented by pyOpenVBA")

    @member
    def NumberFormat(self) -> object:
        cell = self.sheet.cell(self.first.top, self.first.left)
        return cell.number_format if cell is not None else "General"

    @setter("NumberFormat")
    def _set_number_format(self, value: object) -> None:
        text = to_text(value)
        for row, column in self.writable_positions():
            cell = self.sheet.cell(row, column, create=True)
            assert cell is not None
            cell.number_format = text
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
        return ",".join(
            book + area.address(rows_fixed=rows_fixed, columns_fixed=columns_fixed, with_sheet=external)
            for area in self.areas
        )

    @member
    def Row(self) -> object:
        return VBAInt(self.first.top, "Long")

    @member
    def Column(self) -> object:
        return VBAInt(self.first.left, "Long")

    @member
    def Count(self) -> object:
        total = sum(area.rows * area.columns for area in self.areas)
        return VBAInt(total, "Long")

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
        )

    @member
    def EntireColumn(self) -> object:
        return Range(
            self.sheet,
            [Area(1, area.left, MAX_ROWS, area.right, area.sheet) for area in self.areas],
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
        return Font(self)

    @member
    def Interior(self) -> object:
        return Interior(self)

    @member
    def HorizontalAlignment(self) -> object:
        cell = self.sheet.cell(self.first.top, self.first.left)
        return VBAInt(cell.horizontal_alignment if cell is not None else -4131, "Long")

    @setter("HorizontalAlignment")
    def _set_horizontal_alignment(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        for row, column in self.writable_positions():
            cell = self.sheet.cell(row, column, create=True)
            assert cell is not None
            cell.horizontal_alignment = wanted

    @member
    def ColumnWidth(self) -> object:
        return self.sheet.column_widths.get(self.first.left, 8.43)

    @setter("ColumnWidth")
    def _set_column_width(self, value: object) -> None:
        width = float(to_number(value))
        for column in range(self.first.left, min(self.first.right, self.first.left + 255) + 1):
            self.sheet.column_widths[column] = width

    @member
    def RowHeight(self) -> object:
        return self.sheet.row_heights.get(self.first.top, 15.0)

    @setter("RowHeight")
    def _set_row_height(self, value: object) -> None:
        height = float(to_number(value))
        for row in range(self.first.top, min(self.first.bottom, self.first.top + 255) + 1):
            self.sheet.row_heights[row] = height

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
        for row, column in self.writable_positions():
            self.sheet.cells_.pop((row, column), None)
            self.sheet.cell_changed(row, column)
        return EMPTY

    @method
    def ClearContents(self) -> object:
        for row, column in self.writable_positions():
            cell = self.sheet.cell(row, column)
            if cell is not None:
                cell.value = EMPTY
                cell.formula = ""
                cell.stale = False
                self.sheet.cell_changed(row, column)
        self.sheet.touched()
        return EMPTY

    @method
    def Delete(self, Shift: object = MISSING) -> object:
        """Delete, which pulls the cells below or to the right up or left."""
        area = self.first
        up = Shift is MISSING or int(to_integer(Shift, "Long")) == -4162  # xlUp
        moved: dict[tuple[int, int], Cell] = {}
        for (row, column), cell in self.sheet.cells_.items():
            if area.contains(row, column):
                continue
            if up and column >= area.left and column <= area.right and row > area.bottom:
                moved[(row - area.rows, column)] = cell
            elif not up and row >= area.top and row <= area.bottom and column > area.right:
                moved[(row, column - area.columns)] = cell
            else:
                moved[(row, column)] = cell
        self.sheet.cells_ = moved
        self.sheet.shape_changed()
        return EMPTY

    @method
    def Insert(self, Shift: object = MISSING, CopyOrigin: object = MISSING) -> object:
        area = self.first
        down = Shift is MISSING or int(to_integer(Shift, "Long")) == -4121  # xlDown
        moved: dict[tuple[int, int], Cell] = {}
        for (row, column), cell in self.sheet.cells_.items():
            if down and area.left <= column <= area.right and row >= area.top:
                moved[(row + area.rows, column)] = cell
            elif not down and area.top <= row <= area.bottom and column >= area.left:
                moved[(row, column + area.columns)] = cell
            else:
                moved[(row, column)] = cell
        self.sheet.cells_ = moved
        self.sheet.shape_changed()
        return EMPTY

    @method
    def Copy(self, Destination: object = MISSING) -> object:
        if Destination is MISSING:
            self.sheet.book.application.cut_copy_mode = VBAInt(1, "Long")
            return True
        if not isinstance(Destination, Range):
            raise error(1004, "Copy needs a range to copy to")
        area = self.bounded()
        target = Destination.first
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                source = self.sheet.cell(row, column)
                if source is None:
                    Destination.sheet.cells_.pop(
                        (target.top + row - area.top, target.left + column - area.left), None
                    )
                    continue
                copy = Cell(**{field_.name: getattr(source, field_.name) for field_ in _CELL_FIELDS})
                if copy.formula:
                    # A copied formula moves with the cell, the same way
                    # one written to a block does.
                    copy.formula = shift_text(copy.formula, target.top - area.top, target.left - area.left)
                    copy.stale = True
                    copy.value = EMPTY
                Destination.sheet.cells_[(target.top + row - area.top, target.left + column - area.left)] = copy
        Destination.sheet.shape_changed()
        return True

    @method
    def End(self, Direction: object = MISSING) -> object:
        """The cell you land on with Ctrl and an arrow key."""
        if Direction is MISSING:
            raise error(449)
        which = int(to_integer(Direction, "Long"))
        row, column = self.first.top, self.first.left
        steps = {-4162: (-1, 0), -4121: (1, 0), -4159: (0, -1), -4161: (0, 1)}
        if which not in steps:
            raise error(1004, "End takes xlUp, xlDown, xlToLeft or xlToRight")
        down, across = steps[which]

        def occupied(at_row: int, at_column: int) -> bool:
            found = self.sheet.cell(at_row, at_column)
            return found is not None and not found.is_blank()

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
    def Find(self, What: object = MISSING, *rest: object) -> object:
        raise VBAUnsupportedError("Range.Find is not implemented by pyOpenVBA")

    @method
    def AutoFit(self) -> object:
        return EMPTY

    @method
    def Merge(self, Across: object = MISSING) -> object:
        raise VBAUnsupportedError("Range.Merge is not implemented by pyOpenVBA")

    # -- Python side

    def vba_iterate(self) -> Iterator[object]:
        area = self.bounded()
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                yield Range(self.sheet, [Area(row, column, row, column, area.sheet)])

    def _read(self, row: int, column: int) -> object:
        cell = self.sheet.cell(row, column)
        if cell is None:
            return EMPTY
        if cell.formula:
            from pyopenvba.apps.excel._calc import as_vba

            return as_vba(self.sheet.book.calculator.value_of(self.sheet.name, row, column), cell)
        return cell.value

    def writable_positions(self) -> list[tuple[int, int]]:
        """Which cells a write touches, refusing to fill a whole column."""
        out: list[tuple[int, int]] = []
        for area in self.areas:
            if area.rows * area.columns > 1048576:
                raise error(
                    ERR_APPLICATION_DEFINED,
                    "writing to a whole row or column at once is not something pyOpenVBA does",
                )
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    out.append((row, column))
        return out

    def _write(self, value: object) -> None:
        if isinstance(value, VBAArray):
            self._write_array(value)
            return
        if isinstance(value, VBAObject):
            value = value.vba_value()
        if isinstance(value, str) and value.startswith("="):
            self._set_formula(value)
            return
        stored = as_cell_value(value)
        for row, column in self.writable_positions():
            cell = self.sheet.cell(row, column, create=True)
            assert cell is not None
            cell.value = stored
            cell.formula = ""
            cell.stale = False
            self.sheet.cell_changed(row, column)

    def _write_array(self, array: VBAArray) -> None:
        area = self.first
        if array.dimensions == 1:
            # A flat array is one row, laid across and repeated down
            # every row of the range, which is what Excel does with it.
            items = array.elements()
            for offset, item in enumerate(items):
                column = area.left + offset
                if column > area.right:
                    break
                for row in range(area.top, area.bottom + 1):
                    self._put(row, column, item)
        else:
            rows, columns = array.bounds[0], array.bounds[1]
            for row_index in range(rows[0], rows[1] + 1):
                for column_index in range(columns[0], columns[1] + 1):
                    row = area.top + row_index - rows[0]
                    column = area.left + column_index - columns[0]
                    if row > area.bottom or column > area.right:
                        continue
                    self._put(row, column, array.get([row_index, column_index]))
        self.sheet.touched()

    def _put(self, row: int, column: int, value: object) -> None:
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        if isinstance(value, str) and value.startswith("="):
            cell.formula = value
            cell.stale = True
            cell.value = EMPTY
            self.sheet.cell_changed(row, column)
            return
        cell.value = as_cell_value(value)
        cell.formula = ""
        cell.stale = False
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


# --- formatting -------------------------------------------------------------------------------


class Font(ExcelObject):
    """A range's font.  Reading answers for the first cell, writing sets all."""

    vba_type_name = "Font"

    def __init__(self, target: Range) -> None:
        self.target = target

    def _first(self) -> Cell:
        return self.target.sheet.cell(self.target.first.top, self.target.first.left) or Cell()

    def _each(self) -> Iterator[Cell]:
        for row, column in self.target.writable_positions():
            cell = self.target.sheet.cell(row, column, create=True)
            assert cell is not None
            yield cell

    @member
    def Bold(self) -> object:
        return self._first().font_bold

    @setter("Bold")
    def _set_bold(self, value: object) -> None:
        for cell in self._each():
            cell.font_bold = to_bool(value)

    @member
    def Italic(self) -> object:
        return self._first().font_italic

    @setter("Italic")
    def _set_italic(self, value: object) -> None:
        for cell in self._each():
            cell.font_italic = to_bool(value)

    @member
    def Name(self) -> object:
        return self._first().font_name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        for cell in self._each():
            cell.font_name = to_text(value)

    @member
    def Size(self) -> object:
        return self._first().font_size

    @setter("Size")
    def _set_size(self, value: object) -> None:
        for cell in self._each():
            cell.font_size = float(to_number(value))

    @member
    def Color(self) -> object:
        return VBAInt(self._first().font_color, "Long")

    @setter("Color")
    def _set_color(self, value: object) -> None:
        for cell in self._each():
            cell.font_color = int(to_integer(value, "Long"))

    @member
    def Parent(self) -> object:
        return self.target


class Interior(ExcelObject):
    """A range's fill."""

    vba_type_name = "Interior"

    def __init__(self, target: Range) -> None:
        self.target = target

    @member
    def Color(self) -> object:
        # Excel hands a colour back as a Double, not as a Long.
        cell = self.target.sheet.cell(self.target.first.top, self.target.first.left)
        return float(cell.interior_color if cell and cell.interior_color is not None else 16777215)

    @setter("Color")
    def _set_color(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        for row, column in self.target.writable_positions():
            cell = self.target.sheet.cell(row, column, create=True)
            assert cell is not None
            cell.interior_color = wanted

    @member
    def ColorIndex(self) -> object:
        cell = self.target.sheet.cell(self.target.first.top, self.target.first.left)
        return VBAInt(-4142 if cell is None or cell.interior_color is None else 0, "Long")

    @setter("ColorIndex")
    def _set_color_index(self, value: object) -> None:
        if int(to_integer(value, "Long")) == -4142:  # xlColorIndexNone
            for row, column in self.target.writable_positions():
                cell = self.target.sheet.cell(row, column)
                if cell is not None:
                    cell.interior_color = None
            return
        raise VBAUnsupportedError("Interior.ColorIndex takes a palette pyOpenVBA does not model; set Color instead")

    @member
    def Parent(self) -> object:
        return self.target


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


class Names(VBACollection, ExcelObject):
    """A workbook's defined names."""

    vba_type_name = "Names"

    def __init__(self, book: Workbook) -> None:
        self.book = book
        self.entries: list[NameEntry] = []
        #: False until something adds, changes or removes a name, so an
        #: untouched workbook's names are never rewritten.
        self.changed = False

    def vba_items(self) -> list[object]:
        return [DefinedName(entry) for entry in self.entries]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            found = self.find(index)
            if found is None:
                raise error(1004, f"there is no name called {index}")
            return found
        return super().vba_lookup(index, items)

    def find(self, name: str) -> DefinedName | None:
        for entry in self.entries:
            if entry.name.lower() == name.lower():
                return DefinedName(entry)
        return None

    @method
    def Add(self, Name: object = MISSING, RefersTo: object = MISSING, *rest: object) -> object:
        if Name is MISSING or RefersTo is MISSING:
            raise error(449)
        wanted = to_text(Name)
        refers = RefersTo
        # A name refers to a sheet, not to a workbook: RefersTo reads
        # =Sheet1!$A$1:$B$2, with no [Book1] in front of it.
        text = (
            "=" + refers.first.address(with_sheet=True)
            if isinstance(refers, Range)
            else to_text(refers)
        )
        self.entries = [entry for entry in self.entries if entry.name.lower() != wanted.lower()]
        entry = NameEntry(wanted, text, self.book)
        self.entries.append(entry)
        self.changed = True
        self.book.saved = False
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
    def RefersTo(self) -> object:
        return self.entry.refers_to

    @setter("RefersTo")
    def _set_refers_to(self, value: object) -> None:
        self.entry.refers_to = to_text(value)
        self.entry.book.names_.changed = True
        self.entry.book.saved = False

    @member(default=True)
    def Value(self) -> object:
        return self.entry.refers_to

    @member
    def RefersToRange(self) -> object:
        return self.refers_to_range()

    @method
    def Delete(self) -> object:
        book = self.entry.book
        book.names_.entries = [one for one in book.names_.entries if one is not self.entry]
        book.names_.changed = True
        book.saved = False
        return EMPTY

    def refers_to_range(self) -> Range:
        text = self.entry.refers_to.lstrip("=")
        book = self.entry.book
        try:
            areas = parse_reference(text)
        except ValueError:
            raise error(1004, f"{self.entry.name} refers to {text!r}, which is not a range") from None
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
    """The worksheet functions a macro calls through Application.

    A short list on purpose: each one here behaves as Excel's does, and
    anything else says it is not implemented rather than guessing.
    """

    vba_type_name = "WorksheetFunction"

    def __init__(self, application: Application) -> None:
        self.application = application

    def vba_get(self, name: str, args: Any = (), named: Any = None) -> object:
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

    @method
    def Sum(self, *args: object) -> object:
        return sum((float(to_number(one)) for one in _numbers(args)), 0.0)

    @method
    def Average(self, *args: object) -> object:
        values = [float(to_number(one)) for one in _numbers(args)]
        if not values:
            raise error(1004, "Average has nothing to average")
        return sum(values) / len(values)

    @method
    def Max(self, *args: object) -> object:
        values = [float(to_number(one)) for one in _numbers(args)]
        return max(values) if values else 0.0

    @method
    def Min(self, *args: object) -> object:
        values = [float(to_number(one)) for one in _numbers(args)]
        return min(values) if values else 0.0

    @method
    def Count(self, *args: object) -> object:
        # A worksheet function hands back a Double even when counting.
        return float(len(list(_numbers(args))))

    @method
    def CountA(self, *args: object) -> object:
        return float(sum(1 for one in _flatten(args) if one is not EMPTY and one != ""))

    @method
    def Trim(self, Arg1: object = MISSING) -> object:
        return " ".join(to_text(Arg1).split())

    @method
    def Proper(self, Arg1: object = MISSING) -> object:
        return to_text(Arg1).title()

    @method
    def VLookup(
        self,
        Arg1: object = MISSING,
        Arg2: object = MISSING,
        Arg3: object = MISSING,
        Arg4: object = MISSING,
    ) -> object:
        if not isinstance(Arg2, Range):
            raise error(1004, "VLookup needs a range to look in")
        column = int(to_integer(Arg3, "Long"))
        exact = Arg4 is not MISSING and not to_bool(Arg4)
        area = Arg2.bounded()
        if column < 1 or column > area.columns:
            raise error(1004, "VLookup was asked for a column outside the range")
        wanted = Arg1
        for row in range(area.top, area.bottom + 1):
            cell = Arg2.sheet.cell(row, area.left)
            value = cell.value if cell is not None else EMPTY
            if _same(value, wanted):
                found = Arg2.sheet.cell(row, area.left + column - 1)
                return found.value if found is not None else EMPTY
        if exact or True:
            raise error(1004, "VLookup found nothing")
        return EMPTY


def _flatten(args: tuple[object, ...]) -> Iterator[object]:
    for one in args:
        if isinstance(one, Range):
            area = one.bounded()
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    cell = one.sheet.cell(row, column)
                    yield cell.value if cell is not None else EMPTY
        elif isinstance(one, VBAArray):
            yield from one.elements()
        elif one is not MISSING:
            yield one


def _numbers(args: tuple[object, ...]) -> Iterator[object]:
    for value in _flatten(args):
        if isinstance(value, bool):
            continue
        if isinstance(value, (int, float, VBADate)) or isinstance(value, type(EMPTY)) and value is not EMPTY:
            yield value
        elif hasattr(value, "__float__") and not isinstance(value, str):
            yield value


def _same(left: object, right: object) -> bool:
    if isinstance(left, str) or isinstance(right, str):
        return to_text(left).lower() == to_text(right).lower()
    try:
        return float(to_number(left)) == float(to_number(right))
    except Exception:  # noqa: BLE001
        return False


# --- helpers ----------------------------------------------------------------------------------------


def _display_text(cell: Cell) -> str:
    """What the cell shows, which is its value through its number format."""
    from pyopenvba.formula._values import ExcelError, number_text

    if cell.stale:
        return ""
    if isinstance(cell.value, ExcelError):
        return cell.value.name
    if cell.number_format in ("General", "") or cell.value is EMPTY:
        if cell.value is EMPTY:
            return ""
        if isinstance(cell.value, float):
            return number_text(cell.value)
        return to_text(cell.value)
    from pyopenvba.access._format import format_value

    value = cell.value.to_datetime() if isinstance(cell.value, VBADate) else cell.value
    return format_value(value, cell.number_format)


def _formula_text(value: object) -> str:
    if value is EMPTY:
        return ""
    return to_text(value)


def as_cell_value(value: object) -> object:
    """A value as a cell holds it.

    A cell keeps every number as a Double, whatever width the macro
    computed it in, and reads a string the way it reads typing: "5"
    becomes the number, "1/2/2020" becomes a date, and "" leaves the
    cell empty.  Both measured against Excel.
    """
    if isinstance(value, bool) or isinstance(value, VBADate):
        return value
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    if isinstance(value, str):
        return _from_text(value)
    return value


def _from_text(text: str) -> object:
    """A string written into a cell, read as Excel reads typing."""
    from pyopenvba.interpreter._values import parse_date_text

    stripped = text.strip()
    if not stripped:
        return EMPTY
    try:
        return float(stripped)
    except ValueError:
        pass
    when = parse_date_text(stripped)
    return when if when is not None else text
