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

import datetime as _dt
import re
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, replace
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any, ClassVar

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_reference
from pyopenvba.exceptions import PowerQueryError, VBAUnsupportedError, VBARuntimeError
from pyopenvba.formula._parse import shift_text
from pyopenvba.formula._prefixes import at_kept, at_shown
from pyopenvba.formula._r1c1 import from_a1, to_a1
from pyopenvba.apps.excel._creator import Created
from pyopenvba.apps.excel._find import FindState
from pyopenvba.apps.excel import _arrays, _column_fit, _dimensions, _events, _merges, _names, _precision
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
    to_date,
    to_integer,
    to_number,
    to_text,
    type_name,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._rich_text import Run
    from pyopenvba.apps.excel._udf import EvaluationOrigin
    from pyopenvba.powerquery._opc import Entry
    from pyopenvba.apps.excel._charts import ChartState
    from pyopenvba.apps.excel._autofilter import SheetFilter
    from pyopenvba.apps.excel._protection import BookProtection, Gate, SheetProtection
    from pyopenvba.apps.excel._clipboard import Clip
    from pyopenvba.apps.excel._data_tables import DataTable
    from pyopenvba.apps.excel._document_properties import Properties
    from pyopenvba.apps.excel._external_links import LinkedBook
    from pyopenvba.apps.excel._hyperlinks import Link
    from pyopenvba.apps.excel._notes import Note
    from pyopenvba.apps.excel._sort import SortState
    from pyopenvba.apps.excel._sparklines import State as SparklineState
    from pyopenvba.apps.excel._spills import Spill
    from pyopenvba.apps.excel._styles import Color, Style, Stylesheet
    from pyopenvba.apps.excel._tables import Table
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable
    from pyopenvba.apps.excel._outline import OutlineState
    from pyopenvba.apps.excel._page_breaks import SheetBreaks
    from pyopenvba.apps.excel._custom_properties import Property as CustomProperty
    from pyopenvba.apps.excel._consolidate import Settings as ConsolidationSettings
    from pyopenvba.apps.excel._page_setup import PageState
    from pyopenvba.apps.excel._typing import Typed
    from pyopenvba.apps.excel._validation import Rule as ValidationRule
    from pyopenvba.apps.excel._conditional import Rule as ConditionalRule
    from pyopenvba.formula._structured import TableShape
    from pyopenvba.interpreter._runtime import Interpreter

_LIBRARY = "excel"
#: The cells Range.Value reads in full: a whole column, 1,048,576 rows.
_WHOLE_COLUMN_CELLS = 1_048_576
#: The longest formula Excel keeps, in characters, and the longest FormulaArray takes
#: (tests/fixtures/formula_refusals.json).
LONGEST_FORMULA = 8192
LONGEST_ARRAY_FORMULA = 255


class ExcelObject(Created):
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
        member = spec.name if spec is not None else name
        self.guard_set(member)
        super().vba_set(name, value, args, named, by_ref=by_ref)
        self.after_set(member)

    def guard_set(self, member: str) -> None:
        """Refuse a property a protected sheet does not let a macro set: the classes of cells say which."""

    def after_set(self, member: str) -> None:
        """What a property a macro set leaves behind besides its value; the classes of formats say what."""


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
    #: The shared formula the cell's formula belongs to, a key of its sheet's shared_groups; None for its own.
    shared: int | None = None
    #: True for a dynamic-array formula, one Formula2 wrote that answers with what spills (see _spills).
    dynamic: bool = False
    #: The first cell of the dynamic-array formula whose answer this cell shows, where it spilled here.
    spilled_from: tuple[int, int] | None = None
    #: A clipped CSE fill can be dirty in memory but saved without a per-cell dirty flag.
    save_dirty: bool | None = None
    #: The runs of fonts the cell's text is in, where it is in more than its own (see _rich_text).
    runs: tuple[Run, ...] | None = None
    #: The error checks Range.Errors(...).Ignore set, one bit each from xlEvaluateToError (see _error_checks).
    ignored_errors: int = 0
    #: Range.ID, which goes with the cell's contents (tests/fixtures/range_small_members.json).
    html_id: str = ""

    @property
    def number_format(self) -> str:
        return self.style.number_format if self.style is not None else "General"

    def is_blank(self) -> bool:
        return self.value is EMPTY and not self.formula and self.style is None and not self.html_id and not self.ignored_errors


# --- Application ------------------------------------------------------------------------


def coalesced(areas: list[Area]) -> list[Area]:
    """Areas joined as Union joins them (see _coalesced)."""
    return _coalesced(areas)


def _coalesced(areas: list[Area]) -> list[Area]:
    """Union's areas as Excel leaves them: two areas merge while one holds the
    other, or they share their rows or their columns and touch or overlap; the
    merged area goes to the end, so A1:A2,C1 with A3 is C1,A1:A3 (native probes)."""
    pending = list(areas)
    while True:
        for first in range(len(pending)):
            for second in range(first + 1, len(pending)):
                merged = _merged(pending[first], pending[second])
                if merged is not None:
                    pending = [one for index, one in enumerate(pending) if index not in (first, second)] + [merged]
                    break
            else:
                continue
            break
        else:
            return pending


def _joined_in_place(areas: list[Area]) -> list[Area]:
    """Intersect's pieces as Excel leaves them: two join as Union's do (see _merged), but the joined area takes the
    earlier one's place, so A2, A2, C2 is A2,C2 (tests/fixtures/intersect_areas.json)."""
    pending = list(areas)
    while True:
        for first in range(len(pending)):
            for second in range(first + 1, len(pending)):
                merged = _merged(pending[first], pending[second])
                if merged is not None:
                    pending = [merged if index == first else one for index, one in enumerate(pending)
                               if index != second]
                    break
            else:
                continue
            break
        else:
            return pending


def _merged(one: Area, other: Area) -> Area | None:
    def holds(outer: Area, inner: Area) -> bool:
        return (outer.top <= inner.top and inner.bottom <= outer.bottom
                and outer.left <= inner.left and inner.right <= outer.right)

    if holds(one, other):
        return one
    if holds(other, one):
        return other
    if (one.left, one.right) == (other.left, other.right) and one.top <= other.bottom + 1 and other.top <= one.bottom + 1:
        return Area(min(one.top, other.top), one.left, max(one.bottom, other.bottom), one.right, one.sheet)
    if (one.top, one.bottom) == (other.top, other.bottom) and one.left <= other.right + 1 and other.left <= one.right + 1:
        return Area(one.top, min(one.left, other.left), one.bottom, max(one.right, other.right), one.sheet)
    return None


#: Application.International's answers in the en-US locale the model works in, by xlApplicationInternational index:
#: country codes, separators, the letters R1C1 and number formats use, currency and date settings, and at 46 the
#: culture's name (tests/fixtures/international.json).
_INTERNATIONAL: dict[int, object] = {
    1: 1.0, 2: 1.0, 3: ".", 4: ",", 5: ",", 6: "R", 7: "C", 8: "r", 9: "c", 10: "[", 11: "]", 12: "{", 13: "}",
    14: ",", 15: ";", 16: "@", 17: "/", 18: ":", 19: "y", 20: "m", 21: "d", 22: "h", 23: "m", 24: "s", 25: "$",
    26: "General", 27: 2.0, 28: 0.0, 29: 2.0, 30: 3.0, 31: 3.0, 32: 0.0, 33: False, 34: False, 35: False,
    36: False, 37: True, 38: False, 39: True, 40: True, 41: False, 42: False, 43: True, 44: True, 45: False,
    46: "en-US",
}


def _separator_kept(name: str, value: object, current: str) -> None:
    """Setting a separator to the one it is; another one would change every number shown, which the model does not
    follow."""
    if to_text(value) != current:
        raise VBAUnsupportedError(f"Application.{name} set to another separator is not implemented")


#: Application.ErrorCheckingOptions as a new Excel session reads them on the measuring machine
#: (tests/fixtures/error_checking_options.json); Range.Errors answers whatever they say.
_ERROR_CHECKS = {"BackgroundChecking": True, "EvaluateToError": True, "TextDate": True, "NumberAsText": True,
                 "InconsistentFormula": True, "OmittedCells": True, "UnlockedFormulaCells": True,
                 "EmptyCellReferences": False, "ListDataValidation": True, "InconsistentTableFormula": True,
                 "MisleadingNumberFormats": True}


def _error_check(key: str) -> tuple[Callable[..., object], Callable[..., None]]:
    """The getter and setter of one ErrorCheckingOptions switch."""

    def read(self: ErrorCheckingOptions) -> object:
        return self.kept[key]

    def write(self: ErrorCheckingOptions, value: object) -> None:
        self.kept[key] = to_bool(value)

    read.__name__ = key
    return member(read), setter(key)(write)


class ErrorCheckingOptions(ExcelObject):
    """Application.ErrorCheckingOptions: the session's error-checking settings, kept for the session."""

    vba_type_name = "ErrorCheckingOptions"

    def __init__(self, application: Application) -> None:
        self.application = application

    @property
    def kept(self) -> dict[str, object]:
        kept: dict[str, object] | None = getattr(self.application, "error_checking", None)
        if kept is None:
            kept = {**_ERROR_CHECKS, "IndicatorColorIndex": -4105}
            setattr(self.application, "error_checking", kept)
        return kept

    BackgroundChecking, _set_background_checking = _error_check("BackgroundChecking")
    EvaluateToError, _set_evaluate_to_error = _error_check("EvaluateToError")
    TextDate, _set_text_date = _error_check("TextDate")
    NumberAsText, _set_number_as_text = _error_check("NumberAsText")
    InconsistentFormula, _set_inconsistent_formula = _error_check("InconsistentFormula")
    OmittedCells, _set_omitted_cells = _error_check("OmittedCells")
    UnlockedFormulaCells, _set_unlocked_formula_cells = _error_check("UnlockedFormulaCells")
    EmptyCellReferences, _set_empty_cell_references = _error_check("EmptyCellReferences")
    ListDataValidation, _set_list_data_validation = _error_check("ListDataValidation")
    InconsistentTableFormula, _set_inconsistent_table_formula = _error_check("InconsistentTableFormula")
    MisleadingNumberFormats, _set_misleading_number_formats = _error_check("MisleadingNumberFormats")

    @member
    def IndicatorColorIndex(self) -> object:
        return VBAInt(int(to_integer(self.kept["IndicatorColorIndex"])), "Long")

    @setter("IndicatorColorIndex")
    def _set_indicator_color_index(self, value: object) -> None:
        number = int(to_integer(value))
        if number != -4105 and not 1 <= number <= 56:
            raise VBAUnsupportedError("ErrorCheckingOptions.IndicatorColorIndex outside the palette is not implemented")
        self.kept["IndicatorColorIndex"] = number

    @member
    def Parent(self) -> object:
        return self.application

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


class Application(ExcelObject):
    """The Excel instance: the root of everything a macro can reach."""

    vba_type_name = "Application"
    #: The events the model raises in an application sink, which the member inventory counts
    #: (see _events).
    raised_events: ClassVar[frozenset[str]] = frozenset({
        "WorkbookOpen", "WorkbookBeforeSave", "WorkbookAfterSave", "WorkbookBeforeClose",
        "WorkbookActivate", "WorkbookDeactivate", "WorkbookNewSheet", "SheetChange", "SheetCalculate",
        "SheetSelectionChange", "SheetActivate", "SheetDeactivate", "SheetPivotTableUpdate"})

    def vba_information_value(self) -> object:
        return self.Name()

    def __init__(self) -> None:
        from pyopenvba.interpreter._locale import EN_US

        #: The Windows Regional format the VBA runs under (ExcelApplication's locale).
        self.locale = EN_US
        self.workbooks_ = Workbooks(self)
        self.screen_updating = True
        self.display_alerts = True
        self.enable_events = True
        self.calculation = -4105  # xlCalculationAutomatic
        self.calculate_before_save = True
        #: Whether a loop of formulas is worked out over and over, and how many times and to what change at most, as
        #: Application.Iteration, MaxIterations and MaxChange have it (see _calc).
        self.iteration = False
        self.max_iterations = 100
        self.max_change = 0.001
        #: xlA1 or xlR1C1, as Application.ReferenceStyle has it.
        self.reference_style = 1
        self.use_system_separators = True
        #: What a macro last set Application.StatusBar to, None until one does.
        self.status_bar: object = None
        #: The settings that only read back what a macro set, by member name (see _application_settings).
        self.settings: dict[str, object] = {}
        #: Whether PageSetup settings take effect at once, and those held back while they do not (see _page_setup).
        self.print_communication = True
        self.held_page_settings: list[tuple[Worksheet, object, Callable[[], None]]] = []
        #: What Copy or Cut last put on the clipboard (see _clipboard).
        self.clipboard: Clip | None = None
        self.user_name = "pyOpenVBA"
        self.interpreter: Interpreter | None = None
        self.executing_projects: list[Interpreter] = []
        #: The cells whose formulas are running a VBA Function now, innermost last (see _udf).
        self.udf_callers: list[Range] = []
        #: Active formula clients and the origin captured by each VBA UDF invocation.
        self.formula_origins: list[EvaluationOrigin] = []
        self.udf_origins: list[EvaluationOrigin | None] = []
        self.active_book: Workbook | None = None
        #: The workbooks whose windows have been to the front, front first (see _windows).
        self.window_order: list[Workbook] = []
        self._this_workbook: Workbook | None = None
        self.worksheet_function = WorksheetFunction(self)
        self.find_state = FindState()
        #: The number the next pivot table a macro makes without naming it takes, counted through the session
        #: whatever workbook the table is in (see _pivot_create.next_table_name).
        self.pivot_table_number = 1
        #: The number format a macro last met in a cell, which the next Subtotal gives its first label in General's
        #: place (see Range.note_seen and _subtotals).
        self.recent_format = "General"
        #: The spelling ConvertFormula first saw each name in, by the name in lower case (see _convert_formula).
        self.convert_spellings: dict[str, str] = {}
        #: Application.CalculationInterruptKey, xlAnyKey until a macro sets it; and whether a workbook's
        #: ForceFullCalculation has held every workbook to full calculation, as it does until Excel quits (see
        #: _calculation_state).
        self.calculation_interrupt_key = 2
        self.full_calculation = False
        #: What a text file Workbooks.Open reads splits at, as the last Format it was given names, None for nothing
        #: (see _text_files.remember_format).
        self.text_delimiter: str | None = "\t"
        from pyopenvba.apps.excel._text_files import ImportSettings

        #: What OpenText and TextToColumns split at, and OpenText's code page, as the session last named them.
        self.text_import = ImportSettings()
        #: Application.SheetsInNewWorkbook, Excel's own default until a macro sets it.
        self.sheets_in_new_workbook = 1
        #: The custom lists a macro added, after Excel's own four (see _custom_lists).
        self.custom_lists: list[tuple[str, ...]] = []

    # -- custom lists

    @member
    def CustomListCount(self) -> object:
        from pyopenvba.apps.excel._custom_lists import lists

        return VBAInt(len(lists(self)), "Long")

    @member
    def GetCustomListContents(self, ListNum: object = MISSING) -> object:
        from pyopenvba.apps.excel._custom_lists import contents

        return contents(self, ListNum)

    @member
    def GetCustomListNum(self, ListArray: object = MISSING) -> object:
        from pyopenvba.apps.excel._custom_lists import find

        return VBAInt(find(self, ListArray), "Long")

    @method
    def AddCustomList(self, ListArray: object = MISSING, ByRow: object = MISSING) -> object:
        from pyopenvba.apps.excel._custom_lists import add

        add(self, ListArray, ByRow)
        return EMPTY

    @method
    def DeleteCustomList(self, ListNum: object = MISSING) -> object:
        from pyopenvba.apps.excel._custom_lists import delete

        delete(self, ListNum)
        return EMPTY

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
    def ErrorCheckingOptions(self) -> object:
        return ErrorCheckingOptions(self)

    @member
    def EnableEvents(self) -> object:
        return self.enable_events

    @setter("EnableEvents")
    def _set_enable_events(self, value: object) -> None:
        self.enable_events = to_bool(value)

    @member
    def Calculation(self) -> object:
        return VBAInt(self.calculation, "Long")

    @member
    def CalculateBeforeSave(self) -> object:
        return self.calculate_before_save

    @setter("CalculateBeforeSave")
    def _set_calculate_before_save(self, value: object) -> None:
        self.calculate_before_save = to_bool(value)

    @member
    def Iteration(self) -> object:
        return self.iteration

    @setter("Iteration")
    def _set_iteration(self, value: object) -> None:
        """Turned on, it works each loop of formulas out at once, as automatic calculation does; turned off, the
        loops keep what they hold (tests/fixtures/iteration.json)."""
        before = self.iteration
        self.iteration = to_bool(value)
        if self.iteration and not before and self.calculation != -4135:  # xlCalculationManual
            for book in self.workbooks_.books:
                if book.calculator.loop_cells():
                    _events.calculated_all(book, book.calculator.calculate_all(volatile=False))

    @member
    def MaxIterations(self) -> object:
        return VBAInt(self.max_iterations, "Long")

    @setter("MaxIterations")
    def _set_max_iterations(self, value: object) -> None:
        """1 to 32767, rounded as a Long is; anything else is error 1004 and leaves the setting as it was."""
        count = int(to_integer(value, "Long"))
        if not 1 <= count <= 32767:
            raise error(1004, "Unable to set the MaxIterations property of the Application class")
        self.max_iterations = count

    @member
    def MaxChange(self) -> object:
        return self.max_change

    @setter("MaxChange")
    def _set_max_change(self, value: object) -> None:
        """Any number at all, 0 and negative ones among them, which then stop no iteration early."""
        self.max_change = float(to_number(value))

    @member
    def International(self, Index: object = MISSING) -> object:
        """A setting of the Regional format, by its xlApplicationInternational index; any other index is error
        1004 (tests/fixtures/international.json, and tests/fixtures/excel_locale_live.json for ja-JP)."""
        if Index is MISSING:
            raise error(449)
        index = int(to_integer(Index, "Long"))
        found = {**_INTERNATIONAL, **dict(self.locale.international)}.get(index, MISSING)
        if found is MISSING:
            raise error(1004, "Application-defined or object-defined error")
        if index == 46 and self.locale.international:
            raise VBAUnsupportedError(f"Application.International(46) under {self.locale.name} is not implemented")
        return found

    @member
    def StandardFont(self) -> object:
        """The font a new workbook's Normal style takes, Aptos Narrow, as today's Excel's (tests/fixtures/
        column_autofit.json)."""
        return "Aptos Narrow"

    @setter("StandardFont")
    def _set_standard_font(self, value: object) -> None:
        if to_text(value).lower() != "aptos narrow":
            raise VBAUnsupportedError("Application.StandardFont set to another font, which new workbooks take, is not "
                                      "implemented")

    @member
    def StandardFontSize(self) -> object:
        """The size a new workbook's Normal font takes, 11."""
        return 11.0

    @setter("StandardFontSize")
    def _set_standard_font_size(self, value: object) -> None:
        if float(to_number(value)) != 11.0:
            raise VBAUnsupportedError("Application.StandardFontSize set to another size, which new workbooks take, is "
                                      "not implemented")

    @member
    def DecimalSeparator(self) -> object:
        return "."

    @setter("DecimalSeparator")
    def _set_decimal_separator(self, value: object) -> None:
        _separator_kept("DecimalSeparator", value, ".")

    @member
    def ThousandsSeparator(self) -> object:
        return ","

    @setter("ThousandsSeparator")
    def _set_thousands_separator(self, value: object) -> None:
        _separator_kept("ThousandsSeparator", value, ",")

    @member
    def UseSystemSeparators(self) -> object:
        return self.use_system_separators

    @setter("UseSystemSeparators")
    def _set_use_system_separators(self, value: object) -> None:
        # The custom separators are the system's, so turning this off changes nothing the model shows.
        self.use_system_separators = to_bool(value)

    @member
    def ReferenceStyle(self) -> object:
        """xlA1 (1) or xlR1C1 (-4150). Formulas and addresses read as A1 either way; Evaluate reads R1C1 in the
        R1C1 style, and a workbook keeps the style it was saved in (tests/fixtures/reference_style.json)."""
        return VBAInt(self.reference_style, "Long")

    @setter("ReferenceStyle")
    def _set_reference_style(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        # Any other number leaves the style as it was, with no error.
        if wanted in (1, -4150):
            self.reference_style = wanted

    @member
    def Caller(self, Index: object = MISSING) -> object:
        """The cell whose formula is running the VBA Function that asks (see _udf)."""
        if self.udf_callers:
            return self.udf_callers[-1]
        # A macro no cell or control called: the error value #REF! (native probe).
        return VBAErrorValue(2023)

    @setter("Calculation")
    def _set_calculation(self, value: object) -> None:
        before = self.calculation
        for book in self.workbooks_.books:
            book.calculator.settle_automatic_cache()
        self.calculation = int(to_integer(value, "Long"))
        if before != self.calculation and self.calculation == -4105:  # xlCalculationAutomatic
            # Back to automatic works out what manual mode left stale, with its
            # Calculate events. Volatile UDFs in the active workbook run too;
            # built-in volatile phases remain a separate conformance surface.
            for book in self.workbooks_.books:
                _events.calculated_all(book, book.calculator.calculate_all(
                    volatile=False, udf_volatile=book is self.active_book))

    @member
    def StatusBar(self) -> object:
        """False, a Boolean, until a macro sets it, and text from then on: what a macro set, a number or a logical as
        Excel spells it, FALSE once a macro gives the bar back to Excel (tests/fixtures/excel_everyday_3.json,
        statusbar_types; tests/fixtures/application_settings.json, status_hidden)."""
        value = self.status_bar
        if value is None:
            return False
        return ("TRUE" if value else "FALSE") if isinstance(value, bool) else to_text(value)

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
    def FindFormat(self) -> object:
        """The format Range.Find and Replace look for with SearchFormat, a new CellFormat over it each time
        (tests/fixtures/find_format.json)."""
        from pyopenvba.apps.excel._cell_format import CellFormat

        return CellFormat(self, self.find_state.find_format)

    @member
    def ReplaceFormat(self) -> object:
        """The format Range.Replace puts in with ReplaceFormat (tests/fixtures/find_format.json)."""
        from pyopenvba.apps.excel._cell_format import CellFormat

        return CellFormat(self, self.find_state.replace_format)

    @member
    def ThisWorkbook(self) -> object:
        from pyopenvba.apps.excel._projects import ExcelInterpreter

        if self.executing_projects:
            project = self.executing_projects[-1]
            if isinstance(project, ExcelInterpreter) and project.bridge.workbook is not None:
                return project.bridge.workbook
        book = self._this_workbook or self.active_book
        return book if book is not None else NOTHING

    @member
    def ActiveSheet(self) -> object:
        book = self.active_book
        return book.active_sheet if book is not None else NOTHING

    @member
    def ActiveWindow(self) -> object:
        from pyopenvba.apps.excel._windows import window_of

        return window_of(self.active_book) if self.active_book is not None else NOTHING

    @member
    def Windows(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._windows import Windows

        windows = Windows(self)
        return windows if Index is MISSING else windows.vba_get("Item", [Index])

    @method
    def Goto(self, Reference: object = MISSING, Scroll: object = MISSING) -> object:
        from pyopenvba.apps.excel._windows import goto

        goto(self, Reference, Scroll)
        return EMPTY

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
        """The active sheet's Range, but cells on another sheet are taken there (tests/fixtures/cross_sheet_ranges.json)."""
        return self._require_sheet().resolved_range(Cell1, Cell2, foreign=True)

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

    @member
    def PathSeparator(self) -> object:
        """The separator of Windows paths, which the model's Excel runs on (tests/fixtures/application_members.json)."""
        return "\\"

    @member
    def Value(self) -> object:
        return "Microsoft Excel"

    @member
    def OperatingSystem(self) -> object:
        """What Excel answers on 64-bit Windows 10 and 11 (tests/fixtures/application_members.json)."""
        return "Windows (64-bit) NT 10.00"

    @member
    def MeasurementUnit(self) -> object:
        """xlInches, as US English Windows has it."""
        return VBAInt(0, "Long")

    @method
    def InchesToPoints(self, Inches: object = MISSING) -> object:
        """Inches * 72 (tests/fixtures/application_members.json)."""
        return _measure(Inches) * 72

    @method
    def CentimetersToPoints(self, Centimeters: object = MISSING) -> object:
        """Centimeters * (72 / 2.54), the factor rounded to a Double first, to the bit over 32 live values
        (tests/fixtures/application_members.json)."""
        return _measure(Centimeters) * (72 / 2.54)

    @member
    def SheetsInNewWorkbook(self) -> object:
        """How many sheets Workbooks.Add with no template makes."""
        return VBAInt(self.sheets_in_new_workbook, "Long")

    @setter("SheetsInNewWorkbook")
    def _set_sheets_in_new_workbook(self, value: object) -> None:
        """1 to 255, rounded as a Long rounds; anything else error 1004 (tests/fixtures/application_members.json)."""
        count = int(to_integer(value, "Long"))
        if not 1 <= count <= 255:
            raise error(1004, "Method 'SheetsInNewWorkbook' of object '_Application' failed")
        self.sheets_in_new_workbook = count

    @member
    def ThisCell(self) -> object:
        """The cell whose formula is running the VBA Function that asks: an array formula's first cell, where Caller
        is the whole array; from anything else error 1004 (tests/fixtures/application_members.json)."""
        if not self.udf_callers:
            raise error(1004, "Method 'ThisCell' of object '_Application' failed")
        caller = self.udf_callers[-1]
        area = caller.first
        return Range(caller.sheet, [Area(area.top, area.left, area.top, area.left, caller.sheet.name)])

    # -- methods

    @method
    def Evaluate(self, Name: object = MISSING) -> object:
        return self.evaluate_text(to_text(Name))

    @method
    def ConvertFormula(self, Formula: object = MISSING, FromReferenceStyle: object = MISSING,
                       ToReferenceStyle: object = MISSING, ToAbsolute: object = MISSING,
                       RelativeTo: object = MISSING) -> object:
        """A formula's references from A1 to R1C1 or back, made absolute or relative, as Excel converts them; the
        error value #VALUE! for one it cannot (see _convert_formula)."""
        from pyopenvba.apps.excel._convert_formula import convert_formula

        return convert_formula(self, Formula, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo)

    @method
    def Run(self, Macro: object = MISSING, *args: object) -> object:
        from pyopenvba.apps.excel._projects import run_macro

        return run_macro(self, to_text(Macro), list(args))

    @method
    def Intersect(self, Arg1: object = MISSING, Arg2: object = MISSING, *rest: object) -> object:
        """The cells every range holds, area by area: each area of the ranges so far against each of the next, in
        that order, the pieces then joined in place (see _joined_in_place); Nothing where none is left
        (tests/fixtures/intersect_areas.json)."""
        ranges = [one for one in (Arg1, Arg2, *rest) if isinstance(one, Range)]
        if len(ranges) < 2:
            raise error(ERR_APPLICATION_DEFINED, "Intersect needs two ranges")
        if any(one.sheet is not ranges[0].sheet for one in ranges):
            raise error(1004, "Method 'Intersect' of object '_Application' failed")
        sheet = ranges[0].sheet
        pieces = [Area(area.top, area.left, area.bottom, area.right) for area in ranges[0].areas]
        for other in ranges[1:]:
            met = [Area(max(one.top, two.top), max(one.left, two.left), min(one.bottom, two.bottom),
                        min(one.right, two.right))
                   for one in pieces for two in other.areas]
            pieces = _joined_in_place([one for one in met if one.top <= one.bottom and one.left <= one.right])
        if not pieces:
            return NOTHING
        return Range(sheet, [Area(one.top, one.left, one.bottom, one.right, sheet.name) for one in pieces])

    @method
    def Union(self, Arg1: object = MISSING, Arg2: object = MISSING, *rest: object) -> object:
        ranges = [one for one in (Arg1, Arg2, *rest) if isinstance(one, Range)]
        if not ranges:
            raise error(ERR_APPLICATION_DEFINED, "Union needs a range")
        if any(one.sheet is not ranges[0].sheet for one in ranges):
            # Union of cells on two sheets is error 1004 (tests/fixtures/cross_sheet_ranges.json).
            raise error(1004, "Method 'Union' of object '_Application' failed")
        areas: list[Area] = []
        for one in ranges:
            areas.extend(one.areas)
        return Range(ranges[0].sheet, _coalesced(areas))

    @method
    def Calculate(self) -> object:
        """Work out what each workbook has to, every formula of one held to full calculation or opened wanting it
        (see _calculation_state)."""
        from pyopenvba.apps.excel._calculation_state import calculated, in_full

        for book in self.workbooks_.books:
            full = in_full(book)
            worked = book.calculator.calculate_all(asked=True, full=full)
            calculated(book, full)
            _events.calculated_all(book, worked)
        return EMPTY

    @method
    def CalculateFull(self) -> object:
        from pyopenvba.apps.excel._calculation_state import calculated

        for book in self.workbooks_.books:
            worked = book.calculator.calculate_all(full=True, asked=True)
            calculated(book, True)
            _events.calculated_all(book, worked)
        return EMPTY

    @method
    def CalculateFullRebuild(self) -> object:
        from pyopenvba.apps.excel._calculation_state import calculated

        for book in self.workbooks_.books:
            book.calculator.rebuild(keep_pending=True)
            worked = book.calculator.calculate_all(full=True, asked=True)
            calculated(book, True)
            _events.calculated_all(book, worked)
        return EMPTY

    @method
    def CalculateUntilAsyncQueriesDone(self) -> object:
        """Waits for queries refreshing in the background, which the model refreshes at once, and works nothing out
        (tests/fixtures/calculation_state.json)."""
        return EMPTY

    @member
    def CalculationState(self) -> object:
        """xlDone, xlCalculating while a formula's VBA Function runs, or xlPending while there is a formula to work
        out under manual calculation (see _calculation_state)."""
        from pyopenvba.apps.excel._calculation_state import state

        return state(self)

    @member
    def CalculationVersion(self) -> object:
        """The calculation engine's version, Excel 365's."""
        from pyopenvba.apps.excel._calculation_state import VERSION

        return VBAInt(VERSION, "Long")

    @member
    def CalculationInterruptKey(self) -> object:
        """xlAnyKey until a macro sets xlEscKey or xlNoKey; the model keeps the setting, as nothing interrupts its
        calculation."""
        return VBAInt(self.calculation_interrupt_key, "Long")

    @setter("CalculationInterruptKey")
    def _set_calculation_interrupt_key(self, value: object) -> None:
        from pyopenvba.apps.excel._calculation_state import set_interrupt_key

        set_interrupt_key(self, value)

    def vba_sink_attached(self) -> None:
        for book in self.workbooks_.books:
            book.vba_sink_attached()

    @method
    def Quit(self) -> object:
        self.workbooks_.books.clear()
        self.window_order.clear()
        self.active_book = None
        return EMPTY

    @method
    def Wait(self, Time: object = MISSING) -> object:
        """True once Time has come. A frozen clock (ExcelApplication.freeze_clock) moves on to Time, as if the wait
        had passed; under the real clock the model returns at once rather than block."""
        interpreter = self.interpreter
        if Time is not MISSING and interpreter is not None and interpreter.frozen_now is not None:
            until = to_date(Time).to_datetime()
            if until > interpreter.frozen_now:
                interpreter.freeze(until)
        return True

    @method
    def Volatile(self, Volatile: object = MISSING) -> object:
        if self.udf_callers:
            caller = self.udf_callers[-1]
            area = caller.first
            caller.sheet.book.calculator.set_udf_volatile(
                caller.sheet.name, area.top, area.left, True if Volatile is MISSING else to_bool(Volatile))
        return EMPTY

    # -- Python side

    def evaluate_text(self, text: str) -> object:
        """``[A1+1]`` and Evaluate("A1+1"), worked out on the active sheet as Excel's Evaluate works them."""
        from pyopenvba.apps.excel._evaluate import evaluated
        from pyopenvba.apps.excel._udf import evaluation_origin

        with evaluation_origin(self, "evaluate"):
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

    def activate_book(self, book: Workbook, *, events: bool = True) -> None:
        from pyopenvba.apps.excel._windows import brought_forward

        before = self.active_book
        brought_forward(self, book)
        self.active_book = book
        if self._this_workbook is None:
            self._this_workbook = book
        if events:
            _events.book_activated(book, before)

    def describe(self, indent: str = "") -> str:
        lines = [f"{indent}Excel {self.vba_get('Version')} ({len(self.workbooks_.books)} workbook(s))"]
        for book in self.workbooks_.books:
            lines.append(book.describe(indent + "  "))
        return "\n".join(lines)


def refuse_save_options(password: object, reservation: object, recommended: object, backup: object) -> None:
    """SaveAs's options the model does not save with: a password to open or to change the file, a recommendation
    to open it read-only, and a backup."""
    for name, value in (("Password", password), ("WriteResPassword", reservation)):
        if value is not MISSING and to_text(value):
            raise VBAUnsupportedError(f"SaveAs with a {name} is not implemented")
    for name, value in (("ReadOnlyRecommended", recommended), ("CreateBackup", backup)):
        if value is not MISSING and to_bool(value):
            raise VBAUnsupportedError(f"SaveAs with {name} True is not implemented")


def _measure(value: object) -> float:
    """A length InchesToPoints or CentimetersToPoints takes, converted as a Double argument converts it."""
    if value is MISSING:
        raise error(449)
    return float(to_number(value))


# --- workbooks -----------------------------------------------------------------------------


class Workbooks(VBACollection, ExcelObject):
    """The open workbooks."""

    vba_type_name = "Workbooks"

    def __init__(self, application: Application) -> None:
        self.application = application
        self.books: list[Workbook] = []
        #: The number the next new workbook of each kind takes: Book for one with no template, Sheet for one made
        #: from xlWBATWorksheet, each counting on its own and taking none back (tests/fixtures/workbook_names.json).
        self._next_numbers = {"Book": 1, "Sheet": 1}

    def next_name(self, stem: str = "Book") -> str:
        while True:
            name = f"{stem}{self._next_numbers[stem]}"
            self._next_numbers[stem] += 1
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
        book = Workbook(self.application, self.next_name("Book" if Template is MISSING else "Sheet"))
        # A workbook with no template has SheetsInNewWorkbook sheets, Sheet1 on (tests/fixtures/application_members.json).
        for number in range(1, (self.application.sheets_in_new_workbook if Template is MISSING else 1) + 1):
            book.add_sheet(f"Sheet{number}")
        book.saved = True
        self.books.append(book)
        from pyopenvba.apps.excel._projects import attach_project

        attach_project(book)
        self.application.activate_book(book)
        return book

    @method
    def Open(self, Filename: object = MISSING, UpdateLinks: object = MISSING, ReadOnly: object = MISSING,
             Format: object = MISSING, Password: object = MISSING, WriteResPassword: object = MISSING,
             IgnoreReadOnlyRecommended: object = MISSING, Origin: object = MISSING, Delimiter: object = MISSING,
             Editable: object = MISSING, Notify: object = MISSING, Converter: object = MISSING,
             AddToMru: object = MISSING, Local: object = MISSING, CorruptLoad: object = MISSING) -> object:
        """A workbook, or a CSV or text file read as Excel reads one (see _text_files). UpdateLinks, Password,
        WriteResPassword, IgnoreReadOnlyRecommended, Editable, Notify, AddToMru and Local change nothing for the
        files the model opens, which it neither updates links in nor decrypts (tests/fixtures/text_files.json)."""
        from pyopenvba.apps.excel._text_files import is_text, remember_format

        if Filename is MISSING:
            raise error(449)
        if Origin is not MISSING and int(to_integer(Origin, "Long")) != 2:
            raise VBAUnsupportedError("Workbooks.Open with an Origin other than xlWindows is not implemented")
        if CorruptLoad is not MISSING and int(to_integer(CorruptLoad, "Long")) != 0:
            raise VBAUnsupportedError("Workbooks.Open with a CorruptLoad other than xlNormalLoad is not implemented")
        if Converter is not MISSING:
            raise VBAUnsupportedError("Workbooks.Open with a Converter is not implemented")
        given = to_text(Filename)
        path = Path(given)
        if not path.is_file():
            raise error(1004, f"Sorry, we couldn't find {given}. Is it possible it was moved, renamed or deleted?")
        if path.suffix.lower() in (".xls", ".xla", ".xlt"):
            raise VBAUnsupportedError("opening an Excel 97-2003 workbook is not implemented")
        format_code = remember_format(self.application, Format, Delimiter)
        read_only = ReadOnly is not MISSING and to_bool(ReadOnly)
        if is_text(path):
            # A text file Workbooks.Open reads puts OpenText's code page back to Windows-1252
            # (tests/fixtures/open_text.json).
            self.application.text_import.origin = 1252
        try:
            return self.open_file(path, read_only=read_only, format_code=format_code)
        except PowerQueryError:
            import zipfile

            if zipfile.is_zipfile(path):
                raise
            raise error(1004, f"Excel cannot open the file '{path.name}' because the file format or file extension "
                              "is not valid. Verify that the file has not been corrupted and that the file extension "
                              "matches the format of the file.") from None

    @method
    def OpenText(self, Filename: object = MISSING, Origin: object = MISSING, StartRow: object = MISSING,
                 DataType: object = MISSING, TextQualifier: object = MISSING, ConsecutiveDelimiter: object = MISSING,
                 Tab: object = MISSING, Semicolon: object = MISSING, Comma: object = MISSING,
                 Space: object = MISSING, Other: object = MISSING, OtherChar: object = MISSING,
                 FieldInfo: object = MISSING, TextVisualLayout: object = MISSING, DecimalSeparator: object = MISSING,
                 ThousandsSeparator: object = MISSING, TrailingMinusNumbers: object = MISSING,
                 Local: object = MISSING) -> object:
        """A text file read into a new workbook with the delimiters, start row, field kinds and code page given,
        each one left out the session's as OpenText last read (see _text_files). TextVisualLayout and Local change
        nothing on the model's US English system (tests/fixtures/open_text.json)."""
        from pyopenvba.apps.excel._text_files import open_text_with

        if Filename is MISSING:
            raise error(449)
        given = to_text(Filename)
        arguments = (Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space,
                     Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator,
                     TrailingMinusNumbers, Local)

        def read(path: Path) -> Workbook:
            return open_text_with(
                self.application, path, origin=Origin, start_row=StartRow, data_type=DataType,
                qualifier=TextQualifier, consecutive=ConsecutiveDelimiter, tab=Tab, semicolon=Semicolon,
                comma=Comma, space=Space, other=Other, other_char=OtherChar, field_info=FieldInfo,
                decimal=DecimalSeparator, thousands=ThousandsSeparator, trailing_minus=TrailingMinusNumbers,
                given=any(argument is not MISSING for argument in arguments))

        if not Path(given).is_file():
            raise error(1004, f"Sorry, we couldn't find {given}. Is it possible it was moved, renamed or deleted?")
        self.open_file(Path(given), reader=read)
        return EMPTY

    def open_file(self, path: Path, *, with_vba: bool = True, read_only: bool = False,
                  format_code: int | None = None, reader: Callable[[Path], Workbook] | None = None) -> Workbook:
        """The workbook at ``path``, opened, or a CSV or text file read into one; ``format_code`` is the Format
        Workbooks.Open was given, and ``reader`` reads the file in place of Open's way (see _text_files)."""
        from pyopenvba.apps.excel._io import load_workbook
        from pyopenvba.apps.excel._projects import attach_project, load_project
        from pyopenvba.apps.excel._text_files import is_text, open_text

        path = path.resolve()
        existing = next((book for book in self.books if book.path and (Path(book.path) / book.name).resolve() == path), None)
        if existing is not None:
            self.application.activate_book(existing)
            return existing
        if any(book.name.casefold() == path.name.casefold() for book in self.books):
            raise error(1004, "A workbook with that name is already open")
        text = reader is not None or is_text(path)
        if reader is not None:
            book = reader(path)
        else:
            book = open_text(self.application, path, format_code) if text else load_workbook(self.application, path)
        book.read_only = read_only
        self.books.append(book)
        try:
            if with_vba and not text:
                load_project(book)
            else:
                attach_project(book)
        except Exception:
            self.books.remove(book)
            from pyopenvba.apps.excel._projects import ExcelInterpreter

            root = self.application.interpreter
            if isinstance(root, ExcelInterpreter) and root.bridge.workbook is book:
                root.bridge.workbook = None
            raise
        before = self.application.active_book
        self.application.activate_book(book, events=False)
        _events.opened(book)
        if self.application.active_book is book:
            _events.book_activated(book, before)
        return book

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application

    def remove(self, book: Workbook) -> None:
        from pyopenvba.apps.excel._windows import front_to_back

        if book in self.books:
            self.books.remove(book)
        if book in self.application.window_order:
            self.application.window_order.remove(book)
        if self.application.active_book is book:
            # The window behind comes to the front, measured.
            self.application.active_book = front_to_back(self.application)[0] if self.books else None


def _session_property(name: str) -> tuple[Callable[..., object], Callable[..., None]]:
    """The getter and setter of one Workbook property kept for the session (see _workbook_settings.SESSION)."""

    def read(self: Workbook) -> object:
        from pyopenvba.apps.excel._workbook_settings import session_value

        return session_value(self, name)

    def write(self: Workbook, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import set_session_value

        set_session_value(self, name, value)

    read.__name__ = name
    return member(read), setter(name)(write)


class Workbook(ExcelObject):
    """One workbook: its sheets, its names, and where it came from."""

    vba_type_name = "Workbook"
    #: The events the model raises in the workbook's module and sinks (see _events).
    raised_events: ClassVar[frozenset[str]] = frozenset({
        "Open", "BeforeSave", "AfterSave", "BeforeClose", "Activate", "Deactivate", "NewSheet",
        "SheetChange", "SheetCalculate", "SheetSelectionChange", "SheetActivate", "SheetDeactivate",
        "SheetPivotTableUpdate"})

    def __init__(self, application: Application, name: str, path: str = "") -> None:
        self.application = application
        self.name = name
        self.path = path
        self.saved = True
        #: Opened with Workbooks.Open's ReadOnly, until a SaveAs gives the workbook a file of its own.
        self.read_only = False
        self.epoch_1904 = False
        self.date_system_changed = False
        #: workbook.xml's workbookPr attributes as the file held them, and those a macro changed since
        #: (_workbook_settings), which a save writes.
        self.workbook_pr: dict[str, str] = {}
        self.workbook_pr_changed: set[str] = set()
        #: The workbook properties a macro set for the session (_workbook_settings.SESSION).
        self.session_properties: dict[str, object] = {}
        self.project_runtime: Interpreter | None = None
        #: A validated SaveAs changes FileFormat even if the subsequent file write fails.
        self.format_code: int | None = None
        #: The folder the workbook's links were read from or last written for, which a DOS path a link names is
        #: saved relative to (see _hyperlinks_file).
        self.links_folder = ""
        #: The other workbooks the file links to, by the number its formulas give each, with the cells it keeps of
        #: them (see _external_links).
        self.linked_books: dict[int, LinkedBook] = {}
        #: Whether the file has links the model leaves out, to OLE or DDE sources.
        self.foreign_links = False
        from pyopenvba.apps.excel._workbook_links import LinkChanges

        #: The links BreakLink broke, which LinkInfo and BreakLink know by name until the workbook closes, and what
        #: the next save writes differently of the links (see _workbook_links).
        self.broken_links: list[LinkedBook] = []
        self.link_changes = LinkChanges()
        #: Workbook.UpdateLinks, xlUpdateLinksUserSetting unless the file says otherwise, and SaveLinkValues.
        self.update_links = 1
        self.save_link_values = True
        #: Workbook.CheckCompatibility, False unless the file says otherwise (see _workbook_options).
        self.check_compatibility = False
        #: DisplayDrawingObjects, DisplayInkComments, RemovePersonalInformation, ReadOnlyRecommended and
        #: EnableAutoRecover, as the file has them (see _workbook_options).
        self.drawing_objects = -4104
        self.ink_comments = True
        self.remove_personal = False
        self.read_only_recommended = False
        self.auto_recover = True
        #: The name the workbook's own module goes by in VBA, as ThisWorkbook.CodeName gives it.
        self.code_name = ""
        self.sheet_name_counter = 1
        #: The number the next sheet made takes (see Worksheet.sheet_id): one past the highest a file gave, and never
        #: taken back when a sheet goes (tests/fixtures/calculate_order.json).
        self.next_sheet_id = 1
        self.used_code_names: set[str] = set()
        self.deleted_document_names: set[str] = set()
        self.pending_document_sources: dict[str, str] = {}
        self.vbide_project: VBAObject | None = None
        self.sheets_: list[Worksheet] = []
        self.pivot_caches: list[PivotCache] = []
        self.names_ = Names(self)
        self.queries_ = Queries(self)
        #: How the workbook first spelled each name its formulas use but it does not define, by lower case;
        #: None until a formula is written, when the formulas it has teach it.
        self.name_spellings: dict[str, str] | None = None
        self.active_sheet_index = 0
        #: The sheets selected with the active one, in tab order, while they are grouped; empty while the active
        #: sheet stands alone (see _sheet_groups).
        self.selected_sheets: list[Worksheet] = []
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
        from pyopenvba.apps.excel._windows import BookViewState

        #: The workbook's window: its tabs, scroll bars and caption (see _windows).
        self.view = BookViewState()
        #: The workbook's document properties, read from its file or a new workbook's once something asks for them
        #: (see _document_properties).
        self.properties: Properties | None = None
        #: When and by whom the workbook was made, a new workbook's creation date and author.
        interpreter = application.interpreter
        self.made_at = (interpreter.clock() if interpreter is not None else _dt.datetime.now()).replace(microsecond=0)
        self.made_by = application.user_name
        from pyopenvba.apps.excel._calculation_state import VERSION

        #: The calculation engine that last worked the workbook out in full, as its calcId has it; whether its
        #: ForceFullCalculation is on; and whether it was opened wanting a full calculation it has not had yet
        #: (see _calculation_state).
        self.calculation_version = VERSION
        self.force_full_calculation = False
        self.full_pending = False
        #: True once a sheet's custom properties change, so a save writes them all (see _custom_properties).
        self.custom_properties_changed = False
        #: Workbook.PrecisionAsDisplayed, calcPr fullPrecision="0" (see _precision).
        self.precision_as_displayed = False

    @property
    def calculator(self) -> Any:
        """The workbook's formulas, and what each one comes to."""
        if self._calculator is None:
            from pyopenvba.apps.excel._calc import Calculator

            self._calculator = Calculator(self)
        return self._calculator

    @property
    def calculation_in_progress(self) -> bool:
        """Inspect active evaluation without allocating the lazy calculator."""
        found = self._calculator
        return found is not None and bool(found._working or found._calculation_depth or found._learning_depth)

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

    # --- settings (_workbook_settings) ----------------------------------------------------------------------

    @member
    def DefaultTableStyle(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import default_style

        return default_style(self, "table")

    @setter("DefaultTableStyle")
    def _set_default_table_style(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import set_default_style

        set_default_style(self, "table", value)

    @member
    def DefaultPivotTableStyle(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import default_style

        return default_style(self, "pivot")

    @setter("DefaultPivotTableStyle")
    def _set_default_pivot_table_style(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import set_default_style

        set_default_style(self, "pivot", value)

    @member
    def DefaultSlicerStyle(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import TableStyle

        return TableStyle(self, "SlicerStyleLight1")

    @setter("DefaultSlicerStyle")
    def _set_default_slicer_style(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import SLICER_STYLES, set_slicer_default

        set_slicer_default(value, SLICER_STYLES, "DefaultSlicerStyle")

    @member
    def DefaultTimelineStyle(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import TableStyle

        return TableStyle(self, "TimeSlicerStyleLight1")

    @setter("DefaultTimelineStyle")
    def _set_default_timeline_style(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import TIMELINE_STYLES, set_slicer_default

        set_slicer_default(value, TIMELINE_STYLES, "DefaultTimelineStyle")

    @member
    def TableStyles(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._workbook_settings import table_style

        if Index is MISSING:
            raise VBAUnsupportedError("Workbook.TableStyles as a collection is not implemented")
        return table_style(self, Index)

    @member
    def ShowPivotTableFieldList(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import flag

        return flag(self, "ShowPivotTableFieldList")

    @setter("ShowPivotTableFieldList")
    def _set_show_pivot_table_field_list(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import set_flag

        set_flag(self, "ShowPivotTableFieldList", value)

    @member
    def ShowPivotChartActiveFields(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import flag

        return flag(self, "ShowPivotChartActiveFields")

    @setter("ShowPivotChartActiveFields")
    def _set_show_pivot_chart_active_fields(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import set_flag

        set_flag(self, "ShowPivotChartActiveFields", value)

    @member
    def InactiveListBorderVisible(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import flag

        return flag(self, "InactiveListBorderVisible")

    @setter("InactiveListBorderVisible")
    def _set_inactive_list_border_visible(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import set_flag

        set_flag(self, "InactiveListBorderVisible", value)

    @member
    def CreateBackup(self) -> object:
        return False

    @setter("CreateBackup")
    def _set_create_backup(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import refused

        refused("CreateBackup", 1004)

    @member
    def HasPassword(self) -> object:
        return False

    @setter("HasPassword")
    def _set_has_password(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import refused

        refused("HasPassword", 1004)

    @member
    def CaseSensitive(self) -> object:
        return False

    @setter("CaseSensitive")
    def _set_case_sensitive(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import refused

        refused("CaseSensitive", 450)

    @member
    def UseWholeCellCriteria(self) -> object:
        return True

    @setter("UseWholeCellCriteria")
    def _set_use_whole_cell_criteria(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import refused

        refused("UseWholeCellCriteria", 450)

    @member
    def UseWildcards(self) -> object:
        return True

    @setter("UseWildcards")
    def _set_use_wildcards(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import refused

        refused("UseWildcards", 450)

    @member
    def Final(self) -> object:
        return False

    @member
    def UpdateRemoteReferences(self) -> object:
        return True

    @setter("UpdateRemoteReferences")
    def _set_update_remote_references(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import unmeasured_set

        unmeasured_set("UpdateRemoteReferences", value, 1004)

    @member
    def ChartDataPointTrack(self) -> object:
        return True

    @setter("ChartDataPointTrack")
    def _set_chart_data_point_track(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_settings import unmeasured_set

        unmeasured_set("ChartDataPointTrack", value, 13)

    AutoSaveOn, _set_auto_save_on = _session_property("AutoSaveOn")
    AutoUpdateFrequency, _set_auto_update_frequency = _session_property("AutoUpdateFrequency")
    ChangeHistoryDuration, _set_change_history_duration = _session_property("ChangeHistoryDuration")
    CompatibilityVersion, _set_compatibility_version = _session_property("CompatibilityVersion")
    ConflictResolution, _set_conflict_resolution = _session_property("ConflictResolution")
    ConnectionsDisabled, _set_connections_disabled = _session_property("ConnectionsDisabled")
    DoNotPromptForConvert, _set_do_not_prompt_for_convert = _session_property("DoNotPromptForConvert")
    EncryptionProvider, _set_encryption_provider = _session_property("EncryptionProvider")
    EnvelopeVisible, _set_envelope_visible = _session_property("EnvelopeVisible")
    Excel8CompatibilityMode, _set_excel8_compatibility_mode = _session_property("Excel8CompatibilityMode")
    HighlightChangesOnScreen, _set_highlight_changes_on_screen = _session_property("HighlightChangesOnScreen")
    IsAddin, _set_is_addin = _session_property("IsAddin")
    IsInplace, _set_is_inplace = _session_property("IsInplace")
    KeepChangeHistory, _set_keep_change_history = _session_property("KeepChangeHistory")
    ListChangesOnNewSheet, _set_list_changes_on_new_sheet = _session_property("ListChangesOnNewSheet")
    MultiUserEditing, _set_multi_user_editing = _session_property("MultiUserEditing")
    PersonalViewListSettings, _set_personal_view_list_settings = _session_property("PersonalViewListSettings")
    PersonalViewPrintSettings, _set_personal_view_print_settings = _session_property("PersonalViewPrintSettings")
    RevisionNumber, _set_revision_number = _session_property("RevisionNumber")
    ShowConflictHistory, _set_show_conflict_history = _session_property("ShowConflictHistory")
    TemplateRemoveExtData, _set_template_remove_ext_data = _session_property("TemplateRemoveExtData")
    VBASigned, _set_vba_signed = _session_property("VBASigned")
    WriteReserved, _set_write_reserved = _session_property("WriteReserved")

    @member
    def AutoUpdateSaveChanges(self) -> object:
        """Error 1004 on a workbook that is not shared (tests/fixtures/workbook_properties.json)."""
        raise error(1004)

    @setter("AutoUpdateSaveChanges")
    def _set_auto_update_save_changes(self, value: object) -> None:
        raise error(1004)

    @member
    def AccuracyVersion(self) -> object:
        from pyopenvba.apps.excel._workbook_settings import ACCURACY_VERSION

        return ACCURACY_VERSION

    @setter("AccuracyVersion")
    def _set_accuracy_version(self, value: object) -> None:
        raise VBAUnsupportedError("Workbook.AccuracyVersion set is not implemented")

    @method
    def DeleteNumberFormat(self, NumberFormat: object = MISSING) -> object:
        """Take a custom number format out of the workbook: every cell, row, column and cell style using it shows
        General, keeping the rest of its format. A built-in format, or "", is error 1004, and so is a protected
        sheet; a format the workbook does not hold, as NumberFormat spells it, is passed over
        (tests/fixtures/delete_number_format.json)."""
        from dataclasses import replace

        from pyopenvba.apps.excel._number_format import BUILTIN, normalized

        failed = "Method 'DeleteNumberFormat' of object '_Workbook' failed"
        if NumberFormat is MISSING:
            raise error(449)
        code = normalized(to_text(NumberFormat), strict=False) if to_text(NumberFormat) else ""
        if not code or code in BUILTIN.values():
            raise error(1004, failed)
        protected = [sheet for sheet in self.sheets_ if getattr(sheet, "protection", None) is not None]
        if protected:
            if protected != [self.active_sheet]:
                raise VBAUnsupportedError("Workbook.DeleteNumberFormat with a protected sheet other than the "
                                          "active one is not implemented")
            raise error(1004, failed)
        from pyopenvba.apps.excel._cell_styles import restyle_book

        stylesheet = self.stylesheet
        restyle_book(self, lambda style: replace(style, number_format="General") if style.number_format == code
                     else None)
        for entry in stylesheet.cell_styles:
            if entry.format.number_format == code:
                entry.format = replace(entry.format, number_format="General")
                stylesheet.restyle_cell_style(entry)
        self.saved = False
        return EMPTY

    @member
    def Colors(self, Index: object = MISSING) -> object:
        """The palette ColorIndex counts in, or one colour of it (see _palette)."""
        from pyopenvba.apps.excel._palette import colors

        return colors(self, Index)

    @setter("Colors")
    def _set_colors(self, Index: object, value: object = MISSING) -> None:
        from pyopenvba.apps.excel._palette import set_colors

        if value is MISSING:
            # Colors = array: the value comes alone.
            Index, value = MISSING, Index
        set_colors(self, Index, value)

    @method
    def ResetColors(self) -> object:
        from pyopenvba.apps.excel._palette import reset_colors

        reset_colors(self)
        return EMPTY

    @member
    def Date1904(self) -> object:
        return self.epoch_1904

    @setter("Date1904")
    def set_Date1904(self, value: object) -> None:
        from pyopenvba.apps.excel._date_system import set_mode

        set_mode(self, value)

    @member
    def CodeName(self) -> object:
        return self.code_name

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
    def BuiltinDocumentProperties(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._document_properties import DocumentProperties

        found = DocumentProperties(self, builtin=True)
        return found if Index is MISSING else found.Item(Index)

    @member
    def CustomDocumentProperties(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._document_properties import DocumentProperties

        found = DocumentProperties(self, builtin=False)
        return found if Index is MISSING else found.Item(Index)

    @member
    def PivotCaches(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivots import PivotCaches

        caches = PivotCaches(self, list(self.pivot_caches))
        return caches if Index is MISSING else caches.Item(Index)

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
    def Windows(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._windows import Windows

        windows = Windows(self.application, self)
        return windows if Index is MISSING else windows.vba_get("Item", [Index])

    @member
    def Names(self, Index: object = MISSING) -> object:
        return self.names_ if Index is MISSING else self.names_.vba_get("Item", [Index])

    @member
    def Connections(self) -> object:
        """The data connections the file keeps (see _connections)."""
        from pyopenvba.apps.excel._connections import Connections

        return Connections(self)

    @member
    def Queries(self, Index: object = MISSING) -> object:
        return self.queries_ if Index is MISSING else self.queries_.vba_get("Item", [Index])

    @member
    def Styles(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._cell_styles import Styles

        styles = Styles(self)
        return styles if Index is MISSING else styles.vba_get("Item", [Index])

    @member
    def IconSets(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._icon_sets import IconSets

        sets = IconSets(self)
        return sets if Index is MISSING else sets.vba_get("Item", [Index])

    @member
    def VBProject(self) -> object:
        from pyopenvba.apps.excel._vbide import project_for

        return project_for(self)

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
        saving = SaveChanges is not MISSING and to_bool(SaveChanges) and not self.saved
        # Filename asks Excel to save first, before it raises BeforeClose.
        if saving and Filename is not MISSING:
            if not self.save_as(Filename):
                return EMPTY
        if not _events.before_close(self):
            return EMPTY
        if SaveChanges is MISSING and not self.saved:
            raise VBAUnsupportedError("Closing a changed workbook requires a save prompt; specify SaveChanges")
        if saving and Filename is MISSING:
            if not self._save():
                return EMPTY
        self.application.workbooks_.remove(self)
        return EMPTY

    @method
    def Save(self) -> object:
        self._save()
        return EMPTY

    def _save(self) -> bool:
        if not self.path:
            if not _events.before_save(self, save_as_ui=True):
                _events.after_save(self, False)
                return False
            raise VBAUnsupportedError("Saving an unnamed workbook requires a save dialog; use SaveAs with a filename")
        from pyopenvba.apps.excel._text_files import text_writer

        if self.read_only:
            # Excel, its alerts off, saves a copy into its default file path and keeps the file as it was.
            raise VBAUnsupportedError("Save of a workbook opened read-only is not implemented; use SaveAs")
        if self.format_code is not None:
            # A text format the model does not write reports itself before any event.
            text_writer(self.format_code)
        return self.save_to(Path(self.path) / self.name)

    @method
    def SaveAs(self, Filename: object = MISSING, FileFormat: object = MISSING, Password: object = MISSING,
               WriteResPassword: object = MISSING, ReadOnlyRecommended: object = MISSING,
               CreateBackup: object = MISSING, AccessMode: object = MISSING, ConflictResolution: object = MISSING,
               AddToMru: object = MISSING, TextCodepage: object = MISSING, TextVisualLayout: object = MISSING,
               Local: object = MISSING, WorkIdentity: object = MISSING) -> object:
        """SaveAs in a workbook format or a text one (see _text_files). ConflictResolution, AddToMru,
        TextCodepage, TextVisualLayout and Local change nothing in a file the model writes: it shares no workbook,
        and Local's list separator is a comma on its US English system (tests/fixtures/text_saves.json)."""
        refuse_save_options(Password, WriteResPassword, ReadOnlyRecommended, CreateBackup)
        if AccessMode is not MISSING and int(to_integer(AccessMode, "Long")) == 2:
            raise VBAUnsupportedError("SaveAs with AccessMode xlShared is not implemented")
        if WorkIdentity is not MISSING:
            raise VBAUnsupportedError("SaveAs with a WorkIdentity is not implemented")
        self.save_as(Filename, FileFormat)
        return EMPTY

    def save_as(self, Filename: object, FileFormat: object = MISSING, sheet: Worksheet | None = None) -> bool:
        """SaveAs; ``sheet`` is the sheet a text format writes, the active one where it is None."""
        from pyopenvba.apps.excel._text_files import text_writer

        if Filename is MISSING:
            if not _events.before_save(self):
                _events.after_save(self, False)
                return False
            raise VBAUnsupportedError("SaveAs without a filename is not implemented; supply Filename")
        target = Path(to_text(Filename))
        chosen = int(to_integer(FileFormat if FileFormat is not MISSING else self.FileFormat(), "Long"))
        writer = text_writer(chosen)
        expected = writer.extension if writer is not None else {51: ".xlsx", 52: ".xlsm"}.get(chosen)
        if expected is None:
            raise VBAUnsupportedError(f"SaveAs FileFormat {chosen} is not implemented by the Excel cell model")
        if not target.suffix:
            target = target.with_suffix(expected)
        if self.read_only and self.path and target.resolve() == (Path(self.path) / self.name).resolve():
            raise error(1004, "Cannot save as that name. Document was opened as read-only.")
        if not _events.before_save(self):
            _events.after_save(self, False)
            return False
        # A text format takes any extension (tests/fixtures/text_saves.json, txt_extension).
        if writer is None and target.suffix.lower() != expected:
            _events.after_save(self, False)
            raise error(1004, "SaveAs filename does not match FileFormat")
        return self.save_to(target, before=False, file_format=chosen, sheet=sheet)

    @method
    def SaveCopyAs(self, Filename: object = MISSING) -> object:
        """The file a save would write, written to Filename; the workbook keeps its name and Saved flag, and no save
        event is raised (see _save_copy)."""
        from pyopenvba.apps.excel._save_copy import save_copy
        from pyopenvba.apps.excel._text_files import TEXT_FORMATS

        if self.format_code in TEXT_FORMATS:
            raise VBAUnsupportedError("SaveCopyAs of a workbook in a CSV or text format is not implemented")
        save_copy(self, Filename)
        return EMPTY

    def save_to(self, target: Path, *, before: bool = True, file_format: int | None = None, create_parents: bool = False,
                sheet: Worksheet | None = None) -> bool:
        from pyopenvba.apps.excel._io import prepare_format, save_workbook
        from pyopenvba.apps.excel._text_files import save_text, saved_as_text, text_writer

        if before and not _events.before_save(self):
            _events.after_save(self, False)
            return False
        # A text format writes one sheet as text, ``sheet`` or the active one; a SaveAs in one names it after the
        # file, Save keeps its name.
        written = file_format if file_format is not None else self.format_code
        text = written is not None and text_writer(written) is not None
        if file_format is not None and not text:
            prepare_format(self, file_format)
        if not create_parents and not target.parent.is_dir():
            raise error(1004, "the save directory does not exist")
        try:
            if text:
                assert written is not None
                save_text(self, target, written, sheet)
            else:
                save_workbook(self, target)
        except OSError as failure:
            raise error(1004, str(failure)) from failure
        self.path = str(target.resolve().parent)
        self.name = target.name
        if text and file_format is not None:
            saved_as_text(self, file_format, sheet)
        # A text file holds one sheet, which leaves a workbook of more to save (tests/fixtures/text_saves.json).
        self.saved = not text or len(self.sheets_) == 1
        # A file of its own is the workbook's to change (tests/fixtures/text_files.json, readonly_saveas_other).
        self.read_only = False
        _events.after_save(self, True)
        return True

    @member
    def FileFormat(self) -> object:
        from pyopenvba.apps.excel._io import workbook_format

        return VBAInt(workbook_format(self), "Long")

    @member
    def ReadOnly(self) -> object:
        """Whether Workbooks.Open opened the workbook read-only, until a SaveAs gives it a file of its own."""
        return self.read_only

    # -- links to other workbooks (see _workbook_links)

    @method
    def LinkSources(self, Type: object = MISSING) -> object:
        from pyopenvba.apps.excel._workbook_links import sources

        return sources(self, Type)

    @method
    def LinkInfo(self, Name: object = MISSING, LinkInfo: object = MISSING, Type: object = MISSING,
                 EditionRef: object = MISSING) -> object:
        from pyopenvba.apps.excel._workbook_links import info

        return info(self, Name, LinkInfo, Type, EditionRef)

    @method
    def BreakLink(self, Name: object = MISSING, Type: object = MISSING) -> object:
        from pyopenvba.apps.excel._workbook_links import break_link

        break_link(self, Name, Type)
        return EMPTY

    @member
    def UpdateLinks(self) -> object:
        return VBAInt(self.update_links, "Long")

    @setter("UpdateLinks")
    def _set_update_links(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_links import set_update_links

        set_update_links(self, value)

    @member
    def SaveLinkValues(self) -> object:
        return self.save_link_values

    @setter("SaveLinkValues")
    def _set_save_link_values(self, value: object) -> None:
        from pyopenvba.apps.excel._workbook_links import set_save_link_values

        set_save_link_values(self, value)

    @member
    def CheckCompatibility(self) -> object:
        return self.check_compatibility

    @setter("CheckCompatibility")
    def _set_check_compatibility(self, value: object) -> None:
        """A Boolean as VBA passes it, saved in the workbookPr; the workbook is no longer saved
        (tests/fixtures/small_members.json)."""
        self.check_compatibility = to_bool(value)
        self.saved = False

    # -- settings the workbook part keeps (see _workbook_options, tests/fixtures/workbook_options.json)

    @member
    def DisplayDrawingObjects(self) -> object:
        return VBAInt(self.drawing_objects, "Long")

    @setter("DisplayDrawingObjects")
    def _set_drawing_objects(self, value: object) -> None:
        """A Long as VBA passes it; the workbook is no longer saved."""
        self.drawing_objects = int(to_integer(value, "Long"))
        self.saved = False

    @member
    def DisplayInkComments(self) -> object:
        return self.ink_comments

    @setter("DisplayInkComments")
    def _set_ink_comments(self, value: object) -> None:
        """A Boolean as VBA passes it; the workbook stays saved."""
        self.ink_comments = to_bool(value)

    @member
    def RemovePersonalInformation(self) -> object:
        return self.remove_personal

    @setter("RemovePersonalInformation")
    def _set_remove_personal(self, value: object) -> None:
        """A Boolean as VBA passes it; the workbook is no longer saved. The model's saves change nothing for it."""
        self.remove_personal = to_bool(value)
        self.saved = False

    @member
    def ReadOnlyRecommended(self) -> object:
        return self.read_only_recommended

    @setter("ReadOnlyRecommended")
    def _set_read_only_recommended(self, value: object) -> None:
        """A Boolean as VBA passes it; the workbook stays saved."""
        self.read_only_recommended = to_bool(value)

    @member
    def EnableAutoRecover(self) -> object:
        return self.auto_recover

    @setter("EnableAutoRecover")
    def _set_auto_recover(self, value: object) -> None:
        """A Boolean as VBA passes it; the workbook stays saved."""
        self.auto_recover = to_bool(value)

    @member
    def HasVBProject(self) -> object:
        """Whether the workbook's project holds code: a module, or a document module with a line of code; a new
        workbook's holds none (tests/fixtures/small_members.json)."""
        interpreter = self.project_runtime
        sources = [] if interpreter is None else [(runtime.parsed.kind, runtime.parsed.source)
                                                  for runtime in interpreter.modules.values()]
        sources += [("document", source) for source in self.pending_document_sources.values()]
        return any(kind != "document" or any(line.strip() and not re.match(r"\s*(Attribute|Option)\b|\s*'", line,
                                                                              re.I) for line in source.splitlines())
                   for kind, source in sources)

    @method
    def Calculate(self) -> object:
        _events.calculated_all(self, self.calculator.calculate_all(asked=True))
        return EMPTY

    @member
    def CalculationVersion(self) -> object:
        """The version of the calculation engine that last worked the workbook out in full: Excel's own, or the
        calcId its file has where Excel keeps the file's values (see _calculation_state)."""
        return VBAInt(self.calculation_version, "Long")

    @member
    def ForceFullCalculation(self) -> object:
        return self.force_full_calculation

    @setter("ForceFullCalculation")
    def _set_force_full_calculation(self, value: object) -> None:
        """True holds every workbook to full calculation until Excel quits (see _calculation_state)."""
        from pyopenvba.apps.excel._calculation_state import set_force

        set_force(self, value)

    @member
    def PrecisionAsDisplayed(self) -> object:
        return self.precision_as_displayed

    @setter("PrecisionAsDisplayed")
    def _set_precision_as_displayed(self, value: object) -> None:
        """True rounds every number as its format shows it, now and from then on (see _precision)."""
        wanted = to_bool(value)
        if wanted and not self.precision_as_displayed:
            from pyopenvba.apps.excel._precision import turned_on

            self.precision_as_displayed = True
            turned_on(self)
        self.precision_as_displayed = wanted

    def vba_sink_attached(self) -> None:
        """A WithEvents variable now hears this workbook: formulas written while
        nobody listened are worked out first, as Excel worked them out when
        they were written, so the next edit's Calculate is only its own.
        Manual calculation retains its existing caches until requested."""
        if self.calculator.automatic:
            self.calculator.calculate_all(volatile=False)

    @method
    def RefreshAll(self) -> object:
        """Evaluate every query and land each one on its sheet, then refresh every pivot cache a table uses, from
        the last to the first, passing over one whose refresh is disabled, and calculate once after them all
        (tests/fixtures/pivot_shared_caches/ refresh_all_three, refresh_all_reordered, refresh_all_disabled;
        tests/fixtures/calculate_order.json refresh_all_once)."""
        from pyopenvba.apps.excel._pivot_refresh import refresh as refresh_cache
        from pyopenvba.apps.excel._refresh import refresh

        for entry in list(self.queries_.entries):
            refresh(self, entry.name)
        used = {id(table.cache) for sheet in self.sheets_ for table in sheet.pivot_tables}
        caches = [cache for cache in reversed(list(self.pivot_caches)) if id(cache) in used and cache.EnableRefresh()]
        for cache in caches:
            refresh_cache(cache, calculate=False)
        if caches and self.calculator.automatic:
            # The calculation that ends a refresh works out what the queries changed as well.
            _events.refresh_calculated(self)
        else:
            self.calculator.calculate_all()
        return EMPTY

    # -- Python side

    @property
    def active_sheet(self) -> Worksheet | None:
        if not self.sheets_:
            return None
        return self.sheets_[min(self.active_sheet_index, len(self.sheets_) - 1)]

    def activate_sheet(self, sheet: Worksheet) -> None:
        from pyopenvba.apps.excel._table_filters import refocus

        self.active_sheet_index = self.sheets_.index(sheet)
        # A sheet of the group stays in it as it becomes active; any other stands alone
        # (tests/fixtures/sheet_groups.json, activate_in_group and activate_outside).
        if not any(one is sheet for one in self.selected_sheets):
            self.selected_sheets = []
        refocus(sheet)

    def selected(self) -> list[Worksheet]:
        """The sheets the workbook's window has selected, in tab order: the group the active sheet is in, or the
        active sheet alone."""
        active = self.active_sheet
        if active is None:
            return []
        if not any(one is active for one in self.selected_sheets):
            return [active]
        return [sheet for sheet in self.sheets_ if any(one is sheet for one in self.selected_sheets)]

    def add_sheet(self, name: str = "", *, at: int | None = None) -> Worksheet:
        """A new worksheet, named the way Excel names one when asked."""
        chosen = name or self._next_sheet_name()
        if any(_names.same_sheet_name(sheet.name, chosen) for sheet in self.sheets_):
            raise error(1004, f"a sheet called {chosen} is already there")
        sheet = Worksheet(self, chosen)
        if at is None:
            self.sheets_.append(sheet)
        else:
            self.sheets_.insert(at, sheet)
        self.saved = False
        self.sheet_name_counter = max(self.sheet_name_counter, len(self.sheets_))
        return sheet

    def _next_sheet_name(self) -> str:
        taken = {sheet.name.lower() for sheet in self.sheets_}
        number = self.sheet_name_counter + 1
        while f"sheet{number}" in taken:
            number += 1
        self.sheet_name_counter = number
        return f"Sheet{number}"

    def allocate_sheet_names(self, count: int) -> list[str]:
        return [self._next_sheet_name() for _ in range(count)]

    def sheet_named(self, name: str) -> Worksheet:
        """The sheet a name names where Excel keeps sheets apart, as a formula's sheet is found: 'q1AESOP'!A1 is
        #REF! beside a sheet q1Aesop with the ash (tests/fixtures/name_equality.json)."""
        for sheet in self.sheets_:
            if _names.same_sheet_name(sheet.name, name):
                return sheet
        raise error(ERR_SUBSCRIPT_OUT_OF_RANGE, f"there is no sheet called {name}")

    def sheet_found(self, name: str) -> Worksheet:
        """The sheet Worksheets(name) and Sheets(name) find: the first in tab order whose name word sort cannot tell
        from it, case aside, so AESOP finds Aesop with the ash ahead of a later aesop
        (tests/fixtures/name_equality.json)."""
        for sheet in self.sheets_:
            if _names.same_name(sheet.name, name):
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
            return self.book.sheet_found(index)
        if isinstance(index, VBAArray):
            from pyopenvba.apps.excel._sheet_groups import group

            return group(self.book, index)
        return super().vba_lookup(index, items)

    @method
    def FillAcrossSheets(self, Range: object = MISSING, Type: object = MISSING) -> object:
        """Every sheet of the workbook's, as a group's (see _sheet_groups)."""
        from pyopenvba.apps.excel._sheet_groups import fill_across

        fill_across(list(self.book.sheets_), Range, Type)
        return EMPTY

    @method
    def Add2(self, Before: object = MISSING, After: object = MISSING, Count: object = MISSING,
             NewLayout: object = MISSING) -> object:
        """Charts.Add2's; on a workbook's sheets, whatever is given, error 445 (tests/fixtures/sheets_add2.json)."""
        raise error(445, "Object doesn't support this action")

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
        if len(self.book.selected()) > 1:
            # Excel adds no sheet while sheets are grouped (tests/fixtures/sheet_groups.json, add_in_group).
            raise error(1004, "Add method of Sheets class failed")
        count = 1 if Count is MISSING else int(to_number(Count))
        if count <= 0:
            raise error(1004, "Count must be positive")
        if Type is not MISSING and to_integer(Type, "Long") != -4167:
            raise VBAUnsupportedError("Worksheets.Add supports worksheet Type only")
        at: int | None = None
        # Measured: the sheet the new one is placed beside is the one that deactivates, else the active one.
        left = self.book.active_sheet
        if isinstance(Before, Worksheet):
            at = self.book.sheets_.index(Before)
            left = Before
        elif isinstance(After, Worksheet):
            at = self.book.sheets_.index(After) + 1
            left = After
        elif self.book.active_sheet is not None:
            at = self.book.sheets_.index(self.book.active_sheet)
        names = self.book.allocate_sheet_names(count)
        added: list[Worksheet] = []
        for offset, name in enumerate(reversed(names)):
            sheet = self.book.add_sheet(name, at=None if at is None else at + offset)
            added.append(sheet)
            self.book.activate_sheet(sheet)
            _events.new_sheet(self.book, sheet, left)
        self.book.activate_sheet(added[0])
        return added[0]

    def _all(self) -> Any:
        """Every sheet as one group, which the collection's own Select, Copy, Move, Delete and Visible act on as
        Sheets(Array(...))'s do (tests/fixtures/sheet_groups.json, all_*)."""
        from pyopenvba.apps.excel._sheet_groups import SheetGroup

        return SheetGroup(self.book, list(self.book.sheets_))

    @method
    def Select(self, Replace: object = MISSING) -> object:
        return self._all().Select(Replace)

    @method
    def Copy(self, Before: object = MISSING, After: object = MISSING) -> object:
        return self._all().Copy(Before, After)

    @method
    def Move(self, Before: object = MISSING, After: object = MISSING) -> object:
        return self._all().Move(Before, After)

    @method
    def Delete(self) -> object:
        return self._all().Delete()

    @member
    def Visible(self) -> object:
        return self._all().Visible()

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        self._all().vba_set("Visible", value)

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Parent(self) -> object:
        return self.book


class Worksheet(ExcelObject):
    """One worksheet and its cells."""

    vba_type_name = "Worksheet"
    #: The events the model raises in the sheet's module and sinks (see _events).
    raised_events: ClassVar[frozenset[str]] = frozenset({
        "Change", "Calculate", "SelectionChange", "Activate", "Deactivate", "PivotTableUpdate"})
    #: True once Delete took the sheet out of its workbook.
    deleted = False

    def vba_member(self, name: str) -> MemberSpec | None:
        """A member of a deleted sheet fails with Excel's automation error, the member named; TypeName and Is still
        answer (tests/fixtures/deleted_sheet.json)."""
        if self.deleted:
            spec = super(ExcelObject, self).vba_member(name)
            raise error(-2147221080, f"Method '{spec.name if spec is not None else name}' of object '_Worksheet' "
                                     "failed")
        return super().vba_member(name)

    def __init__(self, book: Workbook, name: str) -> None:
        self.book = book
        self.name = name
        #: The number the sheet was made with among the workbook's, which a file keeps as its sheetId, whatever the
        #: tabs' order. Sheets calculate and raise Calculate in this order, and a loop of formulas is iterated in it
        #: (see _calc, _events.calculated_all).
        self.sheet_id = book.next_sheet_id
        book.next_sheet_id += 1
        #: The name the sheet goes by in VBA, Sheet1 in Sheet1.Range("A1"), and its own module's; "" for a sheet a
        #: macro added, which Excel leaves without one.
        self.code_name = ""
        self.cells_: dict[tuple[int, int], Cell] = {}
        #: The shared formulas: each one's block, which moves with inserts and deletes as a reference does,
        #: by the key its cells' Cell.shared holds (see _shared).
        self.shared_groups: dict[int, Area] = {}
        #: The array formulas: each one's block by its first cell, which holds the formula (see _arrays).
        self.array_formulas: dict[tuple[int, int], Area] = {}
        #: The what-if data tables Range.Table made, by the first of the cells each works out (see _data_tables).
        self.data_tables: dict[tuple[int, int], DataTable] = {}
        #: What each dynamic-array formula's answer wants to spill into, by its first cell (see _spills).
        self.spills: dict[tuple[int, int], Spill] = {}
        #: The sheet's tables, in the order its file lists them (see _tables).
        self.tables: list[Table] = []
        self.pivot_tables: list[PivotTable] = []
        self.outline_state: OutlineState | None = None
        self.page_setup_state: PageState | None = None
        #: Worksheet.DisplayPageBreaks, which is not saved (see _application_settings).
        self.display_page_breaks = False
        #: The settings a sheet keeps for the session, by member name, ScrollArea's text among them; whether its
        #: formulas are worked out; and its Transition settings, read when a macro first reaches them, and whether a
        #: macro changed them (see _worksheet_settings).
        self.settings: dict[str, object] = {}
        self.enable_calculation = True
        self.transitions: dict[str, bool] | None = None
        self.transitions_changed = False
        self.page_print_names_owned: set[str] = set()
        #: The sheet's data validation rules, in the order its file lists them (see _validation).
        self.validations: list[ValidationRule] = []
        #: True once a rule is added, changed or removed, so a save writes the rules again.
        self.validations_changed = False
        self.conditional_formats: list[ConditionalRule] = []
        #: The priority and dxf of each rule a macro took away, whose dxf Excel's saves keep where the rule's was
        #: though nothing names it, until the workbook is opened again (_conditional.keep_dxf, _dxf_table).
        self.kept_dxfs: list[tuple[int, int]] = []
        self.conditional_changed = False
        self.conditional_color_slots: dict[int, ConditionalRule] = {}
        self.conditional_color_groups: dict[int, tuple[list[ConditionalRule], dict[ConditionalRule, int]]] | None = None
        self.conditional_color_next = 0
        #: Cache against the exact OPC entry; replacing a part invalidates it.
        self.conditional_extensions: bool | None = None
        self.conditional_extension_entry: Entry | None = None
        #: True when the sheet's Excel 2010 extension holds conditional formats the model does not read
        #: (_conditional_extension), which then keeps that extension as the file had it.
        self.conditional_foreign = False
        self.merged_areas: list[Area] = []
        self.merges_dirty = False
        self.visible = -1  # xlSheetVisible
        #: Row heights, column widths and what is hidden.
        self.dims = _dimensions.SheetDimensions(self)
        from pyopenvba.apps.excel._windows import SheetViewState

        #: How the window shows the sheet: zoom, panes, where it is scrolled to (see _windows).
        self.view = SheetViewState()
        self.selection_range: Range | None = None
        #: The settings Worksheet.Sort holds, made the first time it is asked for.
        self.sort_state: SortState | None = None
        #: The sheet's sparkline groups, read from its part the first time they are reached (see _sparklines).
        self.sparklines: SparklineState | None = None
        #: The sheet's AutoFilter: its range and the criteria on its columns.
        self.auto_filter: SheetFilter | None = None
        #: FilterMode as a table's filter at the active cell left it, None while the sheet's own answers for it
        #: (see _table_filters).
        self.table_filter_mode: bool | None = None
        #: The tab's colour, None for none, and whether a macro changed it, so a save writes it (see _sheet_tabs).
        self.tab_color: Color | None = None
        self.tab_changed = False
        #: The sheet's manual page breaks, read from its file when first needed (see _page_breaks).
        self.page_breaks: SheetBreaks | None = None
        #: The sheet's custom properties, read from its file when first needed (see _custom_properties).
        self.custom_properties: list[CustomProperty] | None = None
        #: Whether the sheet's part holds ignoredErrors, None until asked; and whether a macro changed what the
        #: sheet ignores, so a save writes it (see _error_checks).
        self.ignored_from_file: bool | None = None
        self.ignored_changed = False
        #: The sheet's last Range.Consolidate, None before one (see _consolidate).
        self.consolidation: ConsolidationSettings | None = None
        #: The sheet's scenarios, read from its part when first needed, and whether a macro changed them, so a save
        #: writes them (see _scenarios).
        self.scenarios: Any = None
        self.scenarios_changed = False
        #: An AdvancedFilter has no dropdowns but retains its filtered list.
        self.advanced_filter: Area | None = None
        self.advanced_filter_changed = False
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
        self.chart_states: dict[int, ChartState] = {}
        self.deleted_chart_parts: set[str] = set()
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
        #: The notes on the sheet's cells, which VBA calls comments, by cell in the order they were made (see
        #: _notes); True once one is added, changed, moved or removed, so a save writes them again.
        self.notes: dict[tuple[int, int], Note] = {}
        self.notes_changed = False
        #: The sheet's hyperlinks in the order Worksheet.Hyperlinks lists them (see _hyperlinks); True once one is
        #: added, changed, moved or removed, so a save writes them again.
        self.hyperlinks: list[Link] = []
        self.links_changed = False
        #: The block of ids the sheet's VML part numbers its notes and form controls from, 0 until it has one, and
        #: the last id it gave (see _notes.next_shape_id).
        self.vml_block = 0
        self.vml_last_id = 0

    def touched(self) -> None:
        self.dirty = True
        self.book.saved = False

    def declares_revisions(self) -> bool:
        """Whether the sheet's part declares the revision namespace, where an autoFilter's xr:uid lives."""
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
        """One cell's contents changed: tell the calculator what to redo, and a table whose header it is."""
        self.touched()
        if self.tables:
            from pyopenvba.apps.excel._tables import edited

            edited(self, row, column)
        self.dims.fonts_changed(row)
        calculator = self.book.calculator
        cell = self.cells_.get((row, column))
        if cell is not None:
            calculator.remember(self.name, row, column, cell)
        else:
            calculator.forget(self.name, row, column)
        calculator.wrote(self.name, row, column)
        if cell is None or not cell.formula:
            from pyopenvba.apps.excel._controls import cell_changed

            cell_changed(self, row, column, cell.value if cell is not None else EMPTY)

    def shape_changed(self) -> None:
        """Rows or columns moved, so every formula has to be read again."""
        self.touched()
        self.book.calculator.rebuild()

    # -- consolidation (see _consolidate)

    @member
    def ConsolidationFunction(self) -> object:
        from pyopenvba.apps.excel._consolidate import settings

        return VBAInt(settings(self).function, "Long")

    @member
    def ConsolidationOptions(self) -> object:
        from pyopenvba.apps.excel._consolidate import options

        return options(self)

    @member
    def ConsolidationSources(self) -> object:
        from pyopenvba.apps.excel._consolidate import sources

        return sources(self)

    # -- identity

    @member
    def CodeName(self) -> object:
        return self.code_name

    @member
    def Name(self) -> object:
        return self.name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse_structure

        refuse_structure(self.book, APPLICATION_DEFINED)
        wanted = to_text(value)
        from pyopenvba.apps.excel._sparklines import refuse as refuse_sparklines

        refuse_sparklines(self, "Renaming")
        # Native probes: empty, past 31 characters, any of \ / ? * [ ] :, an
        # apostrophe first or last, and History in any case are refused.
        if (not wanted or len(wanted) > 31 or any(char in wanted for char in "\\/?*[]:")
                or wanted[0] == "'" or wanted[-1] == "'" or wanted.lower() == "history"):
            raise error(1004, "Method 'Name' of object '_Worksheet' failed")
        if any(sheet is not self and _names.same_sheet_name(sheet.name, wanted) for sheet in self.book.sheets_):
            raise error(1004, "That name is already taken. Try a different one.")
        from pyopenvba.apps.excel._worksheet_rename import rename

        rename(self, wanted)

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
        """Hide or show the sheet; hiding the last visible sheet is error 1004, and hiding the active one makes the
        next visible sheet after it active, or the last visible one before it (tests/fixtures/sheet_groups.json)."""
        from pyopenvba.apps.excel._sheet_groups import set_visible

        set_visible(self.book, [self], value, "Unable to set the Visible property of the Worksheet class")

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
        """This sheet's cells only: a Range, a qualified reference or a name on another sheet is error 1004
        (tests/fixtures/cross_sheet_ranges.json). Application.Range takes them (see resolved_range)."""
        return self.resolved_range(Cell1, Cell2, foreign=False)

    def resolved_range(self, Cell1: object, Cell2: object = MISSING, *, foreign: bool) -> Range:
        """Cell1, or the block from Cell1 to Cell2. A name or a qualified reference can point at another sheet,
        and so can a Range; where ``foreign``, as Application.Range has it, the Range then belongs to that sheet,
        but its two corners have to share one."""
        if Cell1 is MISSING:
            raise error(449)
        owner, first = self._located(Cell1)
        other, second = (owner, first) if Cell2 is MISSING else self._located(Cell2)
        if other is not owner or (owner is not self and not foreign):
            raise error(1004, "Method 'Range' of object '_Worksheet' failed")
        if Cell2 is MISSING:
            return Range(owner, first)
        top = min(first[0].top, second[0].top)
        left = min(first[0].left, second[0].left)
        bottom = max(first[-1].bottom, second[-1].bottom)
        right = max(first[-1].right, second[-1].right)
        return Range(owner, [Area(top, left, bottom, right, owner.name)])

    def _located(self, value: object) -> tuple[Worksheet, list[Area]]:
        """The sheet a Range argument's cells are on, and the cells."""
        if isinstance(value, Range):
            return value.sheet, list(value.areas)
        areas = self._as_area(value)
        return (self.book.sheet_named(areas[0].sheet) if areas[0].sheet else self), areas

    def _as_area(self, value: object) -> list[Area]:
        if isinstance(value, Range):
            return list(value.areas)
        text = to_text(value)
        named = self.book.names_.find(text, scope=self)
        if named is not None:
            return list(named.refers_to_range().areas)
        if any(sheet.tables for sheet in self.book.sheets_):
            from pyopenvba.apps.excel._tables import range_areas

            found = range_areas(self, text)
            if found is not None:
                return found
        if text.endswith("#"):
            return self._spill_areas(text[:-1])
        try:
            return parse_reference(text, sheet=self.name)
        except ValueError:
            raise error(1004, f"{text!r} is not a reference this sheet knows") from None

    def _spill_areas(self, text: str) -> list[Area]:
        """What spilled from the cell ``text`` names, Range("E1#"); error 1004 where nothing spilled from it
        (tests/fixtures/dynamic_arrays.json)."""
        areas = self._as_area(text)
        cell = areas[0]
        if len(areas) != 1 or cell.rows != 1 or cell.columns != 1:
            raise error(1004, f"{text}# is not a cell's spill")
        owner = self.book.sheet_named(cell.sheet) if cell.sheet else self
        assert owner is not None
        owner.book.calculator.value_of(owner.name, cell.top, cell.left)
        spill = owner.spills.get((cell.top, cell.left))
        if spill is None or spill.blocked:
            raise error(1004, f"nothing spilled from {text}")
        return [Area(spill.area.top, spill.area.left, spill.area.bottom, spill.area.right, owner.name)]

    @member
    def Shapes(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._shapes import Shapes as ShapesCollection

        shapes = ShapesCollection(self)
        return shapes if Index is MISSING else shapes.vba_get("Item", [Index])

    @member
    def ChartObjects(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._charts import ChartObjects

        charts = ChartObjects(self)
        return charts if Index is MISSING else charts.vba_get('Item', [Index])

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
            # Rows("3") is the whole row, the way Rows("3:5") is; several, Rows("1:2,5:5"), are a type mismatch
            # (tests/fixtures/range_lines.json).
            if "," in Index:
                raise error(13, "Type mismatch")
            wanted = Index if ":" in Index else f"{Index}:{Index}"
            return Range(self, parse_reference(wanted, sheet=self.name), whole="rows")
        row = int(to_integer(Index, "Long"))
        return Range(self, [Area(row, 1, row, MAX_COLUMNS, self.name)], whole="rows")

    @member
    def Columns(self, Index: object = MISSING) -> object:
        if Index is MISSING:
            return Range(self, [Area(1, 1, MAX_ROWS, MAX_COLUMNS, self.name)], whole="columns")
        if isinstance(Index, str):
            # Columns("C") is the whole column, the way Columns("C:D") is; several are a type mismatch, as Rows' are.
            if "," in Index:
                raise error(13, "Type mismatch")
            wanted = Index if ":" in Index else f"{Index}:{Index}"
            return Range(self, parse_reference(wanted, sheet=self.name), whole="columns")
        column = int(to_integer(Index, "Long"))
        return Range(self, [Area(1, column, MAX_ROWS, column, self.name)], whole="columns")

    @member
    def Comments(self, Index: object = MISSING) -> object:
        """The sheet's notes, row by row (see _notes)."""
        from pyopenvba.apps.excel._notes import Comments

        notes = Comments(self)
        return notes if Index is MISSING else notes.vba_get("Item", [Index])

    @member
    def Hyperlinks(self, Index: object = MISSING) -> object:
        """The sheet's hyperlinks (see _hyperlinks)."""
        from pyopenvba.apps.excel._hyperlinks import Hyperlinks

        links = Hyperlinks(self)
        return links if Index is MISSING else links.vba_get("Item", [Index])

    @member
    def Tab(self) -> object:
        """The sheet's tab and its colour (see _sheet_tabs)."""
        from pyopenvba.apps.excel._sheet_tabs import Tab

        return Tab(self)

    @member
    def UsedRange(self) -> object:
        bounds = self.used_bounds()
        if bounds is None:
            return Range(self, [Area(1, 1, 1, 1, self.name)])
        top, left, bottom, right = bounds
        return Range(self, [Area(top, left, bottom, right, self.name)])

    @member
    def CircularReference(self) -> object:
        """A cell of a loop of formulas on the sheet, the one whose formula closed the latest loop still there; Nothing
        for none, and with Application.Iteration on (tests/fixtures/iteration.json)."""
        found = self.book.calculator.circular_cell(self)
        return NOTHING if found is None else Range(self, [Area(found[0], found[1], found[0], found[1], self.name)])

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
    def ListObjects(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._tables import ListObjects

        tables = ListObjects(self)
        return tables if Index is MISSING else tables.vba_get("Item", [Index])

    @member
    def PivotTables(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivots import PivotTable, PivotTables

        tables = PivotTables(self, [PivotTable(self, table.cache, table.root, report_owner=table)
                                   for table in self.pivot_tables])
        return tables if Index is MISSING else tables.Item(Index)

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
        from pyopenvba.apps.excel._table_filters import filter_mode

        return filter_mode(self)

    @member
    def Outline(self) -> object:
        from pyopenvba.apps.excel._outline import Outline

        return Outline(self)

    @member
    def PageSetup(self) -> object:
        from pyopenvba.apps.excel._page_setup import PageSetup

        return PageSetup(self)

    @method
    def Scenarios(self, Index: object = MISSING) -> object:
        """The sheet's scenarios, or one of them (see _scenarios)."""
        from pyopenvba.apps.excel._scenarios import Scenarios, find

        return Scenarios(self) if Index is MISSING else find(self, Index)

    @member
    def CustomProperties(self) -> object:
        from pyopenvba.apps.excel._custom_properties import CustomProperties

        return CustomProperties(self)

    @member
    def HPageBreaks(self) -> object:
        from pyopenvba.apps.excel._page_breaks import HPageBreaks

        return HPageBreaks(self)

    @member
    def VPageBreaks(self) -> object:
        from pyopenvba.apps.excel._page_breaks import VPageBreaks

        return VPageBreaks(self)

    @method
    def ResetAllPageBreaks(self) -> object:
        from pyopenvba.apps.excel._page_breaks import reset

        reset(self)
        return EMPTY

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
        before = self.book.active_sheet
        self.book.activate_sheet(self)
        self.book.application.activate_book(self.book)
        _events.activated(self.book, before, self)
        return EMPTY

    @method
    def Select(self, Replace: object = MISSING) -> object:
        """The sheet active and alone; with Replace:=False the sheet joins the selection, the active sheet staying
        (tests/fixtures/sheet_groups.json, select_added)."""
        if Replace is not MISSING and not to_bool(Replace):
            from pyopenvba.apps.excel._sheet_groups import select

            select(self.book, [self], False)
            return EMPTY
        result = self.Activate()
        self.book.selected_sheets = []
        return result

    @method
    def Paste(self, Destination: object = MISSING, Link: object = MISSING) -> object:
        """What Copy or Cut put on the clipboard, at Destination or the selection."""
        from pyopenvba.apps.excel._clipboard import paste

        return paste(self, Destination, Link)

    @method
    def Delete(self) -> object:
        """Take the sheet out: the formulas and names that read it read #REF! instead, and where it was active the
        next visible sheet after it is, or the last visible one before it (tests/fixtures/sheet_groups.json)."""
        from pyopenvba.apps.excel._sparklines import refuse as refuse_sparklines
        from pyopenvba.apps.excel._sheet_groups import delete

        refuse_sparklines(self, "Deleting")
        delete(self.book, [self], "Cannot delete the only worksheet")
        return EMPTY

    @method
    def Copy(self, Before: object = MISSING, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import COPYING
        from pyopenvba.apps.excel._sheet_copy import copy_sheet

        self.guard_structure(Before, After, COPYING, "Copying")
        return copy_sheet(self, Before, After)

    @method
    def SaveAs(self, Filename: object = MISSING, FileFormat: object = MISSING, Password: object = MISSING,
               WriteResPassword: object = MISSING, ReadOnlyRecommended: object = MISSING,
               CreateBackup: object = MISSING, AddToMru: object = MISSING, TextCodepage: object = MISSING,
               TextVisualLayout: object = MISSING, Local: object = MISSING) -> object:
        """The workbook's SaveAs, but a text format writes this sheet, which takes the file's name, active or not
        (see _text_files, tests/fixtures/sheet_save_as.json)."""
        refuse_save_options(Password, WriteResPassword, ReadOnlyRecommended, CreateBackup)
        self.book.save_as(Filename, FileFormat, sheet=self)
        return EMPTY

    @method
    def Move(self, Before: object = MISSING, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import MOVING
        from pyopenvba.apps.excel._sheet_move import move_sheet

        self.guard_structure(Before, After, MOVING, "Moving")
        return move_sheet(self, Before, After)

    def guard_structure(self, before: object, after: object, message: str, what: str) -> None:
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
        """Calculate this sheet against foreign caches and deliver pending notifications; every formula of it while
        a workbook has held every workbook to full calculation (see _calculation_state)."""
        self.book.calculator.settle_automatic_cache()
        _events.calculated_all(self.book, self.book.calculator.calculate_all(
            self, asked=True, full=self.book.application.full_calculation))
        return EMPTY

    def vba_sink_attached(self) -> None:
        self.book.vba_sink_attached()

    @method
    def Evaluate(self, Name: object = MISSING) -> object:
        from pyopenvba.apps.excel._evaluate import evaluated
        from pyopenvba.apps.excel._udf import evaluation_origin

        with evaluation_origin(self.book.application, "evaluate"):
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

    @method
    def ClearArrows(self) -> object:
        """Takes away tracer arrows, which the model never draws; on a protected sheet error 1004
        (tests/fixtures/clear_arrows.json)."""
        if self.protection is not None and self.protection.contents:
            raise error(1004, "ClearArrows method of Worksheet class failed")
        return EMPTY

    @method
    def ClearCircles(self) -> object:
        """Takes away validation circles, which the model never draws, protected sheet or not
        (tests/fixtures/clear_arrows.json)."""
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
        if cell.value is not EMPTY or cell.formula or cell.html_id or cell.ignored_errors:
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
        self.book.stylesheet.meet(found.style)
        if not found.formula:
            # A constant given a format takes the precision it shows, where the workbook keeps numbers so.
            _precision.settle(self, row, column)
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
        for area in [*self.merged_areas, *self.array_formulas.values()]:
            live.extend(((area.top, area.left), (area.bottom, area.right)))
        # A note's cell counts, empty or not (tests/fixtures/excel_model/, notes.json).
        live.extend(self.notes)
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


#: The properties of a range that set its cells' format, as a protected sheet guards them.
_FORMAT_PROPERTIES = frozenset({"NumberFormat", "NumberFormatLocal", "HorizontalAlignment", "VerticalAlignment",
                                "WrapText", "ShrinkToFit", "IndentLevel", "AddIndent", "Orientation", "ReadingOrder",
                                "Locked", "FormulaHidden"})


class Range(ExcelObject):
    """A rectangle of cells, or several of them."""

    vba_type_name = "Range"

    def vba_information_value(self) -> object:
        if self.sheet not in self.sheet.book.sheets_ or self.sheet.book not in self.sheet.book.application.workbooks_.books:
            return self
        try:
            return self.vba_get("Value")
        except VBARuntimeError:
            # Native information functions retain object identity when a
            # deleted range can no longer supply its default value.
            return self

    def __init__(self, sheet: Worksheet, areas: list[Area], *, whole: str = "", lines: str = "") -> None:
        self.sheet = sheet
        self.areas = areas or [Area(1, 1, 1, 1, sheet.name)]
        #: "rows" or "columns" for a range made as whole rows or columns.
        #: Every row and every column is the same block of cells, but
        #: Rows("1:1048576").Hidden hides rows and Columns("A:XFD").Hidden
        #: columns, where Cells.Hidden is an error.
        self.whole = whole
        #: "rows" or "columns" for a range Rows or Columns made, whole or not: its Count, its items and For Each go by
        #: its first area's rows or columns, every area's for For Each, and AutoFit and Hidden take its rows or
        #: columns; its other members act on its cells (tests/fixtures/range_lines.json).
        self.lines = lines or whole

    def guard_set(self, member: str) -> None:
        from pyopenvba.apps.excel._protection import check_range_set

        check_range_set(self, member)

    # A macro reaching a cell leaves Excel the cell's number format, which the next Subtotal gives its first label
    # in General's place (tests/fixtures/subtotals/, register_*): any property of a range read or written, and
    # ClearContents, where its first cell is one the sheet holds; any format set, whatever the cell; and Style set
    # leaves General. Whole rows and columns, an empty cell and the other methods, Sort among them, leave it.

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        spec = self.vba_member(name)
        if spec is not None and (spec.kind == "property" or spec.name == "ClearContents"):
            self.note_seen()
        return super().vba_get(name, args, named)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        self.note_seen()
        super().vba_set(name, value, args, named, by_ref=by_ref)

    def after_set(self, member: str) -> None:
        if member in _FORMAT_PROPERTIES:
            self.note_formatted()
        elif member == "Style":
            self.sheet.book.application.recent_format = "General"

    def note_seen(self) -> None:
        """The range's first cell's number format as the one Subtotal's first label takes, where the sheet holds
        the cell."""
        area = self.first
        if area.whole_rows or area.whole_columns:
            return
        cell = self.sheet.cells_.get((area.top, area.left))
        if cell is not None:
            self.sheet.book.application.recent_format = cell.number_format

    def note_formatted(self) -> None:
        """A format set on the range: its first cell's number format after it, held or not, as the one."""
        area = self.first
        self.sheet.book.application.recent_format = self.sheet.style_at(area.top, area.left).number_format

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

    def vba_default_call(self, args: list[object], named: dict[str, object]) -> object:
        """``rng(i)`` and ``rng(r, c)``: Range's default member with arguments
        is Item, so Range("B2:C3")(5) is B4 (native probes)."""
        return self.vba_get("Item", args, named)

    @member(default=True)
    def Value(self, RangeValueDataType: object = MISSING) -> object:
        return self._values(raw=False)

    @setter("Value")
    def _set_value(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import writing

        _within_length(value)
        if not _arrays.write_admitted(self):
            return
        with writing(self.sheet):
            self._write(value)
        self._tables_follow()
        self._links_emptied(value)
        _events.after_edit(self)

    def _tables_follow(self) -> None:
        """A table around what was written follows it: it grows over a row or column written next to it, and a
        formula written into one of its columns can make the column calculated (see _table_edits)."""
        if self.sheet.tables:
            from pyopenvba.apps.excel import _table_edits

            _table_edits.calculated(self.sheet, list(self.areas))
            _table_edits.grow(self.sheet, list(self.areas))

    @member
    def Value2(self) -> object:
        """Value without the Date and Currency a cell's format makes of its number."""
        return self._values(raw=True)

    @setter("Value2")
    def _set_value2(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import writing

        _within_length(value)
        if not _arrays.write_admitted(self):
            return
        with writing(self.sheet):
            self._write(value, raw=True)
        self._tables_follow()
        self._links_emptied(value)
        _events.after_edit(self)

    def _values(self, *, raw: bool) -> object:
        if self.single:
            return self._read(self.first.top, self.first.left, raw=raw)
        # The array has the range's own shape, empty cells and all: Range("A1:B2").Value
        # of an empty sheet is 2 by 2 (native probes). Past a whole column's worth of
        # cells it is clipped to what the sheet uses, as walking it would mean nothing.
        area = self.first if self.first.rows * self.first.columns <= _WHOLE_COLUMN_CELLS else self.bounded()
        used = self.sheet.used_bounds()
        items: list[object] = []
        # VBA lays a two-dimensional array out column by column.
        for column in range(area.left, area.right + 1):
            for row in range(area.top, area.bottom + 1):
                if used is None or not (used[0] <= row <= used[2] and used[1] <= column <= used[3]):
                    items.append(EMPTY)
                else:
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
        # Read the current value (the cache in manual mode) before formatting;
        # read raw, since a Date too large to hold still shows.
        row, column = self.first.top, self.first.left
        application = self.sheet.book.application
        caller = application.udf_callers[-1] if application.udf_callers else None
        self_text = (application.udf_origins and application.udf_origins[-1] == "worksheet"
                     and caller is not None and caller.sheet is self.sheet
                     and caller.first.top == row and caller.first.left == column
                     and self.sheet.book.calculator.is_working(self.sheet.name, row, column))
        if not self_text:
            self._read(row, column, raw=True)
        cell = self.sheet.cell(row, column)
        if cell is None:
            return ""
        if self_text and cell.value is EMPTY:
            # A new formula displays a provisional zero. Later calls format its
            # actual cache without re-entering the caller's unfinished formula.
            cell = replace(cell, value=0.0)
        if self.sheet.conditional_formats:
            from pyopenvba.apps.excel._conditional_eval import conditional_number_format, value_hidden

            if value_hidden(self.sheet, row, column, cell.value):
                return ""
            style = self.sheet.style_at(self.first.top, self.first.left)
            cell = replace(cell, style=replace(style, number_format=conditional_number_format(
                self.sheet, self.first.top, self.first.left)))
        # What the cell shows, in what its column has room for (tests/fixtures/narrow_text.json).
        return _column_fit.narrowed(self.sheet, row, column, cell)

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
        _within_length(value)
        self._formula_written(value)

    def _formula_written(self, value: object, *, dynamic: bool = False) -> None:
        """What Formula writes; ``dynamic`` for a dynamic-array formula Formula2 writes, which spills (see
        _spills)."""
        from pyopenvba.apps.excel._protection import writing

        if not _arrays.write_admitted(self):
            return
        with writing(self.sheet):
            self._write_formula(value, dynamic=dynamic)
        self._tables_follow()
        self._links_emptied(value)
        _events.after_edit(self)

    def _write_formula(self, value: object, *, dynamic: bool = False) -> None:
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
                    # Excel reads the formula once and writes it out again, as written_formula does.
                    spelled = written_formula(self.sheet, value, anchor.top, anchor.left, at=True)
                    spelled = at_kept(spelled) if dynamic else legacy_formula(self.sheet, spelled)
                if self._put_formula(row, column, shift_text(spelled, row - anchor.top, column - anchor.left),
                                     dynamic=dynamic):
                    placed.append((row, column))
            else:
                self._type_into(row, column, value)
        if spelled is not None:
            self._bring_format(placed, spelled, anchor.top, anchor.left)
            if not dynamic:
                self._share(placed, spelled)

    @member
    def HasFormula(self) -> object:
        """True when every cell holds a formula, False when none does, and Null for a mix.

        Every cell of an array formula holds the array's.
        """
        area = self.first if self.single else self.bounded()
        arrays = bool(self.sheet.array_formulas)
        found = {bool(cell is not None and cell.formula)
                 or (arrays and _arrays.array_at(self.sheet, row, column) is not None)
                 for row in range(area.top, area.bottom + 1)
                 for column in range(area.left, area.right + 1) for cell in [self.sheet.cell(row, column)]}
        return NULL if len(found) > 1 else found.pop()

    @member
    def FormulaR1C1(self) -> object:
        return self._read_formulas(r1c1=True)

    @member
    def Formula2(self) -> object:
        return self._read_formulas(r1c1=False, at=True)

    @member
    def Formula2R1C1(self) -> object:
        return self._read_formulas(r1c1=True, at=True)

    @setter("Formula2")
    def _set_formula2(self, value: object) -> None:
        _within_length(value)
        written, dynamic = self._legacy(value, r1c1=False)
        self._formula_written(written, dynamic=dynamic)

    @setter("Formula2R1C1")
    def _set_formula2_r1c1(self, value: object) -> None:
        _within_length(value)
        written, dynamic = self._legacy(value, r1c1=True)
        self._formula_written(written, dynamic=dynamic)

    @member
    def HasSpill(self) -> object:
        """Whether the range meets what a dynamic-array formula spilled (see _spills)."""
        from pyopenvba.apps.excel._spills import has_spill

        return has_spill(self)

    @member
    def Precedents(self) -> object:
        """The cells on the sheet the range's formulas read, at every level, as Excel finds and orders them (see
        _precedents)."""
        from pyopenvba.apps.excel._precedents import answer

        return answer(self, dependents=False, direct=False)

    @member
    def DirectPrecedents(self) -> object:
        from pyopenvba.apps.excel._precedents import answer

        return answer(self, dependents=False, direct=True)

    @member
    def Dependents(self) -> object:
        """The formulas on the sheet reading the range, at every level, as Excel finds and orders them (see
        _precedents)."""
        from pyopenvba.apps.excel._precedents import answer

        return answer(self, dependents=True, direct=False)

    @member
    def DirectDependents(self) -> object:
        from pyopenvba.apps.excel._precedents import answer

        return answer(self, dependents=True, direct=True)

    @member
    def SpillParent(self) -> object:
        """The cell of the dynamic-array formula whose answer spilled over the range's first cell, or Nothing."""
        from pyopenvba.apps.excel._spills import spill_parent

        return spill_parent(self)

    @member
    def SpillingToRange(self) -> object:
        """The block the dynamic-array formula in the range's first cell spilled into, or Nothing."""
        from pyopenvba.apps.excel._spills import spilling_to

        return spilling_to(self)

    def _legacy(self, value: object, *, r1c1: bool) -> tuple[object, bool]:
        """A formula Formula2 writes, and whether it is a dynamic-array formula. One that Formula would write so
        that its Formula2 reads back as written is written as Formula writes it, @ and all
        (tests/fixtures/implicit_intersection.json, at_sign.json); any other is a dynamic-array formula, as written,
        which spills (tests/fixtures/dynamic_arrays.json). A value is written as Formula writes one."""
        from pyopenvba.formula._formula2 import UnreadFormula2Error, array_evaluation, formula2, written
        from pyopenvba.formula._parse import FormulaError

        if isinstance(value, VBAObject):
            value = value.vba_value()
        if isinstance(value, VBAArray):
            raise VBAUnsupportedError("Range.Formula2 given an array is not implemented")
        if not isinstance(value, str) or not _is_formula(value):
            return value, False
        anchor = self.first
        if r1c1:
            try:
                text = to_a1(value, anchor.top, anchor.left, names=True)
            except ValueError as exc:
                raise error(1004, str(exc)) from None
            spelled = spelled_formula(self.sheet, text, anchor.top, anchor.left, at=True, r1c1=True)
        else:
            spelled = written_formula(self.sheet, value, anchor.top, anchor.left, at=True)
        spelled = at_kept(spelled)
        legacy = ""
        try:
            legacy = written(spelled, self._named)
            same = formula2(legacy, self._named) == spelled and not array_evaluation(spelled, self._named)
        except (FormulaError, UnreadFormula2Error):
            same = False
        return (legacy, False) if same else (spelled, True)

    def _named(self, name: str) -> str | None:
        """A defined name's formula as this sheet finds the name, or None where it finds none."""
        return defined_formula(self.sheet, name)

    def _formula2(self, formula: str, *, whole: bool) -> str:
        """A formula as Formula2 reads it: with an @ where it cuts cells to one value, none in an array formula's
        top, which is worked out whole."""
        from pyopenvba.formula._formula2 import UnreadFormula2Error, formula2
        from pyopenvba.formula._parse import FormulaError

        try:
            return formula2(formula, self._named, whole=whole)
        except UnreadFormula2Error as exc:
            raise VBAUnsupportedError(str(exc)) from None
        except FormulaError:
            raise VBAUnsupportedError(f"Formula2 of {formula} is not implemented") from None

    def _read_formulas(self, *, r1c1: bool, at: bool = False) -> object:
        from pyopenvba.apps.excel._data_tables import table_at

        def read(row: int, column: int) -> object:
            table = table_at(self.sheet, row, column) if self.sheet.data_tables else None
            if table is not None:
                # A data table's cells say =TABLE(), Formula2 included (see _data_tables).
                return table.formula(r1c1=r1c1)
            found = _arrays.array_at(self.sheet, row, column) if self.sheet.array_formulas else None
            if found is not None:
                # Every cell of an array formula reports the array's formula.
                row, column = found[0]
            cell = self.sheet.cell(row, column)
            if cell is None:
                return ""
            if cell.formula:
                text = shown_formula(self.sheet, row, column, cell.formula)
                if at and not cell.dynamic:
                    # A dynamic-array formula is kept as Formula2 writes it, with no @ to add.
                    text = self._formula2(text, whole=found is not None)
                # A call of SINGLE is shown as an @, even where the @ would hold less than SINGLE does.
                text = at_shown(text)
                # R1C1 can spell a formula longer than Excel's 8192 characters, and reads only that many
                # (tests/fixtures/long_chains.json).
                return from_a1(text, row, column)[:LONGEST_FORMULA] if r1c1 else text
            # A cell a formula spilled into has neither a formula nor a constant of its own.
            return "" if cell.spilled_from is not None else _formula_text(cell.value)

        area = self.first
        if self.single:
            return read(area.top, area.left)
        return VBAArray([(1, area.rows), (1, area.columns)], items=[
            read(row, column) for column in range(area.left, area.right + 1)
            for row in range(area.top, area.bottom + 1)
        ])

    @member
    def FormulaArray(self) -> object:
        """The array formula the range is part of, or the formula its cells share, or Null (see _arrays)."""
        return _arrays.formula_array(self, lambda row, column: str(Range(self.sheet, [Area(row, column, row,
                                                                                            column)]).Formula()))

    @setter("FormulaArray")
    def _set_formula_array(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import whole

        text = to_text(value.vba_value() if isinstance(value, VBAObject) else value)
        if not _is_formula(text):
            raise VBAUnsupportedError("FormulaArray set to a value rather than a formula is not implemented")
        if len(text) > LONGEST_ARRAY_FORMULA:
            raise error(1004, "Unable to set the FormulaArray property of the Range class")
        if self.sheet.data_tables:
            from pyopenvba.apps.excel._data_tables import met as tables_met

            tables, split = tables_met(self.sheet, self.areas)
            if split:
                # Part of a data table takes no array formula (tests/fixtures/data_tables.json).
                raise error(1004, "Unable to set the FormulaArray property of the Range class")
            if tables:
                raise VBAUnsupportedError("FormulaArray written over a whole data table is not implemented")
        whole(self.sheet, self.areas, "Changing")
        _arrays.put(self, _array_formula_text(self.sheet, text, self.first.top, self.first.left))
        if not self.sheet.book.calculator.automatic:
            found = _arrays.array_at(self.sheet, self.first.top, self.first.left)
            assert found is not None
            self.sheet.book.calculator.value_of(self.sheet.name, *found[0], force=True)
        _events.after_edit(self)

    @member
    def HasArray(self) -> object:
        return _arrays.has_array(self)

    @member
    def SavedAsArray(self) -> object:
        """Whether the cell is part of an array: one entered with Ctrl+Shift+Enter, or what a dynamic-array formula
        spills (tests/fixtures/range_small_reads.json)."""
        from pyopenvba.apps.excel._spills import has_spill

        whole, part = _arrays.met(self.sheet, self.areas)
        if not self.single:
            if not whole and not part and not has_spill(self):
                return False
            raise VBAUnsupportedError("SavedAsArray of several cells meeting an array is not implemented")
        return bool(whole or part) or has_spill(self) is True

    @member
    def HasRichDataType(self) -> object:
        """False: the model holds no linked or rich data types (tests/fixtures/range_small_reads.json)."""
        return False

    @member
    def LinkedDataTypeState(self) -> object:
        """xlLinkedDataTypeStateNone (tests/fixtures/range_small_reads.json)."""
        return VBAInt(0, "Long")

    @member
    def MDX(self) -> object:
        """An OLAP pivot cell's; empty for one cell elsewhere, error 1004 for several
        (tests/fixtures/range_small_reads.json)."""
        if not self.single:
            raise error(1004)
        from pyopenvba.apps.excel._pivot_cells import owner

        if owner(self.sheet, self.first.top, self.first.left) is not None:
            raise VBAUnsupportedError("Range.MDX in a pivot table is not implemented")
        return ""

    @member
    def Summary(self) -> object:
        """Error 1004 outside a list of the kind Summary reads (tests/fixtures/range_small_reads.json)."""
        raise error(1004)

    @member
    def QueryTable(self) -> object:
        """Error 1004: the model holds no query tables (tests/fixtures/range_small_reads.json)."""
        raise error(1004)

    @member
    def ServerActions(self) -> object:
        """Error 1004 outside a server's list (tests/fixtures/range_small_reads.json)."""
        raise error(1004)

    @member
    def CommentThreaded(self) -> object:
        """Nothing: the model holds no threaded comments; several cells are error 5
        (tests/fixtures/range_small_reads.json)."""
        if not self.single:
            raise error(5)
        return NOTHING

    @member
    def CurrentArray(self) -> object:
        return Range(self.sheet, [_arrays.current_array(self)])

    @method
    def Table(self, RowInput: object = MISSING, ColumnInput: object = MISSING) -> object:
        """A what-if data table over the range, which Excel answers True (see _data_tables)."""
        from pyopenvba.apps.excel._data_tables import make

        make(self, RowInput, ColumnInput)
        return True

    @member
    def Validation(self) -> object:
        """The data validation rule the range's cells have (see _validation)."""
        from pyopenvba.apps.excel._validation import Validation

        return Validation(self)

    @member
    def FormatConditions(self) -> object:
        from pyopenvba.apps.excel._conditional import FormatConditions

        return FormatConditions(self)

    @member
    def DisplayFormat(self) -> object:
        from pyopenvba.apps.excel._conditional_display import DisplayFormat

        application = self.sheet.book.application
        if application.udf_origins and application.udf_origins[-1] == "worksheet":
            raise error(1004)
        return DisplayFormat(self)

    @member
    def ListObject(self) -> object:
        """The table the range's first cell is part of, or Nothing."""
        from pyopenvba.apps.excel._tables import table_at, view

        table = table_at(self.sheet, self.first.top, self.first.left)
        return NOTHING if table is None else view(table)

    @member
    def SparklineGroups(self) -> object:
        """The sparkline groups the range reaches, newest first (see _sparklines)."""
        from pyopenvba.apps.excel._sparklines import SparklineGroups

        return SparklineGroups(self)

    @member
    def PivotCell(self) -> object:
        """What the range's first cell stands for in the pivot table it is part of (see _pivot_cells)."""
        from pyopenvba.apps.excel._pivot_cells import pivot_cell

        return pivot_cell(self)

    @member
    def LocationInTable(self) -> object:
        """Where in a pivot table the range's cells are (see _pivot_cells)."""
        from pyopenvba.apps.excel._pivot_cells import location_in_table

        return location_in_table(self)

    @member
    def PivotTable(self) -> object:
        """The pivot table the range's first cell is part of, its page fields too."""
        from pyopenvba.apps.excel._pivot_cells import pivot_table

        return pivot_table(self)

    @member
    def PivotField(self) -> object:
        """The pivot field the range's first cell names, a value's data field among them."""
        from pyopenvba.apps.excel._pivot_cells import pivot_field

        return pivot_field(self)

    @member
    def PivotItem(self) -> object:
        """The pivot item the range's first cell names, a value's data field's among them."""
        from pyopenvba.apps.excel._pivot_cells import pivot_item

        return pivot_item(self)

    @setter("FormulaR1C1")
    def _set_formula_r1c1(self, value: object) -> None:
        from pyopenvba.apps.excel._protection import writing

        _within_length(value)
        if not _arrays.write_admitted(self):
            return
        with writing(self.sheet):
            self._write_formula_r1c1(value)
        self._tables_follow()
        self._links_emptied(value)
        _events.after_edit(self)

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
                a1 = to_a1(value, row, column, names=True)
            except ValueError as exc:
                raise error(1004, str(exc)) from None
            formula = legacy_formula(self.sheet, spelled_formula(self.sheet, a1, row, column, at=True, r1c1=True))
            first = first or (formula, row, column)
            if self._put_formula(row, column, formula):
                placed.append((row, column))
        if first is not None:
            self._bring_format(placed, *first)
            self._share(placed, first[0])

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
                            item = (to_a1(item, row + extra_row, column + extra_column, names=True) if r1c1
                                    else shift_text(item, extra_row, extra_column))
                        except ValueError as exc:
                            raise error(1004, str(exc)) from None
                    self._put(row, column, item, r1c1=r1c1)

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
        _column_fit.widen(self.sheet, self.areas)

    @member
    def Style(self) -> object:
        from pyopenvba.apps.excel._cell_styles import style_of

        return style_of(self)

    @setter("Style")
    def _set_style(self, value: object) -> None:
        from pyopenvba.apps.excel._cell_styles import set_style

        set_style(self, value)
        _column_fit.widen(self.sheet, self.areas)

    # The Local spellings are in the language of the user's settings: the same text as the plain ones under en-US,
    # as an English Excel has them (tests/fixtures/excel_model/probes.txt); NumberFormatLocal follows the
    # Regional format (_number_format.to_local), FormulaLocal stays English.

    @member
    def NumberFormatLocal(self) -> object:
        from pyopenvba.apps.excel._number_format import to_local

        code = self.NumberFormat()
        return to_local(code, self.sheet.book.application.locale) if isinstance(code, str) else code

    @setter("NumberFormatLocal")
    def _set_number_format_local(self, value: object) -> None:
        from pyopenvba.apps.excel._number_format import from_local

        self._set_number_format(from_local(to_text(value), self.sheet.book.application.locale))

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

    @member
    def Formula2Local(self) -> object:
        return self.Formula2()

    @setter("Formula2Local")
    def _set_formula2_local(self, value: object) -> None:
        self._set_formula2(value)

    @member
    def Formula2R1C1Local(self) -> object:
        return self.Formula2R1C1()

    @setter("Formula2R1C1Local")
    def _set_formula2_r1c1_local(self, value: object) -> None:
        self._set_formula2_r1c1(value)

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
        texts = [area.address(rows_fixed=rows_fixed, columns_fixed=columns_fixed) for area in self.areas]
        if ReferenceStyle is not MISSING and int(to_integer(ReferenceStyle, "Long")) == -4150:  # xlR1C1
            # Relative parts count from RelativeTo, and from A1 without it, whatever is selected.
            anchor = RelativeTo.first if isinstance(RelativeTo, Range) else Area(1, 1, 1, 1)
            texts = [from_a1("=" + text, anchor.top, anchor.left)[1:] for text in texts]
        prefix = ""
        if external:
            from pyopenvba._a1 import quote_sheet

            owner, book = self.sheet.name, self.sheet.book.name
            joined = f"[{book}]{owner}"
            prefix = (("'" + joined.replace("'", "''") + "'")
                      if quote_sheet(owner) != owner or re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", book) is None else joined) + "!"
        return prefix + ",".join(texts)

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
        # A range of rows or columns counts its first area's alone (tests/fixtures/range_lines.json).
        if self.lines == "rows":
            return self.first.rows
        if self.lines == "columns":
            return self.first.columns
        return sum(area.rows * area.columns for area in self.areas)

    def _line(self, index: object) -> Range:
        """The row or column ``index`` of a range of rows or columns, counted from its first area's first and past
        its last if need be; error 1004 before the first."""
        area = self.first
        position = int(to_integer(index, "Long"))
        if position < 1:
            raise error(1004, "Application-defined or object-defined error")
        if self.lines == "rows":
            row = area.top + position - 1
            if row > MAX_ROWS:
                raise error(1004, "Application-defined or object-defined error")
            line = Area(row, area.left, row, area.right, area.sheet)
        else:
            column = area.left + position - 1
            if column > MAX_COLUMNS:
                raise error(1004, "Application-defined or object-defined error")
            line = Area(area.top, column, area.bottom, column, area.sheet)
        return Range(self.sheet, [line], whole=self.whole, lines=self.lines)

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
        if calculator.automatic:
            # Not a calculation that works data tables out in semi-automatic calculation
            # (tests/fixtures/data_tables.json).
            _events.calculated_all(self.sheet.book, calculator.calculate_all())
        else:
            # Deliver earlier notifications, then retain this limited
            # calculation's notification until another calculation request.
            _events.calculated_all(self.sheet.book, [])
            worked = calculator.calculate_all(self.sheet, areas=tuple(self.areas), force=True, volatile=False)
            calculator.pending_notifications.update(worked)
            for (name, row, column), compiled in calculator.compiled.items():
                if name == self.sheet.name.lower() \
                        and any(area.contains(row, column) for area in self.areas):
                    cell = self.sheet.cells_.get((row, column))
                    if cell is not None:
                        cell.stale = True
                    compiled.udf_pending = compiled.calls_user
        return NULL

    @method
    def CalculateRowMajorOrder(self) -> object:
        """Under manual calculation, each formula of the range worked out once, a row at a time, reading what the
        others hold then; under automatic calculation as Calculate. Excel answers Null (see _calculation_state)."""
        if self.sheet.book.calculator.automatic:
            return self.Calculate()
        from pyopenvba.apps.excel._calculation_state import row_major

        row_major(self)
        return NULL

    @method
    def Dirty(self) -> object:
        """The range's formulas, and what reads its cells, to work out again: at once under automatic calculation,
        at the next calculation under manual (see _calculation_state)."""
        from pyopenvba.apps.excel._calculation_state import dirty

        dirty(self)
        return EMPTY

    @member
    def Rows(self, Index: object = MISSING, Second: object = MISSING) -> object:
        """The range as its rows, or its Item(Index), a row (see lines); a second index is error 1004 there."""
        rows = Range(self.sheet, list(self.areas), whole=self.whole if self.whole == "rows" else "", lines="rows")
        return rows if Index is MISSING else rows.Item(Index, Second)

    @member
    def Columns(self, Index: object = MISSING, Second: object = MISSING) -> object:
        """The range as its columns, or its Item(Index), a column (see lines); a second index is error 1004 there."""
        columns = Range(self.sheet, list(self.areas), whole=self.whole if self.whole == "columns" else "",
                        lines="columns")
        return columns if Index is MISSING else columns.Item(Index, Second)

    @member
    def Cells(self, RowIndex: object = MISSING, ColumnIndex: object = MISSING) -> object:
        """The range as its cells, or one of them, whatever Rows or Columns made it (tests/fixtures/range_lines.json)."""
        cells = Range(self.sheet, list(self.areas)) if self.lines else self
        if RowIndex is MISSING:
            return cells
        return cells.vba_get("Item", [RowIndex, ColumnIndex])

    @member
    def Item(self, RowIndex: object = MISSING, ColumnIndex: object = MISSING) -> object:
        area = self.first
        if RowIndex is MISSING:
            raise error(449)
        if self.lines:
            # A range of rows or columns takes one index, a row or a column (tests/fixtures/range_lines.json).
            if ColumnIndex is not MISSING and ColumnIndex is not NOTHING:
                raise error(1004, "Application-defined or object-defined error")
            return self._line(RowIndex)
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
    def Range(self, Cell1: object = MISSING, Cell2: object = MISSING) -> object:
        """A reference read as though the range's first cell were A1, as recorded macros use it:
        Range("B2:D5").Range("B2:C3") is C3:D4, a dollar sign making no difference and a name moving as its cells
        would. Two Ranges are corners read the same way; one alone is error 1004, as is anything that lands off
        the sheet (tests/fixtures/excel_model/)."""
        if Cell1 is MISSING:
            raise error(449)
        if isinstance(Cell1, Range) and Cell2 is MISSING:
            raise error(1004, "Range.Range takes a reference, not a Range on its own")
        found = self.sheet.Range(Cell1, Cell2)
        assert isinstance(found, Range)
        if found.sheet is not self.sheet:
            raise VBAUnsupportedError("Range.Range given a reference on another sheet is not implemented")
        down, across = self.first.top - 1, self.first.left - 1
        areas = [Area(area.top + down, area.left + across, area.bottom + down, area.right + across, area.sheet)
                 for area in found.areas]
        if any(area.bottom > MAX_ROWS or area.right > MAX_COLUMNS for area in areas):
            raise error(1004, "Range.Range would leave the sheet")
        return Range(self.sheet, areas)

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
        # A range of rows or columns stays one (tests/fixtures/range_lines.json).
        return Range(self.sheet, areas, lines=self.lines)

    @member
    def Resize(self, RowSize: object = MISSING, ColumnSize: object = MISSING) -> object:
        area = self.first
        rows = area.rows if RowSize is MISSING else int(to_integer(RowSize, "Long"))
        columns = area.columns if ColumnSize is MISSING else int(to_integer(ColumnSize, "Long"))
        if rows < 1 or columns < 1:
            raise error(ERR_APPLICATION_DEFINED, "Resize needs at least one row and column")
        if area.top + rows - 1 > MAX_ROWS or area.left + columns - 1 > MAX_COLUMNS:
            # Past the sheet's last row or column is error 1004 (native probes).
            raise error(ERR_APPLICATION_DEFINED, "Resize goes past the sheet")
        return Range(self.sheet, [Area(area.top, area.left, area.top + rows - 1, area.left + columns - 1, area.sheet)],
                     lines=self.lines)

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
    def Characters(self, Start: object = MISSING, Length: object = MISSING) -> object:
        """Part of a cell's text, its font among it (see _rich_text)."""
        from pyopenvba.apps.excel._rich_text import characters

        return characters(self, Start, Length)

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
    def Errors(self) -> object:
        """The cell's error checks (see _error_checks)."""
        from pyopenvba.apps.excel._error_checks import Errors

        return Errors(self)

    @member
    def AllowEdit(self) -> object:
        """False where the sheet's contents are protected, for the user alone too, and any cell is locked
        (tests/fixtures/range_small_members.json)."""
        from pyopenvba.apps.excel._protection import any_locked

        protection = self.sheet.protection
        return not (protection is not None and protection.contents and any_locked(self.sheet, self.areas))

    @member
    def ID(self) -> object:
        """The cell's ID, "" for none; a range of more than one cell is error 1004."""
        top = self._single_cell()
        cell = self.sheet.cells_.get((top.top, top.left))
        return cell.html_id if cell is not None else ""

    @setter("ID")
    def _set_id(self, value: object) -> None:
        top = self._single_cell()
        text = to_text(value)
        cell = self.sheet.cells_.get((top.top, top.left))
        if cell is None:
            if not text:
                return
            cell = self.sheet.cells_[(top.top, top.left)] = Cell()
        cell.html_id = text
        self.sheet.settle(top.top, top.left)

    def _single_cell(self) -> Area:
        if len(self.areas) != 1 or self.first.rows != 1 or self.first.columns != 1:
            raise error(1004)
        return self.first

    @member
    def ListHeaderRows(self) -> object:
        """How many header rows Excel guesses the first area has (see _sort.header_rows)."""
        from pyopenvba.apps.excel._sort import header_rows

        return VBAInt(header_rows(self.sheet, self.first), "Long")

    @member
    def Locked(self) -> object:
        return self._format("Locked")

    @setter("Locked")
    def _set_locked(self, value: object) -> None:
        from pyopenvba.interpreter._values import NULL

        if value is NULL:
            return  # Excel takes Null without a word and leaves the cells as they were.
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
        before = _dimensions.visibility_state(self)
        _dimensions.write_row_height(self, value)
        _dimensions.recalculate_visibility(self, before)

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

    @method
    def Group(self, start: object = MISSING, end: object = MISSING, by: object = MISSING,
              periods: object = MISSING) -> object:
        from pyopenvba.apps.excel._outline import group

        if self.sheet.pivot_tables:
            from pyopenvba.apps.excel._pivot_groups import group as group_items, table_at

            if table_at(self) is not None:
                return group_items(self, start, end, by, periods)
        if any(value is not MISSING for value in (start, end, by, periods)):
            raise VBAUnsupportedError('Group with pivot/date/number arguments is not implemented')
        return group(self, 'Group')

    @method
    def Ungroup(self) -> object:
        from pyopenvba.apps.excel._outline import group

        if self.sheet.pivot_tables:
            from pyopenvba.apps.excel._pivot_groups import table_at, ungroup

            if table_at(self) is not None:
                return ungroup(self)
        return group(self, 'Ungroup')

    @method
    def AutoOutline(self) -> object:
        """Group rows and columns from the formulas that sum the cells beside them (see _auto_outline)."""
        from pyopenvba.apps.excel._auto_outline import auto_outline

        auto_outline(self)
        return True

    @method
    def ClearOutline(self) -> object:
        from pyopenvba.apps.excel._outline import group

        return group(self, 'ClearOutline')

    @method
    def Subtotal(self, GroupBy: object = MISSING, Function: object = MISSING, TotalList: object = MISSING,
                 Replace: object = MISSING, PageBreaks: object = MISSING,
                 SummaryBelowData: object = MISSING) -> object:
        """Subtotal rows put into the list around or in the range, as _subtotals has it."""
        from pyopenvba.apps.excel._subtotals import subtotal

        return subtotal(self, GroupBy, Function, TotalList, Replace, PageBreaks, SummaryBelowData)

    @method
    def RemoveSubtotal(self) -> object:
        """The list's subtotal rows taken out, and the list out of the outline."""
        from pyopenvba.apps.excel._subtotals import remove_subtotal

        return remove_subtotal(self)

    @member
    def PageBreak(self) -> object:
        from pyopenvba.apps.excel._page_breaks import page_break

        return page_break(self)

    @setter("PageBreak")
    def _set_page_break(self, value: object) -> None:
        from pyopenvba.apps.excel._page_breaks import set_page_break

        set_page_break(self, value)

    @member
    def OutlineLevel(self) -> object:
        from pyopenvba.apps.excel._outline import read_level

        return read_level(self)

    @setter('OutlineLevel')
    def _set_outline_level(self, value: object) -> None:
        from pyopenvba.apps.excel._outline import set_level

        set_level(self, value)

    @member
    def ShowDetail(self) -> object:
        from pyopenvba.apps.excel._outline import read_detail

        return read_detail(self)

    @setter('ShowDetail')
    def _set_show_detail(self, value: object) -> None:
        from pyopenvba.apps.excel._outline import set_detail

        set_detail(self, value)

    @member
    def Hidden(self) -> object:
        return _dimensions.read_hidden(self)

    @setter("Hidden")
    def _set_hidden(self, value: object) -> None:
        before = _dimensions.visibility_state(self)
        _dimensions.write_hidden(self, value)
        _dimensions.recalculate_visibility(self, before)

    @member
    def UseStandardHeight(self) -> object:
        return _dimensions.read_use_standard_height(self)

    @setter("UseStandardHeight")
    def _set_use_standard_height(self, value: object) -> None:
        before = _dimensions.visibility_state(self)
        _dimensions.write_use_standard_height(self, value)
        _dimensions.recalculate_visibility(self, before)

    @member
    def UseStandardWidth(self) -> object:
        return _dimensions.read_use_standard_width(self)

    @setter("UseStandardWidth")
    def _set_use_standard_width(self, value: object) -> None:
        _dimensions.write_use_standard_width(self, value)

    # -- methods

    @method
    def Select(self) -> object:
        from pyopenvba.apps.excel._windows import check_shown

        check_shown(self, "Select")
        self.select_as_excel_does()
        return EMPTY

    def select_as_excel_does(self) -> None:
        """Make the range the sheet's selection, its first cell the active one; SelectionChange when that moves it."""
        before = self.sheet.selection_range
        shown = [(one.top, one.left, one.bottom, one.right) for one in before.areas] if before is not None \
            else [(1, 1, 1, 1)]
        # The selection is the range's cells, whatever Rows or Columns made it (tests/fixtures/range_lines.json).
        self.sheet.selection_range = self if self.lines == self.whole else Range(self.sheet, list(self.areas),
                                                                                 whole=self.whole)
        self.sheet.active_cell_range = Range(
            self.sheet, [Area(self.first.top, self.first.left, self.first.top, self.first.left, self.first.sheet)]
        )
        self.sheet.view.changed = True
        from pyopenvba.apps.excel._table_filters import refocus

        refocus(self.sheet)
        if shown != [(one.top, one.left, one.bottom, one.right) for one in self.areas]:
            _events.selected(self.sheet, self)

    @method
    def Activate(self) -> object:
        from pyopenvba.apps.excel._windows import activate

        activate(self)
        return EMPTY

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
        _arrays.clear_admitted(target)
        clear_formats(target)
        target._clear_contents()
        header_cleared(target)
        if self.sheet.notes:
            from pyopenvba.apps.excel._notes import clear

            # Clear takes a cell's note too; ClearContents and ClearFormats leave it (tests/fixtures/excel_model/).
            clear(self.sheet, list(target.areas))
        if self.sheet.validations:
            from pyopenvba.apps.excel._validation import remove

            # Clear takes the cells out of their validation rules too; ClearContents leaves them.
            remove(self.sheet, [Area(area.top, area.left, area.bottom, area.right) for area in target.areas])
        if self.sheet.hyperlinks:
            from pyopenvba.apps.excel._hyperlinks import cleared

            cleared(self.sheet, list(target.areas))
        from pyopenvba.apps.excel._sparklines import cleared as sparklines_cleared

        # Clear takes the range's sparklines too; ClearContents leaves them (tests/fixtures/sparklines.json).
        sparklines_cleared(target)
        _events.after_edit(self)
        return EMPTY

    @method
    def ClearContents(self) -> object:
        self.clear_contents_checked()
        _events.after_edit(self)
        return EMPTY

    def clear_contents_checked(self) -> None:
        """Validate and clear contents before the caller publishes its edit events."""
        from pyopenvba.apps.excel._autofilter import header_cleared
        from pyopenvba.apps.excel._protection import PROTECTED, any_locked, enforced
        from pyopenvba.apps.excel._visible import visible

        target = visible(self)
        _merges.validate_clear(target)
        if enforced(self.sheet) is not None and any_locked(self.sheet, target.areas):
            # Unlike a write, a clear that meets a locked cell leaves every cell as it was.
            raise error(1004, PROTECTED)
        _arrays.clear_admitted(target)
        target._clear_contents()
        header_cleared(target)
        if self.sheet.hyperlinks:
            from pyopenvba.apps.excel._hyperlinks import cleared

            # ClearContents takes the links too, the formats they gave staying (tests/fixtures/hyperlinks/).
            cleared(self.sheet, list(target.areas))

    # -- notes, which VBA calls comments (see _notes)

    @method
    def AddComment(self, Text: object = MISSING) -> object:
        from pyopenvba.apps.excel._notes import add_comment

        return add_comment(self, Text)

    @member
    def Comment(self) -> object:
        from pyopenvba.apps.excel._notes import comment_of

        return comment_of(self)

    @method
    def ClearComments(self) -> object:
        from pyopenvba.apps.excel._notes import clear

        clear(self.sheet, list(self.areas))
        return EMPTY

    @method
    def TextToColumns(self, Destination: object = MISSING, DataType: object = MISSING,
                      TextQualifier: object = MISSING, ConsecutiveDelimiter: object = MISSING, Tab: object = MISSING,
                      Semicolon: object = MISSING, Comma: object = MISSING, Space: object = MISSING,
                      Other: object = MISSING, OtherChar: object = MISSING, FieldInfo: object = MISSING,
                      DecimalSeparator: object = MISSING, ThousandsSeparator: object = MISSING,
                      TrailingMinusNumbers: object = MISSING) -> object:
        from pyopenvba.apps.excel._text_columns import text_to_columns

        return text_to_columns(self, Destination, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon,
                               Comma, Space, Other, OtherChar, FieldInfo, DecimalSeparator, ThousandsSeparator,
                               TrailingMinusNumbers)

    @method
    def ClearNotes(self) -> object:
        """The same as ClearComments, as Excel does it (tests/fixtures/excel_model/)."""
        return self.ClearComments()

    # -- hyperlinks (see _hyperlinks)

    @member
    def Hyperlinks(self, Index: object = MISSING) -> object:
        """The links the range holds: those inside one of its areas, and for one area the link it lies in."""
        from pyopenvba.apps.excel._hyperlinks import Hyperlinks

        links = Hyperlinks(self.sheet, list(self.areas))
        return links if Index is MISSING else links.vba_get("Item", [Index])

    @method
    def ClearHyperlinks(self) -> object:
        from pyopenvba.apps.excel._hyperlinks import clear_links

        clear_links(self)
        return EMPTY

    def _links_emptied(self, value: object) -> None:
        """A value that empties the cells takes their links away, as ClearContents does."""
        if self.sheet.hyperlinks:
            from pyopenvba.apps.excel._hyperlinks import cleared, empties

            if empties(value):
                cleared(self.sheet, list(self.areas))

    @method
    def NoteText(self, Text: object = MISSING, Start: object = MISSING, Length: object = MISSING) -> object:
        from pyopenvba.apps.excel._notes import note_text

        return note_text(self, Text, Start, Length)

    def _clear_contents(self) -> None:
        """Every value and formula gone; a cell left with nothing to say goes with them."""
        for row, column in self.cell_positions():
            cell = self.sheet.cells_.get((row, column))
            if cell is None or cell.spilled_from is not None and self.sheet.book.calculator.automatic:
                # Automatic mode restores a cleared child; manual mode invalidates its cached spill.
                continue
            cell.value = EMPTY
            cell.formula = ""
            cell.stale = False
            cell.dynamic = False
            if cell.ignored_errors:
                # A cell cleared checks for errors again (tests/fixtures/range_errors.json).
                cell.ignored_errors = 0
                self.sheet.ignored_changed = True
            cell.html_id = ""
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
        _events.after_edit(self, calculate_first=True)
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
        # xlFormatFromRightOrBelow: the new lines take the formats, height or width of the line after them, not
        # before (tests/fixtures/excel_everyday_2.json, insert_from_below_*).
        below = CopyOrigin is not MISSING and int(to_integer(CopyOrigin, "Long")) == 1
        filtered = filtering(self.sheet)
        if area.whole_rows or area.whole_columns:
            from pyopenvba.apps.excel._editing import edit

            runs = visible_rows(self) if filtered and area.whole_rows and len(self.areas) == 1 else None
            if runs is not None:
                count = sum(bottom - top + 1 for top, bottom in runs)
                area = Area(area.top, 1, area.top + count - 1, MAX_COLUMNS, area.sheet)
                inserted = Range(self.sheet, [area], whole="rows")
                edit(inserted, delete=False, below=below)
                _events.after_edit(inserted, calculate_first=False)
                return EMPTY
            edit(self, delete=False, below=below)
            _events.after_edit(self, calculate_first=False)
            return EMPTY
        if filtered:
            raise error(1004, "Insert method of Range class failed")
        if len(self.areas) != 1:
            raise VBAUnsupportedError("inserting cells in several areas is not implemented")
        from pyopenvba.apps.excel._editing import shift_cells

        self._shifting_cells()
        # With no Shift a range taller than it is wide pushes cells right, any other range down.
        down = area.rows <= area.columns if Shift is MISSING else int(to_integer(Shift, "Long")) == -4121  # xlDown
        shift_cells(self, area, delete=False, vertical=down, below=below)
        _events.after_edit(self, calculate_first=False)
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
        from pyopenvba.apps.excel._clipboard import admit_write

        if not admit_write(Destination.sheet):
            return False
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
        from pyopenvba.apps.excel._outline import guard_axis_copy

        guard_axis_copy(self.sheet, area, Destination.sheet, written)
        from pyopenvba.apps.excel._conditional_transfer import plan as conditional_plan_for
        from pyopenvba.apps.excel._protection import whole

        conditional_plan = conditional_plan_for(self.sheet, [area], Destination.sheet, [written])
        whole(Destination.sheet, [written], "Copying")
        if Destination.sheet.tables:
            from pyopenvba.apps.excel._tables import refuse_growth

            refuse_growth(Destination.sheet, written, "Copying into")
        if any(_merges.intersects(area, one) for one in self.sheet.merged_areas) or any(
            _merges.intersects(written, one) for one in Destination.sheet.merged_areas
        ):
            raise VBAUnsupportedError("Copy involving merged cells is not implemented")
        if rows * columns > 1048576:
            raise VBAUnsupportedError("Copy destinations larger than 1048576 cells are not implemented")
        covered, cut = _arrays.met(Destination.sheet, [written])
        if cut:
            raise error(1004, _arrays.PART_OF_AN_ARRAY)
        if Destination.sheet.data_tables:
            from pyopenvba.apps.excel._data_tables import refuse

            refuse(Destination.sheet, [written], "Copying onto")
        # Snapshot before clearing or writing: source and destination can overlap.
        fragments = _arrays.copy_fragments(self.sheet, area)
        arrays = [fragment for _, fragment in fragments]
        origins = {(fragment.top - area.top, fragment.left - area.left): original for original, fragment in fragments}
        from pyopenvba.apps.excel._spills import CachedSpill, cached, restore
        from pyopenvba.formula._calc.values import Array as FormulaArray

        source_caches = cached(self.sheet)
        destination_caches = source_caches if Destination.sheet is self.sheet else cached(Destination.sheet)
        by_cell = {id(cache.cell): cache for cache in source_caches}
        copied_caches: list[CachedSpill] = []
        sources = {(row - area.top, column - area.left): cell
                   for (row, column), cell in self.sheet.cells_.items() if area.contains(row, column)}
        for position, original in origins.items():
            anchor = self.sheet.cells_.get((original.top, original.left))
            assert anchor is not None
            sources[position] = replace(sources.get(position, Cell()), formula=anchor.formula, shared=None, dynamic=False)
        if self.sheet.data_tables:
            from pyopenvba.apps.excel._data_tables import as_values

            # A data table's cells are copied as the values they show (tests/fixtures/data_tables.json).
            as_values(self.sheet, sources, (area.top, area.left))
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
        names_brought = bool(plan.entries)
        if plan.entries:
            Destination.sheet.book.names_.entries.extend(plan.entries)
            Destination.sheet.book.names_.changed = True
        for position in list(Destination.sheet.cells_):
            if written.contains(*position):
                del Destination.sheet.cells_[position]
        for anchor in covered:
            del Destination.sheet.array_formulas[anchor]
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
                    copy.shared = None
                    # Spill children are cached output, not independent cell contents. A copied anchor
                    # calculates a fresh spill; copying only its children copies empty contents.
                    if copy.spilled_from is not None:
                        copy.value, copy.spilled_from = EMPTY, None
                    if copy.formula:
                        original = origins.get((down, across))
                        row_shift = tile_row - area.top if original is None else tile_row + down - original.top
                        column_shift = tile_column - area.left if original is None else tile_column + across - original.left
                        copy.formula = shift_text(formulas.get((down, across), copy.formula), row_shift, column_shift)
                        copy.stale = True
                        if Destination.sheet.book.calculator.automatic:
                            copy.value = EMPTY
                    source_cache = by_cell.get(id(source))
                    if source_cache is not None:
                        height = min(source_cache.area.bottom, area.bottom) - source_cache.area.top + 1
                        width = min(source_cache.area.right, area.right) - source_cache.area.left + 1
                        answer = None if source_cache.answer is None else FormulaArray(
                            [line[:width] for line in source_cache.answer.rows[:height]])
                        copied_caches.append(CachedSpill(copy, Area(1, 1, height, width), answer,
                                                        source_cache.value, source_cache.why))
                        copy.value = source_cache.value
                    Destination.sheet.cells_[(tile_row + down, tile_column + across)] = copy
                    Destination.sheet.settle(tile_row + down, tile_column + across)
                # A copied array fragment becomes an array of its own where it lands.
                for block in arrays:
                    down, across = tile_row - area.top, tile_column - area.left
                    Destination.sheet.array_formulas[(block.top + down, block.left + across)] = Area(
                        block.top + down, block.left + across, block.bottom + down, block.right + across)
        lone = sources.get((0, 0))
        if area.rows == area.columns == 1 and lone is not None and lone.formula and rows * columns > 1 and not arrays:
            from pyopenvba.apps.excel import _shared

            # One cell's formula copied over a block is one shared formula there, as Excel keeps it.
            if _shared.shares(lone.formula):
                _shared.group(Destination.sheet, written, [
                    Destination.sheet.cells_[(row, column)] for row in range(written.top, written.bottom + 1)
                    for column in range(written.left, written.right + 1)])
        # A position that showed its row's or column's format keeps showing it where the destination's differs.
        for plan in shown:
            for (row, column), style in plan:
                if style != Destination.sheet.inherited_style(row, column):
                    Destination.sheet.restyle(row, column, style)
        if self.sheet.validations or Destination.sheet.validations:
            from pyopenvba.apps.excel._validation import copied

            # The copy takes the validation rules of the cells copied along, each tile of it.
            for tile_row in range(target.top, bottom + 1, area.rows):
                for tile_column in range(target.left, right + 1, area.columns):
                    copied(self.sheet, Area(area.top, area.left, area.bottom, area.right), Destination.sheet,
                           Area(tile_row, tile_column, tile_row + area.rows - 1, tile_column + area.columns - 1))
        if self.sheet.hyperlinks or Destination.sheet.hyperlinks:
            from pyopenvba.apps.excel import _hyperlinks

            # Each tile brings the links the copy meets, and takes off the links it lands on
            # (tests/fixtures/hyperlinks/).
            for tile_row in range(target.top, bottom + 1, area.rows):
                for tile_column in range(target.left, right + 1, area.columns):
                    _hyperlinks.copied(self.sheet, Area(area.top, area.left, area.bottom, area.right),
                                       Destination.sheet, Area(tile_row, tile_column, tile_row + area.rows - 1,
                                                               tile_column + area.columns - 1))
        if self.sheet.notes or Destination.sheet.notes:
            from pyopenvba.apps.excel import _notes

            # Each cell copied takes its note along, and one without takes off the note it lands on
            # (tests/fixtures/excel_model/).
            notes = {(tile_row + row - area.top, tile_column + column - area.left): note
                     for (row, column), note in self.sheet.notes.items() if area.contains(row, column)
                     for tile_row in range(target.top, bottom + 1, area.rows)
                     for tile_column in range(target.left, right + 1, area.columns)}
            _notes.copied(Destination.sheet, notes, [written])
        if conditional_plan is not None:
            conditional_plan.commit()
        restore(Destination.sheet, copied_caches + destination_caches)
        if names_brought:
            # Names the copy brought can change what any formula means.
            Destination.sheet.shape_changed()
        else:
            Destination.sheet.touched()
            Destination.sheet.book.calculator.wrote_area(Destination.sheet.name, written)
        return True

    @method
    def Consolidate(self, Sources: object = MISSING, Function: object = MISSING, TopRow: object = MISSING,
                    LeftColumn: object = MISSING, CreateLinks: object = MISSING) -> object:
        """Sum or otherwise combine source ranges into this one (see _consolidate)."""
        from pyopenvba.apps.excel._consolidate import consolidate

        consolidate(self, Sources, Function, TopRow, LeftColumn, CreateLinks)
        return True

    @method
    def ApplyNames(self, Names: object = MISSING, IgnoreRelativeAbsolute: object = MISSING,
                   UseRowColumnNames: object = MISSING, OmitColumn: object = MISSING, OmitRow: object = MISSING,
                   Order: object = MISSING, AppendLast: object = MISSING) -> object:
        """Turn references in the range's formulas into the names that cover them (see _apply_names)."""
        from pyopenvba.apps.excel._apply_names import apply_names

        apply_names(self, Names, IgnoreRelativeAbsolute)
        return EMPTY

    @method
    def AutoFill(self, Destination: object = MISSING, Type: object = MISSING) -> object:
        from pyopenvba.apps.excel._autofill import autofill

        answer = autofill(self, Destination, Type)
        if isinstance(Destination, Range) and len(Destination.areas) == 1:
            # Excel selects what it filled, then reports the cells it filled as changed.
            if Destination.sheet is self.sheet.book.active_sheet:
                Destination.select_as_excel_does()
            filled = _beyond(self.first, Destination.first)
            if filled is not None:
                _events.after_edit(Range(Destination.sheet, [filled]))
        return answer

    @method
    def FillDown(self) -> object:
        answer = self._fill(1, 0)
        _events.after_edit(self)
        return answer

    @method
    def FillUp(self) -> object:
        answer = self._fill(-1, 0)
        _events.after_edit(self)
        return answer

    @method
    def FillRight(self) -> object:
        answer = self._fill(0, 1)
        _events.after_edit(self)
        return answer

    @method
    def FillLeft(self) -> object:
        answer = self._fill(0, -1)
        _events.after_edit(self)
        return answer

    def _fill(self, down: int, across: int) -> object:
        """Copy each area's first row or column, in the direction of the fill, over the rest of it.

        Everything the source holds goes -- values, formulas moved with
        each row or column, formats, and blanks over what was there. An
        area one row deep takes its source from the row before it, and one
        with no row before it is error 1004, before anything changes.
        On a filtered sheet the visible cells fill, as _visible has it.
        """
        from pyopenvba.apps.excel._sparklines import refuse as refuse_sparklines

        refuse_sparklines(self.sheet, "Filling on")
        from pyopenvba.apps.excel._visible import visible_areas

        if self.sheet.tables:
            from pyopenvba.apps.excel._tables import refuse_growth

            for area in self.areas:
                refuse_growth(self.sheet, area, "Filling into")
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
        if self.sheet.data_tables:
            from pyopenvba.apps.excel._data_tables import refuse as refuse_tables

            refuse_tables(self.sheet, [one for plan in plans for one in plan], "Filling from or into")
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
        answer = range_sort(self, [(Key1, Order1, DataOption1), (Key2, Order2, DataOption2),
                                   (Key3, Order3, DataOption3)], Header, OrderCustom, MatchCase, Orientation)
        # A sort changes no cell as far as Change goes; what it moved is worked out again.
        _events.recalculated(self.sheet.book)
        return answer

    @method
    def AutoFilter(self, Field: object = MISSING, Criteria1: object = MISSING, Operator: object = MISSING,
                   Criteria2: object = MISSING, VisibleDropDown: object = MISSING, SubField: object = MISSING) -> object:
        from pyopenvba.apps.excel._autofilter import range_autofilter
        from pyopenvba.apps.excel._protection import FILTERING, refuse

        # AllowFiltering lets the user filter, not a macro.
        refuse(self.sheet, FILTERING)
        return range_autofilter(self, Field, Criteria1, Operator, Criteria2, VisibleDropDown, SubField)

    @method
    def AdvancedFilter(self, Action: object, CriteriaRange: object = MISSING, CopyToRange: object = MISSING,  # noqa: N803 -- Excel argument names
                       Unique: object = MISSING) -> object:  # noqa: N803
        from pyopenvba.apps.excel._advanced_filter import advanced_filter

        return advanced_filter(self, Action, CriteriaRange, CopyToRange, Unique)

    @method
    def RemoveDuplicates(self, Columns: object = MISSING, Header: object = MISSING) -> object:
        from pyopenvba.apps.excel._duplicates import remove_duplicates
        from pyopenvba.apps.excel._protection import APPLICATION_DEFINED, refuse

        refuse(self.sheet, APPLICATION_DEFINED)
        return remove_duplicates(self, Columns, Header)

    @method
    def DataSeries(self, Rowcol: object = MISSING, Type: object = MISSING, Date: object = MISSING,
                   Step: object = MISSING, Stop: object = MISSING, Trend: object = MISSING) -> object:
        from pyopenvba.apps.excel._data_series import data_series

        return data_series(self, Rowcol, Type, Date, Step, Stop, Trend)

    @method
    def InsertIndent(self, InsertAmount: object = MISSING) -> object:
        """Each cell's IndentLevel raised by a Long, as IndentLevel sets it, a centred cell made left; 0, or a level
        that would go outside 0 to 250, is error 1004 and changes nothing (tests/fixtures/small_members.json)."""
        if InsertAmount is MISSING:
            raise error(449)
        amount = int(to_integer(InsertAmount, "Long"))
        failed = error(1004, "Application-defined or object-defined error")
        if amount == 0:
            raise failed
        level = self.IndentLevel()
        if level is not NULL:
            if not 0 <= int(to_integer(level, "Long")) + amount <= 250:
                raise failed
            self.vba_set("IndentLevel", int(to_integer(level, "Long")) + amount)
            return EMPTY
        cells = [(row, column) for area in self.areas for row in range(area.top, area.bottom + 1)
                 for column in range(area.left, area.right + 1)]
        if len(cells) > 100_000:
            raise VBAUnsupportedError("InsertIndent over more than 100,000 cells of mixed indents is not implemented")
        levels = [int(to_integer(Range(self.sheet, [Area(row, column, row, column)]).IndentLevel(), "Long"))
                  for row, column in cells]
        if any(not 0 <= one + amount <= 250 for one in levels):
            raise failed
        for (row, column), one in zip(cells, levels, strict=True):
            Range(self.sheet, [Area(row, column, row, column)]).vba_set("IndentLevel", one + amount)
        return EMPTY

    @method
    def CreateNames(self, Top: object = MISSING, Left: object = MISSING, Bottom: object = MISSING,
                    Right: object = MISSING) -> object:
        from pyopenvba.apps.excel._create_names import create_names

        return create_names(self, Top, Left, Bottom, Right)

    @method
    def ListNames(self) -> object:
        from pyopenvba.apps.excel._create_names import list_names

        return list_names(self)

    @member
    def AutoComplete(self, String: object = MISSING) -> object:
        from pyopenvba.apps.excel._auto_complete import auto_complete

        return auto_complete(self, String)

    @method
    def Parse(self, ParseLine: object = MISSING, Destination: object = MISSING) -> object:
        from pyopenvba.apps.excel._text_columns import parse

        return parse(self, ParseLine, Destination)

    @method
    def RowDifferences(self, Comparison: object = MISSING) -> object:
        from pyopenvba.apps.excel._special_cells import differences

        return differences(self, Comparison, rows=True)

    @method
    def ColumnDifferences(self, Comparison: object = MISSING) -> object:
        from pyopenvba.apps.excel._special_cells import differences

        return differences(self, Comparison, rows=False)

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

        # The cell and its neighbour decide which of the two journeys this
        # is: from a filled cell along a run of filled cells to its last
        # one, or across a gap, from an empty cell too, to the next filled
        # cell, or to the edge of the sheet if there is none.
        step_row, step_column = row + down, column + across
        if occupied(row, column) and inside(step_row, step_column) and occupied(step_row, step_column):
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
        calculated: set[Worksheet] = set()
        answer = replace(self, What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat,
                         ReplaceFormat, FormulaVersion, calculated=calculated)
        # Each cell replaced was reported as it changed; the calculation comes after them all.
        _events.recalculated(self.sheet.book, already_calculated=calculated)
        return answer

    def replaced_in(self, row: int, column: int, text: str) -> bool:
        """Enter the text Replace left in one cell; False, changing nothing, for a formula Excel cannot read.

        A cell that shows a prefix character keeps its text as text; any
        other cell, a Text cell's included, takes a formula as a formula.
        """
        from pyopenvba.apps.excel._typing import replaced

        if not _merges.writable(self.sheet, row, column):
            return True
        if self.sheet.data_tables:
            from pyopenvba.apps.excel._data_tables import refuse

            refuse(self.sheet, [Area(row, column, row, column)], "Replace in")
        cell = self.sheet.cell(row, column)
        array = _arrays.array_at(self.sheet, row, column)
        if array is not None:
            anchor, block = array
            if anchor != (row, column):
                return True
            target = Range(self.sheet, [block])
            if _is_formula(text):
                try:
                    formula = spelled_formula(self.sheet, text, row, column)
                except VBARuntimeError:
                    return False
                _arrays.put(target, formula, limit=LONGEST_FORMULA)
                self.sheet.book.calculator.value_of(self.sheet.name, *anchor, force=True)
            else:
                _arrays.clear_admitted(target)
                for each_row, each_column in target.cell_positions():
                    target._store(each_row, each_column, replaced(text, self.sheet.style_at(each_row, each_column).number_format,
                                                                prefixed=False, epoch_1904=self.sheet.book.epoch_1904))
            _events.changed(target)
            return True
        prefixed = cell is not None and not cell.formula and isinstance(cell.value, str) \
            and cell.style is not None and cell.style.quote_prefix
        if _is_formula(text) and not prefixed:
            try:
                formula = spelled_formula(self.sheet, text, row, column)
            except VBARuntimeError:
                return False
            self._put_formula(row, column, formula)
            self.sheet.book.calculator.settle_entries()
        else:
            self._store(row, column, replaced(text, self.sheet.style_at(row, column).number_format,
                                              prefixed=prefixed, epoch_1904=self.sheet.book.epoch_1904))
        # Replace reports each cell as it changes it.
        _events.changed(Range(self.sheet, [Area(row, column, row, column, self.sheet.name)]))
        return True

    @method
    def AutoFit(self) -> object:
        from pyopenvba.apps.excel._protection import refuse, whole_lines

        allow = "AllowFormattingRows" if (self.lines or whole_lines(self)) == "rows" else "AllowFormattingColumns"
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
        if self.lines:
            yield from self._walk_lines()
            return
        for row, column in self.positions():
            yield Range(self.sheet, [Area(row, column, row, column, self.sheet.name)])

    def _walk_lines(self) -> Iterator[Range]:
        """For Each over a range of rows or columns: every area's, one at a time; the rows of an area spanning every
        row of the sheet only as far down as the sheet uses."""
        for area in self.areas:
            if self.lines == "columns":
                for column in range(area.left, area.right + 1):
                    yield Range(self.sheet, [Area(area.top, column, area.bottom, column, area.sheet)], whole=self.whole,
                                lines="columns")
                continue
            bottom = area.bottom
            if area.rows >= MAX_ROWS:
                used = self.sheet.used_bounds()
                bottom = area.top if used is None else min(area.bottom, max(used[2], area.top))
            for row in range(area.top, bottom + 1):
                yield Range(self.sheet, [Area(row, area.left, row, area.right, area.sheet)], whole=self.whole,
                            lines="rows")

    def _read(self, row: int, column: int, *, raw: bool = False) -> object:
        """A cell's value as VBA reads it; ``raw`` is Value2, which leaves a number a Double whatever its format."""
        from pyopenvba.apps.excel._calc import as_vba

        if self.sheet.array_formulas and _arrays.array_at(self.sheet, row, column) is not None:
            # A cell of an array formula shows its item of the array, worked out first if it has to be.
            value = self.sheet.book.calculator.value_of(self.sheet.name, row, column)
            return as_vba(value, None if raw else self.sheet.cell(row, column), epoch_1904=self.sheet.book.epoch_1904)
        cell = self.sheet.cell(row, column)
        if cell is None:
            return EMPTY
        value = self.sheet.book.calculator.value_of(self.sheet.name, row, column) if cell.formula else cell.value
        return as_vba(value, None if raw else cell, epoch_1904=self.sheet.book.epoch_1904)

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

    def _put(self, row: int, column: int, value: object, *, raw: bool = False, r1c1: bool = False) -> None:
        """Write one cell as Value does; ``r1c1`` for a formula already turned from R1C1 into A1."""
        if not _merges.writable(self.sheet, row, column):
            return
        if isinstance(value, str) and _is_formula(value) and not self._keeps_text(row, column):
            formula = legacy_formula(self.sheet, spelled_formula(self.sheet, value, row, column, at=True, r1c1=True)
                                     if r1c1 else written_formula(self.sheet, value, row, column, at=True))
            if self._put_formula(row, column, formula):
                self._bring_format([(row, column)], formula, row, column)
        else:
            self._type_into(row, column, value, raw=raw)

    def _keeps_text(self, row: int, column: int) -> bool:
        """Whether a position is a Text cell, which keeps whatever string is written to it as text."""
        from pyopenvba.formula._display import keeps_text

        return keeps_text(self.sheet.style_at(row, column).number_format)

    def _put_formula(self, row: int, column: int, formula: str, *, dynamic: bool = False) -> bool:
        """Write one formula into one cell, ``dynamic`` for a dynamic-array formula; False when a protected sheet
        held the cell back."""
        gate = self.sheet.write_gate
        if gate is not None and not gate.admits(row, column):
            return False
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        cell.formula = formula
        cell.stale = True
        cell.save_dirty = None
        cell.value = EMPTY
        cell.shared = None
        cell.dynamic = dynamic
        cell.spilled_from = None
        self.sheet.cell_changed(row, column)
        self.sheet.book.calculator.entered(self.sheet.name, row, column)
        return True

    def _share(self, placed: list[tuple[int, int]], formula: str) -> None:
        """Make a formula written to a whole block at once one shared formula over the block, as Excel keeps it."""
        from pyopenvba.apps.excel import _shared

        area = self.first
        if len(self.areas) != 1 or len(placed) < 2 or len(placed) != area.rows * area.columns \
                or not _shared.shares(formula):
            return
        _shared.group(self.sheet, area, [self.sheet.cells_[position] for position in placed])

    def _bring_format(self, placed: list[tuple[int, int]], formula: str, row: int, column: int) -> None:
        """Give the cells a formula was just written to the number format it brings, where they are General.

        Excel works the format out once, from the formula as written at the
        top left of the range, and gives it to every cell of the write.
        """
        from pyopenvba.apps.excel._formula_format import brought_format
        from pyopenvba.apps.excel._hyperlinks import styles_as_link

        if placed and styles_as_link(self.sheet, formula):
            from pyopenvba.apps.excel._cell_styles import give_link_style

            # A formula whose first function is HYPERLINK gives its cells the Hyperlink style
            # (tests/fixtures/hyperlinks/formulas.xlsx).
            give_link_style(Range(self.sheet, [Area(one_row, one_column, one_row, one_column)
                                               for one_row, one_column in placed]))
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

        from pyopenvba.apps.excel._spills import keeps

        gate = self.sheet.write_gate
        if gate is not None and not gate.admits(row, column):
            return
        if keeps(self.sheet, row, column, value):
            # An empty string or nothing written where a formula spilled leaves the spill as it is.
            return
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        style = cell.style or self.sheet.book.stylesheet.default
        self._store(row, column, typed(value, style.number_format, raw=raw, epoch_1904=self.sheet.book.epoch_1904))

    def store_typed(self, row: int, column: int, result: Typed) -> None:
        """Put what a feature typed into one cell, as a write puts it (see _store): TextToColumns's fields."""
        self._store(row, column, result)

    def put_as_value(self, row: int, column: int, value: object) -> None:
        """Write one cell as Value writes it, a formula as a formula (see _put)."""
        self._put(row, column, value)

    def _store(self, row: int, column: int, result: Typed) -> None:
        """Put what typing made of a write into one cell: its value, the format it ends with, and a prefix."""
        cell = self.sheet.cell(row, column, create=True)
        assert cell is not None
        style = cell.style or self.sheet.book.stylesheet.default
        cell.value = result.value
        cell.formula = ""
        cell.stale = False
        cell.shared = None
        cell.dynamic = False
        cell.spilled_from = None
        # A value written, the same text too, is in the cell's own font (tests/fixtures/characters.json).
        cell.runs = None
        if cell.ignored_errors:
            # And is checked for errors again (tests/fixtures/range_errors.json).
            cell.ignored_errors = 0
            self.sheet.ignored_changed = True
        changed = style
        if result.number_format is not None:
            changed = applying(changed, "number_format", number_format=result.number_format)
        if result.prefix and not changed.quote_prefix:
            changed = replace(changed, quote_prefix=True)
        if changed != style:
            self.sheet.restyle(row, column, changed)
        _precision.settle(self.sheet, row, column)
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

    @member
    def Parent(self) -> object:
        """The range of every area (tests/fixtures/object_parents.json)."""
        return self.target

    @member
    def Application(self) -> object:
        return self.target.sheet.book.application


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
    #: Macro metadata (see _name_macro): 0 for an ordinary name, 1 a function, 2 a command; the category's
    #: functionGroupId; and whether any of it changed, so a name read from a file keeps its attributes as written.
    macro_type: int = 0
    shortcut_key: str = ""
    category: str = ""
    group: int = 0
    macro_changed: bool = False


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

    @member
    def Parent(self) -> object:
        """The workbook, or for a sheet's Names the sheet (tests/fixtures/small_members.json)."""
        return self.sheet if self.sheet is not None else self.book

    @member
    def Application(self) -> object:
        return self.book.application

    def vba_items(self) -> list[object]:
        from pyopenvba._a1 import split_sheet

        self._check_alive()
        entries = [entry for entry in self.entries if self.sheet is None or split_sheet(entry.name)[0].casefold() == self.sheet.name.casefold()]
        from pyopenvba.apps.excel._names import collation

        def order(entry: NameEntry) -> tuple[object, ...]:
            scope, bare = split_sheet(entry.name)
            return (collation(bare), 0 if scope else 1, collation(scope))

        return [DefinedName(entry) for entry in sorted(entries, key=order)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            found = self._item(index)
            if found is None:
                raise error(1004, "Application-defined or object-defined error")
            return found
        return super().vba_lookup(index, items)

    def _item(self, name: str) -> DefinedName | None:
        """A name as Names(...) finds it by its text: a sheet's Names only the sheet's own; the workbook's one
        qualified with its sheet, or else the first sheet's own before the workbook's, whichever sheet is active,
        and never another sheet's (tests/fixtures/names_add/, lookup_*)."""
        from pyopenvba._a1 import split_sheet

        owner, bare = split_sheet(name)
        if self.sheet is not None:
            scopes = [self.sheet.name] if not owner or _names.same_sheet_name(owner, self.sheet.name) else []
        elif owner and owner.casefold() != self.book.name.casefold():
            scopes = [owner]
        elif owner and any(_names.same_sheet_name(sheet.name, owner) for sheet in self.book.sheets_):
            # A book Workbooks.Add(xlWBATWorksheet) made is named for its sheet, Sheet1: Sheet1!Local is the sheet's
            # own name first (tests/fixtures/small_members.json).
            scopes = [owner, ""]
        else:
            scopes = ([self.book.sheets_[0].name] if self.book.sheets_ and not owner else []) + [""]
        for scope in scopes:
            for entry in self.entries:
                entry_scope, entry_bare = split_sheet(entry.name)
                if _names.same_name(entry_bare, bare) and _names.same_sheet_name(entry_scope, scope):
                    return DefinedName(entry)
        return None

    def find(self, name: str, *, scope: Worksheet | None = None) -> DefinedName | None:
        from pyopenvba._a1 import split_sheet

        owner, bare = split_sheet(name)
        workbook_qualified = owner.casefold() == self.book.name.casefold()
        context = self.sheet or scope or self.book.active_sheet
        wanted_scope = owner or (context.name if context is not None else "")
        fallback = None
        for entry in self.entries:
            entry_scope, entry_bare = split_sheet(entry.name)
            if not _names.same_name(entry_bare, bare):
                continue
            if workbook_qualified:
                if not entry_scope:
                    return DefinedName(entry)
                continue
            if _names.same_sheet_name(entry_scope, wanted_scope):
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
        wanted = _names.canonical(self.book, to_text(Name), self.sheet)
        owner = self.sheet or self.book.active_sheet
        if owner is None:
            raise error(1004, "No worksheet is available for the name")
        if self.sheet is None and "!" not in wanted:
            from pyopenvba._a1 import quote_sheet

            # The workbook's Names take a name the active sheet has one of its own of as that one
            # (tests/fixtures/names_add/, active_local_then_book).
            local = quote_sheet(owner.name) + "!" + wanted
            if any(_names.same_entry(one.name, local) for one in self.entries):
                wanted = local
        refers = RefersTo if RefersTo is not MISSING else RefersToLocal
        if refers is MISSING:
            rc = RefersToR1C1 if RefersToR1C1 is not MISSING else RefersToR1C1Local
            if rc is MISSING:
                raise error(449)
            refers = to_a1(to_text(rc), 1, 1)
        elif not isinstance(refers, Range):
            refused = _names.refused_structured(to_text(refers), owner)
            if refused is not None:
                raise error(1004, refused)
        if isinstance(refers, Range):
            from pyopenvba._a1 import quote_sheet

            text = "=" + ",".join(quote_sheet(refers.sheet.name) + "!" + area.address() for area in refers.areas)
        else:
            text = _names.qualify(to_text(refers), owner.name)
        # A name word sort cannot tell from one there is redefines it, and takes the spelling given where that is
        # the old one's as Excel keeps names apart, ECLAIR for eclair but not aesop for the ash's
        # (tests/fixtures/name_equality.json).
        from pyopenvba._a1 import split_sheet

        entry = next((one for one in self.entries if _names.same_entry(one.name, wanted)), None)
        if entry is None:
            entry = NameEntry(wanted, text, self.book)
            self.entries.append(entry)
        elif _names.same_sheet_name(split_sheet(entry.name)[1], split_sheet(wanted)[1]):
            entry.name = wanted
        entry.refers_to = text
        if Visible is not MISSING:
            entry.visible = to_bool(Visible)
        from pyopenvba.apps.excel._name_macro import added

        added(entry, MacroType, ShortcutKey, Category, CategoryLocal)
        _names.changed(self.book, entry.name)
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
        if any(one is not self.entry and _names.same_entry(one.name, wanted) for one in book.names_.entries):
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
        before, self.entry.name = self.entry.name, wanted
        _names.changed(book, before, wanted)

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
    def MacroType(self) -> object:
        """xlFunction or xlCommand, or xlNone for an ordinary name (see _name_macro)."""
        from pyopenvba.apps.excel._name_macro import macro_type

        return VBAInt(macro_type(self.entry), "Long")

    @setter("MacroType")
    def _set_macro_type(self, value: object) -> None:
        from pyopenvba.apps.excel._name_macro import set_macro_type

        set_macro_type(self.entry, value)
        _names.changed(self.entry.book)

    @member
    def ShortcutKey(self) -> object:
        from pyopenvba.apps.excel._name_macro import shortcut_key

        return shortcut_key(self.entry)

    @setter("ShortcutKey")
    def _set_shortcut_key(self, value: object) -> None:
        from pyopenvba.apps.excel._name_macro import set_shortcut_key

        set_shortcut_key(self.entry, value)
        _names.changed(self.entry.book)

    @member
    def Category(self) -> object:
        from pyopenvba.apps.excel._name_macro import category

        return category(self.entry)

    @setter("Category")
    def _set_category(self, value: object) -> None:
        from pyopenvba.apps.excel._name_macro import set_category

        set_category(self.entry, value)
        _names.changed(self.entry.book)

    @member
    def CategoryLocal(self) -> object:
        return self.Category()

    @setter("CategoryLocal")
    def _set_category_local(self, value: object) -> None:
        self._set_category(value)

    @member
    def WorkbookParameter(self) -> object:
        """False: Excel refuses every setting with error 5 (tests/fixtures/name_macro.json)."""
        return False

    @setter("WorkbookParameter")
    def _set_workbook_parameter(self, value: object) -> None:
        raise error(5)

    @member
    def ValidWorkbookParameter(self) -> object:
        """True for a workbook's own name for one cell written all absolute, $A$1, once the workbook has been saved
        (tests/fixtures/name_macro.json); never for a macro's name."""
        from pyopenvba._a1 import split_sheet

        if not self.entry.book.path or split_sheet(self.entry.name)[0] or self.entry.macro_type:
            return False
        sheet, bare = split_sheet(self.entry.refers_to.lstrip("="))
        return bool(sheet) and re.fullmatch(r"\$[A-Za-z]{1,3}\$[0-9]+", bare) is not None

    @member
    def Parent(self) -> object:
        """The workbook, or for a sheet's own name the sheet (tests/fixtures/names_add/, lookup_*)."""
        from pyopenvba._a1 import split_sheet

        scope = split_sheet(self.entry.name)[0]
        sheet = self.entry.book.sheet_named(scope) if scope else None
        return sheet if sheet is not None else self.entry.book

    @member
    def Application(self) -> object:
        return self.entry.book.application

    @member
    def Index(self) -> object:
        """The name's place among the workbook's names in their order, read through a sheet's Names too
        (tests/fixtures/small_members.json)."""
        items = self.entry.book.names_.vba_items()
        return VBAInt(next(at for at, item in enumerate(items, start=1)
                           if isinstance(item, DefinedName) and item.entry is self.entry), "Long")

    def shown_refers_to(self) -> str:
        """The name's formula as RefersTo shows it: another workbook by its path (see _external_links)."""
        from pyopenvba.apps.excel._external_links import shown

        return shown(self.entry.book, self.entry.refers_to)

    @member
    def RefersToR1C1(self) -> object:
        return from_a1(self.shown_refers_to(), 1, 1)

    @setter("RefersToR1C1")
    def _set_refers_r1c1(self, value: object) -> None:
        self._set_refers_to(to_a1(to_text(value), 1, 1))

    @member
    def RefersTo(self) -> object:
        return self.shown_refers_to()

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
        _names.changed(self.entry.book, self.entry.name)

    @member(default=True)
    def Value(self) -> object:
        return self.shown_refers_to()

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
        _names.changed(book, self.entry.name)
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

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

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

    @member(default=True, name="_Default")
    def _Default(self) -> object:
        """Its name (tests/fixtures/collection_members.json)."""
        return self.entry.name

    @member
    def Parent(self) -> object:
        return self.book.queries_

    @member
    def Application(self) -> object:
        return self.book.application

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
            # A sheet function WorksheetFunction does not carry, If or Sin, is no member either (error 438).
            if member_exists("WorksheetFunction", name, "excel"):
                raise VBAUnsupportedError(
                    f"WorksheetFunction.{name} is a real Excel function that pyOpenVBA "
                    f"does not implement"
                )
            raise error(438, f"WorksheetFunction has no member named {name}")
        return super().vba_get(name, args, named)

    @member
    def Parent(self) -> object:
        return self.application

    @member
    def Application(self) -> object:
        return self.application


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


def _register_settings() -> None:
    """A member for each setting kept apart from the classes' own (see _application_settings and
    _worksheet_settings)."""
    from pyopenvba.apps.excel._application_settings import application_members, worksheet_members
    from pyopenvba.apps.excel._worksheet_settings import members as sheet_settings

    for spec in application_members():
        Application.vba_add_member(spec)
    for spec in [*worksheet_members(), *sheet_settings()]:
        Worksheet.vba_add_member(spec)


_register_settings()


# --- helpers ----------------------------------------------------------------------------------------


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

    def __init__(self, sheet: Worksheet, row: int = 0, column: int = 0) -> None:
        self.book = sheet.book
        self.home = sheet.name
        self.at = (sheet, row, column)

    def sheet(self, name: str) -> str | None:
        """A sheet's own name, or another workbook's sheet as the file names it by its link, [1]Sheet1 (see
        _external_links)."""
        found = next((one.name for one in self.book.sheets_ if one.name.lower() == name.lower()), None)
        if found is None and ("[" in name or "\\" in name or name.lower().endswith((".xlsx", ".xlsm", ".xlsb", ".xls"))):
            from pyopenvba.apps.excel._external_links import linked_sheet

            return linked_sheet(self.book, name)
        return found

    def order(self) -> list[str]:
        return [one.name for one in self.book.sheets_]

    def defined(self, name: str, sheet: str) -> str | None:
        from pyopenvba._a1 import split_sheet

        scope = next((one for one in self.book.sheets_ if one.name.lower() == sheet.lower()), None)
        found = self.book.names_.find(name, scope=scope)
        if found is None:
            # A table's name is the workbook's too, and stands for the table's data.
            table = self.table(name)
            return None if table is None else table.name
        return split_sheet(found.entry.name)[1]

    def table(self, name: str) -> TableShape | None:
        return self.book.calculator.table(name)

    def here(self) -> TableShape | None:
        from pyopenvba.apps.excel._tables import shape, table_at

        sheet, row, column = self.at
        found = table_at(sheet, row, column)
        return None if found is None else shape(found)

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


def _within_length(value: object) -> None:
    """Error 7, out of memory, for a formula longer than Excel keeps, written through Formula, Formula2,
    FormulaR1C1 or Value (tests/fixtures/formula_refusals.json)."""
    if isinstance(value, str) and _is_formula(value) and len(value) > LONGEST_FORMULA:
        raise error(7, "Out of memory")


def _beyond(source: Area, destination: Area) -> Area | None:
    """The cells a fill from ``source`` over ``destination`` filled: the destination past the source."""
    if destination.bottom > source.bottom:
        return Area(source.bottom + 1, destination.left, destination.bottom, destination.right)
    if destination.top < source.top:
        return Area(destination.top, destination.left, source.top - 1, destination.right)
    if destination.right > source.right:
        return Area(destination.top, source.right + 1, destination.bottom, destination.right)
    if destination.left < source.left:
        return Area(destination.top, destination.left, destination.bottom, source.left - 1)
    return None


#: A name that is an R1C1 reference, which A1 will not take as a name: R1C1, RC2, R3C, R, C.
def written_formula(sheet: Worksheet, formula: str, row: int, column: int, *, whole: bool = False,
                    at: bool = False) -> str:
    """A formula a macro writes in A1, as Excel spells it: read as R1C1 where A1 cannot read it, as Excel reads
    =R[-1]C+1 written to C8 as =C7+1 and =R1C1 as =$A$1, what looks like a cell being a name there, 'A1'
    (tests/fixtures/formula_notation.json). ``whole`` and ``at`` as for :func:`spelled_formula`."""
    try:
        return spelled_formula(sheet, formula, row, column, whole=whole, at=at)
    except VBARuntimeError:
        pass
    try:
        converted = to_a1(formula, row, column, names=True)
    except ValueError:
        raise error(1004, "Application-defined or object-defined error") from None
    return spelled_formula(sheet, converted, row, column, whole=whole, at=at, r1c1=True)


def legacy_formula(sheet: Worksheet, formula: str) -> str:
    """A formula written through Range.Formula, spelled, as Excel keeps it: a call of SINGLE an @, an @ left out
    where the formula cuts to one value anyway, and where one is kept every @ Formula2 shows written out
    (see pyopenvba.formula._formula2.written)."""
    from pyopenvba.formula._formula2 import UnreadFormula2Error, written
    from pyopenvba.formula._parse import FormulaError

    try:
        return written(formula, lambda name: defined_formula(sheet, name))
    except UnreadFormula2Error as exc:
        raise VBAUnsupportedError(str(exc)) from None
    except FormulaError:
        raise VBAUnsupportedError(f"Range.Formula given {formula} is not implemented") from None


def defined_formula(sheet: Worksheet, name: str) -> str | None:
    """A defined name's formula as a formula on ``sheet`` finds the name, or None where it finds none."""
    found = sheet.book.names_.find(name, scope=sheet)
    return None if found is None else found.entry.refers_to


def _array_formula_text(sheet: Worksheet, formula: str, row: int, column: int) -> str:
    """What FormulaArray is given, as Excel spells it: A1, or R1C1 where A1 cannot read it, =SUM(R1C1:R3C1)."""
    try:
        return written_formula(sheet, formula, row, column, whole=True)
    except VBARuntimeError:
        raise error(1004, "Unable to set the FormulaArray property of the Range class") from None


def shown_formula(sheet: Worksheet, row: int, column: int, formula: str) -> str:
    """A cell's formula as Range.Formula shows it: another workbook by its path (see _external_links), and in a
    table, without the table's name where Excel leaves it out (see pyopenvba.formula._structured)."""
    if "[" not in formula:
        return formula
    from pyopenvba.apps.excel._external_links import shown as linked_shown

    formula = linked_shown(sheet.book, formula)
    if not sheet.tables:
        return formula
    from pyopenvba.apps.excel._tables import table_at
    from pyopenvba.formula._structured import shown

    table = table_at(sheet, row, column)
    return formula if table is None else shown(formula, table.name)


def spelled_formula(sheet: Worksheet, formula: str, row: int = 0, column: int = 0, *, whole: bool = False,
                    at: bool = False, r1c1: bool = False) -> str:
    """A formula a macro writes to ``sheet``, at the cell in ``row`` and ``column`` where it has one, as Excel
    spells it back; error 1004 where Excel refuses it. ``whole`` for an array formula, ``at`` for one written
    through Formula2, which may hold an @, ``r1c1`` for one read in R1C1 and turned into A1."""
    from pyopenvba.formula._parse import FormulaError
    from pyopenvba.formula._spell import UnmodelledFormulaError, spelled

    try:
        return spelled(formula, _BookNames(sheet, row, column), whole=whole, at=at, r1c1=r1c1)
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
