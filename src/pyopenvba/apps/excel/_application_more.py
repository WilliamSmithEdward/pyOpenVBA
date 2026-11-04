"""Members of Application a macro meets with no screen and no one at it, measured in live Excel
(scripts/measure_application_undo.py, tests/fixtures/application_undo.json).

Nothing is left to undo or repeat when a macro runs, so Undo and Repeat raise 1004. OnUndo and OnRepeat name what the
menu would offer, and OnKey a key's macro, and with no menu or keyboard here they are checked and held nowhere."""
from __future__ import annotations

import datetime as _dt
import re
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._creator import Created
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member
from pyopenvba.interpreter._values import (EMPTY, MISSING, NULL, VBADate, VBAErrorValue, VBAInt, error, to_bool,
                                           to_date, to_text)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application

_KEY_NAMES = frozenset({
    'BACKSPACE', 'BS', 'BREAK', 'CAPSLOCK', 'CLEAR', 'DELETE', 'DEL', 'DOWN', 'END', 'ENTER', 'ESCAPE', 'ESC', 'HELP',
    'HOME', 'INSERT', 'INS', 'LEFT', 'NUMLOCK', 'PGDN', 'PGUP', 'RETURN', 'RIGHT', 'SCROLLLOCK', 'TAB', 'UP', '~',
    *(f'F{number}' for number in range(1, 16))})


#: How many procedures one idle moment runs before the model calls it a loop.
TIMERS_AT_ONCE = 10000
NUMBER = re.compile(r"[+-]?\d+(\.\d*)?")


def required(*values: object) -> None:
    """A required argument left out: 449."""
    if any(value is MISSING for value in values):
        raise error(449, 'Argument not optional')


def checked_key(key: object) -> None:
    """A key as OnKey's Key spells it: shift, control and alt signs, then one character or a name in braces."""
    required(key)
    found = re.fullmatch(r'[+^%]*(?:\{(.+)\}|([^{}]))', to_text(key))
    if found is None or (found.group(1) is not None and found.group(1).upper() not in _KEY_NAMES):
        raise error(1004, 'OnKey method of Application class failed')


class EmptyCollection(Created):
    """A collection the model never holds anything in."""

    def __init__(self, type_name: str) -> None:
        self.vba_type_name = type_name

    @member
    def Count(self) -> object:
        return VBAInt(0, 'Long')


#: What the objects behind Application's options read as the measuring machine's Excel had them
#: (tests/fixtures/application_objects.json). Nothing here was set: setting one changes a user's profile, so a setter
#: reports itself as a member the model does not have.
PROFILE: dict[str, dict[str, object]] = {
    'AutoRecover': {'Enabled': False, 'Time': VBAInt(10, 'Long')},
    'MultiThreadedCalculation': {'Enabled': True},
    'SpellingOptions': {'ArabicModes': VBAInt(2, 'Long'), 'IgnoreCaps': True, 'IgnoreMixedDigits': True,
                        'IgnoreFileNames': True, 'SuggestMainOnly': False, 'UserDict': 'RoamingCustom.dic',
                        'DictLang': VBAInt(1033, 'Long')},
    'AutoCorrect': {'CapitalizeNamesOfDays': True, 'CorrectCapsLock': True, 'CorrectSentenceCap': True,
                    'ReplaceText': True, 'TwoInitialCapitals': True, 'AutoExpandListRange': True,
                    'AutoFillFormulasInLists': True, 'DisplayAutoCorrectOptions': True},
    'DefaultWebOptions': {'AllowPNG': True, 'DownloadComponents': False, 'Encoding': VBAInt(1252, 'Long'),
                          'LocationOfComponents': '', 'OrganizeInFolder': True, 'PixelsPerInch': VBAInt(96, 'Long'),
                          'RelyOnCSS': True, 'RelyOnVML': False, 'SaveHiddenData': True,
                          'SaveNewWebPagesAsWebArchives': True, 'ScreenSize': VBAInt(4, 'Long'),
                          'TargetBrowser': VBAInt(4, 'Long'), 'UpdateLinksOnSave': True, 'UseLongFileNames': True,
                          'AlwaysSaveInDefaultEncoding': False, 'CheckIfOfficeIsHTMLEditor': True,
                          'LoadPictures': True},
    'DefaultPivotTableLayoutOptions': {},
}


class HeldNowhere(VBACollection, Created):
    """A list of what a machine has installed or has used, which the model has none of: asking for an item is 9
    (tests/fixtures/application_rest.json)."""

    def __init__(self, application: Application, type_name: str) -> None:
        self.application = application
        self.vba_type_name = type_name

    def vba_items(self) -> list[object]:
        return []

    def vba_lookup(self, index: object, items: list[object]) -> object:
        raise error(9, 'Subscript out of range')

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application


class UsedObjects(HeldNowhere):
    def vba_lookup(self, index: object, items: list[object]) -> object:
        raise VBAUnsupportedError('an item of UsedObjects, which answers from what the session has used, is not '
                                  'implemented')


class RecentFiles(HeldNowhere):
    @member
    def Maximum(self) -> object:
        """As the measuring machine's Excel has it; setting it would change a user's options."""
        return VBAInt(50, 'Long')


class ProfileObject(Created):
    """One of Application's objects of options, read only (see PROFILE)."""

    def __init__(self, application: Application, type_name: str) -> None:
        self.application = application
        self.vba_type_name = type_name

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        for known, value in PROFILE[self.vba_type_name].items():
            if known.casefold() == name.casefold():
                return value
        return super().vba_get(name, args, named)

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application


@dataclass
class Timer:
    """One procedure Application.OnTime has waiting."""

    at: VBADate
    procedure: str
    latest: VBADate | None


def on_time(application: Application, earliest: object, procedure: object, latest: object, schedule: object) -> None:
    """Application.OnTime (tests/fixtures/on_time.json): a time VBA makes no Date of and an empty name are 1004, a Null
    time does nothing. Schedule:=False takes away every procedure of that time and name, in any case, and is 1004 where
    there is none."""
    required(earliest, procedure)
    if earliest is NULL:
        return
    try:
        at = to_date(earliest)
        last = None if latest is MISSING else to_date(latest)
    except VBARuntimeError as failure:
        raise error(1004, "Method 'OnTime' of object '_Application' failed") from failure
    name = to_text(procedure)
    if not name:
        raise error(1004, "Method 'OnTime' of object '_Application' failed")
    if schedule is not MISSING and not to_bool(schedule):
        kept = [one for one in application.timers
                if one.at.serial != at.serial or one.procedure.casefold() != name.casefold()]
        if len(kept) == len(application.timers):
            raise error(1004, "Method 'OnTime' of object '_Application' failed")
        application.timers[:] = kept
        return
    application.timers.append(Timer(at, name, last))
    interpreter = application.executing_projects[-1] if application.executing_projects else application.interpreter
    if interpreter is not None and application.pump_timers not in interpreter.pumped:
        interpreter.pumped.append(application.pump_timers)


def run_due(application: Application) -> None:
    """What is due, the last scheduled first, and what those schedule that is due: Excel runs them whatever their times
    say of the order. One past its LatestTime is dropped."""
    interpreter = application.interpreter
    if interpreter is None:
        return
    ran = 0
    while True:
        # A second is as fine as a Date reads; a time just set may round up to the next one.
        now = interpreter.frozen_now or _dt.datetime.now().replace(microsecond=0) + _dt.timedelta(seconds=1)
        due = [one for one in application.timers if one.at.to_datetime() <= now]
        if not due:
            return
        timer = due[-1]
        application.timers.remove(timer)
        if timer.latest is not None and timer.latest.to_datetime() < now:
            continue
        ran += 1
        if ran > TIMERS_AT_ONCE:
            raise VBAUnsupportedError("procedures that keep scheduling themselves for now would run for ever")
        _run_timer(application, timer.procedure)


def _run_timer(application: Application, text: str) -> None:
    """A name, a name with its workbook, or the name and its arguments in single quotes, which the model reads where
    they are numbers or text."""
    from pyopenvba.apps.excel._projects import run_macro

    args: list[object] = []
    if len(text) > 1 and text[0] == "'" and text[-1] == "'":
        text, _, given = text[1:-1].partition(" ")
        for one in filter(None, (part.strip() for part in given.split(","))):
            if one.startswith('"') and one.endswith('"') and len(one) > 1:
                args.append(one[1:-1])
            elif NUMBER.fullmatch(one):
                args.append(float(one))
            else:
                raise VBAUnsupportedError(f"an OnTime procedure's argument {one} is not implemented")
    run_macro(application, text, args)


_CELLS = re.compile(r"(?s)(?P<sheet>.*)!R(?P<row>\d+)C(?P<column>\d+)(?::R(?P<end_row>\d+)C(?P<end_column>\d+))?")
_BARE = re.compile(r"(?i)R(\[?-?\d*\]?)C(\[?-?\d*\]?)|[A-Z]{1,3}\d+|.*![A-Z$]{1,4}[$\d]+")
_FOLDERED = re.compile(r"(?s)(?P<folder>.*)[\\/]\[(?P<book>[^\]]+)\](?P<sheet>.*)")
_BOOKED = re.compile(r"(?s)\[(?P<book>[^\]]+)\](?P<sheet>.*)")
_REF_ERROR = VBAErrorValue(2023)
_UNREADABLE = 2015
_NAME_ERROR = 2029
_EXCEL4_FUNCTION = re.compile(r"[A-Za-z]+\.[A-Za-z.]+\(")


def execute_excel4(application: Application, text: str) -> object:
    """Application.ExecuteExcel4Macro for what macros use it for, reading a cell or working out an expression
    (tests/fixtures/execute_excel4.json).

    A reference is R1C1 and names its sheet, even if only with "!" for the active one; without, or in A1, it is 1004. It
    answers its last cell, an empty one as 0, numbers as Doubles, and #REF! for a sheet or file that is not there. A
    closed workbook is read from its file. An unknown function is #NAME?, but one spelled as Excel 4's are, GET.CELL
    and its like, reports itself: those work in Excel and the model has none of them. An empty text answers "=", as
    measured."""
    if not text:
        return "="
    if _BARE.fullmatch(text):
        raise error(1004, "Method 'ExecuteExcel4Macro' of object '_Application' failed")
    found = _CELLS.fullmatch(text)
    if found is not None:
        return _read_cell(application, found)
    if "[" in text:
        raise VBAUnsupportedError("ExecuteExcel4Macro over another workbook's cells within an expression is not "
                                  "implemented")
    from pyopenvba.formula._r1c1 import to_a1

    answer = application.evaluate_text(to_a1("=" + text, 1, 1)[1:])
    if isinstance(answer, VBAErrorValue) and answer.number == _UNREADABLE or isinstance(answer, VBACollection):
        raise error(1004, "Method 'ExecuteExcel4Macro' of object '_Application' failed")
    if isinstance(answer, VBAErrorValue) and answer.number == _NAME_ERROR and _EXCEL4_FUNCTION.search(text):
        raise VBAUnsupportedError(f"the Excel 4 function in ExecuteExcel4Macro({text!r}) is not implemented")
    return _as_excel4(answer)


def _as_excel4(value: object) -> object:
    if value is EMPTY:
        return 0.0
    if isinstance(value, VBAObject):
        value = value.vba_get("Value2")
        return 0.0 if value is EMPTY else _as_excel4(value)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return value
    return float(value)


def _read_cell(application: Application, found: re.Match[str]) -> object:
    from pyopenvba.apps.excel._io import load_workbook
    from pyopenvba.apps.excel._model import Range

    name = found["sheet"]
    if len(name) > 1 and name[0] == name[-1] == "'":
        name = name[1:-1].replace("''", "'")
    foldered = _FOLDERED.fullmatch(name)
    booked = _BOOKED.fullmatch(name)
    book = application._require_book()  # pyright: ignore[reportPrivateUsage]
    if not name and book.active_sheet is not None:
        name = book.active_sheet.name
    if foldered is not None:
        name = foldered["sheet"]
        path = Path(foldered["folder"].replace("\\", "/")) / foldered["book"]
        opened = next((one for one in application.workbooks_.books if one.name.casefold() == path.name.casefold()
                       and one.path and Path(one.path) == path.parent), None)
        if opened is not None:
            book = opened
        else:
            application.outside("ExecuteExcel4Macro reading a closed workbook")
            if not path.is_file():
                return _REF_ERROR
            book = load_workbook(application, path)
    elif booked is not None:
        name = booked["sheet"]
        opened = next((one for one in application.workbooks_.books if one.name.casefold() == booked["book"].casefold()),
                      None)
        if opened is None:
            return _REF_ERROR
        book = opened
    sheet = next((one for one in book.sheets_ if one.name.casefold() == name.casefold()), None)
    if sheet is None:
        return _REF_ERROR
    row, column = int(found["end_row"] or found["row"]), int(found["end_column"] or found["column"])
    return _as_excel4(Range(sheet, [Area(row, column, row, column, sheet.name)]))
