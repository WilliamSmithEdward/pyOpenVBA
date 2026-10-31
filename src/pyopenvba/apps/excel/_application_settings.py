"""Application settings everyday macros switch, and Worksheet.DisplayPageBreaks: what each reads as Excel starts,
and what each setting a macro gives it reads back as.

Measured in live Excel (scripts/measure_application_settings.py, tests/fixtures/application_settings.json):

- The switches take any value VBA turns into a Boolean, Empty as False.
  Null and text that is no Boolean leave a switch as it was, with no
  error. DisplayFullScreen and FixedDecimal start False, UserControl False
  as in Excel a program started, AskToUpdateLinks True as Excel ships it
  (the measuring harness turns it off), and the rest of the first fifteen
  True. The 31 since (AutoPercentEntry, ExtendList, EnableSound and the
  rest) start as the measuring machine's Excel had them. ConstrainNumeric
  and FlashFillMode stay False whatever is set. ShowMenuFloaties and
  ShowSelectionFloaties are left out: their setters are turned round, and
  measuring them turned the measuring user's mini toolbar off.
- Cursor, EnableCancelKey, MoveAfterReturnDirection, DisplayCommentIndicator
  and FixedDecimalPlaces take a number rounded as VBA rounds to a Long,
  Empty as 0, Null and text that is no number changing nothing. A Cursor
  of 0 is xlDefault and one Excel does not name changes nothing. An
  EnableCancelKey Excel does not name becomes xlInterrupt from xlDisabled
  and xlDisabled from anything else. MoveAfterReturnDirection takes 1 to
  4 as down, right, up and left, and reads any other number back as it
  was set, 65536 added where it is not positive. A DisplayCommentIndicator
  Excel does not name changes nothing. FixedDecimalPlaces takes -300 to
  300.
- DisplayNoteIndicator reads whether DisplayCommentIndicator shows
  anything; True sets it to show indicators only, False to show none.
- Caption reads after the active window's caption and a dash, kept to its
  first 200 characters; an empty one is Excel's own, Excel.
- DisplayPageBreaks starts False and is not saved.

The model draws nothing and takes no keystrokes, so none of these but
PrintCommunication (see _page_setup) changes what it does; a macro reads
back what it set. FixedDecimal changes no value a macro writes, measured.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._objects import MemberSpec
from pyopenvba.interpreter._values import NULL, VBAInt, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Worksheet

#: Each switch and the value it has as Excel starts.
SWITCHES = {'DisplayStatusBar': True, 'DisplayFormulaBar': True, 'DisplayScrollBars': True,
            'DisplayFullScreen': False, 'EnableAnimations': True, 'Interactive': True, 'AskToUpdateLinks': True,
            'AlertBeforeOverwriting': True, 'CellDragAndDrop': True, 'CopyObjectsWithCells': True,
            'EditDirectlyInCell': True, 'EnableAutoComplete': True, 'FixedDecimal': False, 'MoveAfterReturn': True,
            'UserControl': False,
            'AutoFormatAsYouTypeReplaceHyperlinks': True, 'AutoPercentEntry': True, 'ChartDataPointTrack': True,
            'ConstrainNumeric': False, 'ControlCharacters': False, 'DeferAsyncQueries': False,
            'DisplayFormulaAutoComplete': True, 'DisplayFunctionToolTips': True, 'DisplayInsertOptions': True,
            'DisplayPasteOptions': True, 'EnableCheckFileExtensions': True, 'EnableLargeOperationAlert': True,
            'EnableLivePreview': True, 'EnableMacroAnimations': False, 'EnableSound': False, 'ExtendList': True,
            'FlashFill': True, 'FlashFillMode': False, 'FormatStaleValues': True, 'GenerateGetPivotData': True,
            'MapPaperSize': True, 'PivotTableSelection': True, 'PromptForSummaryInfo': False, 'RollZoom': False,
            'ShowChartTipNames': True, 'ShowChartTipValues': True, 'ShowConvertToDataType': True,
            'ShowQuickAnalysis': True, 'ShowToolTips': True, 'TransitionNavigKeys': False,
            'WarnOnFunctionNameConflict': True}
#: Switches that stay as they are whatever a macro sets, with no error: pen input's ConstrainNumeric and Flash
#: Fill's FlashFillMode read False after every setting (tests/fixtures/application_settings.json).
FIXED = frozenset({'ConstrainNumeric', 'FlashFillMode'})
CURSORS = (1, 2, 3, -4143)
#: MoveAfterReturnDirection's directions, by the place Excel also takes for each: down, right, up and left.
DIRECTIONS = (-4121, -4161, -4162, -4159)
INDICATORS = (0, 1, -1)
STARTS: dict[str, int] = {'Cursor': -4143, 'EnableCancelKey': 1, 'MoveAfterReturnDirection': -4121,
                          'DisplayCommentIndicator': -1, 'FixedDecimalPlaces': 2}


def flag(value: object) -> bool | None:
    """A switch's value as Excel takes it, or None where VBA makes no Boolean of it."""
    if value is NULL:
        return None
    try:
        return to_bool(value)
    except VBARuntimeError:
        return None


def number(value: object) -> int | None:
    """A setting's value as Excel takes it, rounded as VBA rounds to a Long, or None where VBA makes no number of
    it."""
    if value is NULL:
        return None
    try:
        return int(to_integer(value, 'Long'))
    except VBARuntimeError:
        return None


def _numbered(name: str, current: int, given: int) -> int:
    """What a numbered setting holds once ``given`` is set over ``current``."""
    if name == 'Cursor':
        return -4143 if given == 0 else given if given in CURSORS else current
    if name == 'EnableCancelKey':
        return given if given in (0, 1, 2) else 1 if current == 0 else 0
    if name == 'MoveAfterReturnDirection':
        return DIRECTIONS[given - 1] if 1 <= given <= 4 else given
    if name == 'DisplayCommentIndicator':
        return given if given in INDICATORS else current
    return given if -300 <= given <= 300 else current


def _switch(name: str) -> MemberSpec:
    def get(application: Application) -> object:
        return application.settings.get(name, SWITCHES[name])

    def put(application: Application, value: object) -> None:
        found = flag(value)
        if found is not None and name not in FIXED:
            application.settings[name] = found

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _numbered_setting(name: str) -> MemberSpec:
    def get(application: Application) -> object:
        held = application.settings.get(name, STARTS[name])
        assert isinstance(held, int)
        if name == 'MoveAfterReturnDirection' and held not in DIRECTIONS and held <= 0:
            held += 65536
        return VBAInt(held, 'Long')

    def put(application: Application, value: object) -> None:
        given = number(value)
        if given is not None:
            current = application.settings.get(name, STARTS[name])
            assert isinstance(current, int)
            application.settings[name] = _numbered(name, current, given)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _note_indicator() -> MemberSpec:
    def get(application: Application) -> object:
        return application.settings.get('DisplayCommentIndicator', STARTS['DisplayCommentIndicator']) != 0

    def put(application: Application, value: object) -> None:
        found = flag(value)
        if found is not None:
            application.settings['DisplayCommentIndicator'] = -1 if found else 0

    return MemberSpec(name='DisplayNoteIndicator', kind='property', getter=get, setter=put)


def _caption() -> MemberSpec:
    def get(application: Application) -> object:
        from pyopenvba.apps.excel._windows import window_of

        own = application.settings.get('Caption') or 'Excel'
        book = application.active_book
        return f'{window_of(book).caption()} - {own}' if book is not None else own

    def put(application: Application, value: object) -> None:
        if value is not NULL:
            application.settings['Caption'] = to_text(value)[:200]

    return MemberSpec(name='Caption', kind='property', getter=get, setter=put)


def _print_communication() -> MemberSpec:
    def get(application: Application) -> object:
        return application.print_communication

    def put(application: Application, value: object) -> None:
        from pyopenvba.apps.excel._page_setup import flush

        found = flag(value)
        if found is None:
            return
        application.print_communication = found
        if found:
            flush(application)

    return MemberSpec(name='PrintCommunication', kind='property', getter=get, setter=put)


def application_members() -> list[MemberSpec]:
    """Every Application setting this module keeps, as a member."""
    return [*(_switch(name) for name in SWITCHES), *(_numbered_setting(name) for name in STARTS), _note_indicator(),
            _caption(), _print_communication()]


def worksheet_members() -> list[MemberSpec]:
    """Worksheet.DisplayPageBreaks, which reads back what a macro set and is not saved."""

    def get(sheet: Worksheet) -> object:
        return sheet.display_page_breaks

    def put(sheet: Worksheet, value: object) -> None:
        found = flag(value)
        if found is not None:
            sheet.display_page_breaks = found

    return [MemberSpec(name='DisplayPageBreaks', kind='property', getter=get, setter=put)]
