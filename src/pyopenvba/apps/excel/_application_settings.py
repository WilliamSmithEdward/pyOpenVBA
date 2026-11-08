"""Application settings everyday macros switch, and Worksheet.DisplayPageBreaks: what each reads as Excel starts,
and what each setting a macro gives it reads back as.

Measured in live Excel (scripts/measure_application_settings.py, tests/fixtures/application_settings.json):

- The switches take any value VBA turns into a Boolean, Empty as False;
  Null is 94 and text that is no Boolean 13, as for every setting here,
  VBA converting first as these are bound early. DisplayFullScreen and FixedDecimal start False, UserControl False
  as in Excel a program started, AskToUpdateLinks True as Excel ships it
  (the measuring harness turns it off), and the rest of the first fifteen
  True. The 31 since (AutoPercentEntry, ExtendList, EnableSound and the
  rest) start as the measuring machine's Excel had them. FlashFillMode
  stays False whatever is set, and ConstrainNumeric refuses every value
  with 1004. ShowMenuFloaties and
  ShowSelectionFloaties are left out: their setters are turned round, and
  measuring them turned the measuring user's mini toolbar off.
- Cursor, EnableCancelKey, MoveAfterReturnDirection, DisplayCommentIndicator
  and FixedDecimalPlaces take a number rounded as VBA rounds to a Long,
  Empty as 0. A Cursor of 0 is xlDefault and one Excel does not name is
  1004. An
  EnableCancelKey Excel does not name becomes xlInterrupt from xlDisabled
  and xlDisabled from anything else. MoveAfterReturnDirection takes 1 to
  4 as down, right, up and left, and reads any other number back as it
  was set, 65536 added where it is not positive. A DisplayCommentIndicator
  Excel does not name changes nothing. FixedDecimalPlaces takes -300 to
  300, and anything else is 1004.
- MergeInstances (True) and HighQualityModeForGraphics (False) are
  switches like the rest. GenerateTableRefs takes 0 and 1, and anything
  else is 5; CursorMovement takes 1 and 2 and turns any other
  number into 1; DefaultSheetDirection takes xlRTL and xlLTR, takes
  xlContext and 0 as xlLTR, and keeps anything else out;
  LargeOperationCellThousandCount takes 1 and up and ODBCTimeout 0 and
  up, and anything else is 5. FormulaBarHeight is left out: how tall it may be follows the
  window, which the model does not have.
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

import re
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._objects import BINDING, MemberSpec
from pyopenvba.interpreter._values import MISSING, NULL, VBAInt, error, to_bool, to_integer, to_text

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
            'WarnOnFunctionNameConflict': True, 'MergeInstances': True, 'HighQualityModeForGraphics': False}
#: A switch that stays as it is whatever a macro sets, with no error: Flash Fill's FlashFillMode reads False after
#: every setting (tests/fixtures/application_settings.json).
FIXED = frozenset({'FlashFillMode'})
#: A switch that refuses every setting: pen input's ConstrainNumeric.
REFUSING = frozenset({'ConstrainNumeric'})
CURSORS = (1, 2, 3, -4143)
#: MoveAfterReturnDirection's directions, by the place Excel also takes for each: down, right, up and left.
DIRECTIONS = (-4121, -4161, -4162, -4159)
INDICATORS = (0, 1, -1)
#: LargeOperationCellThousandCount and ODBCTimeout start as the measuring machine's Excel had them.
STARTS: dict[str, int] = {'Cursor': -4143, 'EnableCancelKey': 1, 'MoveAfterReturnDirection': -4121,
                          'DisplayCommentIndicator': -1, 'FixedDecimalPlaces': 2, 'GenerateTableRefs': 1,
                          'CursorMovement': 1, 'DefaultSheetDirection': -5003,
                          'LargeOperationCellThousandCount': 33554, 'ODBCTimeout': 45}


def flag(value: object) -> bool:
    """A switch's value. Bound early VBA converts it first, so text is 13 and Null 94 (tests/fixtures/application_settings.json)."""
    return bool(to_bool(value))


def number(value: object) -> int:
    """A setting's value, rounded as VBA rounds to a Long; text is 13 and Null 94 (tests/fixtures/application_settings.json)."""
    return int(to_integer(value, 'Long'))


def _failed(name: str) -> VBARuntimeError:
    """The 1004 a setting Excel refuses raises."""
    return error(1004, f"Method '{name}' of object '_Application' failed")


def _check_number(name: str, given: int) -> None:
    """Error 1004 or 5 where Excel refuses ``given`` for this setting (tests/fixtures/application_settings.json)."""
    if name == 'Cursor' and given != 0 and given not in CURSORS:
        raise _failed(name)
    if name == 'FixedDecimalPlaces' and not -300 <= given <= 300:
        raise _failed(name)
    if (name == 'GenerateTableRefs' and given not in (0, 1)) or (name == 'ODBCTimeout' and given < 0) \
            or (name == 'LargeOperationCellThousandCount' and given < 1):
        raise error(5)


def _numbered(name: str, current: int, given: int) -> int:
    """What a numbered setting holds once ``given``, which Excel takes, is set over ``current``."""
    if name == 'Cursor':
        return -4143 if given == 0 else given
    if name == 'EnableCancelKey':
        return given if given in (0, 1, 2) else 1 if current == 0 else 0
    if name == 'MoveAfterReturnDirection':
        return DIRECTIONS[given - 1] if 1 <= given <= 4 else given
    if name == 'DisplayCommentIndicator':
        return given if given in INDICATORS else current
    if name == 'CursorMovement':
        return given if given in (1, 2) else 1
    if name == 'DefaultSheetDirection':
        return given if given in (-5004, -5003) else -5003 if given in (-5002, 0) else current
    return given


def _switch(name: str) -> MemberSpec:
    def get(application: Application) -> object:
        return application.settings.get(name, SWITCHES[name])

    def put(application: Application, value: object) -> None:
        found = flag(value)
        if name in REFUSING:
            raise _failed(name)
        if name not in FIXED:
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
        _check_number(name, given)
        current = application.settings.get(name, STARTS[name])
        assert isinstance(current, int)
        application.settings[name] = _numbered(name, current, given)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _note_indicator() -> MemberSpec:
    def get(application: Application) -> object:
        return application.settings.get('DisplayCommentIndicator', STARTS['DisplayCommentIndicator']) != 0

    def put(application: Application, value: object) -> None:
        application.settings['DisplayCommentIndicator'] = -1 if flag(value) else 0

    return MemberSpec(name='DisplayNoteIndicator', kind='property', getter=get, setter=put)


def _caption() -> MemberSpec:
    def get(application: Application) -> object:
        from pyopenvba.apps.excel._windows import window_of

        own = application.settings.get('Caption') or 'Excel'
        book = application.active_book
        return f'{window_of(book).caption()} - {own}' if book is not None else own

    def put(application: Application, value: object) -> None:
        application.settings['Caption'] = to_text(value)[:200]

    return MemberSpec(name='Caption', kind='property', getter=get, setter=put)


def _print_communication() -> MemberSpec:
    def get(application: Application) -> object:
        return application.print_communication

    def put(application: Application, value: object) -> None:
        from pyopenvba.apps.excel._page_setup import flush

        found = flag(value)
        application.print_communication = found
        if found:
            flush(application)

    return MemberSpec(name='PrintCommunication', kind='property', getter=get, setter=put)


#: Settings read as Excel starts, as the measuring machine's Excel had them, and flags Excel keeps itself; read only
#: in the model. The writable ones' setters are left out: some apply only once Excel starts again, so measuring them
#: would change the measuring user's profile (tests/fixtures/application_settings_more.json).
READ_ONLY: dict[str, object] = {
    'DisplayClipboardWindow': False, 'DisplayDocumentActionTaskPane': False, 'DisplayDocumentInformationPanel': False,
    'DisplayExcel4Menus': False, 'DisplayRecentFiles': True, 'ShowDevTools': True, 'ShowStartupDialog': False,
    'UseClusterConnector': False, 'CommandUnderlines': VBAInt(1, 'Long'), 'DataEntryMode': VBAInt(-4146, 'Long'),
    'DefaultSaveFormat': VBAInt(51, 'Long'),
    'CanPlaySounds': True, 'CanRecordSounds': False, 'MathCoprocessorAvailable': True, 'MouseAvailable': True,
    'Ready': True, 'ArbitraryXMLSupportAvailable': True, 'IsSandboxed': False, 'RecordRelative': False,
    # Names a fresh install gives no value, and the product code of the Office the measurements ran on
    # (tests/fixtures/application_identity.json).
    'NetworkTemplatesPath': '', 'OrganizationName': '', 'ProductCode': '{90160000-000F-0000-1000-0000000FF1CE}',
    # Folders in a user's profile, chosen as Path is: a real one names the machine and its user.
    'StartupPath': '<PROFILE>\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART',
    'TemplatesPath': '<PROFILE>\\AppData\\Roaming\\Microsoft\\Templates\\',
    'UserLibraryPath': '<PROFILE>\\AppData\\Roaming\\Microsoft\\AddIns\\',
}
#: The profile the chosen folders above sit in.
PROFILE = 'C:\\Users\\User'
READ_ONLY.update({name: value.replace('<PROFILE>', PROFILE) for name, value in READ_ONLY.items() if isinstance(value, str)})
#: The printer ActivePrinter names: one every Windows install has. The model prints to no printer, so a set of any is refused.
PRINTER = 'Microsoft Print to PDF on Ne00:'


def _read_only(name: str, value: object) -> MemberSpec:
    def get(application: Application) -> object:
        return value

    return MemberSpec(name=name, kind='property', getter=get)


def _feature_install() -> MemberSpec:
    """msoFeatureInstallNone to msoFeatureInstallOnDemandWithUI, rounded as VBA rounds; 1004 beyond, True as -1 among them, and
    Null 13 bound late (tests/fixtures/application_identity.json)."""

    def get(application: Application) -> object:
        return VBAInt(int(application.settings.get('FeatureInstall', 0)), 'Long')  # type: ignore[arg-type]

    def put(application: Application, value: object) -> None:
        if value is NULL and BINDING.late:
            raise error(13)
        given = number(value)
        if not 0 <= given <= 2:
            raise _failed('FeatureInstall')
        application.settings['FeatureInstall'] = given

    return MemberSpec(name='FeatureInstall', kind='property', getter=get, setter=put)


def _active_printer() -> MemberSpec:
    """The printer; a set names one the model does not have, which is 1004 in Excel for any name it lacks. Null is ignored
    bound late (tests/fixtures/application_identity.json)."""

    def get(application: Application) -> object:
        return PRINTER

    def put(application: Application, value: object) -> None:
        if value is NULL:
            if BINDING.late:
                return
            raise error(94)
        raise error(1004, "Method 'ActivePrinter' of object '_Application' failed")

    return MemberSpec(name='ActivePrinter', kind='property', getter=get, setter=put)


#: The folder Excel was started from, spelled as the automation server's registration names it, which a bare name in
#: AltStartupPath goes under (tests/fixtures/application_folders.json).
EXCEL_FOLDER = 'C:\\Program Files\\Microsoft Office\\Root\\Office16'


def _folder_text(value: object) -> str | None:
    """The text a folder setting takes, or None for a Null Excel ignores bound late; bound early it is 94."""
    if value is NULL:
        if BINDING.late:
            return None
        raise error(94)
    return to_text(value)


def _default_file_path() -> MemberSpec:
    """The folder a macro sets, kept only where it is one that is there: a full path as Windows spells it in full, slashes,
    dots and a trailing space resolved and a trailing backslash kept, and a relative one as typed. Anything else changes
    nothing, a drive with no folder too; empty text empties it (tests/fixtures/application_folders.json). A check that a
    folder is there reaches the machine, so the sandbox refuses any set but an empty one."""

    def get(application: Application) -> object:
        return application.settings.get('DefaultFilePath', PROFILE + '\\Documents')

    def put(application: Application, value: object) -> None:
        text = _folder_text(value)
        if text is None:
            return
        if text:
            application.outside('DefaultFilePath')
            text = _kept_folder(application, text)
            if text is None:
                return
        application.settings['DefaultFilePath'] = text

    return MemberSpec(name='DefaultFilePath', kind='property', getter=get, setter=put)


def _kept_folder(application: Application, text: str) -> str | None:
    """A folder named by text as DefaultFilePath keeps it, or None where none is there."""
    import os

    if re.fullmatch(r'[A-Za-z]:', text) or not text.strip():
        return None
    folder = application.interpreter.files.folder if application.interpreter is not None else os.getcwd()
    if not os.path.isdir(os.path.join(folder, text)):
        return None
    if not os.path.isabs(text):
        return text
    kept = os.path.normpath(text.rstrip(' '))
    return kept + os.sep if re.search(r'[\\/]$', text) and not kept.endswith(os.sep) else kept


def _alt_startup_path() -> MemberSpec:
    """Any text, in capitals and with one trailing backslash taken off: a bare name goes under the folder Excel was started
    from and a name from the root onto its drive, with no check that the folder is there (tests/fixtures/application_folders.json)."""

    def get(application: Application) -> object:
        return application.settings.get('AltStartupPath', '')

    def put(application: Application, value: object) -> None:
        text = _folder_text(value)
        if text is None:
            return
        if text:
            text = text.upper()
            if text.startswith('\\') and not text.startswith('\\\\'):
                text = EXCEL_FOLDER[:2] + text
            elif not re.match(r'[A-Z]:|\\\\', text):
                text = EXCEL_FOLDER + '\\' + text
            text = text.removesuffix('\\')
        application.settings['AltStartupPath'] = text

    return MemberSpec(name='AltStartupPath', kind='property', getter=get, setter=put)


def _sharepoint_version() -> MemberSpec:
    """The SharePoint version a site at a URL runs: 0 for one it cannot reach, -1 for text that is no URL. The model reaches
    no site (tests/fixtures/application_identity.json)."""

    def call(application: Application, bstrUrl: object = MISSING) -> object:  # noqa: N803 - the parameter's name in Excel
        if bstrUrl is MISSING:
            raise error(449, 'Argument not optional')
        return VBAInt(0 if re.match(r'https?://', to_text(bstrUrl), re.IGNORECASE) else -1, 'Long')

    return MemberSpec(name='SharePointVersion', kind='method', getter=call, parameters=('bstrUrl',))


def application_members() -> list[MemberSpec]:
    """Every Application setting this module keeps, as a member."""
    return [*(_switch(name) for name in SWITCHES), *(_numbered_setting(name) for name in STARTS), _note_indicator(),
            _caption(), _print_communication(), *(_read_only(name, value) for name, value in READ_ONLY.items()),
            _feature_install(), _active_printer(), _sharepoint_version(), _default_file_path(), _alt_startup_path()]


def worksheet_members() -> list[MemberSpec]:
    """Worksheet.DisplayPageBreaks, which reads back what a macro set and is not saved."""

    def get(sheet: Worksheet) -> object:
        return sheet.display_page_breaks

    def put(sheet: Worksheet, value: object) -> None:
        # Bound early VBA converts first, so text is 13 and Null 94 (tests/fixtures/settable_audit.json).
        sheet.display_page_breaks = bool(to_bool(value))

    return [MemberSpec(name='DisplayPageBreaks', kind='property', getter=get, setter=put)]
