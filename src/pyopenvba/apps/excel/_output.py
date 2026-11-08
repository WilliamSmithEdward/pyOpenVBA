"""ExportAsFixedFormat and PrintOut: what they answer themselves, measured with no viewer and no paper (scripts/measure_output_members.py, tests/fixtures/output_members.json), and then
the function set with ExcelApplication.set_hook (see _hooks). The model makes no PDF, XPS or print of its own.

- A bad argument is refused before BeforePrint, and the file is looked at after it.
- ExportAsFixedFormat: Type is 0 or 1, and text that is not a number 13, Null 94 where it was bound to the parameter and 13 otherwise. Quality is 0 or 1. Anything that is not a number or
  a flag where one is wanted is 5. Cancelled in BeforePrint, it is 1004. It answers Empty.
- A name left out, empty or Null is the workbook's, in the current folder: not the workbook's folder, nor the default file path. A relative one is in the current folder too. Any
  name but one ending in the type's extension has it added, so x.xps as a PDF is x.xps.pdf. A number or True is its text.
- A folder not there or a character a file cannot hold is 1004, a read-only file -2147018887, and a name past what a path can hold 5.
- PrintOut: a bad argument is 1004, Null passes as if left out, and any printer name passes. It answers True; cancelled in BeforePrint, it goes no further, raises nothing and answers False
  (tests/fixtures/print_returns.json).
- Window.PrintOut is always 1004, so it never reaches the function.
- Chart.Export, on a chart sheet or a chart in a frame (tests/fixtures/output_pictures.json): the name is the file's, in the current folder if relative, with
  nothing added. FilterName, or else the name's extension, picks the kind: PNG, GIF, JPG, JPEG and BMP, in any case. Any other, or none, and Excel writes an
  empty file and answers True, as the model does without the function. Null as the name is 94 typed and 13 late, as the filter -2147352571; an empty name and a
  folder not there -2147287037, a folder or a read-only file -2147287035, a character a file cannot hold -2147286917, a path too long -2147024809, each with
  Excel's text about a chart type. It answers True. A fresh automated Excel exports a chart in a frame as an empty file until some chart object has been
  activated or selected in its session (tests/fixtures/frame_drawn.json), a state of its drawing, not of the chart, which the model does not copy: the
  function is called for a chart in a frame as for a chart sheet.
- CopyPicture on a Range, a Shape, a chart sheet or one in a frame, and a ChartObject: each class checks Appearance and Format its own way, below. A Range of
  several areas is 1004. A hidden shape copies nothing. Range answers True bound late, the rest Empty. CutCopyMode reads 0 after, and what a paste of the
  picture makes is not modelled.
- A chart in a frame prints as a chart sheet does. It exports nothing: past its argument errors, From or To below 1, a name too long and a read-only file, -2147417848 bound late, shown or hidden and active or not, and with no name -2147024773, and no file. Bound
  typed, Excel's macro stops there, which is not modelled: it answers as bound late (tests/fixtures/output_embedded_chart.json).

Chosen, not measured: From or To below 1 going on to the file check's stage rather than before the event, as it reads as a failure to write; and the length a path can hold, the 259
characters of Windows. Measured only on one printer: what a print to a file writes, which a spooler writes late, and From past the last page, as the model has no pages.
PrToFileName empty, Null or left out puts up a form in Excel, which is not measured: it takes the next answer queued with answer_dialogs, and reports itself with none.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _events
from pyopenvba.apps.excel._hooks import Hook, call, gate, hook_of
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.interpreter._objects import BINDING
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, error, to_bool, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet

_NOT_SAVED = 'Document not saved. The document may be open, or an error may have been encountered when saving.'
_NOTHING = "We didn't find anything to print."
_INVALID = 'Invalid procedure call or argument'
_READ_ONLY = -2147018887
#: What a path can hold on Windows, whose answer this is.
_LONGEST_PATH = 259
_BAD_CHARACTERS = frozenset('?*<>|"')
_EXTENSIONS = {0: '.pdf', 1: '.xps'}
#: The typed names of what PrintOut fails on.
_INTERFACES = {'Worksheet': '_Worksheet', 'Workbook': '_Workbook', 'Chart': '_Chart', 'Sheets': 'Sheets'}


def _number(value: object) -> float | None:
    if value is NULL or value is MISSING:
        return None
    try:
        return float(to_number(value))
    except (VBAError, ValueError, TypeError, OverflowError):
        return None


def _fixed_format_type(value: object, owner: str) -> int:
    if value is MISSING:
        raise error(450, 'Wrong number of arguments or invalid property assignment')
    if value is NULL:
        raise error(94 if owner != 'Range' and not BINDING.late else 13)
    found = _number(value)
    if found is None:
        raise error(13)
    # Rounded as a Long is, so 0.5 is 0.
    kind = int(round(found))
    if kind not in _EXTENSIONS or isinstance(value, bool):
        raise error(5, _INVALID)
    return kind


def _export_arguments(owner: str, given: dict[str, object]) -> int:
    kind = _fixed_format_type(given.get('Type', MISSING), owner)
    if 'FixedFormatExtClassPtr' in given or 'WorkIdentity' in given:
        raise error(5, _INVALID)
    if 'Quality' in given and _number(given['Quality']) not in (0.0, 1.0):
        raise error(5, _INVALID)
    for name in ('IncludeDocProperties', 'IgnorePrintAreas', 'From', 'To'):
        if name in given and _number(given[name]) is None:
            raise error(5, _INVALID)
    return kind


def given_arguments(**arguments: object) -> dict[str, object]:
    return {name: value for name, value in arguments.items() if value is not MISSING}


def _export_path(book: Workbook, kind: int, given: dict[str, object]) -> Path:
    interpreter = book.application.interpreter
    assert interpreter is not None
    name = given.get('Filename', '')
    text = '' if name is NULL or name is MISSING else str(name).upper() if isinstance(name, bool) else to_text(name)
    extension = _EXTENSIONS[kind]
    if not text:
        text = os.path.splitext(book.name)[0]
    if not text.lower().endswith(extension):
        text += extension
    return Path(os.path.normpath(interpreter.files.full(text)))


def _check_writable(path: Path, *, failed: tuple[int, str], read_only: tuple[int, str]) -> None:
    """After the gate, so a refused path says nothing of what is there."""
    if _BAD_CHARACTERS & set(path.name) or not path.parent.is_dir() or path.is_dir():
        raise error(*failed)
    if path.exists() and not os.access(path, os.W_OK):
        raise error(*read_only)


def _refuse_frame(book: Workbook, kind: int, given: dict[str, object]) -> None:
    """Excel makes no file of a chart in a frame (tests/fixtures/output_embedded_chart.json), so there is nothing for a function to stand in for. What it does look at it answers for first."""
    if any(name in given and (_number(given[name]) or 0) < 1 for name in ('From', 'To')):
        raise error(1004, _NOT_SAVED)
    name = given.get('Filename', '')
    if name is NULL or name == '' or name is MISSING:
        raise error(-2147024773, 'Document not saved.')
    book.application.outside('ExportAsFixedFormat writing a file')
    path = _export_path(book, kind, given)
    if len(str(path)) > _LONGEST_PATH:
        raise error(5, _INVALID)
    if os.access(path, os.F_OK) and not os.access(path, os.W_OK):
        raise error(_READ_ONLY, 'Document not saved.')
    raise error(-2147417848, 'Automation error\nThe object invoked has disconnected from its clients.')


def export_fixed_format(target: object, owner: str, book: Workbook, sheet: Worksheet | None, given: dict[str, object], *, in_frame: bool = False) -> object:
    application = book.application
    kind = _export_arguments(owner, given)
    if in_frame:
        _refuse_frame(book, kind, given)
    if sheet is not None and owner == 'Worksheet' and sheet.visible != -1:
        raise error(5, _INVALID)
    hook = hook_of(application, 'ExportAsFixedFormat')
    if hook is None:
        raise VBAUnsupportedError(f'{owner}.ExportAsFixedFormat needs a function that makes the file, which ExcelApplication.set_hook registers; the model makes no PDF of its own')
    if not _events.before_print(book):
        raise error(1004, _NOT_SAVED)
    if any(name in given and (_number(given[name]) or 0) < 1 for name in ('From', 'To')):
        raise error(1004, _NOT_SAVED)
    if sheet is not None and owner == 'Worksheet' and not sheet.cells_ and not sheet.shapes_:
        raise error(1004, _NOTHING)
    path = _export_path(book, kind, given)
    gate(application, hook, 'ExportAsFixedFormat', path)
    if len(str(path)) > _LONGEST_PATH:
        raise error(5, _INVALID)
    _check_writable(path, failed=(1004, _NOT_SAVED), read_only=(_READ_ONLY, 'Document not saved.'))
    call(application, hook, 'ExportAsFixedFormat', owner, target, book, given, path)
    return EMPTY


_EXPORT_TEXT = 'The specified dimension is not valid for the current chart type.'
_FILTERS = frozenset({'PNG', 'GIF', 'JPG', 'JPEG', 'BMP'})


def _export_chart_path(book: Workbook, given: dict[str, object]) -> Path:
    interpreter = book.application.interpreter
    assert interpreter is not None
    name = given.get('Filename', MISSING)
    if name is MISSING:
        raise error(450, 'Wrong number of arguments or invalid property assignment')
    if name is NULL:
        raise error(13 if BINDING.late else 94)
    if given.get('FilterName', MISSING) is NULL:
        raise error(-2147352571, _EXPORT_TEXT)
    text = str(name).upper() if isinstance(name, bool) else to_text(name)
    if not text:
        raise error(-2147287037, _EXPORT_TEXT)
    return Path(os.path.normpath(interpreter.files.full(text)))


def _check_export_path(path: Path) -> None:
    """After the gate, so a refused path says nothing of what is there."""
    if len(str(path)) > _LONGEST_PATH:
        raise error(-2147024809, _EXPORT_TEXT)
    if _BAD_CHARACTERS & set(path.name):
        raise error(-2147286917, _EXPORT_TEXT)
    if not path.parent.is_dir():
        raise error(-2147287037, _EXPORT_TEXT)
    if path.is_dir() or (path.exists() and not os.access(path, os.W_OK)):
        raise error(-2147287035, _EXPORT_TEXT)


def export_chart(target: object, book: Workbook, given: dict[str, object]) -> object:
    """Chart.Export, as the module's docstring has it: a kind Excel draws goes to the function, any other is the empty file Excel writes."""
    application = book.application
    path = _export_chart_path(book, given)
    filter_name = given.get('FilterName', MISSING)
    kind = (to_text(filter_name) if filter_name is not MISSING else path.suffix.removeprefix('.')).upper()
    hook = hook_of(application, 'Export')
    if kind in _FILTERS:
        if hook is None:
            raise VBAUnsupportedError('Chart.Export needs a function that draws the chart, which ExcelApplication.set_hook registers; the model draws no picture of its own')
        gate(application, hook, 'Export', path)
        _check_export_path(path)
        call(application, hook, 'Export', 'Chart', target, book, given, path)
        return True
    if hook is not None:
        gate(application, hook, 'Export', path)
    else:
        application.outside('Chart.Export writing a file')
    _check_export_path(path)
    path.write_bytes(b'')
    return True


#: Appearance xlScreen and xlPrinter; Format xlPicture and xlBitmap.
_APPEARANCES = (1.0, 2.0)
_FORMATS = (-4147.0, 2.0)
#: What a class says of a value it does not take.
_PICTURE_FAILED = {'Range': (1004, 'CopyPicture method of Range class failed'), 'Shape': (1004, 'Application-defined or object-defined error'),
                   'ChartObject': (5, _INVALID)}


def _picture_value(owner: str, value: object, allowed: tuple[float, ...]) -> float | None:
    """One of CopyPicture's arguments as ``owner`` takes it, None for its default. A chart takes any number; the rest only the two each argument
    has, and refuse the rest as _PICTURE_FAILED says. Null and True, and text that is no number, differ by class and binding."""
    if value is MISSING:
        return None
    failed = _PICTURE_FAILED.get(owner)
    if value is NULL:
        if owner == 'Shape' or (owner == 'Range' and BINDING.late):
            return None
        raise error(13 if BINDING.late else 94)
    if isinstance(value, bool):
        if owner == 'Chart' or (owner == 'Range' and BINDING.late):
            return None
        assert failed is not None
        raise error(*failed)
    found = _number(value)
    if found is None:
        if failed is not None and (owner == 'Shape' or (owner == 'Range' and BINDING.late)):
            raise error(*failed)
        raise error(13)
    if failed is not None and found not in allowed:
        raise error(*failed)
    return found


def copy_picture(target: object, owner: str, book: Workbook, given: dict[str, object], *, shown: bool = True) -> object:
    """CopyPicture, as the module's docstring has it. The function set for it makes the picture; the model keeps only that the clipboard holds one."""
    application = book.application
    appearance = _picture_value(owner, given.get('Appearance', MISSING), _APPEARANCES)
    picture_format = _picture_value(owner, given.get('Format', MISSING), _FORMATS)
    if 'Size' in given:
        _picture_value(owner, given['Size'], _APPEARANCES)
    if owner in _PICTURE_FAILED and appearance == 2.0 and picture_format == 2.0:
        # A bitmap as printed is one only a chart makes.
        raise error(*_PICTURE_FAILED[owner])
    hook = hook_of(application, 'CopyPicture')
    if hook is None:
        raise VBAUnsupportedError(f'{owner}.CopyPicture needs a function that draws the picture, which ExcelApplication.set_hook registers; the model draws no picture of its own')
    application.clipboard = None
    application.copied_chart = None
    application.copied_picture = None
    if shown:
        call(application, hook, 'CopyPicture', owner, target, book, given)
        application.copied_picture = (application.clipboard,)
    return True if owner == 'Range' and BINDING.late else EMPTY


def _print_failed(owner: str) -> tuple[int, str]:
    interface = _INTERFACES.get(owner)
    if interface is None or BINDING.late:
        return 1004, f'PrintOut method of {owner} class failed'
    return 1004, f"Method 'PrintOut' of object '{interface}' failed"


def _print_arguments(owner: str, given: dict[str, object]) -> None:
    for name in ('From', 'To', 'Copies'):
        value = given.get(name, NULL)
        if value is NULL:
            continue
        found = _number(value)
        if found is None or found < 1:
            raise error(*_print_failed(owner))
    for name in ('Collate', 'IgnorePrintAreas'):
        if name in given and given[name] is not NULL and _number(given[name]) is None:
            raise error(*_print_failed(owner))


def _print_path(book: Workbook, hook: Hook, owner: str, given: dict[str, object]) -> Path | None:
    if 'PrintToFile' not in given or not to_bool(given['PrintToFile']):
        return None
    application = book.application
    interpreter = application.interpreter
    assert interpreter is not None
    name = given.get('PrToFileName', '')
    text = '' if name is NULL or name is MISSING else to_text(name)
    if not text:
        # Excel puts up a form for the name here, which nobody has measured an answer for.
        answer = interpreter.next_answer('PrintOut', 'Output file name')
        if not isinstance(answer, str) or not answer:
            raise VBAUnsupportedError('PrintOut with no file name, its form cancelled, is not implemented')
        text = answer
    path = Path(os.path.normpath(interpreter.files.full(text)))
    gate(application, hook, 'PrintOut', path)
    failed = _print_failed(owner)
    _check_writable(path, failed=failed, read_only=failed)
    return path


def print_out(target: object, owner: str, book: Workbook, sheet: Worksheet | None, given: dict[str, object]) -> object:
    """A print: all a macro sees of it is BeforePrint, and the function set for it hears of it. With none, it goes nowhere. A sheet not active in its workbook is active for the event and the
    function, and made active again without one, and a hidden sheet is refused (tests/fixtures/more_events.json print_*)."""
    application = book.application
    hook = hook_of(application, 'PrintOut')
    _print_arguments(owner, given)
    if hook is None:
        if 'Preview' in given and to_bool(given['Preview']):
            raise VBAUnsupportedError('PrintOut with Preview, which waits for someone at the screen, is not implemented')
        if 'ActivePrinter' in given:
            raise VBAUnsupportedError('PrintOut naming a printer is not implemented')
    if sheet is not None and sheet.visible != -1:  # xlSheetVisible
        raise error(*_print_failed(owner))
    active = book.active_sheet
    moving = sheet is not None and active is not sheet
    if sheet is not None and moving:
        if application.active_book is not book:
            raise VBAUnsupportedError('PrintOut of a sheet not active in a workbook behind is not implemented')
        sheet.Activate()
    try:
        if not _events.before_print(book):
            return False
        if hook is None:
            if 'PrintToFile' in given and to_bool(given['PrintToFile']):
                raise VBAUnsupportedError('PrintOut to a file, which a printer driver writes, is not implemented')
            return True
        path = _print_path(book, hook, owner, given)
        call(application, hook, 'PrintOut', owner, target, book, given, path)
        return True
    finally:
        if moving and active is not None:
            book.active_sheet_index = book.sheets_.index(active)

