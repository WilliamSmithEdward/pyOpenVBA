"""Excel's dialogs, which take their answers from the caller's queue (ExcelApplication.answer_dialogs), as nobody is at the screen.

Measured in live Excel, its forms answered by window handle (scripts/measure_dialogs.py, tests/fixtures/dialogs.json):

- Application.InputBox reads what was typed by its Type, and a form that turns text down stays up for another try. Cancel is
  False whatever the Type, Type 128 and negative ones are 1004 before any form, and a Type that is not a number is #VALUE!.
- GetOpenFilename and GetSaveAsFilename give a path as text and open nothing. Cancel is False, and with MultiSelect a chosen
  file comes in a 1-based array. A filter of an odd number of parts is 1004 before any form.
- FileDialog is one object: asking for another type turns it into that type, and whoever holds it sees the change, with the
  title, button, multi-select, filters and selection back at that type's. The view, file name and filter index stay.
- Dialogs(n).Show puts up one of Excel's own forms.

The answer is what the dialog returned, so a save dialog's filter adds no extension to it, and no folder is looked at.
Unsupported, and reporting itself: other Type sums, Dialogs(n).Show answered True, as what OK does in each form is not modelled,
and FileDialog.Execute on a saving dialog. An InitialFileName in a folder that is not there is kept, where Excel moves it to
the user's Documents.
"""
from __future__ import annotations

import ntpath
from collections.abc import Iterator
from typing import TYPE_CHECKING, cast

from pyopenvba.apps.excel._creator import ItemDefault
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._typing import typed_text
from pyopenvba.exceptions import VBADialogError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    VBAArray,
    VBAErrorValue,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application

_VALUE_ERROR = VBAErrorValue(2015)
_ERROR_NUMBERS = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029, "#NUM!": 2036, "#N/A": 2042}
#: The Type values measured: one kind each, and text or logical beside numbers.
_TYPES = frozenset({0, 1, 2, 3, 4, 5, 8, 16, 64})
_REFUSED = object()

_OPEN_FILTERS: tuple[tuple[str, str], ...] = (
    ('All Files', '*.*'),
    ('All Excel Files', '*.xl*, *.xlsx, *.xlsm, *.xlsb, *.xlam, *.xltx, *.xltm, *.xls, *.xlt, *.htm, *.html, *.mht, *.mhtml, *.xml, *.xla, *.xlm, *.xlw, *.odc, *.ods'),
    ('Excel Files', '*.xl*, *.xlsx, *.xlsm, *.xlsb, *.xlam, *.xltx, *.xltm, *.xls, *.xla, *.xlt, *.xlm, *.xlw'),
    ('All Web Pages', '*.htm, *.html, *.mht, *.mhtml'),
    ('XML Files', '*.xml'),
    ('Text Files', '*.prn, *.txt, *.csv'),
    ('All Data Sources', '*.odc, *.udl, *.dsn, *.mdb, *.mde, *.accdb, *.accde, *.dbc, *.iqy, *.dqy, *.rqy, *.oqy, *.cub, *.atom, *.atomsvc'),
    ('Access Databases', '*.mdb, *.mde, *.accdb, *.accde'),
    ('Query Files', '*.iqy, *.dqy, *.oqy, *.rqy'),
    ('dBase Files', '*.dbf'),
    ('Microsoft Excel 4.0 Macros', '*.xlm, *.xla'),
    ('Microsoft Excel 4.0 Workbooks', '*.xlw'),
    ('Worksheets', '*.xlsx, *.xlsm, *.xlsb, *.xls'),
    ('Workspaces', '*.xlw'),
    ('Templates', '*.xltx, *.xltm, *.xlt'),
    ('Add-ins', '*.xlam, *.xla, *.xll'),
    ('Toolbars', '*.xlb'),
    ('SYLK Files', '*.slk'),
    ('Data Interchange Format', '*.dif'),
    ('Backup Files', '*.xlk, *.bak'),
    ('OpenDocument Spreadsheet', '*.ods'),
)
_SAVE_FILTERS: tuple[tuple[str, str], ...] = (
    ('Excel Workbook', '*.xlsx'),
    ('Excel Macro-Enabled Workbook', '*.xlsm'),
    ('Excel Binary Workbook', '*.xlsb'),
    ('Excel 97-2003 Workbook', '*.xls'),
    ('CSV UTF-8 (Comma delimited)', '*.csv'),
    ('XML Data', '*.xml'),
    ('Single File Web Page', '*.mht, *.mhtml'),
    ('Web Page', '*.htm, *.html'),
    ('Excel Template', '*.xltx'),
    ('Excel Macro-Enabled Template', '*.xltm'),
    ('Excel 97-2003 Template', '*.xlt'),
    ('Text (Tab delimited)', '*.txt'),
    ('Unicode Text', '*.txt'),
    ('XML Spreadsheet 2003', '*.xml'),
    ('Microsoft Excel 5.0/95 Workbook', '*.xls'),
    ('CSV (Comma delimited)', '*.csv'),
    ('Formatted Text (Space delimited)', '*.prn'),
    ('Text (Macintosh)', '*.txt'),
    ('Text (MS-DOS)', '*.txt'),
    ('CSV (Macintosh)', '*.csv'),
    ('CSV (MS-DOS)', '*.csv'),
    ('DIF (Data Interchange Format)', '*.dif'),
    ('SYLK (Symbolic Link)', '*.slk'),
    ('Excel Add-in', '*.xlam'),
    ('Excel 97-2003 Add-in', '*.xla'),
    ('PDF', '*.pdf'),
    ('XPS Document', '*.xps'),
    ('Strict Open XML Spreadsheet', '*.xlsx'),
    ('OpenDocument Spreadsheet', '*.ods'),
)
#: The numbers below 3000 that Dialogs takes, the xlDialog constants among them.
_DIALOG_NUMBERS = frozenset({
    1, 2, 5, 6, 7, 8, 9, 12, 13, 14, 17, 23, 26, 27, 28, 32, 35, 36, 39, 40, 41, 42, 43, 44, 45, 46, 47, 52, 53, 54,
    55, 61, 62, 63, 64, 67, 68, 69, 70, 71, 72, 73, 76, 78, 80, 84, 85, 86, 87, 88, 89, 91, 94, 95, 103, 108, 110, 111,
    119, 127, 130, 132, 133, 134, 137, 142, 145, 147, 150, 154, 159, 161, 166, 170, 171, 185, 186, 188, 189, 190, 191,
    192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 207, 208, 210, 212, 213, 214, 217, 218, 219, 220,
    222, 223, 225, 226, 228, 229, 249, 251, 256, 259, 261, 262, 269, 272, 273, 276, 281, 282, 283, 284, 285, 288, 293,
    300, 301, 302, 305, 307, 308, 311, 312, 313, 318, 319, 320, 321, 322, 323, 325, 328, 330, 336, 339, 342, 344, 350,
    354, 355, 356, 370, 379, 380, 381, 382, 384, 386, 388, 390, 392, 394, 398, 415, 417, 420, 421, 422, 423, 433, 434,
    435, 436, 437, 438, 445, 447, 450, 458, 460, 461, 462, 463, 464, 465, 466, 470, 472, 473, 474, 475, 476, 478, 480,
    481, 482, 485, 489, 491, 493, 494, 504, 505, 506, 509, 510, 511, 525, 526, 527, 530, 540, 541, 552, 553, 557, 558,
    567, 568, 570, 572, 583, 590, 596, 620, 647, 653, 656, 666, 667, 673, 674, 683, 684, 685, 686, 687, 689, 709, 723,
    724, 731, 732, 753, 754, 755, 759, 771, 773, 796, 817, 832, 833, 834, 835, 860, 862, 907, 908, 962, 970, 977, 978,
    1022, 1107, 1108, 1109, 1133, 1134, 1135, 1153, 1179, 1182, 1183, 1184, 1203, 1208, 1216, 1217, 1218, 1245, 1254,
    1256, 1257, 1258, 1266, 1267, 1270, 1271, 1272, 1276, 1300, 1390, 1391, 1413, 1415, 1461, 1494,
})
_DIALOG_COUNT = 1531
_KINDS = {1: "msoFileDialogOpen", 2: "msoFileDialogSaveAs", 3: "msoFileDialogFilePicker", 4: "msoFileDialogFolderPicker"}
_TITLES = {1: "File Open", 2: "File Save", 3: "Browse", 4: "Browse"}
_BUTTONS = {1: "&Open", 2: "&Save", 3: "OK", 4: "OK"}


def _next(application: Application, kind: str, prompt: str = "", title: str = "", style: int = 0) -> object:
    interpreter = application.interpreter
    assert interpreter is not None
    return interpreter.next_answer(kind, prompt, title, style)


def _wrong(kind: str, answer: object, wanted: str) -> VBADialogError:
    return VBADialogError(f"{kind} cannot take the answer {answer!r}; it takes {wanted}, or False for Cancel")


# --- Application.InputBox --------------------------------------------------------------------------------------------


def _reference(application: Application, text: str) -> Range | None:
    try:
        found = application.Range(text.removeprefix("="))
    except VBARuntimeError:
        return None
    return found if isinstance(found, Range) else None


def _first(found: Range) -> object:
    return found.vba_get("Cells", [1, 1]).vba_get("Value")  # type: ignore[union-attr]


def _worked_out(application: Application, text: str) -> tuple[object, bool]:
    """What the text comes to, and whether it was a formula: a formula's first cell if it is a reference, and otherwise
    what typing it in a cell reads it as."""
    if not text.startswith("="):
        value = typed_text(text).value
        if isinstance(value, ExcelError):
            return VBAErrorValue(_ERROR_NUMBERS.get(str(value), 2015)), False
        return value, False
    value = application.evaluate_text(text)
    return _first(value) if isinstance(value, Range) else value, True


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _as_formula(application: Application, text: str) -> object:
    """Type 0: the formula as R1C1 from the active cell, and anything else made into one."""
    if text.startswith("="):
        converted = application.ConvertFormula(text, 1, -4150, MISSING, application.vba_get("ActiveCell"))
        book = application.active_book
        return converted.replace(f"[{book.name}]", "") if isinstance(converted, str) and book is not None else converted
    value = typed_text(text).value
    if value is EMPTY:
        return _REFUSED
    if _is_number(value) or isinstance(value, (bool, ExcelError)):
        return "=" + text.strip().upper()
    return '="' + text.replace('"', '""') + '"'


def _as_array(application: Application, text: str) -> object:
    """Type 64: an array as it is, a block's values, and anything else one item in a 1-based array."""
    reference = _reference(application, text)
    if reference is not None and text.strip():
        value = reference.vba_get("Value")
        return value if isinstance(value, VBAArray) else VBAArray([(1, 1)], items=[value])
    if text.startswith("{"):
        made = application.evaluate_text(text)
        return made if isinstance(made, VBAArray) else _REFUSED
    value, _ = _worked_out(application, text)
    if value is EMPTY or isinstance(value, (VBAErrorValue, str)):
        return _REFUSED
    return VBAArray([(1, 1)], items=[float(value) if _is_number(value) else value])  # type: ignore[arg-type]


def _read(application: Application, kind: int, text: str) -> object:
    """What the form gives back for this text, or _REFUSED where it stays up."""
    if kind == 0:
        return _as_formula(application, text)
    if kind == 8:
        found = _reference(application, text) if text.strip() else None
        return _REFUSED if found is None else found
    if kind == 64:
        return _as_array(application, text)
    value, formula = _worked_out(application, text)
    if kind == 16:
        if value is EMPTY:
            return _REFUSED
        return value if isinstance(value, VBAErrorValue) else _VALUE_ERROR
    if isinstance(value, VBAErrorValue):
        return text if kind == 2 and not formula else _REFUSED
    if kind == 2:
        return to_text(value) if formula else text
    numeric = _is_number(value)
    if kind == 1:
        return float(value) if numeric or isinstance(value, bool) else _REFUSED  # type: ignore[arg-type]
    if kind == 4:
        return bool(value) if numeric or isinstance(value, bool) else _REFUSED
    if kind == 3:
        return float(value) if numeric or isinstance(value, bool) else (to_text(value) if formula else text)  # type: ignore[arg-type]
    if isinstance(value, bool):  # Type 5
        return value
    return float(value) if numeric else _REFUSED  # type: ignore[arg-type]


def _shown(default: object) -> str:
    """What the box starts with: a range is its value."""
    if default is MISSING:
        return ""
    return to_text(_first(default) if isinstance(default, Range) else default)


def input_box(application: Application, prompt: object, title: object, default: object, type_value: object) -> object:
    kind = 2
    if type_value is not MISSING:
        try:
            kind = int(to_integer(type_value, "Long"))
        except VBARuntimeError:
            return _VALUE_ERROR
    if kind < 0 or kind >= 128:
        raise error(1004, "Method 'InputBox' of object '_Application' failed")
    if kind not in _TYPES:
        raise VBAUnsupportedError(f"Application.InputBox with Type:={kind}, a sum not measured, is not implemented")
    caption = "Input" if title is MISSING else to_text(title)
    answer = _next(application, "Application.InputBox", to_text(prompt), caption, kind)
    while True:
        if answer is False:
            return False
        if answer is True:
            answer = "TRUE"
        if answer is not None and not isinstance(answer, (str, int, float)):
            raise _wrong("Application.InputBox", answer, "text, a number or None")
        value = _read(application, kind, _shown(default) if answer is None else to_text(answer))
        if value is not _REFUSED:
            return value
        interpreter = application.interpreter
        assert interpreter is not None
        answer = interpreter.next_retry(f"Application.InputBox, having turned down {answer!r} for Type:={kind},")


# --- the file name dialogs -------------------------------------------------------------------------------------------


def _check_filter(file_filter: object, name: str) -> None:
    if file_filter is not MISSING and to_text(file_filter) and to_text(file_filter).count(",") % 2 == 0:
        raise error(1004, f"Method '{name}' of object '_Application' failed")


def _paths(kind: str, answer: object) -> list[str]:
    given: list[object] = cast("list[object]", answer) if isinstance(answer, list | tuple) else [answer]
    paths = [path for path in given if isinstance(path, str)]
    if not paths or len(paths) != len(given):
        raise _wrong(kind, cast("object", answer), "a path or a list of paths")
    return paths


def _moved_to(application: Application, path: str) -> None:
    """The folder of the chosen file is the current folder now. This looks at nothing out there."""
    folder = ntpath.dirname(path)
    if folder and application.interpreter is not None:
        application.interpreter.files.folder = folder


def get_open_filename(application: Application, file_filter: object, title: object, multi_select: object) -> object:
    _check_filter(file_filter, "GetOpenFilename")
    answer = _next(application, "GetOpenFilename", "", "" if title is MISSING else to_text(title))
    if answer is False:
        return False
    paths = _paths("GetOpenFilename", answer)
    if multi_select is not MISSING and to_bool(multi_select):
        return VBAArray([(1, len(paths))], items=list(paths))
    if len(paths) > 1:
        raise _wrong("GetOpenFilename without MultiSelect", answer, "one path")
    return paths[0]


def get_save_as_filename(application: Application, initial: object, file_filter: object, title: object) -> object:
    _check_filter(file_filter, "GetSaveAsFilename")
    answer = _next(application, "GetSaveAsFilename", "", "" if title is MISSING else to_text(title))
    if answer is False:
        return False
    if answer is None and initial is not MISSING:
        answer = to_text(initial)
    if not isinstance(answer, str):
        raise _wrong("GetSaveAsFilename", answer, "a path")
    _moved_to(application, answer)
    return answer


# --- FileDialog ------------------------------------------------------------------------------------------------------


def _filters_of(kind: int) -> list[FileDialogFilter]:
    return [FileDialogFilter(*pair) for pair in {1: _OPEN_FILTERS, 2: _SAVE_FILTERS}.get(kind, ())]


class FileDialogFilter(ExcelObject):
    vba_type_name = "FileDialogFilter"
    vba_library = "office"

    def __init__(self, description: str, extensions: str) -> None:
        self.description = description
        self.extensions = extensions
        self.owner: FileDialog | None = None

    @member
    def Description(self) -> object:
        return self.description

    @member
    def Extensions(self) -> object:
        return self.extensions

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.application if self.owner is not None else None


class FileDialogFilters(VBACollection, ItemDefault, ExcelObject):
    vba_type_name = "FileDialogFilters"
    vba_library = "office"

    def __init__(self, dialog: FileDialog) -> None:
        self.dialog = dialog

    def vba_items(self) -> list[object]:
        return list(self.dialog.filters)

    @member
    def Parent(self) -> object:
        return self.dialog

    @member
    def Application(self) -> object:
        return self.dialog.application

    @method
    def Add(self, Description: object = MISSING, Extensions: object = MISSING, Position: object = MISSING) -> object:  # noqa: N803
        """A filter's parts each start `*.`, which is 5 even where Add is not there: a saving dialog's filters and a
        folder's are fixed, which is 438. The same filter again changes nothing."""
        extensions = to_text(Extensions)
        if extensions and not all(part.strip().startswith("*.") for part in extensions.replace(",", ";").split(";")):
            raise error(5, "Invalid procedure call or argument")
        if self.dialog.kind in (2, 4):
            raise error(438, "Object doesn't support this property or method")
        filters = self.dialog.filters
        where = len(filters) if Position is MISSING else int(to_integer(Position, "Long")) - 1
        if not 0 <= where <= len(filters):
            raise error(9, "Subscript out of range")
        description = to_text(Description)
        for held in filters:
            if held.description == description and held.extensions == extensions:
                return held
        made = FileDialogFilter(description, extensions)
        made.owner = self.dialog
        filters.insert(where, made)
        return made

    @method
    def Clear(self) -> object:
        if self.dialog.kind == 2:
            raise error(438, "Object doesn't support this property or method")
        self.dialog.filters.clear()
        return EMPTY

    @method
    def Delete(self, Filter: object = MISSING) -> object:  # noqa: N803
        if self.dialog.kind == 2:
            raise error(438, "Object doesn't support this property or method")
        where = 1 if Filter is MISSING else int(to_integer(Filter, "Long"))
        if not 1 <= where <= len(self.dialog.filters):
            raise error(9, "Subscript out of range")
        del self.dialog.filters[where - 1]
        return EMPTY


class FileDialogSelectedItems(VBACollection, ItemDefault, ExcelObject):
    vba_type_name = "FileDialogSelectedItems"
    vba_library = "office"

    def __init__(self, dialog: FileDialog) -> None:
        self.dialog = dialog

    def vba_items(self) -> list[object]:
        return list(self.dialog.selected)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, "Long"))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(5, "Invalid procedure call or argument")

    @member
    def Parent(self) -> object:
        return self.dialog

    @member
    def Application(self) -> object:
        return self.dialog.application


class FileDialog(ExcelObject):
    """The application's one file dialog, of whichever type was asked for last."""

    vba_type_name = "FileDialog"
    vba_library = "office"

    def __init__(self, application: Application, kind: int) -> None:
        self.application = application
        self.view = 1
        self.file_name = ""
        #: Where the last chosen file was, which a name with no folder goes in ahead of the current folder.
        self.chosen_folder = ""
        self.filter_index = 0
        self.turn_into(kind)

    def turn_into(self, kind: int) -> None:
        self.kind = kind
        self.title = _TITLES[kind]
        self.button = _BUTTONS[kind]
        self.multi = kind in (1, 3)
        self.filters = _filters_of(kind)
        for held in self.filters:
            held.owner = self
        self.selected: list[str] = []

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application

    @member
    def Item(self) -> object:
        return f"FileDialog({_KINDS[self.kind]})"

    @member
    def DialogType(self) -> object:
        return VBAInt(self.kind, "Long")

    @member
    def Filters(self) -> object:
        return FileDialogFilters(self)

    @member
    def SelectedItems(self) -> object:
        return FileDialogSelectedItems(self)

    @member
    def Title(self) -> object:
        return self.title

    @setter("Title")
    def set_title(self, value: object) -> None:
        self.title = to_text(value) or ("File Save" if self.kind == 2 else "File Open")

    @member
    def ButtonName(self) -> object:
        return self.button

    @setter("ButtonName")
    def set_button(self, value: object) -> None:
        self.button = to_text(value) or _BUTTONS[self.kind]

    @member
    def AllowMultiSelect(self) -> object:
        return self.multi

    @setter("AllowMultiSelect")
    def set_multi(self, value: object) -> None:
        if self.kind in (1, 3):
            self.multi = to_bool(value)

    @member
    def InitialView(self) -> object:
        return VBAInt(self.view, "Long")

    @setter("InitialView")
    def set_view(self, value: object) -> None:
        view = int(to_integer(value, "Long"))
        if not 1 <= view <= 9:
            raise error(5, "Invalid procedure call or argument")
        self.view = view

    @member
    def FilterIndex(self) -> object:
        return VBAInt(self.filter_index, "Long")

    @setter("FilterIndex")
    def set_filter_index(self, value: object) -> None:
        index = int(to_integer(value, "Long"))
        if not self.filters and self.kind == 1:
            raise error(5, "Invalid procedure call or argument")
        self.filter_index = max(1, min(index, len(self.filters)))

    @member
    def InitialFileName(self) -> object:
        return self.file_name

    @setter("InitialFileName")
    def set_file_name(self, value: object) -> None:
        """A name with no folder is in the folder last chosen, or else the current one, which is outside the engine."""
        name = to_text(value)
        if ntpath.isabs(name) or ntpath.splitdrive(name)[0] or "://" in name:
            self.file_name = name
            return
        interpreter = self.application.interpreter
        assert interpreter is not None
        interpreter.outside("FileDialog.InitialFileName with no folder")
        folder = self.chosen_folder or interpreter.files.folder
        self.file_name = ntpath.join(folder, name) if name else folder + "\\"

    @method
    def Show(self) -> object:
        answer = _next(self.application, "FileDialog", "", self.title, self.kind)
        self.selected = []
        self.filter_index = max(1, min(self.filter_index, len(self.filters)))
        if answer is False:
            return VBAInt(0, "Long")
        paths = _paths("FileDialog", answer)
        if len(paths) > 1 and not self.multi:
            raise _wrong("A FileDialog without AllowMultiSelect", answer, "one path")
        self.selected = paths
        if self.kind != 4:
            _moved_to(self.application, paths[0])
            self.chosen_folder = ntpath.dirname(paths[0])
            self.file_name = self.file_name or self.chosen_folder + "\\"
        return VBAInt(-1, "Long")

    @method
    def Execute(self) -> object:
        """Nothing with nothing chosen. An opening dialog opens what was chosen, through Workbooks.Open and its sandbox."""
        if not self.selected:
            return EMPTY
        if self.kind in (3, 4):
            raise error(438, "Object doesn't support this property or method")
        if self.kind == 2:
            raise VBAUnsupportedError("FileDialog.Execute on a saving dialog is not implemented")
        for path in self.selected:
            self.application.workbooks_.Open(path)
        return EMPTY


def file_dialog(application: Application, kind_value: object) -> FileDialog:
    kind = int(to_integer(kind_value, "Long"))
    if kind not in _KINDS:
        raise error(-2147467259, "Method 'FileDialog' of object '_Application' failed")
    dialog = application.file_dialog
    if dialog is None:
        dialog = application.file_dialog = FileDialog(application, kind)
    elif dialog.kind != kind:
        dialog.turn_into(kind)
    return dialog


# --- Dialogs ---------------------------------------------------------------------------------------------------------


class Dialog(ExcelObject):
    vba_type_name = "Dialog"

    def __init__(self, application: Application, number: int) -> None:
        self.application = application
        self.number = number

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application

    @method
    def Show(self, *arguments: object) -> object:
        answer = _next(self.application, "Dialog", "", "", self.number)
        if answer is False:
            return False
        if answer is True:
            raise VBAUnsupportedError(f"Dialogs({self.number}).Show answered OK is not implemented, as what OK does in that form is not modelled")
        raise _wrong("Dialogs.Show", answer, "True")


class Dialogs(VBACollection, ItemDefault, ExcelObject):
    """Item takes the number of one of Excel's forms. Excel's own For Each over them is 1004."""

    vba_type_name = "Dialogs"

    def __init__(self, application: Application) -> None:
        self.application = application

    @member
    def Application(self) -> object:
        return self.application

    @member
    def Parent(self) -> object:
        return self.application

    @member
    def Count(self) -> object:
        return VBAInt(_DIALOG_COUNT, "Long")

    def vba_items(self) -> list[object]:
        return []

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = int(to_integer(index, "Long"))
        if number not in _DIALOG_NUMBERS:
            raise error(1004, "Application-defined or object-defined error")
        return Dialog(self.application, number)

    def vba_iterate(self) -> Iterator[object]:
        raise error(1004, "Application-defined or object-defined error")
