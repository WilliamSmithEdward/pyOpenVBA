"""Scripting.FileSystemObject's files, folders and text streams on Python's, as tests/fixtures/file_system_object.json
measured them in live Excel. All of it is refused with the sandbox on.

Its errors are its own: where a statement says 75 or 55, this says 70. A TextStream is written at Close, so a file being
written reads as empty from outside until then. ANSI text is Windows-1252, and a character it lacks is error 5, where
Print # writes a question mark. A line ends at an LF, with or without a CR before it, and a CR alone is a character.

Drive objects, ShortName, ShortPath and Type come from the machine and are refused."""
from __future__ import annotations

import datetime as _dt
import fnmatch
import os
import random
import shutil
import tempfile
from collections.abc import Iterator
from typing import TYPE_CHECKING, Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBADate,
    VBAInt,
    error,
    to_bool,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._files import FileTable
    from pyopenvba.interpreter._runtime import Interpreter

ERR_INVALID = 5
ERR_MISMATCH = 13
ERR_NOT_FOUND = 53
ERR_BAD_MODE = 54
ERR_EXISTS = 58
ERR_PAST_END = 62
ERR_DENIED = 70
ERR_NO_PATH = 76
ERR_NO_OBJECT = 91
_READING, _WRITING, _APPENDING = 1, 2, 8
_ANSI, _UNICODE, _DEFAULT = 0, -1, -2
_BOM: Final = b"\xff\xfe"
#: The bits Attributes can set: read-only, hidden, system, archive.
_SETTABLE: Final = 1 | 2 | 4 | 32


def _decode_ansi(data: bytes) -> str:
    return "".join(bytes([byte]).decode("cp1252", errors="ignore") or chr(byte) for byte in data)


def _given(value: object) -> str:
    """A path as given: Null is a type mismatch, and nothing at all is error 5."""
    if value is NULL:
        raise error(ERR_MISMATCH)
    text = to_text(value)
    if not text:
        raise error(ERR_INVALID)
    return text


def _wild(name: str) -> bool:
    return "*" in name or "?" in name


def _matches(name: str, pattern: str) -> bool:
    if pattern in ("*", "*.*"):
        return True
    return fnmatch.fnmatchcase(name.casefold(), pattern.casefold().replace("[", "[[]"))


def _ends_with_separator(text: str) -> bool:
    return text.endswith(("\\", "/"))


class TextStream(VBAObject):
    """A text file open for reading, writing or appending, its text in memory until it closes."""

    vba_type_name = "TextStream"
    vba_library = "scripting"

    def __init__(self, table: FileTable, path: str, mode: int, unicode: bool, text: str, existing: bytes) -> None:
        self.table, self.path, self.mode, self.unicode = table, path, mode, unicode
        self.text, self.at = text, 0
        self.written = bytearray()
        self.existing = existing
        self.closed = False
        self.line, self.column = 1, 1
        if mode == _APPENDING:
            # It goes on from where the file ends.
            self.column = len(text) - text.rfind("\n")
        table.streams.append(self)

    def _open(self) -> None:
        self.table.interpreter.outside("TextStream")
        if self.closed:
            raise error(ERR_NO_OBJECT)

    def _reading(self) -> None:
        self._open()
        if self.mode != _READING:
            raise error(ERR_BAD_MODE)

    def _writing(self) -> None:
        self._open()
        if self.mode == _READING:
            raise error(ERR_BAD_MODE)

    def _advance(self, text: str) -> None:
        cut = text.rfind("\n")
        if cut >= 0:
            self.line += text.count("\n")
            self.column = len(text) - cut
        else:
            self.column += len(text)

    def _take(self, count: int) -> str:
        piece = self.text[self.at:self.at + count]
        self.at += len(piece)
        self._advance(piece)
        return piece

    def _line_end(self) -> tuple[int, int]:
        """Where the line goes up to, and where the next starts."""
        found = self.text.find("\n", self.at)
        if found < 0:
            return len(self.text), len(self.text)
        return (found - 1 if found > self.at and self.text[found - 1] == "\r" else found), found + 1

    @member
    def AtEndOfStream(self) -> object:
        self._reading()
        return self.at >= len(self.text)

    @member
    def AtEndOfLine(self) -> object:
        self._reading()
        return self.at >= len(self.text) or self.text[self.at] == "\n" or self.text[self.at:self.at + 2] == "\r\n"

    @member
    def Line(self) -> object:
        self._open()
        return VBAInt(self.line, "Long")

    @member
    def Column(self) -> object:
        self._open()
        return VBAInt(self.column, "Long")

    @method
    def Read(self, Characters: object = MISSING) -> object:
        self._reading()
        count = int(to_integer(Characters, "Long"))
        if count < 0:
            raise error(ERR_INVALID)
        if count and self.at >= len(self.text):
            raise error(ERR_PAST_END)
        return self._take(count)

    @method
    def ReadLine(self) -> object:
        self._reading()
        if self.at >= len(self.text):
            raise error(ERR_PAST_END)
        end, after = self._line_end()
        text = self.text[self.at:end]
        self._take(after - self.at)
        return text

    @method
    def ReadAll(self) -> object:
        self._reading()
        if self.at >= len(self.text):
            raise error(ERR_PAST_END)
        return self._take(len(self.text))

    @method
    def Skip(self, Characters: object = MISSING) -> object:
        self._reading()
        count = int(to_integer(Characters, "Long"))
        if count < 0:
            raise error(ERR_INVALID)
        if count and self.at >= len(self.text):
            raise error(ERR_PAST_END)
        self._take(count)
        return EMPTY

    @method
    def SkipLine(self) -> object:
        self._reading()
        if self.at >= len(self.text):
            raise error(ERR_PAST_END)
        self._take(self._line_end()[1] - self.at)
        return EMPTY

    def _put(self, text: str) -> None:
        if self.unicode:
            data = text.encode("utf-16-le", errors="surrogatepass")
        else:
            try:
                data = text.encode("cp1252")
            except UnicodeEncodeError:
                raise error(ERR_INVALID) from None
        self.written.extend(data)
        self._advance(text)

    @staticmethod
    def _text_of(value: object) -> str:
        if value is NULL:
            raise error(ERR_MISMATCH)
        return "" if value is EMPTY or value is MISSING else to_text(value)

    @method
    def Write(self, Text: object = MISSING) -> object:
        self._writing()
        self._put(self._text_of(Text))
        return EMPTY

    @method
    def WriteLine(self, Text: object = MISSING) -> object:
        self._writing()
        self._put(self._text_of(Text) + "\r\n")
        return EMPTY

    @method
    def WriteBlankLines(self, Lines: object = MISSING) -> object:
        self._writing()
        count = int(to_integer(Lines, "Long"))
        if count < 0:
            raise error(ERR_INVALID)
        self._put("\r\n" * count)
        return EMPTY

    @method
    def Close(self) -> object:
        self.table.interpreter.outside("TextStream.Close")
        if not self.closed:
            self.closed = True
            self.table.streams.remove(self)
            self._flush()
        return EMPTY

    def _flush(self) -> None:
        if self.mode == _READING:
            return
        head = self.existing
        if self.unicode:
            if not head:
                head = _BOM
            elif len(head) % 2:
                head += b"\0"
        with open(self.path, "wb") as stream:
            stream.write(head + bytes(self.written))


class FileSystemMembers(VBAObject):
    """What FileSystemObject does on the disk, beside its path functions."""

    interpreter: Interpreter

    def _table(self, what: str) -> FileTable:
        self.interpreter.outside(f"FileSystemObject.{what}")
        return self.interpreter.files

    @method
    def FileExists(self, FileSpec: object = MISSING) -> object:
        table = self._table("FileExists")
        if FileSpec is NULL:
            raise error(ERR_MISMATCH)
        given = to_text(FileSpec)
        return bool(given) and os.path.isfile(table.full(given))

    @method
    def FolderExists(self, FolderSpec: object = MISSING) -> object:
        table = self._table("FolderExists")
        if FolderSpec is NULL:
            raise error(ERR_MISMATCH)
        given = to_text(FolderSpec)
        return bool(given) and os.path.isdir(table.full(given))

    @method
    def DriveExists(self, DriveSpec: object = MISSING) -> object:
        self._table("DriveExists")
        drive = to_text(DriveSpec)
        return bool(drive) and len(drive) <= 3 and drive[0].isalpha() and os.path.exists(f"{drive[0]}:\\")

    @method
    def GetFile(self, FilePath: object = MISSING) -> object:
        table = self._table("GetFile")
        path = table.full(_given(FilePath))
        if not os.path.isfile(path):
            raise error(ERR_NOT_FOUND)
        return FsoFile(table, _actual(path))

    @method
    def GetFolder(self, FolderPath: object = MISSING) -> object:
        table = self._table("GetFolder")
        path = table.full(_given(FolderPath))
        if not os.path.isdir(path):
            raise error(ERR_NO_PATH)
        return FsoFolder(table, _actual(path))

    @method
    def CreateFolder(self, Path: object = MISSING) -> object:
        table = self._table("CreateFolder")
        path = table.full(_given(Path))
        return _make_folder(table, path)

    @method
    def CreateTextFile(self, FileName: object = MISSING, Overwrite: object = MISSING,
                       Unicode: object = MISSING) -> object:
        table = self._table("CreateTextFile")
        return _create_stream(table, table.full(_given(FileName)), Overwrite, Unicode)

    @method
    def OpenTextFile(self, FileName: object = MISSING, IOMode: object = MISSING, Create: object = MISSING,
                     Format: object = MISSING) -> object:
        table = self._table("OpenTextFile")
        return _open_stream(table, table.full(_given(FileName)), IOMode, Create, Format)

    @method
    def DeleteFile(self, FileSpec: object = MISSING, Force: object = MISSING) -> object:
        table = self._table("DeleteFile")
        path = table.full(_given(FileSpec))
        folder, pattern = os.path.split(path)
        # The name of a folder is no error, and deletes nothing.
        if pattern and not _wild(pattern) and os.path.isdir(path):
            return EMPTY
        names = _listed(folder, pattern, folders=False) if _wild(pattern) else [pattern] if os.path.isfile(path) else []
        if not names:
            raise error(ERR_NOT_FOUND)
        for name in names:
            _delete_file(table, os.path.join(folder, name), Force)
        return EMPTY

    @method
    def DeleteFolder(self, FolderSpec: object = MISSING, Force: object = MISSING) -> object:
        table = self._table("DeleteFolder")
        given = _given(FolderSpec)
        path = table.full(given)
        folder, pattern = os.path.split(path)
        names = _listed(folder, pattern, folders=True) if _wild(pattern) else [
            pattern] if pattern and os.path.isdir(path) else []
        if not names:
            raise error(ERR_NO_PATH)
        for name in names:
            _delete_folder(table, os.path.join(folder, name), Force)
        return EMPTY

    @method
    def CopyFile(self, Source: object = MISSING, Destination: object = MISSING,
                 OverWriteFiles: object = MISSING) -> object:
        table = self._table("CopyFile")
        given, target = _given(Source), _given(Destination)
        source = table.full(given)
        folder, pattern = os.path.split(source)
        if _wild(pattern):
            names = _listed(folder, pattern, folders=False)
            if not names:
                raise error(ERR_NOT_FOUND)
            place = table.full(target)
            if not os.path.isdir(place):
                raise error(ERR_NO_PATH)
            for name in names:
                _copy_file(table, os.path.join(folder, name), os.path.join(place, name), OverWriteFiles)
            return EMPTY
        if not os.path.isfile(source):
            raise error(ERR_NOT_FOUND)
        _copy_to(table, source, target, OverWriteFiles)
        return EMPTY

    @method
    def CopyFolder(self, Source: object = MISSING, Destination: object = MISSING,
                   OverWriteFiles: object = MISSING) -> object:
        table = self._table("CopyFolder")
        given, target = _given(Source), _given(Destination)
        source = table.full(given)
        folder, pattern = os.path.split(source)
        if _wild(pattern):
            names = _listed(folder, pattern, folders=True)
            place = table.full(target)
            if not names or not os.path.isdir(place):
                raise error(ERR_NO_PATH)
            for name in names:
                _copy_tree(table, os.path.join(folder, name), os.path.join(place, name), OverWriteFiles)
            return EMPTY
        if not os.path.isdir(source) or _ends_with_separator(given):
            raise error(ERR_NO_PATH)
        _copy_folder_to(table, source, target, OverWriteFiles)
        return EMPTY

    @method
    def MoveFile(self, Source: object = MISSING, Destination: object = MISSING) -> object:
        table = self._table("MoveFile")
        given, target = _given(Source), _given(Destination)
        source = table.full(given)
        folder, pattern = os.path.split(source)
        if _wild(pattern):
            names = _listed(folder, pattern, folders=False)
            if not names:
                raise error(ERR_NOT_FOUND)
            for name in names:
                _move(table, os.path.join(folder, name), os.path.join(table.full(target), name))
            return EMPTY
        if not os.path.isfile(source):
            raise error(ERR_NOT_FOUND)
        _move(table, source, _into(table, source, target))
        return EMPTY

    @method
    def MoveFolder(self, Source: object = MISSING, Destination: object = MISSING) -> object:
        table = self._table("MoveFolder")
        given, target = _given(Source), _given(Destination)
        source = table.full(given)
        folder, pattern = os.path.split(source)
        if _wild(pattern):
            names = _listed(folder, pattern, folders=True)
            if not names:
                raise error(ERR_NO_PATH)
            for name in names:
                _move(table, os.path.join(folder, name), os.path.join(table.full(target), name))
            return EMPTY
        if not os.path.isdir(source):
            raise error(ERR_NO_PATH)
        place = _into(table, source, target)
        if _same(source, place):
            return EMPTY
        if _within(place, source):
            raise error(ERR_INVALID)
        _move(table, source, place)
        return EMPTY

    @method
    def GetAbsolutePathName(self, Path: object = MISSING) -> object:
        table = self._table("GetAbsolutePathName")
        given = to_text(Path)
        return os.path.normpath(table.full(given)) if given else table.folder

    @method
    def GetTempName(self) -> object:
        self._table("GetTempName")
        return f"rad{random.randrange(0x100000):05X}.tmp"

    @method
    def GetSpecialFolder(self, SpecialFolder: object = MISSING) -> object:
        table = self._table("GetSpecialFolder")
        kind = int(to_integer(SpecialFolder, "Long"))
        if kind == 2:
            return FsoFolder(table, tempfile.gettempdir())
        if kind not in (0, 1):
            raise error(ERR_INVALID)
        root = os.environ.get("SystemRoot")
        if root is None:
            raise VBAUnsupportedError("GetSpecialFolder of Windows's own folders is not implemented elsewhere")
        return FsoFolder(table, _actual(root if kind == 0 else os.path.join(root, "System32")))


def _actual(path: str) -> str:
    """The path as the disk spells its last name, with no separator at the end."""
    path = os.path.normpath(path)
    folder, name = os.path.split(path)
    if name:
        try:
            spelled = next((one for one in os.listdir(folder) if one.casefold() == name.casefold()), name)
        except OSError:
            spelled = name
        return os.path.join(folder, spelled)
    return path


def _listed(folder: str, pattern: str, *, folders: bool) -> list[str]:
    try:
        names = sorted(os.listdir(folder), key=str.casefold)
    except OSError:
        return []
    return [name for name in names if _matches(name, pattern) and os.path.isdir(os.path.join(folder, name)) == folders]


def _make_folder(table: FileTable, path: str) -> FsoFolder:
    table.interpreter.outside("FileSystemObject")
    path = os.path.normpath(path)
    try:
        os.mkdir(path)
    except FileExistsError:
        raise error(ERR_EXISTS) from None
    except FileNotFoundError:
        raise error(ERR_NO_PATH) from None
    except OSError:
        raise error(ERR_DENIED) from None
    return FsoFolder(table, path)


def _busy(table: FileTable, path: str) -> bool:
    key = os.path.normcase(os.path.abspath(path))
    return table.is_open(path) or any(os.path.normcase(os.path.abspath(one.path)) == key for one in table.streams)


def _writer_on(table: FileTable, path: str) -> bool:
    key = os.path.normcase(os.path.abspath(path))
    return table.is_open(path, for_writing=True) or any(
        os.path.normcase(os.path.abspath(one.path)) == key and one.mode != _READING for one in table.streams)


def _unicode_flag(value: object) -> bool:
    return value is not MISSING and to_bool(value)


def _create_stream(table: FileTable, path: str, overwrite: object, unicode: object) -> TextStream:
    table.interpreter.outside("FileSystemObject")
    if os.path.exists(path) and not os.path.isdir(path) and overwrite is not MISSING and not to_bool(overwrite):
        raise error(ERR_EXISTS)
    if os.path.isdir(path) or table.read_only(path) or _writer_on(table, path):
        raise error(ERR_DENIED)
    wide = _unicode_flag(unicode)
    try:
        with open(path, "wb"):
            pass
    except FileNotFoundError:
        raise error(ERR_NO_PATH) from None
    except OSError:
        raise error(ERR_DENIED) from None
    return TextStream(table, path, _WRITING, wide, "", b"")


def _open_stream(table: FileTable, path: str, mode_value: object, create: object, form: object) -> TextStream:
    table.interpreter.outside("FileSystemObject")
    mode = _READING if mode_value is MISSING else int(to_integer(mode_value, "Long"))
    kind = _ANSI if form is MISSING else int(to_integer(form, "Long"))
    if mode not in (_READING, _WRITING, _APPENDING) or kind not in (_ANSI, _UNICODE, _DEFAULT):
        raise error(ERR_INVALID)
    if os.path.isdir(path) or _writer_on(table, path):
        raise error(ERR_DENIED)
    if mode != _READING and table.read_only(path):
        raise error(ERR_DENIED)
    exists = os.path.exists(path)
    if not exists and not _unicode_flag(create):
        raise error(ERR_NOT_FOUND if os.path.isdir(os.path.dirname(path)) else ERR_NO_PATH)
    try:
        data = b""
        if exists and mode != _WRITING:
            with open(path, "rb") as stream:
                data = stream.read()
        elif not exists or mode == _WRITING:
            with open(path, "wb"):
                pass
    except FileNotFoundError:
        raise error(ERR_NO_PATH) from None
    except OSError:
        raise error(ERR_DENIED) from None
    wide = kind == _UNICODE or (kind == _DEFAULT and mode != _WRITING and data.startswith(_BOM))
    if mode == _READING:
        body = data[2:] if wide and data.startswith(_BOM) else data
        text = body.decode("utf-16-le", errors="surrogatepass") if wide else _decode_ansi(body)
        return TextStream(table, path, mode, wide, text, data)
    if mode == _APPENDING:
        shown = data[2:].decode("utf-16-le", errors="replace") if wide else _decode_ansi(data)
        stream_ = TextStream(table, path, mode, wide, shown, data)
        stream_.text = ""
        return stream_
    return TextStream(table, path, mode, wide, "", b"")


def _delete_file(table: FileTable, path: str, force: object) -> None:
    table.interpreter.outside("FileSystemObject")
    if _busy(table, path):
        raise error(ERR_DENIED)
    if table.read_only(path):
        if force is MISSING or not to_bool(force):
            raise error(ERR_DENIED)
        table.set_attributes(path, VBAInt(0, "Integer"))
    try:
        os.remove(path)
    except OSError:
        raise error(ERR_DENIED) from None
    table.attributes.pop(os.path.normcase(os.path.abspath(path)), None)


def _delete_folder(table: FileTable, path: str, force: object) -> None:
    table.interpreter.outside("FileSystemObject")
    for root, folders, files in os.walk(path, topdown=False):
        for name in files:
            _delete_file(table, os.path.join(root, name), force)
        for name in folders:
            os.rmdir(os.path.join(root, name))
    os.rmdir(path)


def _copy_file(table: FileTable, source: str, target: str, overwrite: object) -> None:
    table.interpreter.outside("FileSystemObject")
    if os.path.isdir(target):
        raise error(ERR_DENIED)
    if os.path.exists(target) and overwrite is not MISSING and not to_bool(overwrite):
        raise error(ERR_EXISTS)
    if os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(target)):
        return
    if table.read_only(target) or _writer_on(table, target) or _writer_on(table, source):
        raise error(ERR_DENIED)
    try:
        shutil.copyfile(source, target)
    except FileNotFoundError:
        raise error(ERR_NO_PATH) from None
    except OSError:
        raise error(ERR_DENIED) from None
    bits = table.attributes_of(source)
    if bits != 32:
        table.set_attributes(target, VBAInt(bits | 32, "Integer"))


def _copy_to(table: FileTable, source: str, given: str, overwrite: object) -> None:
    target = table.full(given)
    if _ends_with_separator(given):
        if not os.path.isdir(target):
            raise error(ERR_NO_PATH)
        target = os.path.join(target, os.path.basename(source))
    _copy_file(table, source, target, overwrite)


def _copy_tree(table: FileTable, source: str, target: str, overwrite: object) -> None:
    table.interpreter.outside("FileSystemObject")
    if os.path.isfile(target):
        raise error(ERR_DENIED)
    if not os.path.isdir(target):
        try:
            os.mkdir(target)
        except FileNotFoundError:
            raise error(ERR_NO_PATH) from None
    for name in sorted(os.listdir(source), key=str.casefold):
        from_path, to_path = os.path.join(source, name), os.path.join(target, name)
        if os.path.isdir(from_path):
            _copy_tree(table, from_path, to_path, overwrite)
        else:
            _copy_file(table, from_path, to_path, overwrite)


def _same(one: str, other: str) -> bool:
    return os.path.normcase(os.path.abspath(one)) == os.path.normcase(os.path.abspath(other))


def _within(path: str, folder: str) -> bool:
    inside = os.path.normcase(os.path.abspath(path)) + os.sep
    return inside.startswith(os.path.normcase(os.path.abspath(folder)) + os.sep)


def _copy_folder_to(table: FileTable, source: str, given: str, overwrite: object) -> None:
    target = table.full(given)
    if _ends_with_separator(given):
        if not os.path.isdir(target):
            raise error(ERR_NO_PATH)
        target = os.path.join(target, os.path.basename(source))
    if _same(source, target):
        return
    if _within(target, source):
        raise error(ERR_INVALID)
    if not os.path.isdir(os.path.dirname(os.path.normpath(target))):
        raise error(ERR_NO_PATH)
    _copy_tree(table, source, target, overwrite)


def _into(table: FileTable, source: str, given: str) -> str:
    target = table.full(given)
    if _ends_with_separator(given):
        if not os.path.isdir(target):
            raise error(ERR_NO_PATH)
        return os.path.join(target, os.path.basename(os.path.normpath(source)))
    return target


def _move(table: FileTable, source: str, target: str) -> None:
    table.interpreter.outside("FileSystemObject")
    if os.path.exists(target):
        raise error(ERR_EXISTS)
    if not os.path.isdir(os.path.dirname(os.path.normpath(target))):
        raise error(ERR_NO_PATH)
    if _busy(table, source):
        raise error(ERR_DENIED)
    try:
        os.rename(source, target)
    except OSError:
        raise error(ERR_DENIED) from None
    kept = table.attributes.pop(os.path.normcase(os.path.abspath(source)), None)
    if kept is not None:
        table.attributes[os.path.normcase(os.path.abspath(target))] = kept


def _date(when: float) -> VBADate:
    return VBADate.from_datetime(_dt.datetime.fromtimestamp(when).replace(microsecond=0))


class _Entry(VBAObject):
    """What a File and a Folder share."""

    vba_library = "scripting"
    #: The error when it is gone.
    gone: int = ERR_NOT_FOUND

    def __init__(self, table: FileTable, path: str) -> None:
        self.table, self.path = table, path

    def _exists(self) -> bool:
        return os.path.exists(self.path)

    def _check(self, what: str) -> None:
        self.table.interpreter.outside(f"{self.vba_type_name}.{what}")
        if not self._exists():
            raise error(self.gone)

    @member
    def Path(self) -> object:
        self._check("Path")
        return self.path

    def vba_value(self) -> object:
        self._check("Path")
        return self.path

    @member
    def Name(self) -> object:
        self._check("Name")
        return os.path.basename(self.path)

    @setter("Name")
    def _rename(self, value: object) -> None:
        self._check("Name")
        name = to_text(value)
        if not name or "\\" in name or "/" in name:
            raise error(ERR_INVALID)
        target = os.path.join(os.path.dirname(self.path), name)
        _move(self.table, self.path, target)
        self.path = target

    @member
    def Attributes(self) -> object:
        self._check("Attributes")
        return VBAInt(self.table.attributes_of(self.path), "Long")

    @setter("Attributes")
    def _attributes(self, value: object) -> None:
        self._check("Attributes")
        bits = int(to_integer(value, "Long")) & _SETTABLE
        self.table.set_attributes(self.path, VBAInt(bits, "Integer"))
        self.table.attributes[os.path.normcase(os.path.abspath(self.path))] = bits | (
            16 if os.path.isdir(self.path) else 0)

    @member
    def DateCreated(self) -> object:
        self._check("DateCreated")
        return _date(os.path.getctime(self.path))

    @member
    def DateLastModified(self) -> object:
        self._check("DateLastModified")
        return _date(os.path.getmtime(self.path))

    @member
    def DateLastAccessed(self) -> object:
        self._check("DateLastAccessed")
        return _date(os.path.getatime(self.path))

    @member
    def ParentFolder(self) -> object:
        self._check("ParentFolder")
        parent = os.path.dirname(self.path)
        return NOTHING if parent == self.path else FsoFolder(self.table, parent)


class FsoFile(_Entry):
    vba_type_name = "File"

    def _exists(self) -> bool:
        return os.path.isfile(self.path)

    @member
    def Size(self) -> object:
        self._check("Size")
        return VBAInt(os.path.getsize(self.path), "Long")

    @method
    def Delete(self, Force: object = MISSING) -> object:
        self._check("Delete")
        _delete_file(self.table, self.path, Force)
        return EMPTY

    @method
    def Copy(self, Destination: object = MISSING, OverWriteFiles: object = MISSING) -> object:
        self._check("Copy")
        _copy_to(self.table, self.path, _given(Destination), OverWriteFiles)
        return EMPTY

    @method
    def Move(self, Destination: object = MISSING) -> object:
        self._check("Move")
        target = _into(self.table, self.path, _given(Destination))
        _move(self.table, self.path, target)
        self.path = target
        return EMPTY

    @method
    def OpenAsTextStream(self, IOMode: object = MISSING, Format: object = MISSING) -> object:
        self._check("OpenAsTextStream")
        return _open_stream(self.table, self.path, IOMode, MISSING, Format)


class FsoFolder(_Entry):
    vba_type_name = "Folder"
    gone = ERR_NO_PATH

    def _exists(self) -> bool:
        return os.path.isdir(self.path)

    @member
    def Size(self) -> object:
        self._check("Size")
        total = 0
        for root, _, files in os.walk(self.path):
            total += sum(os.path.getsize(os.path.join(root, name)) for name in files)
        return VBAInt(total, "Long")

    @member
    def IsRootFolder(self) -> object:
        self._check("IsRootFolder")
        return os.path.dirname(self.path) == self.path

    @member
    def Files(self, Key: object = MISSING) -> object:
        self._check("Files")
        found = FsoFiles(self.table, self.path)
        return found if Key is MISSING else found.vba_lookup(Key, [])

    @member
    def SubFolders(self, Key: object = MISSING) -> object:
        self._check("SubFolders")
        found = FsoFolders(self.table, self.path)
        return found if Key is MISSING else found.vba_lookup(Key, [])

    @method
    def CreateTextFile(self, FileName: object = MISSING, Overwrite: object = MISSING,
                       Unicode: object = MISSING) -> object:
        self._check("CreateTextFile")
        return _create_stream(self.table, os.path.join(self.path, _given(FileName)), Overwrite, Unicode)

    @method
    def Delete(self, Force: object = MISSING) -> object:
        self._check("Delete")
        _delete_folder(self.table, self.path, Force)
        return EMPTY

    @method
    def Copy(self, Destination: object = MISSING, OverWriteFiles: object = MISSING) -> object:
        self._check("Copy")
        _copy_folder_to(self.table, self.path, _given(Destination), OverWriteFiles)
        return EMPTY

    @method
    def Move(self, Destination: object = MISSING) -> object:
        self._check("Move")
        target = _into(self.table, self.path, _given(Destination))
        _move(self.table, self.path, target)
        self.path = target
        return EMPTY


class _Entries(VBACollection):
    """A folder's files or subfolders, read from the disk each time they are asked for."""

    vba_library = "scripting"
    folders = False
    missing = ERR_NOT_FOUND

    def __init__(self, table: FileTable, path: str) -> None:
        self.table, self.path = table, path

    def _make(self, path: str) -> VBAObject:
        return FsoFolder(self.table, path) if self.folders else FsoFile(self.table, path)

    def vba_items(self) -> list[object]:
        self.table.interpreter.outside(f"{self.vba_type_name}")
        return [self._make(os.path.join(self.path, name)) for name in _listed(self.path, "*", folders=self.folders)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        del items
        if not isinstance(index, str):
            raise error(ERR_INVALID)
        for name in _listed(self.path, "*", folders=self.folders):
            if name.casefold() == index.casefold():
                return self._make(os.path.join(self.path, name))
        raise error(self.missing)

    def vba_iterate(self) -> Iterator[object]:
        yield from self.vba_items()

    @member
    def Count(self) -> object:
        return VBAInt(len(self.vba_items()), "Long")

    @method
    def Item(self, Key: object = MISSING) -> object:
        return self.vba_lookup(Key, [])


class FsoFiles(_Entries):
    vba_type_name = "Files"


class FsoFolders(_Entries):
    vba_type_name = "Folders"
    folders = True
    missing = ERR_NO_PATH

    @method
    def Add(self, Name: object = MISSING) -> object:
        self.table.interpreter.outside("Folders.Add")
        name = _given(Name)
        if "\\" in name or "/" in name:
            raise error(ERR_INVALID)
        return _make_folder(self.table, os.path.join(self.path, name))
