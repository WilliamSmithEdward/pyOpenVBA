"""VBA's text files and file system statements on Python's, as tests/fixtures/file_statements.json measured them in live
Excel. All of it is refused with the sandbox on.

A file opened is held in memory and written back at Close, so a file being written reads as empty from outside until
then, as the measurements show. Text goes through Windows-1252. A CR ends a line and a lone LF does not, Ctrl-Z ends the
file, and a NUL in a line is dropped. Dir takes its list when it is given a path, so what later goes from the folder
still comes back. SetAttr's hidden, system and archive bits are kept for the session, and read-only reaches the file.

A Binary or Random file reads and writes through _file_records. ChDrive is refused for now."""
from __future__ import annotations

import datetime as _dt
import fnmatch
import math
import os
import re
import shutil
import stat
import weakref
from dataclasses import dataclass
from pathlib import Path
from decimal import Decimal
from typing import TYPE_CHECKING, Any, Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._file_records import ERR_RECORD_LENGTH, Reader, decode, encode, fill_array, fill_record
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBADecimal,
    VBAErrorValue,
    VBAInt,
    date_text,
    error,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter

ERR_BAD_NUMBER = 52
ERR_NOT_FOUND = 53
ERR_BAD_MODE = 54
ERR_ALREADY_OPEN = 55
ERR_EXISTS = 58
ERR_PAST_END = 62
ERR_BAD_RECORD = 63
ERR_DENIED = 70
ERR_NO_MODE = ERR_BAD_MODE
ERR_ACCESS = 75
ERR_NO_PATH = 76
ERR_INVALID = 5
ERR_MISMATCH = 13
ERR_OVERFLOW = 6
ERR_NO_OBJECT = 91

_CTRL_Z: Final = 0x1A
_ZONE: Final = 14
_MODES: Final = {"input": 1, "output": 2, "random": 4, "append": 8, "binary": 32}
_DEFAULT_RECORD: Final = 128
#: The modes that may share a file: Input with Input, and Binary and Random with each other.
_FAMILIES: Final = {"input": 0, "binary": 1, "random": 1}
#: SetAttr's bits: read-only, hidden, system, archive.
_SETTABLE: Final = 1 | 2 | 4 | 32
_DIR_BITS: Final = 1 | 2 | 4 | 8 | 16 | 32 | 64
_VOLUME: Final = 8
_NUMBER: Final = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?")
_RADIX: Final = re.compile(r"&([hHoO])([0-9a-fA-F]+)&?")
_MARK: Final = re.compile(r"#([^#]*)#$")
_FLOATS: Final = {"single", "double", "currency", "decimal"}
_WHOLES: Final = {"byte", "integer", "long", "longlong", "longptr"}


@dataclass
class TabTo:
    """Tab(n) in a Print #: None moves to the next zone."""

    column: int | None


@dataclass
class Spaces:
    """Spc(n) in a Print #."""

    count: int


def _encode(text: str) -> bytes:
    return text.encode("cp1252", errors="replace")


def _decode(data: bytes) -> str:
    """Windows-1252; the five bytes it leaves undefined keep their value."""
    return "".join(bytes([byte]).decode("cp1252", errors="ignore") or chr(byte) for byte in data)


def native_path(path: str) -> str:
    """A path a macro gives, in VBA's backslashes, as this system spells it: off Windows, folder & "\\book.xlsx" is a file in
    that folder, not a file with a backslash in its name."""
    return path.replace("\\", os.sep) if os.sep != "\\" else path


def respell_file(path: str) -> None:
    """Rename the file that differs from ``path`` only in case to ``path``. Where a file has that exact spelling there is nothing to do."""
    folder, name = os.path.split(path)
    listed = os.listdir(folder or ".")
    if name not in listed:
        for one in listed:
            if one.casefold() == name.casefold():
                os.replace(os.path.join(folder, one), path)
                return


def on_disk(path: str) -> str:
    """The file a macro's path means on this system, in its spelling and with its case."""
    return spelled(native_path(path))


def spelled(path: str, separator: str = os.sep) -> str:
    """The path as the disk spells it, since VBA's paths have no case and some systems' files do. A part that matches
    nothing stays as given, for what is about to be made."""
    if separator == "\\":
        return path
    current = separator if path.startswith(separator) else ""
    for part in path.split(separator):
        if not part:
            continue
        # By the folder's list, not by whether the name exists: a Mac's disk ignores case, so ALPHA.XLSX exists there,
        # and a workbook is named after the file as the disk spells it.
        try:
            listed = os.listdir(current or ".")
        except OSError:
            listed = []
        found = part if part in listed else next((one for one in listed if one.casefold() == part.casefold()), part)
        current = os.path.join(current, found)
    return current + separator if path.endswith(separator) else current


def _wildcard(name: str) -> bool:
    return "*" in name or "?" in name


def _matches(name: str, pattern: str) -> bool:
    """Windows's wildcards, whose case does not count and where *.* takes a name with no dot too."""
    if pattern in ("*", "*.*"):
        return True
    return fnmatch.fnmatchcase(name.casefold(), pattern.casefold().replace("[", "[[]"))


class OpenFile:
    """A file number and what stands behind it."""

    def __init__(self, path: str, mode: str, data: bytearray, position: int, relative: int) -> None:
        self.path, self.mode, self.data, self.position = path, mode, data, position
        #: Where Seek and Loc count from: an Append reports from where it started, until a Seek #.
        self.relative = relative
        self.column = 0
        self.width = 0
        self.changed = False
        #: What the Access clause, or a read-only file, rules out.
        self.can_read, self.can_write = True, True
        self.record = _DEFAULT_RECORD
        #: A Get or a read of text that ran off the end, which is what EOF reports in these modes.
        self.ran_out = False
        self.locks: list[tuple[int, int]] = []

    def write(self, data: bytes) -> None:
        end = self.position + len(data)
        if end > len(self.data):
            self.data.extend(b"\0" * (end - len(self.data)))
        self.data[self.position:end] = data
        self.position = end
        self.changed = True

    def exhausted(self) -> bool:
        return self.position >= len(self.data) or (self.mode == "input" and self.data[self.position] == _CTRL_Z)

    def at_end(self) -> bool:
        if self.mode in ("binary", "random"):
            return self.ran_out
        if self.mode != "input":
            return True
        return self.position >= len(self.data) or self.data[self.position] == _CTRL_Z

    def flush(self) -> None:
        if self.changed:
            with open(self.path, "wb") as stream:
                stream.write(self.data)
            self.changed = False


def _flush_all(handles: dict[int, OpenFile]) -> None:
    for handle in handles.values():
        try:
            handle.flush()
        except OSError:
            pass


class FileTable:
    """The file numbers, the current folder, Dir's list and the attributes of one interpreter."""

    def __init__(self, interpreter: Interpreter) -> None:
        self.interpreter = interpreter
        self.handles: dict[int, OpenFile] = {}
        self.folder = os.getcwd()
        self.listing: list[str] | None = None
        self.attributes: dict[str, int] = {}
        #: The TextStreams open, which Kill finds it cannot get past.
        self.streams: list[Any] = []
        #: VBA closes what is left open when the project goes, so nothing written is lost when the object does.
        weakref.finalize(self, _flush_all, self.handles)

    # -- paths

    def full(self, path: object) -> str:
        return spelled(os.path.join(self.folder, native_path(to_text(path))))

    def _key(self, path: str) -> str:
        return os.path.normcase(os.path.abspath(path))

    def attributes_of(self, path: str) -> int:
        kept = self.attributes.get(self._key(path))
        if kept is not None:
            return kept
        found = os.stat(path)
        bits = 16 if stat.S_ISDIR(found.st_mode) else 32
        return bits | (0 if found.st_mode & stat.S_IWRITE else 1)

    def read_only(self, path: str) -> bool:
        return os.path.exists(path) and bool(self.attributes_of(path) & 1)

    def is_open(self, path: str, *, for_writing: bool = False) -> bool:
        key = self._key(path)
        return any(self._key(one.path) == key and (one.mode != "input" or not for_writing)
                   for one in self.handles.values())

    # -- file numbers

    def number(self, value: object) -> int:
        """A file number: at most an Integer, 1 to 512 (tests/fixtures/file_statements.json)."""
        wanted = int(to_integer(value, "Integer"))
        if not 1 <= wanted <= 512:
            raise error(ERR_BAD_NUMBER)
        return wanted

    def handle(self, value: object) -> OpenFile:
        found = self.handles.get(int(to_integer(value, "Integer")))
        if found is None:
            raise error(ERR_BAD_NUMBER)
        return found

    def reading(self, value: object) -> OpenFile:
        found = self.handle(value)
        if found.mode not in ("input", "binary"):
            raise error(ERR_BAD_MODE)
        return found

    def writing(self, value: object) -> OpenFile:
        found = self.handle(value)
        if found.mode not in ("output", "append"):
            raise error(ERR_BAD_MODE)
        return found

    def open(self, path_value: object, mode: str, number_value: object, *, access: str = "",
             length: object = MISSING) -> None:
        self.interpreter.outside("Open")
        if mode not in _MODES:
            raise VBAUnsupportedError(f"Open for {mode} is not implemented")
        record = _DEFAULT_RECORD if length is MISSING else int(to_integer(length, "Integer"))
        if record == 0:
            raise error(ERR_INVALID)
        # A negative Len was taken as it stood in 16 bits (tests/fixtures/binary_files.json).
        record &= 0xFFFF
        given = to_text(path_value)
        number = self.number(number_value)
        if number in self.handles:
            raise error(ERR_ALREADY_OPEN)
        if not given:
            raise error(ERR_ACCESS)
        path = self.full(given)
        if self._conflicts(path, mode):
            raise error(ERR_ALREADY_OPEN)
        shared = mode in ("binary", "random")
        if os.path.isdir(path) or (mode != "input" and not shared and self.read_only(path)):
            raise error(ERR_ACCESS)
        if shared:
            self._open_shared(number, path, mode, access, record)
            return
        try:
            if mode == "input":
                with open(path, "rb") as stream:
                    data = bytearray(stream.read())
                self.handles[number] = OpenFile(path, mode, data, 0, 0)
                return
            data = bytearray()
            if mode == "append" and os.path.exists(path):
                with open(path, "rb") as stream:
                    data = bytearray(stream.read())
            else:
                with open(path, "wb"):
                    pass
        except FileNotFoundError:
            parent = os.path.dirname(path)
            raise error(ERR_NOT_FOUND if mode == "input" and os.path.isdir(parent) else ERR_NO_PATH) from None
        except OSError:
            raise error(ERR_ACCESS) from None
        self.handles[number] = OpenFile(path, mode, data, len(data), len(data))

    def _conflicts(self, path: str, mode: str) -> bool:
        """Whether a file open already rules this mode out."""
        key = self._key(path)
        family = _FAMILIES.get(mode)
        return any(self._key(one.path) == key and (family is None or _FAMILIES.get(one.mode) != family)
                   for one in self.handles.values())

    def _open_shared(self, number: int, path: str, mode: str, access: str, record: int) -> None:
        self.interpreter.outside("Open")
        read_only = self.read_only(path)
        try:
            if os.path.exists(path):
                with open(path, "rb") as stream:
                    data = bytearray(stream.read())
            else:
                data = bytearray()
                with open(path, "wb"):
                    pass
        except FileNotFoundError:
            raise error(ERR_NO_PATH) from None
        except OSError:
            raise error(ERR_ACCESS) from None
        handle = OpenFile(path, mode, data, 0, 0)
        handle.record = record
        handle.can_read = access != "write"
        handle.can_write = access != "read" and not read_only
        self.handles[number] = handle

    def record_handle(self, number_value: object) -> OpenFile:
        found = self.handle(number_value)
        if found.mode not in ("binary", "random"):
            raise error(ERR_BAD_MODE)
        return found

    def _place(self, handle: OpenFile, where: object) -> int:
        """Where a Get or Put starts: a byte in a Binary file, a record in a Random one."""
        if where is MISSING:
            return handle.position
        place = int(to_integer(where, "Long"))
        if place < 1:
            raise error(ERR_BAD_RECORD)
        return place - 1 if handle.mode == "binary" else (place - 1) * handle.record

    def put(self, handle: OpenFile, where: object, value: object, declared: str) -> None:
        self.interpreter.outside("Put #")
        random = handle.mode == "random"
        place = self._place(handle, where)
        if random and isinstance(value, VBAArray) and not value.fixed and len(value.bounds) > 1:
            # Its dimensions take eight bytes each, whose order was not measured: only a record too short answers.
            body = b"".join(encode(one, value.element_type, nested=True) for one in value.items)
            if 2 + 8 * len(value.bounds) + len(body) > handle.record:
                raise error(ERR_RECORD_LENGTH)
            raise VBAUnsupportedError("Put of a dynamic array of more than one dimension is not implemented")
        data = encode(value, declared, nested=random)
        if not handle.can_write:
            raise error(ERR_ACCESS)
        if random and len(data) > handle.record:
            raise error(ERR_RECORD_LENGTH)
        handle.position = place
        handle.write(data)
        handle.ran_out = False
        handle.relative = 0
        if random:
            handle.position = place + handle.record

    def get(self, handle: OpenFile, where: object, value: object, declared: str) -> object:
        self.interpreter.outside("Get #")
        random = handle.mode == "random"
        place = self._place(handle, where)
        if not handle.can_read:
            raise error(ERR_ACCESS)
        reader = Reader(handle.data, place, handle.record if random else None)
        handle.position = place
        result: object = None
        if isinstance(value, VBAArray):
            plan: list[tuple[VBAArray, int, object]] = []
            fill_array(reader, value, describe=random, plan=plan)
            for array, index, item in plan:
                array.items[index] = item
        elif hasattr(value, "fields") and hasattr(value, "defined"):
            fill_record(reader, value)
        else:
            result = decode(reader, declared or "variant", value, nested=random)
        handle.position = place + handle.record if random else reader.at
        handle.ran_out = reader.short
        return result

    def lock(self, locking: bool, number_value: object, first_value: object, last_value: object,
             has_range: bool) -> None:
        self.interpreter.outside("Lock #" if locking else "Unlock #")
        handle = self.handle(number_value)
        first = 1 if first_value is MISSING else int(to_integer(first_value, "Long"))
        last = 0x7FFFFFFF if not has_range else first if last_value is MISSING else int(to_integer(last_value, "Long"))
        if first < 1 or last < first:
            raise error(ERR_BAD_RECORD)
        span = (first, last)
        if locking:
            if any(one[0] <= last and first <= one[1] for one in handle.locks):
                raise error(ERR_DENIED)
            handle.locks.append(span)
        elif span in handle.locks:
            handle.locks.remove(span)
        else:
            raise error(ERR_DENIED)

    def close(self, numbers: list[object]) -> None:
        self.interpreter.outside("Close")
        wanted = [self.number(one) for one in numbers] if numbers else list(self.handles)
        for number in wanted:
            found = self.handles.pop(number, None)
            if found is not None:
                found.flush()

    # -- Print # and Write #

    def _emit(self, handle: OpenFile, text: str) -> None:
        handle.write(_encode(text))
        cut = max(text.rfind("\r"), text.rfind("\n"))
        handle.column = len(text) - cut - 1 if cut >= 0 else handle.column + len(text)

    def print_item(self, handle: OpenFile, item: object) -> None:
        """One item of a Print #: a number stands between spaces, a Date has one after it."""
        if isinstance(item, TabTo):
            self._tab(handle, item.column)
            return
        if isinstance(item, Spaces):
            self._emit(handle, " " * max(item.count, 0))
            return
        text = self._printed(item)
        if handle.width and handle.column and handle.column + len(text) > handle.width:
            self._emit(handle, "\r\n")
        self._emit(handle, text)

    def _tab(self, handle: OpenFile, column: int | None) -> None:
        target = (handle.column // _ZONE + 1) * _ZONE if column is None else max(column, 1) - 1
        if target < handle.column:
            self._emit(handle, "\r\n")
        self._emit(handle, " " * (target - handle.column))

    def print_zone(self, handle: OpenFile) -> None:
        self._tab(handle, None)

    def _printed(self, item: object) -> str:
        if item is NOTHING:
            raise error(ERR_NO_OBJECT)
        if isinstance(item, VBAObject):
            item = item.vba_value()
        if isinstance(item, VBAArray):
            raise error(ERR_MISMATCH)
        if item is NULL:
            return "Null"
        if item is EMPTY or item is MISSING:
            return ""
        if isinstance(item, VBAErrorValue):
            return f"Error {item.number}"
        if isinstance(item, VBADate):
            return date_text(item) + " "
        if isinstance(item, bool) or not isinstance(item, (int, float, Decimal)):
            return to_text(item)
        text = to_text(item)
        return (text if text.startswith("-") else f" {text}") + " "

    def end_line(self, handle: OpenFile) -> None:
        self._emit(handle, "\r\n")

    def written(self, item: object) -> str:
        """One item of a Write #, as Input # will read it back."""
        if item is NOTHING:
            raise error(ERR_NO_OBJECT)
        if isinstance(item, VBAObject):
            item = item.vba_value()
        if isinstance(item, VBAArray):
            raise error(ERR_MISMATCH)
        if item is NULL:
            return "#NULL#"
        if item is EMPTY or item is MISSING:
            return ""
        if isinstance(item, VBAErrorValue):
            return f"#ERROR {item.number}#"
        if isinstance(item, bool):
            return "#TRUE#" if item else "#FALSE#"
        if isinstance(item, VBADate):
            when = item.to_datetime()
            whole = item.serial == int(item.serial)
            return f"#{when:%Y-%m-%d}#" if whole else f"#{when:%Y-%m-%d %H:%M:%S}#"
        if isinstance(item, (int, float, Decimal)):
            return re.sub(r"^(-?)0\.", r"\1.", to_text(item))
        doubled = to_text(item).replace('"', '""')
        return f'"{doubled}"'

    def write_items(self, handle: OpenFile, items: list[str], hold: bool) -> None:
        text = ",".join(items)
        self._emit(handle, f"{text}," if hold else f"{text}\r\n")

    def set_width(self, number_value: object, width_value: object) -> None:
        self.interpreter.outside("Width #")
        handle = self.handle(number_value)
        if handle.mode == "input":
            raise error(ERR_BAD_MODE)
        width = int(to_integer(width_value, "Integer"))
        if not 0 <= width <= 255:
            raise error(ERR_INVALID)
        handle.width = width

    # -- Line Input # and Input #

    def read_line(self, number_value: object) -> str:
        self.interpreter.outside("Line Input #")
        handle = self.reading(number_value)
        if handle.exhausted():
            raise error(ERR_PAST_END)
        data, position = handle.data, handle.position
        out = bytearray()
        while position < len(data) and data[position] != _CTRL_Z:
            byte = data[position]
            position += 1
            if byte == 13:
                if position < len(data) and data[position] == 10:
                    position += 1
                break
            if byte:
                out.append(byte)
        handle.position = position
        handle.ran_out = position >= len(data)
        return _decode(bytes(out))

    def read_chars(self, count_value: object, number_value: object) -> str:
        self.interpreter.outside("Input")
        count = int(to_integer(count_value, "Long"))
        handle = self.reading(number_value)
        if count < 0:
            raise error(ERR_INVALID)
        if count and handle.exhausted():
            raise error(ERR_PAST_END)
        piece = handle.data[handle.position:handle.position + count]
        if len(piece) < count:
            raise error(ERR_PAST_END)
        handle.position += count
        return _decode(bytes(piece))

    def read_field(self, handle: OpenFile) -> tuple[bool, str]:
        """The next field of an Input #: whether it was quoted, and its text. A quote ends its field at once, and a
        number with spaces after it ends there too, so what follows either is the next field."""
        data = handle.data
        if handle.exhausted():
            raise error(ERR_PAST_END)
        position = handle.position
        while position < len(data) and data[position] == 32:
            position += 1
        if position < len(data) and data[position] == 34:
            end = data.find(b'"', position + 1)
            end = len(data) if end < 0 else end
            text = _decode(bytes(data[position + 1:end]))
            handle.position = self._past_delimiter(data, end + 1)
            return True, text
        end = position
        while end < len(data) and data[end] not in (44, 13, 10):
            end += 1
        raw = _decode(bytes(data[position:end]))
        number = _NUMBER.match(raw)
        if number and raw[number.end():number.end() + 1] == " " and raw[number.end():].strip():
            after = position + number.end()
            while data[after] == 32:
                after += 1
            handle.position = after
            return False, number.group()
        handle.position = self._past_delimiter(data, end)
        handle.ran_out = end >= len(data)
        return False, raw.rstrip(" ")

    @staticmethod
    def _past_delimiter(data: bytearray, position: int) -> int:
        while position < len(data) and data[position] == 32:
            position += 1
        if position < len(data) and data[position] == 44:
            return position + 1
        if position < len(data) and data[position] == 13:
            position += 1
            return position + 1 if position < len(data) and data[position] == 10 else position
        if position < len(data) and data[position] == 10:
            return position + 1
        return position

    def field_value(self, quoted: bool, text: str, declared: str) -> object:
        """What a field gives a variable of this type (tests/fixtures/file_statements.json)."""
        kind = declared.lower()
        if kind == "string":
            return text
        mark = None if quoted else self._mark(text)
        number = None if quoted else self._number(text)
        if kind == "variant":
            if quoted:
                return text
            if not text:
                return EMPTY
            if mark is not None:
                return mark
            return text if number is None else number
        if kind == "boolean":
            if isinstance(mark, bool):
                return mark
            if number is not None:
                raise error(ERR_OVERFLOW)
            return False
        if kind == "date":
            if isinstance(mark, VBADate):
                return mark
            if number is not None:
                raise error(ERR_OVERFLOW)
            return VBADate(0.0)
        if kind in _WHOLES or kind in _FLOATS:
            return VBAInt(0, "Integer") if number is None else number
        raise VBAUnsupportedError(f"Input # into a {declared} is not implemented")

    @staticmethod
    def _mark(text: str) -> object:
        found = _MARK.match(text)
        if not found or text[0] != "#":
            return None
        inner = found.group(1).strip().upper()
        if inner == "TRUE":
            return True
        if inner == "FALSE":
            return False
        if inner == "NULL":
            return NULL
        if inner.startswith("ERROR "):
            return VBAErrorValue(int(inner[6:]))
        for form in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%H:%M:%S", "%H:%M"):
            try:
                when = _dt.datetime.strptime(inner, form)
            except ValueError:
                continue
            return VBADate.from_datetime(when if "%Y" in form else when.time())
        return None

    @staticmethod
    def _number(text: str) -> object:
        """A number at the start of the text, in the narrowest type: Integer, Long, Double; a fraction is Currency to
        four places and Decimal past them, and with an exponent, Double."""
        radix = _RADIX.match(text)
        if radix:
            value = int(radix.group(2), 16 if radix.group(1).lower() == "h" else 8)
            return VBAInt(value, "Integer" if value <= 0x7FFF else "Long")
        found = _NUMBER.match(text)
        if not found:
            return None
        word = found.group().lower().replace("d", "e")
        if "e" in word:
            return float(word)
        value = Decimal(word)
        if value == value.to_integral_value():
            whole = int(value)
            if -32768 <= whole <= 32767:
                return VBAInt(whole, "Integer")
            return VBAInt(whole, "Long") if -2147483648 <= whole <= 2147483647 else float(whole)
        places = -value.normalize().as_tuple().exponent  # type: ignore[operator]
        return VBACurrency(value) if places <= 4 else VBADecimal(value)

    # -- where a file stands

    def is_at_end(self, number_value: object) -> bool:
        self.interpreter.outside("EOF")
        return self.handle(number_value).at_end()

    def length(self, number_value: object) -> int:
        self.interpreter.outside("LOF")
        return len(self.handle(number_value).data)

    def location(self, number_value: object) -> int:
        self.interpreter.outside("Loc")
        handle = self.handle(number_value)
        if handle.mode == "input":
            return max(1, math.ceil(handle.position / 128))
        if handle.mode == "binary":
            return handle.position
        if handle.mode == "random":
            return handle.position // handle.record
        return (handle.position - handle.relative) // 128

    def seek_read(self, number_value: object) -> int:
        self.interpreter.outside("Seek")
        handle = self.handle(number_value)
        if handle.mode == "random":
            return handle.position // handle.record + 1
        return handle.position - handle.relative + 1

    def seek_to(self, number_value: object, position_value: object) -> None:
        self.interpreter.outside("Seek #")
        handle = self.handle(number_value)
        position = int(to_integer(position_value, "Long"))
        if position < 1:
            raise error(ERR_BAD_RECORD)
        handle.position = (position - 1) * (handle.record if handle.mode == "random" else 1)
        handle.relative = 0
        handle.ran_out = False

    def free_number(self, range_value: object) -> int:
        self.interpreter.outside("FreeFile")
        high = range_value is not MISSING and int(to_integer(range_value, "Integer")) != 0
        for number in range(256, 512) if high else range(1, 256):
            if number not in self.handles:
                return number
        raise error(67)

    def file_attribute(self, number_value: object, kind_value: object) -> int:
        self.interpreter.outside("FileAttr")
        handle = self.handle(number_value)
        if kind_value is not MISSING and int(to_integer(kind_value, "Integer")) != 1:
            raise VBAUnsupportedError("FileAttr of anything but the mode is not implemented")
        return _MODES[handle.mode]

    # -- the file system

    def kill(self, path_value: object) -> None:
        self.interpreter.outside("Kill")
        path = self.full(path_value)
        folder, pattern = os.path.split(path)
        if _wildcard(pattern):
            try:
                names = [name for name in sorted(os.listdir(folder), key=str.casefold)
                         if _matches(name, pattern) and not os.path.isdir(os.path.join(folder, name))]
            except OSError:
                names = []
            targets = [os.path.join(folder, name) for name in names]
        else:
            targets = [path] if to_text(path_value) and os.path.isfile(path) else []
        if not targets:
            raise error(ERR_NOT_FOUND)
        for target in targets:
            if self.is_open(target):
                raise error(ERR_ALREADY_OPEN)
            if any(self._key(one.path) == self._key(target) for one in self.streams):
                raise error(ERR_DENIED)
            if self.read_only(target):
                raise error(ERR_ACCESS)
            try:
                os.remove(target)
            except OSError:
                raise error(ERR_ACCESS) from None
            self.attributes.pop(self._key(target), None)

    def make_folder(self, path_value: object) -> None:
        self.interpreter.outside("MkDir")
        if not to_text(path_value):
            raise error(ERR_NO_PATH)
        try:
            os.mkdir(self.full(path_value))
        except FileNotFoundError:
            raise error(ERR_NO_PATH) from None
        except OSError:
            raise error(ERR_ACCESS) from None

    def remove_folder(self, path_value: object) -> None:
        self.interpreter.outside("RmDir")
        path = self.full(path_value)
        if not to_text(path_value) or not os.path.isdir(path):
            raise error(ERR_NO_PATH)
        try:
            os.rmdir(path)
        except OSError:
            raise error(ERR_ACCESS) from None

    def change_folder(self, path_value: object) -> None:
        self.interpreter.outside("ChDir")
        path = self.full(path_value)
        if not to_text(path_value) or not os.path.isdir(path):
            raise error(ERR_NO_PATH)
        self.folder = os.path.normpath(path)

    def current_folder(self, drive_value: object) -> str:
        self.interpreter.outside("CurDir")
        drive = "" if drive_value is MISSING else to_text(drive_value)[:1]
        if not drive or drive.casefold() == self.folder[:1].casefold():
            return self.folder
        if not drive.isalpha():
            raise error(ERR_INVALID)
        raise VBAUnsupportedError("CurDir of another drive is not implemented")

    def copy(self, source_value: object, target_value: object) -> None:
        self.interpreter.outside("FileCopy")
        source, target = self.full(source_value), self.full(target_value)
        if _wildcard(os.path.basename(source)):
            raise error(ERR_BAD_NUMBER)
        if not os.path.isfile(source):
            raise error(ERR_NOT_FOUND)
        if self.is_open(source, for_writing=True) or self.is_open(target):
            raise error(ERR_ALREADY_OPEN)
        if self.read_only(target):
            raise error(ERR_ACCESS)
        try:
            shutil.copyfile(source, target)
        except FileNotFoundError:
            raise error(ERR_NO_PATH) from None
        except OSError:
            raise error(ERR_ACCESS) from None

    def rename(self, old_value: object, new_value: object) -> None:
        self.interpreter.outside("Name")
        old, new = self.full(old_value), self.full(new_value)
        if not os.path.exists(old):
            raise error(ERR_NOT_FOUND)
        if self.is_open(old):
            raise error(ERR_ALREADY_OPEN)
        if os.path.exists(new):
            raise error(ERR_EXISTS)
        try:
            os.rename(old, new)
        except FileNotFoundError:
            raise error(ERR_NOT_FOUND) from None
        except OSError:
            raise error(ERR_ACCESS) from None
        kept = self.attributes.pop(self._key(old), None)
        if kept is not None:
            self.attributes[self._key(new)] = kept

    def file_length(self, path_value: object) -> int:
        self.interpreter.outside("FileLen")
        path = self.full(path_value)
        if not os.path.exists(path):
            raise error(ERR_NOT_FOUND)
        return 0 if os.path.isdir(path) else os.path.getsize(path)

    def file_time(self, path_value: object) -> VBADate:
        self.interpreter.outside("FileDateTime")
        path = self.full(path_value)
        if not os.path.exists(path):
            raise error(ERR_NOT_FOUND)
        return VBADate.from_datetime(_dt.datetime.fromtimestamp(os.path.getmtime(path)).replace(microsecond=0))

    def get_attributes(self, path_value: object) -> int:
        self.interpreter.outside("GetAttr")
        path = self.full(path_value)
        if not os.path.exists(path):
            raise error(ERR_NOT_FOUND)
        return self.attributes_of(path)

    def set_attributes(self, path_value: object, bits_value: object) -> None:
        self.interpreter.outside("SetAttr")
        path = self.full(path_value)
        bits = int(to_integer(bits_value, "Integer"))
        if not os.path.exists(path):
            raise error(ERR_NOT_FOUND)
        if bits & ~_SETTABLE:
            raise error(ERR_INVALID)
        os.chmod(path, stat.S_IREAD if bits & 1 else stat.S_IREAD | stat.S_IWRITE)
        self.attributes[self._key(path)] = bits

    def next_in_listing(self, path_value: object, bits_value: object) -> str:
        """Dir: given a path it takes the whole list there and then; without one it hands out the next, "" at the end,
        and error 5 once that has been said."""
        self.interpreter.outside("Dir")
        if path_value is MISSING:
            if self.listing is None:
                raise error(ERR_INVALID)
            if self.listing:
                return self.listing.pop(0)
            self.listing = None
            return ""
        if path_value is NULL:
            raise error(ERR_MISMATCH)
        bits = 0 if bits_value is MISSING else int(to_integer(bits_value, "Integer"))
        if bits & ~_DIR_BITS:
            raise error(ERR_INVALID)
        self.listing = self._listed(to_text(path_value), bits)
        if self.listing:
            return self.listing.pop(0)
        self.listing = None
        return ""

    def _listed(self, given: str, bits: int) -> list[str]:
        if bits == _VOLUME:
            return []
        path = self.full(given)
        folder, pattern = os.path.split(path) if given else (self.folder, "*")
        pattern = pattern or "*"
        try:
            names = sorted(os.listdir(folder), key=str.casefold)
        except OSError:
            return []
        wants_folders = bool(bits & 16)
        found = [name for name in (".", "..") if wants_folders and pattern in ("*", "*.*")]
        for name in names:
            if not _matches(name, pattern):
                continue
            full = os.path.join(folder, name)
            if os.path.isdir(full) and not wants_folders:
                continue
            kept = self.attributes.get(self._key(full), 0)
            if kept & 2 and not bits & 2 or kept & 4 and not bits & 4:
                continue
            found.append(name)
        return found

    def environment(self, value: object) -> str:
        self.interpreter.outside("Environ")
        if isinstance(value, str):
            if not value:
                raise error(ERR_INVALID)
            return os.environ.get(value, "")
        place = int(to_integer(value, "Integer"))
        if place < 1:
            raise error(ERR_INVALID)
        entries = list(os.environ.items())
        return "" if place > len(entries) else "=".join(entries[place - 1])


#: Names Windows gives to devices rather than files, whatever follows the dot.
_DEVICES = frozenset({'CON', 'PRN', 'AUX', 'NUL', 'CONIN$', 'CONOUT$', *(f'COM{at}' for at in range(1, 10)), *(f'LPT{at}' for at in range(1, 10))})


def _real(path: str) -> Path:
    return Path(os.path.realpath(spelled(path)))


def final_path(path: str) -> str:
    """The file a macro's path names, as this system opens it: a backslash a separator off Windows too, each part in the case the disk
    spells it, then links and .. resolved. A gate checks this and the read, write or open uses this, so nothing turns the path round between
    them: off Windows a backslash is a character of a name, and "grant/..\\away" checked as given is one file inside grant that a read
    turning its backslashes round would find outside."""
    return os.path.realpath(on_disk(path))


def is_inside(path: Path | str, folders: tuple[str, ...]) -> bool:
    """Whether a file lies in one of the folders once links and .. are resolved. Parts compare whole, so C:\\out2 is not in C:\\out; a UNC or device path has an anchor of its own, so it
    is in nothing but a folder of that kind; and the names by which Windows reaches a device or a file's other streams are in nothing."""
    real = Path(final_path(os.fspath(path)))
    if os.name == 'nt' and (':' in real.name or real.name.split('.')[0].rstrip(' ').upper() in _DEVICES):
        return False
    return any(real != base and real.is_relative_to(base) for base in (_real(one) for one in folders))
