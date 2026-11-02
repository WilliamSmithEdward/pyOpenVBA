"""ADODB.Stream, as tests/fixtures/adodb_stream.json measured it in live Excel: bytes in memory, and text through a
charset.

It reaches nothing until it saves or loads, so only SaveToFile and LoadFromFile are refused with the sandbox on.

A stream is its bytes and a place in them, whichever Type it has, and Type only says how they are read. Text written to
an empty stream starts with a byte order mark in UTF-8 and UTF-16. SaveToFile saves all of it, wherever the place is,
and leaves the place at 0, and LoadFromFile of text adds the mark to a file without one. Only the charsets below are
known: the two _autodetect names and x-user-defined answer by the machine, so they are refused."""
from __future__ import annotations

import unicodedata
from typing import TYPE_CHECKING, Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBAInt,
    error,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter

_BINARY, _TEXT = 1, 2
_READ_ALL, _READ_LINE = -1, -2
_CRLF, _LF, _CR = -1, 10, 13
_SEPARATORS: Final = {_CRLF: "\r\n", _LF: "\n", _CR: "\r"}
_MODES: Final = {0, 1, 2, 3, 4, 8, 12, 16, 0x400000}
_UTF16: Final = {"unicode", "utf-16", "unicodefffe"}
_CODECS: Final = {"iso-8859-1": "latin-1", "shift_jis": "cp932", "windows-1252": "cp1252"}

_ARGS = (3001, "Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.")
_OPENS = (3002, "File could not be opened.")
_WRITES = (3004, "Write to file failed.")
_CONTEXT = (3219, "Operation is not allowed in this context.")
_CLOSED = (3704, "Operation is not allowed when the object is closed.")
_IS_OPEN = (3705, "Operation is not allowed when the object is open.")
_MISMATCH = (-2147352571, "Type mismatch.")
_PARAMETER = (-2147024809, "The parameter is incorrect.")


def _known(charset: str) -> str:
    name = charset.lower()
    if name in _UTF16 or name == "utf-8" or name == "us-ascii" or name in _CODECS or name.startswith("windows-"):
        return name
    raise VBAUnsupportedError(f"ADODB.Stream with the charset {charset} is not implemented")


def _ascii(text: str) -> bytes:
    """US-ASCII by Windows's best fit: a letter loses its accent, and what has no fit is a question mark."""
    out = bytearray()
    for char in text:
        base = unicodedata.normalize("NFKD", char)[:1]
        out.append(ord(base) if base.isascii() and base else 63)
    return bytes(out)


def _utf8(data: bytes) -> tuple[str, list[int]]:
    """UTF-8 and the bytes each character took. A byte that cannot start a character is U+FFFD by itself. A sequence
    the wrong bytes break ends the text there, and one the end of the data cuts short is U+FFFD."""
    chars: list[str] = []
    widths: list[int] = []
    at = 0
    while at < len(data):
        byte = data[at]
        width = 1 if byte < 0x80 else 0
        if 0xC0 <= byte < 0xF5:
            width = 2 if byte < 0xE0 else 3 if byte < 0xF0 else 4
        if width == 0:
            chars.append("�")
            widths.append(1)
            at += 1
            continue
        piece = data[at:at + width]
        if any(not 0x80 <= one < 0xC0 for one in piece[1:]):
            widths.append(len(data) - at)
            break
        if len(piece) < width:
            chars.append("�")
            widths.append(len(piece))
            break
        try:
            chars.append(piece.decode("utf-8"))
        except UnicodeDecodeError:
            widths.append(len(data) - at)
            break
        widths.append(width)
        at += width
    return "".join(chars), widths


class AdoStream(VBAObject):
    vba_type_name = "Stream"

    def __init__(self, interpreter: Interpreter) -> None:
        self.interpreter = interpreter
        self.kind, self.mode, self.charset, self.separator = _TEXT, 0, "Unicode", _CRLF
        self.is_open = False
        self.data = bytearray()
        self.at = 0

    def _opened(self) -> None:
        if not self.is_open:
            raise error(*_CLOSED)

    def _text_only(self) -> None:
        self._opened()
        if self.kind != _TEXT:
            raise error(*_CONTEXT)

    @member
    def State(self) -> object:
        return VBAInt(1 if self.is_open else 0, "Long")

    @member
    def Type(self) -> object:
        return VBAInt(self.kind, "Long")

    @setter("Type")
    def _type(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if wanted not in (_BINARY, _TEXT):
            raise error(*_ARGS)
        if self.is_open and self.at:
            raise error(*_CONTEXT)
        self.kind = wanted

    @member
    def Mode(self) -> object:
        return VBAInt(self.mode, "Long")

    @setter("Mode")
    def _mode(self, value: object) -> None:
        wanted = int(to_integer(value, "Long"))
        if self.is_open:
            raise error(*_IS_OPEN)
        if wanted not in _MODES:
            raise error(*_ARGS)
        self.mode = wanted

    @member
    def Charset(self) -> object:
        if self.kind != _TEXT:
            raise error(*_CONTEXT)
        return self.charset

    @setter("Charset")
    def _charset(self, value: object) -> None:
        if self.kind != _TEXT or (self.is_open and self.at):
            raise error(*_CONTEXT)
        wanted = to_text(value)
        _known(wanted)
        self.charset = wanted

    @member
    def LineSeparator(self) -> object:
        if self.kind != _TEXT:
            raise error(*_CONTEXT)
        return VBAInt(self.separator, "Long")

    @setter("LineSeparator")
    def _line_separator(self, value: object) -> None:
        if self.kind != _TEXT:
            raise error(*_CONTEXT)
        wanted = int(to_integer(value, "Long"))
        if wanted not in _SEPARATORS:
            raise error(*_ARGS)
        self.separator = wanted

    @member
    def Position(self) -> object:
        self._opened()
        return VBAInt(self.at, "Long")

    @setter("Position")
    def _position(self, value: object) -> None:
        self._opened()
        wanted = int(to_integer(value, "Long"))
        if wanted < 0:
            raise error(*_ARGS)
        if wanted > len(self.data):
            raise error(*_PARAMETER)
        self.at = wanted

    @member
    def Size(self) -> object:
        self._opened()
        return VBAInt(len(self.data), "Long")

    @member
    def EOS(self) -> object:
        self._opened()
        return self.at >= len(self.data)

    @method
    def Open(self, Source: object = MISSING, Mode: object = MISSING, Options: object = MISSING,
             UserName: object = MISSING, Password: object = MISSING) -> object:
        del Mode, Options, UserName, Password
        if Source is not MISSING:
            raise error(*_ARGS)
        if self.is_open:
            raise error(*_IS_OPEN)
        self.is_open = True
        self.data, self.at = bytearray(), 0
        return EMPTY

    @method
    def Close(self) -> object:
        self._opened()
        self.is_open = False
        self.data, self.at = bytearray(), 0
        return EMPTY

    @method
    def Flush(self) -> object:
        self._opened()
        return EMPTY

    @method
    def Cancel(self) -> object:
        self._opened()
        return EMPTY

    @method
    def SetEOS(self) -> object:
        self._opened()
        del self.data[self.at:]
        return EMPTY

    def _put(self, data: bytes) -> None:
        end = self.at + len(data)
        self.data[self.at:end] = data
        self.at = end

    def _encoded(self, text: str) -> bytes:
        name = _known(self.charset)
        if name in _UTF16:
            return text.encode("utf-16-le", errors="surrogatepass")
        if name == "utf-8":
            return text.encode("utf-8", errors="surrogatepass")
        if name == "us-ascii":
            return _ascii(text)
        return text.encode(_CODECS.get(name, name.replace("windows-", "cp")), errors="replace")

    def _mark(self) -> bytes:
        name = _known(self.charset)
        return b"\xff\xfe" if name in _UTF16 else b"\xef\xbb\xbf" if name == "utf-8" else b""

    def _decoded(self) -> tuple[str, list[int]]:
        """The text from here on, and how many bytes each character takes, its byte order mark counted to the first."""
        name = _known(self.charset)
        data = bytes(self.data[self.at:])
        if name == "utf-8":
            skip = 3 if self.at == 0 and data.startswith(b"\xef\xbb\xbf") else 0
            text, widths = _utf8(data[skip:])
            return text, _with_mark(widths, skip)
        if name in _UTF16:
            big = self.at == 0 and data.startswith(b"\xfe\xff")
            skip = 2 if self.at == 0 and data[:2] in (b"\xff\xfe", b"\xfe\xff") else 0
            body = data[skip:]
            whole = body[:len(body) // 2 * 2]
            text = whole.decode("utf-16-be" if big else "utf-16-le", errors="surrogatepass")
            widths = [2] * len(text)
            return text, _with_mark(widths, skip)
        if name == "us-ascii":
            return "".join(chr(byte & 0x7F) for byte in data), [1] * len(data)
        codec = _CODECS.get(name, name.replace("windows-", "cp"))
        text = data.decode(codec, errors="replace")
        return text, [len(char.encode(codec, errors="replace")) for char in text]

    def _consume(self, widths: list[int], count: int) -> None:
        self.at = min(self.at + sum(widths[:count]), len(self.data))

    @method
    def ReadText(self, NumChars: object = MISSING) -> object:
        wanted = _READ_ALL if NumChars is MISSING else int(to_integer(NumChars, "Long"))
        self._text_only()
        if wanted < _READ_LINE:
            raise error(*_ARGS)
        text, widths = self._decoded()
        if wanted == _READ_ALL:
            self.at = len(self.data)
            return text
        if wanted == _READ_LINE:
            mark = _SEPARATORS[self.separator]
            found = text.find(mark)
            if found < 0:
                self.at = len(self.data)
                return text
            self._consume(widths, found + len(mark))
            return text[:found]
        self._consume(widths, wanted)
        return text[:wanted]

    @method
    def SkipLine(self) -> object:
        self._text_only()
        text, widths = self._decoded()
        mark = _SEPARATORS[self.separator]
        found = text.find(mark)
        if found < 0:
            self.at = len(self.data)
        else:
            self._consume(widths, found + len(mark))
        return EMPTY

    @method
    def WriteText(self, Data: object = MISSING, Options: object = MISSING) -> object:
        if Data is NULL:
            raise error(*_MISMATCH)
        # A Boolean goes as a number, -1 or 0.
        text = "" if Data is MISSING or Data is EMPTY else str(-int(Data)) if isinstance(Data, bool) else to_text(Data)
        line = False if Options is MISSING else int(to_integer(Options, "Long"))
        if line not in (0, 1):
            raise error(*_ARGS)
        self._text_only()
        if not self.data:
            self._put(self._mark())
        self._put(self._encoded(text + (_SEPARATORS[self.separator] if line else "")))
        return EMPTY

    @method
    def Read(self, NumBytes: object = MISSING) -> object:
        wanted = _READ_ALL if NumBytes is MISSING else int(to_integer(NumBytes, "Long"))
        self._opened()
        if self.kind != _BINARY:
            raise error(*_CONTEXT)
        if wanted < _READ_ALL:
            raise error(*_ARGS)
        piece = bytes(self.data[self.at:] if wanted == _READ_ALL else self.data[self.at:self.at + wanted])
        if not piece:
            return NULL
        self.at += len(piece)
        return VBAArray([(0, len(piece) - 1)], element_type="Byte", items=[VBAInt(byte, "Byte") for byte in piece])

    @method
    def Write(self, Buffer: object = MISSING) -> object:
        if not isinstance(Buffer, VBAArray) or Buffer.element_type != "Byte" or not Buffer.items:
            raise error(*_ARGS)
        self._opened()
        if self.kind != _BINARY:
            raise error(*_CONTEXT)
        self._put(bytes(int(to_integer(item, "Byte")) for item in Buffer.items))
        return EMPTY

    @method
    def SaveToFile(self, FileName: object = MISSING, Options: object = MISSING) -> object:
        self.interpreter.outside("ADODB.Stream.SaveToFile")
        if FileName is NULL:
            raise error(*_MISMATCH)
        given = to_text(FileName)
        option = 1 if Options is MISSING else int(to_integer(Options, "Long"))
        if not given or option not in (1, 2):
            raise error(*_ARGS)
        self._opened()
        table = self.interpreter.files
        path = table.full(given)
        if "://" in given or table.is_open(path) or table.read_only(path):
            raise error(*_WRITES)
        try:
            with open(path, "xb" if option == 1 else "wb") as stream:
                stream.write(self.data)
        except OSError:
            raise error(*_WRITES) from None
        self.at = 0
        return EMPTY

    @method
    def LoadFromFile(self, FileName: object = MISSING) -> object:
        self.interpreter.outside("ADODB.Stream.LoadFromFile")
        if FileName is NULL:
            raise error(*_MISMATCH)
        given = to_text(FileName)
        if not given:
            raise error(*_ARGS)
        self._opened()
        table = self.interpreter.files
        path = table.full(given)
        if table.is_open(path, for_writing=True):
            raise error(*_OPENS)
        try:
            with open(path, "rb") as stream:
                self.data = bytearray(stream.read())
        except OSError:
            raise error(*_OPENS) from None
        if self.kind == _TEXT:
            # A text stream holds its byte order mark, and adds one to a file without.
            mark = self._mark()
            big = _known(self.charset) in _UTF16 and self.data.startswith(b"\xfe\xff")
            if mark and not big and not self.data.startswith(mark):
                self.data[0:0] = mark
        self.at = 0
        return EMPTY

    @method
    def CopyTo(self, DestStream: object = MISSING, CharNumber: object = MISSING) -> object:
        if DestStream is NOTHING:
            raise error(*_ARGS)
        if not isinstance(DestStream, AdoStream):
            raise error(*_MISMATCH)
        self._opened()
        if not DestStream.is_open:
            raise error(*_ARGS)
        wanted = -1 if CharNumber is MISSING else int(to_integer(CharNumber, "Long"))
        if self.kind == _BINARY and DestStream.kind == _TEXT:
            raise error(*_CONTEXT)
        if self.kind == _TEXT and DestStream.kind == _TEXT:
            if wanted != -1:
                raise VBAUnsupportedError("CopyTo of some of a text stream to another is not implemented")
            text, _ = self._decoded()
            self.at = len(self.data)
            if not DestStream.data:
                DestStream._put(DestStream._mark())
            DestStream._put(DestStream._encoded(text))
            return EMPTY
        piece = bytes(self.data[self.at:] if wanted < 0 else self.data[self.at:self.at + wanted])
        self.at += len(piece)
        DestStream._put(piece)
        return EMPTY


def _with_mark(widths: list[int], skip: int) -> list[int]:
    if not skip:
        return widths
    return [widths[0] + skip, *widths[1:]] if widths else [skip]

