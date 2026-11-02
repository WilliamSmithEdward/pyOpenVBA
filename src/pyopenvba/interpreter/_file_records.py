"""How Put lays out a value and Get reads it back, as tests/fixtures/binary_files.json measured them in live Excel.

Numbers go little-endian at their own size. A String at the top of a Binary file is its bytes alone, sized on reading by
what the variable holds; anywhere else it has a two-byte length before it. A fixed-length string is always just its
bytes. A Variant has its two-byte type first. An array at the top of a Binary file, or of a fixed size, is its
elements; a dynamic one anywhere else, or one in a Variant, has its dimensions first. A read past the end of the file
gets zeros."""
from __future__ import annotations

import struct
from decimal import Decimal
from typing import Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import (
    EMPTY,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBADecimal,
    VBAErrorValue,
    VBAInt,
    VBASingle,
    error,
    to_text,
)

ERR_RECORD_LENGTH = 59
ERR_NO_OBJECT = 91
ERR_AUTOMATION_TYPE = 458
#: The fixed-size types, by the name a declaration gives: the struct format, and the Variant type number.
_FORMATS: Final = {"byte": ("<B", 17), "integer": ("<h", 2), "long": ("<i", 3), "longptr": ("<q", 20),
                   "longlong": ("<q", 20), "single": ("<f", 4), "double": ("<d", 5), "currency": ("<q", 6),
                   "date": ("<d", 7), "boolean": ("<h", 11)}
_VARIANT_NAMES: Final = {number: name for name, (_, number) in _FORMATS.items() if name != "longptr"}
_EMPTY, _NULL, _STRING, _ERROR, _DECIMAL = 0, 1, 8, 10, 14
_ARRAY_FLAG: Final = 0x2000
_ERROR_BASE: Final = 0x800A0000


def _ansi(text: str) -> bytes:
    return text.encode("cp1252", errors="replace")


def _text(data: bytes) -> str:
    return "".join(bytes([byte]).decode("cp1252", errors="ignore") or chr(byte) for byte in data)


def kind_of(value: object) -> str:
    """The type a value holds, for a literal or a Variant, whose declaration does not say."""
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, VBAInt):
        return value.kind.lower()
    if isinstance(value, VBASingle):
        return "single"
    if isinstance(value, VBACurrency):
        return "currency"
    if isinstance(value, VBADecimal):
        return "decimal"
    if isinstance(value, VBADate):
        return "date"
    if isinstance(value, float):
        return "double"
    if isinstance(value, Decimal):
        return "decimal"
    if isinstance(value, int):
        return "integer" if -32768 <= value <= 32767 else "long"
    if isinstance(value, str):
        return "string"
    if value is EMPTY:
        return "empty"
    if value is NULL:
        return "null"
    if isinstance(value, VBAErrorValue):
        return "error"
    if value is NOTHING:
        raise error(ERR_NO_OBJECT)
    raise VBAUnsupportedError(f"Put of a {type(value).__name__} is not implemented")


def _scalar(value: object, kind: str) -> bytes:
    form = _FORMATS[kind][0]
    if kind == "boolean":
        return struct.pack(form, -1 if value else 0)
    if kind == "date":
        assert isinstance(value, VBADate)
        return struct.pack(form, value.serial)
    if kind == "currency":
        assert isinstance(value, Decimal)
        return struct.pack(form, int(value * 10000))
    if kind in ("single", "double"):
        assert isinstance(value, (int, float, Decimal))
        return struct.pack(form, float(value))
    assert isinstance(value, int)
    return struct.pack(form, int(value))


def _decimal(value: Decimal) -> bytes:
    """A DECIMAL as a Variant holds one, its first two bytes the Variant's type."""
    sign, digits, exponent = value.as_tuple()
    assert isinstance(exponent, int)
    whole = int("".join(map(str, digits)) or "0")
    scale = max(-exponent, 0)
    if exponent > 0:
        whole *= 10 ** exponent
    return struct.pack("<HBBIQ", _DECIMAL, scale, 0x80 if sign else 0, whole >> 64, whole & 0xFFFFFFFFFFFFFFFF)


def _counted(text: str) -> bytes:
    data = _ansi(text)
    return struct.pack("<H", len(data)) + data


def _dimensions(array: VBAArray) -> bytes:
    if len(array.bounds) != 1:
        raise VBAUnsupportedError("Put of a dynamic array of more than one dimension is not implemented")
    low, high = array.bounds[0]
    return struct.pack("<Hii", 1, high - low + 1, low)


def encode_variant(value: object) -> bytes:
    if isinstance(value, VBAArray):
        return _array_in_variant(value)
    if isinstance(value, VBAObject):
        raise VBAUnsupportedError("Put of an object in a Variant is not implemented")
    kind = kind_of(value)
    if kind == "empty":
        return struct.pack("<H", _EMPTY)
    if kind == "null":
        return struct.pack("<H", _NULL)
    if kind == "string":
        assert isinstance(value, str)
        return struct.pack("<H", _STRING) + _counted(value)
    if kind == "error":
        assert isinstance(value, VBAErrorValue)
        return struct.pack("<HI", _ERROR, _ERROR_BASE | value.number)
    if kind == "decimal":
        assert isinstance(value, Decimal)
        return struct.pack("<H", _DECIMAL) + _decimal(value)
    return struct.pack("<H", _FORMATS[kind][1]) + _scalar(value, kind)


def _array_in_variant(array: VBAArray) -> bytes:
    kind = array.element_type.lower()
    known = {"variant": 12, "string": _STRING, **{name: entry[1] for name, entry in _FORMATS.items()}}
    number = known.get(kind)
    if number is None:
        raise VBAUnsupportedError(f"Put of a Variant holding {array.element_type} elements is not implemented")
    head = struct.pack("<H", number | _ARRAY_FLAG) + _dimensions(array)
    return head + b"".join(encode(one, array.element_type, nested=True) for one in array.items)


def encode_array(array: VBAArray, *, describe: bool) -> bytes:
    """An array's elements, with its dimensions first where it is dynamic and not at the top of a Binary file."""
    if not array.bounds:
        return b""
    head = _dimensions(array) if describe and not array.fixed else b""
    return head + b"".join(encode(one, array.element_type, nested=True) for one in array.items)


def _declared_array(slot: object) -> bool:
    """A field declared as an array, as distinct from a Variant one that holds a copy."""
    held = getattr(slot, "value", None)
    return bool(getattr(slot, "array_of", "") or getattr(slot, "fixed_array", False) or (
        isinstance(held, VBAArray) and held.fixed))


def encode_record(record: object) -> bytes:
    """A Type's fields, one after the other with no padding."""
    fields = getattr(record, "fields")
    out = b""
    for slot in fields.values():
        held = slot.get()
        if isinstance(held, VBAArray) and _declared_array(slot):
            out += encode_array(held, describe=True)
        else:
            out += encode(held, slot.declared, nested=True)
    return out


def encode(value: object, declared: str, *, nested: bool) -> bytes:
    """A value of a declared type; ``nested`` where a String has its length: anywhere but the top of a Binary file."""
    kind = declared.lower()
    if kind.replace(" ", "").startswith("string*"):
        assert isinstance(value, str)
        return _ansi(value)
    if kind == "variant":
        return encode_variant(value)
    if hasattr(value, "fields") and hasattr(value, "defined"):
        return encode_record(value)
    if isinstance(value, VBAArray):
        return encode_array(value, describe=nested)
    if isinstance(value, VBAObject):
        raise VBAUnsupportedError("Put of an object is not implemented")
    if not kind:
        kind = kind_of(value)
    if kind == "string":
        text = to_text(value)
        return _counted(text) if nested else _ansi(text)
    if kind == "decimal":
        raise VBAUnsupportedError("Put of a Decimal outside a Variant is not implemented")
    if kind not in _FORMATS:
        raise VBAUnsupportedError(f"Put of a {declared} is not implemented")
    return _scalar(value, kind)


class Reader:
    """Bytes from a place in a file: zeros past its end, and in a Random file error 59 past the record."""

    def __init__(self, data: bytearray, start: int, limit: int | None) -> None:
        self.data, self.start, self.at, self.limit = data, start, start, limit
        self.short = False

    def take(self, count: int) -> bytes:
        if self.limit is not None and self.at + count - self.start > self.limit:
            raise error(ERR_RECORD_LENGTH)
        piece = bytes(self.data[self.at:self.at + count])
        if len(piece) < count:
            self.short = True
            piece += b"\0" * (count - len(piece))
        self.at += count
        return piece

    def number(self, form: str) -> int | float:
        return struct.unpack(form, self.take(struct.calcsize(form)))[0]


def _read_scalar(reader: Reader, kind: str) -> object:
    raw = reader.number(_FORMATS[kind][0])
    if kind == "boolean":
        return raw != 0
    if kind == "date":
        return VBADate(float(raw))
    if kind == "currency":
        return VBACurrency(Decimal(int(raw)) / 10000)
    if kind == "single":
        return VBASingle(float(raw))
    if kind == "double":
        return float(raw)
    assert isinstance(raw, int)
    return VBAInt(raw, {"byte": "Byte", "integer": "Integer", "long": "Long", "longlong": "LongLong",
                        "longptr": "LongLong"}[kind])


def _read_counted(reader: Reader) -> str:
    return _text(reader.take(int(reader.number("<H"))))


def decode_variant(reader: Reader) -> object:
    """A Variant's value. Its type goes first, and a type VBA has no use for is error 458 with the place unmoved."""
    start = reader.at
    tag = int(reader.number("<H"))
    if tag == _EMPTY:
        return EMPTY
    if tag == _NULL:
        return NULL
    if tag == _STRING:
        return _read_counted(reader)
    if tag == _ERROR:
        return VBAErrorValue(int(reader.number("<I")) & 0xFFFF)
    if tag == _DECIMAL:
        _, scale, sign, high, low = struct.unpack("<HBBIQ", reader.take(16))
        value = Decimal((high << 64) | low).scaleb(-scale)
        return VBADecimal(-value if sign else value)
    if tag in _VARIANT_NAMES:
        return _read_scalar(reader, _VARIANT_NAMES[tag])
    if tag & _ARRAY_FLAG:
        raise VBAUnsupportedError("Get of an array into a Variant is not implemented")
    reader.at = start
    raise error(ERR_AUTOMATION_TYPE)


def decode(reader: Reader, declared: str, current: object, *, nested: bool) -> object:
    """The value a variable of this type gets; arrays and records keep their shape and are filled through ``fill``."""
    kind = declared.lower()
    if kind.replace(" ", "").startswith("string*"):
        return _text(reader.take(int(kind.partition("*")[2])))
    if kind == "variant":
        return decode_variant(reader)
    if kind == "string":
        if nested:
            return _read_counted(reader)
        return _text(reader.take(len(to_text(current))))
    if kind in _FORMATS:
        return _read_scalar(reader, kind)
    raise VBAUnsupportedError(f"Get of a {declared} is not implemented")


def fill_array(reader: Reader, array: VBAArray, *, describe: bool, plan: list[tuple[VBAArray, int, object]]) -> None:
    if not array.bounds:
        return
    if describe and not array.fixed:
        raise VBAUnsupportedError("Get of a dynamic array with its dimensions is not implemented")
    for index, one in enumerate(array.items):
        if hasattr(one, "fields") and hasattr(one, "defined"):
            fill_record(reader, one)
        else:
            plan.append((array, index, decode(reader, array.element_type, one, nested=True)))


def fill_record(reader: Reader, record: object) -> None:
    fields = getattr(record, "fields")
    for slot in fields.values():
        held = slot.get()
        if hasattr(held, "fields") and hasattr(held, "defined"):
            fill_record(reader, held)
        elif isinstance(held, VBAArray) and _declared_array(slot):
            plan: list[tuple[VBAArray, int, object]] = []
            fill_array(reader, held, describe=True, plan=plan)
            for array, index, value in plan:
                array.items[index] = value
        else:
            slot.set(decode(reader, slot.declared, held, nested=True))
