"""Row and value codecs for Jet 4 / ACE data pages.

A row is::

    u16  column count the row was written with
    ...  fixed-length column data, at each column's fixed offset
    ...  variable-length column data
    u16  (var count + 1) offsets, stored in REVERSE order: the last one in
         memory is the start of variable column 0, the first in memory is
         the end of the last variable column
    u16  variable column count
    ...  null mask, one bit per column number, set when the column HAS a
         value; Boolean columns live only here

Values are decoded to plain Python types.  Long values (Memo, OLE) decode
to their 12-byte definition, which :mod:`pyopenvba.access._lval` resolves,
because resolving them needs the rest of the file.
"""

from __future__ import annotations

import datetime as _dt
import struct
import uuid
from dataclasses import dataclass
from decimal import Decimal

from pyopenvba.access_read import AccessError
from pyopenvba.access._tdef import (
    TYPE_BIGINT,
    TYPE_BINARY,
    TYPE_BOOLEAN,
    TYPE_BYTE,
    TYPE_COMPLEX,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_EXTENDED_DATETIME,
    TYPE_FLOAT,
    TYPE_GUID,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MEMO,
    TYPE_MONEY,
    TYPE_NUMERIC,
    TYPE_OLE,
    TYPE_TEXT,
    ColumnDef,
    TableDefinition,
)

EPOCH = _dt.datetime(1899, 12, 30)
TEXT_COMPRESSION_MARK = b"\xff\xfe"


@dataclass
class LongValueRef:
    """The 12-byte long-value definition stored in a row."""

    length: int
    kind: int
    inline: bytes
    row: int
    page: int

    KIND_INLINE = 0x80
    KIND_SINGLE_PAGE = 0x40
    KIND_CHAINED = 0x00


@dataclass
class RawRow:
    """A row split into per-column byte slices, before value decoding."""

    column_count: int
    values: dict[int, bytes | None]
    present: dict[int, bool]


def split_row(definition: TableDefinition, row: bytes) -> RawRow:
    if len(row) < 2:
        raise AccessError("row shorter than its column-count word")
    row_columns = struct.unpack_from("<H", row, 0)[0]
    null_mask_length = (row_columns + 7) // 8
    if len(row) < 2 + null_mask_length:
        raise AccessError("row shorter than its null mask")
    null_mask = row[len(row) - null_mask_length :]
    has_var = definition.var_column_count > 0
    var_offsets: list[int] = []
    if has_var:
        count_pos = len(row) - null_mask_length - 2
        if count_pos < 2:
            raise AccessError("row shorter than its variable-column count")
        var_count = struct.unpack_from("<H", row, count_pos)[0]
        table_pos = count_pos - 2 * (var_count + 1)
        if table_pos < 2:
            raise AccessError("row shorter than its variable-column offsets")
        reversed_offsets = struct.unpack_from(f"<{var_count + 1}H", row, table_pos)
        # Memory order is [end, start(n-1), ..., start(0)].
        var_offsets = list(reversed(reversed_offsets))
    values: dict[int, bytes | None] = {}
    present: dict[int, bool] = {}
    for column in definition.columns:
        number = column.number
        if number >= row_columns:
            # Added after this row was written: no data, no mask bit.
            values[number] = None
            present[number] = False
            continue
        has_value = bool(null_mask[number // 8] & (1 << (number % 8)))
        present[number] = has_value
        if not has_value:
            values[number] = None
            continue
        if column.type_code == TYPE_BOOLEAN:
            values[number] = b""
            continue
        if column.is_fixed:
            start = 2 + column.fixed_offset
            values[number] = row[start : start + column.length]
            continue
        if column.var_index + 1 >= len(var_offsets):
            raise AccessError(
                f"column {column.name!r} has variable index {column.var_index} "
                f"but the row holds {max(len(var_offsets) - 1, 0)} variable columns"
            )
        start = var_offsets[column.var_index]
        end = var_offsets[column.var_index + 1]
        if not 2 <= start <= end <= len(row):
            raise AccessError(
                f"column {column.name!r} spans {start}..{end} in a {len(row)}-byte row"
            )
        values[number] = row[start:end]
    return RawRow(row_columns, values, present)


# --- text ------------------------------------------------------------------


def decode_text(raw: bytes) -> str:
    """Jet 4 text: UTF-16LE, or the compressed form that starts FF FE where
    a lone 0x00 byte toggles between one byte per character (Latin-1) and
    two."""
    if raw[:2] != TEXT_COMPRESSION_MARK:
        if len(raw) % 2:
            raise AccessError(f"odd-length UTF-16 text ({len(raw)} bytes)")
        return raw.decode("utf-16-le")
    units = bytearray()
    compressed = True
    i = 2
    while i < len(raw):
        if raw[i] == 0:
            compressed = not compressed
            i += 1
        elif compressed:
            units += bytes((raw[i], 0))
            i += 1
        else:
            if i + 1 >= len(raw):
                raise AccessError("compressed text ends inside a UTF-16 unit")
            units += raw[i : i + 2]
            i += 2
    return units.decode("utf-16-le")


def encode_text(text: str) -> bytes:
    """Compress when every character fits one byte, as Access does."""
    if text and all(1 <= ord(ch) <= 0xFF for ch in text):
        return TEXT_COMPRESSION_MARK + text.encode("latin-1")
    return text.encode("utf-16-le")


# --- fixed-size scalars -----------------------------------------------------


def decode_datetime(raw: bytes) -> _dt.datetime:
    value = struct.unpack("<d", raw)[0]
    days = int(value)
    fraction = abs(value - days)
    return EPOCH + _dt.timedelta(days=days) + _dt.timedelta(days=fraction)


def encode_datetime(value: _dt.datetime) -> bytes:
    delta = value - EPOCH
    days = delta.days
    fraction = (delta - _dt.timedelta(days=days)) / _dt.timedelta(days=1)
    if days < 0:
        # OLE dates keep the day negative and the time positive.
        days += 1
        fraction = 1 - fraction if fraction else 0.0
        number = days - fraction if days <= 0 else days + fraction
    else:
        number = days + fraction
    return struct.pack("<d", number)


def decode_numeric(raw: bytes, scale: int) -> Decimal:
    if len(raw) != 17:
        raise AccessError(f"Decimal value is {len(raw)} bytes, not 17")
    negative = raw[0] != 0
    words = struct.unpack("<4I", raw[1:])
    magnitude = 0
    for word in words:
        magnitude = (magnitude << 32) | word
    value = Decimal(magnitude).scaleb(-scale)
    return -value if negative else value


def encode_numeric(value: Decimal, scale: int) -> bytes:
    scaled = int((value.copy_abs() * (Decimal(10) ** scale)).to_integral_value())
    if scaled >= 1 << 128:
        raise AccessError("Decimal value does not fit 16 bytes")
    words = [(scaled >> shift) & 0xFFFFFFFF for shift in (96, 64, 32, 0)]
    return bytes((0x80 if value < 0 else 0,)) + struct.pack("<4I", *words)


def decode_long_value_ref(raw: bytes) -> LongValueRef:
    if len(raw) < 12:
        raise AccessError(f"long-value definition is {len(raw)} bytes, not 12+")
    length = int.from_bytes(raw[0:3], "little")
    kind = raw[3]
    row = raw[4]
    page = int.from_bytes(raw[5:8], "little")
    inline = raw[12:] if kind == LongValueRef.KIND_INLINE else b""
    if kind == LongValueRef.KIND_INLINE and len(inline) != length:
        raise AccessError(
            f"inline long value declares {length} bytes but carries {len(inline)}"
        )
    return LongValueRef(length, kind, inline, row, page)


def decode_scalar(column: ColumnDef, raw: bytes) -> object:
    """Decode a non-long-value column.  ``raw`` is the present value."""
    code = column.type_code
    if code == TYPE_BOOLEAN:
        return True
    if code == TYPE_BYTE:
        return raw[0]
    if code == TYPE_INT:
        return struct.unpack("<h", raw)[0]
    if code == TYPE_LONG:
        return struct.unpack("<i", raw)[0]
    if code == TYPE_MONEY:
        return Decimal(struct.unpack("<q", raw)[0]).scaleb(-4)
    if code == TYPE_FLOAT:
        return struct.unpack("<f", raw)[0]
    if code == TYPE_DOUBLE:
        return struct.unpack("<d", raw)[0]
    if code == TYPE_DATETIME:
        return decode_datetime(raw)
    if code == TYPE_BINARY:
        return bytes(raw)
    if code == TYPE_TEXT:
        return decode_text(raw)
    if code == TYPE_GUID:
        return uuid.UUID(bytes_le=bytes(raw))
    if code == TYPE_NUMERIC:
        return decode_numeric(raw, column.scale)
    if code == TYPE_COMPLEX:
        return struct.unpack("<i", raw)[0]
    if code == TYPE_BIGINT:
        return struct.unpack("<q", raw)[0]
    if code == TYPE_EXTENDED_DATETIME:
        return bytes(raw)
    if code in (TYPE_OLE, TYPE_MEMO):
        return decode_long_value_ref(raw)
    return bytes(raw)
