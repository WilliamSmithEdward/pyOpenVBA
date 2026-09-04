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
from decimal import ROUND_HALF_EVEN, Decimal

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
    """A row split into per-column byte slices, before value decoding.
    ``raw`` and ``var_offsets`` keep the whole row so a re-encoding can
    carry the bytes of columns the definition no longer has (a dropped
    column's data stays in its rows as a phantom)."""

    column_count: int
    values: dict[int, bytes | None]
    present: dict[int, bool]
    raw: bytes = b""
    var_offsets: tuple[int, ...] = ()


def split_row(definition: TableDefinition, row: bytes) -> RawRow:
    width = definition.layout.count_width
    code = "B" if width == 1 else "H"
    if len(row) < width:
        raise AccessError("row shorter than its column count")
    row_columns = struct.unpack_from(f"<{code}", row, 0)[0]
    null_mask_length = (row_columns + 7) // 8
    if len(row) < width + null_mask_length:
        raise AccessError("row shorter than its null mask")
    null_mask = row[len(row) - null_mask_length :]
    has_var = definition.var_column_count > 0
    var_offsets: list[int] = []
    if has_var:
        count_pos = len(row) - null_mask_length - width
        if count_pos < width:
            raise AccessError("row shorter than its variable-column count")
        var_count = struct.unpack_from(f"<{code}", row, count_pos)[0]
        table_pos = count_pos - width * (var_count + 1)
        if table_pos < width:
            raise AccessError("row shorter than its variable-column offsets")
        reversed_offsets = struct.unpack_from(f"<{var_count + 1}{code}", row, table_pos)
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
            start = width + column.fixed_offset
            values[number] = row[start : start + column.length]
            continue
        if column.var_index + 1 >= len(var_offsets):
            raise AccessError(
                f"column {column.name!r} has variable index {column.var_index} "
                f"but the row holds {max(len(var_offsets) - 1, 0)} variable columns"
            )
        start = var_offsets[column.var_index]
        end = var_offsets[column.var_index + 1]
        if not width <= start <= end <= len(row):
            raise AccessError(
                f"column {column.name!r} spans {start}..{end} in a {len(row)}-byte row"
            )
        values[number] = row[start:end]
    return RawRow(row_columns, values, present, bytes(row), tuple(var_offsets))


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


def decode_column_text(column: ColumnDef, raw: bytes) -> str:
    """Text as the column's own version stores it: UTF-16 (compressed or
    not) in Jet 4, code page bytes in Jet 3."""
    if not column.layout.unicode_text:
        return raw.decode(column.code_page)
    return decode_text(raw)


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


def encode_scalar(column: ColumnDef, value: object, *, compress_text: bool) -> bytes:
    """Encode a present, non-long-value column.  Boolean yields no bytes
    (it lives in the null mask)."""
    code = column.type_code
    name = column.name
    if code == TYPE_BOOLEAN:
        return b""
    if code == TYPE_BYTE:
        if not isinstance(value, int) or not 0 <= value <= 255:
            raise AccessError(f"column {name!r}: {value!r} is not a Byte")
        return bytes((value,))
    if code == TYPE_INT:
        if not isinstance(value, int) or not -32768 <= value <= 32767:
            raise AccessError(f"column {name!r}: {value!r} is not an Integer")
        return struct.pack("<h", value)
    if code in (TYPE_LONG, TYPE_COMPLEX):
        if not isinstance(value, int) or not -(1 << 31) <= value < 1 << 31:
            raise AccessError(f"column {name!r}: {value!r} is not a Long")
        return struct.pack("<i", value)
    if code == TYPE_BIGINT:
        if not isinstance(value, int) or not -(1 << 63) <= value < 1 << 63:
            raise AccessError(f"column {name!r}: {value!r} is not a BigInt")
        return struct.pack("<q", value)
    if code == TYPE_MONEY:
        # Currency is a scaled 64-bit integer with four decimals.  A float
        # goes through its shortest repr and rounds half-even, which is
        # what CCur does to a Double.
        if isinstance(value, float):
            value = Decimal(repr(value))
        if not isinstance(value, (int, Decimal)) or (isinstance(value, Decimal) and not value.is_finite()):
            raise AccessError(f"column {name!r}: {value!r} is not Currency")
        scaled = int(Decimal(value).quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN).scaleb(4))
        if not -(1 << 63) <= scaled < 1 << 63:
            raise AccessError(f"column {name!r}: {value!r} is out of range for Currency")
        return struct.pack("<q", scaled)
    if code == TYPE_FLOAT:
        if not isinstance(value, (int, float)):
            raise AccessError(f"column {name!r}: {value!r} is not a Single")
        return struct.pack("<f", float(value))
    if code == TYPE_DOUBLE:
        if not isinstance(value, (int, float)):
            raise AccessError(f"column {name!r}: {value!r} is not a Double")
        return struct.pack("<d", float(value))
    if code == TYPE_DATETIME:
        # A float is taken as the stored serial itself (days since
        # 1899-12-30, the time as a fraction), which is how a stamp read
        # from another database is reproduced bit for bit.
        if isinstance(value, float):
            return struct.pack("<d", value)
        if not isinstance(value, _dt.datetime):
            raise AccessError(f"column {name!r}: {value!r} is not a datetime")
        return encode_datetime(value)
    if code == TYPE_BINARY:
        if not isinstance(value, (bytes, bytearray)):
            raise AccessError(f"column {name!r}: {value!r} is not bytes")
        if len(value) > column.length:
            raise AccessError(f"column {name!r}: {len(value)} bytes exceed its size {column.length}")
        if column.is_fixed:
            # A fixed-size Binary column always holds its full width; the
            # engine pads with zeros and hands the padded value back.
            return bytes(value).ljust(column.length, b"\x00")
        return bytes(value)
    if code == TYPE_TEXT:
        if not isinstance(value, str):
            raise AccessError(f"column {name!r}: {value!r} is not text")
        if 2 * len(value) > column.length:
            raise AccessError(f"column {name!r}: {len(value)} characters exceed its size {column.length // 2}")
        return encode_text(value) if compress_text else value.encode("utf-16-le")
    if code == TYPE_GUID:
        if not isinstance(value, uuid.UUID):
            raise AccessError(f"column {name!r}: {value!r} is not a UUID")
        return value.bytes_le
    if code == TYPE_NUMERIC:
        if isinstance(value, float):
            # A number written into the statement arrives as a float; take
            # it as the decimal that was written rather than the binary
            # value nearest it, which is what the engine stores.
            value = Decimal(repr(value))
        if not isinstance(value, (int, Decimal)):
            raise AccessError(f"column {name!r}: {value!r} is not a Decimal")
        return encode_numeric(Decimal(value), column.scale)
    if code in (TYPE_OLE, TYPE_MEMO):
        if not isinstance(value, (bytes, bytearray)):
            raise AccessError(f"column {name!r}: a long value must be passed encoded")
        return bytes(value)
    raise AccessError(f"column {name!r}: cannot encode type {column.type_name}")


def encode_row(
    definition: TableDefinition,
    values: dict[int, bytes | None],
    present_booleans: set[int],
    template: RawRow | None = None,
    template_var_count: int | None = None,
) -> bytes:
    """Assemble a row from per-column encoded bytes (``None`` for null),
    keyed by column number.  ``present_booleans`` names the Boolean
    columns that are True.  ``template`` is the row's previous version:
    bytes it holds for columns the definition no longer has -- fixed
    slots, variable slots and null-mask bits -- are carried over, as the
    engine carries them (measured on ALTER COLUMN after a DROP COLUMN).
    ``template_var_count`` is the variable-column count the row is
    expected to have been written with (the definition's, unless a
    variable column was just added), which decides how it is carried."""
    columns = definition.columns_by_number()
    column_count = max((c.number for c in columns), default=-1) + 1
    fixed = [c for c in columns if c.is_fixed]
    fixed_size = max((c.fixed_offset + c.length for c in fixed), default=0)
    fixed_block = bytearray(fixed_size)
    variable: dict[int, bytes] = {}
    null_mask = bytearray((column_count + 7) // 8)
    known = {c.number for c in columns}
    if template is not None and template.raw:
        old_mask_length = (template.column_count + 7) // 8
        old_var_count = max(len(template.var_offsets) - 1, 0)
        expected = definition.var_column_count if template_var_count is None else template_var_count
        if old_var_count == expected or not template.var_offsets:
            # The old row's slots line up with the definition's: the fixed
            # block is the old row's whole pre-variable region (never
            # shorter than the definition's fixed size, junk from an earlier
            # rewrite included), variable slots are carried by index.
            old_fixed_length = (template.var_offsets[0] - 2) if template.var_offsets else len(template.raw) - 2 - old_mask_length
            fixed_block = bytearray(max(fixed_size, old_fixed_length))
            old_fixed = template.raw[2 : 2 + len(fixed_block)]
            fixed_block[: len(old_fixed)] = old_fixed
            taken = {c.var_index for c in columns if not c.is_fixed}
            for var_index in range(old_var_count):
                if var_index not in taken:
                    variable[var_index] = template.raw[template.var_offsets[var_index] : template.var_offsets[var_index + 1]]
        else:
            # Variable columns came or went since the row was written: the
            # engine keeps the old body verbatim, its variable table and
            # count included, and appends a fresh table behind it
            # (measured on ALTER COLUMN after ADD and DROP COLUMN).
            body = bytearray(template.raw[2 : len(template.raw) - old_mask_length])
            if len(body) < fixed_size:
                body += bytes(fixed_size - len(body))
            fixed_block = body
        old_mask = template.raw[len(template.raw) - old_mask_length :]
        for number in range(min(template.column_count, column_count)):
            if number not in known and old_mask[number // 8] & (1 << (number % 8)):
                null_mask[number // 8] |= 1 << (number % 8)

    def mark(number: int) -> None:
        null_mask[number // 8] |= 1 << (number % 8)

    for column in columns:
        if column.type_code == TYPE_BOOLEAN:
            if column.number in present_booleans:
                mark(column.number)
            continue
        value = values.get(column.number)
        if value is None:
            continue
        mark(column.number)
        if column.is_fixed:
            if len(value) != column.length:
                raise AccessError(
                    f"column {column.name!r}: {len(value)} bytes for a {column.length}-byte fixed column"
                )
            fixed_block[column.fixed_offset : column.fixed_offset + column.length] = value
        else:
            variable[column.var_index] = value

    row = bytearray(struct.pack("<H", column_count))
    row += fixed_block
    var_count = definition.var_column_count
    if var_count:
        offsets: list[int] = []
        for var_index in range(var_count):
            offsets.append(len(row))
            row += variable.get(var_index, b"")
        offsets.append(len(row))
        if offsets[-1] > 0x1FFF:
            raise AccessError("row is too long for its offset table")
        for offset in reversed(offsets):
            row += struct.pack("<H", offset)
        row += struct.pack("<H", var_count)
    row += null_mask
    return bytes(row)


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
        return decode_column_text(column, raw)
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
