"""Property blobs: the ``MR2`` value in ``MSysObjects.LvProp``.

Access keeps an object's properties -- a table's Description, a column's
Caption, Format, DecimalPlaces and the rest, the database's own settings
on the MSysDb row -- in one long value per catalog row:

    "MR2\\0"                        signature (Jet 3 wrote "KKD\\0")
    blocks, each: u32 length (counting itself), u16 kind, body
      kind 0x80  names: repeated u16 byte length + UTF-16 text; every
                 record below refers to a name by its index here
      kind 0x00  the object's own properties
      kind 0x01  one column's properties
                 body: u16 name-part length (6 when unnamed), u16 0,
                       u16 name byte length, UTF-16 column name,
                       then records: u16 length, u8 flags, u8 DAO type,
                       u16 name index, u16 value length, value bytes

Value bytes follow the DAO type (``dbText`` 10 is UTF-16, possibly with
the engine's compression marker; the fixed types are the engine's own
row encodings), so :mod:`_rows` decodes and encodes them.  Measured on a
blob DAO wrote for a table Description plus a field Caption and
Description, and read back on every Access-authored blob in the fixtures.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from pyopenvba.access_read import AccessError
from pyopenvba.access._rows import decode_text
from pyopenvba.access._tdef import (
    TYPE_BOOLEAN,
    TYPE_BYTE,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_FLOAT,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MONEY,
    TYPE_TEXT,
)

SIGNATURE = b"MR2\0"
BLOCK_NAMES = 0x80
BLOCK_OBJECT = 0x00
BLOCK_COLUMN = 0x01

# DAO's DataTypeEnum, which the property records use; the scalar values
# coincide with the engine's column type codes.
DB_BOOLEAN = TYPE_BOOLEAN
DB_BYTE = TYPE_BYTE
DB_INTEGER = TYPE_INT
DB_LONG = TYPE_LONG
DB_CURRENCY = TYPE_MONEY
DB_SINGLE = TYPE_FLOAT
DB_DOUBLE = TYPE_DOUBLE
DB_DATE = TYPE_DATETIME
DB_TEXT = TYPE_TEXT
DB_MEMO = 0x0C


@dataclass(frozen=True)
class PropertyValue:
    """One property as stored: its DAO type, flags byte and raw value."""

    type: int
    flags: int
    raw: bytes

    def decode(self) -> object:
        return decode_property_value(self.type, self.raw)


@dataclass
class PropertyBlob:
    """A parsed ``MR2`` blob: the name table and the property records of the
    object itself and of each named column, in stored order."""

    names: list[str] = field(default_factory=lambda: [])
    object_properties: dict[str, PropertyValue] = field(default_factory=lambda: {})
    column_properties: dict[str, dict[str, PropertyValue]] = field(default_factory=lambda: {})
    #: Block kinds in the order they appeared, so a rewrite keeps it.
    block_order: list[tuple[int, str]] = field(default_factory=lambda: [])

    def decoded(self) -> dict[str, object]:
        return {name: value.decode() for name, value in self.object_properties.items()}

    def decoded_column(self, column: str) -> dict[str, object]:
        return {name: value.decode() for name, value in self.column_properties.get(column, {}).items()}


def decode_property_value(dao_type: int, raw: bytes) -> object:
    if dao_type in (DB_TEXT, DB_MEMO):
        return decode_text(raw)
    if dao_type == DB_BOOLEAN:
        return bool(raw and raw[0])
    if dao_type in (DB_BYTE, DB_INTEGER, DB_LONG):
        # Access writes some Integer-typed properties (ColumnWidth,
        # ColumnOrder) with four bytes, so the width decides.
        if len(raw) == 1:
            return raw[0]
        if len(raw) == 2:
            return struct.unpack("<h", raw)[0]
        if len(raw) == 4:
            return struct.unpack("<i", raw)[0]
        if len(raw) == 8:
            return struct.unpack("<q", raw)[0]
        return raw
    if dao_type in (DB_DOUBLE, DB_DATE):
        return struct.unpack("<d", raw)[0] if dao_type == DB_DOUBLE else _decode_date(raw)
    if dao_type == DB_SINGLE:
        return struct.unpack("<f", raw)[0]
    if dao_type == DB_CURRENCY:
        from decimal import Decimal

        return Decimal(struct.unpack("<q", raw)[0]).scaleb(-4)
    return raw


def _decode_date(raw: bytes) -> object:
    from pyopenvba.access._rows import decode_datetime

    return decode_datetime(raw)


def encode_property_value(dao_type: int, value: object) -> bytes:
    """The value bytes for a property of ``dao_type``, as DAO writes them:
    text uncompressed UTF-16, the fixed types in the engine's encodings."""
    import datetime as _dt
    from decimal import Decimal

    if dao_type in (DB_TEXT, DB_MEMO):
        if not isinstance(value, str):
            raise AccessError(f"a text property needs a str, not {value!r}")
        return value.encode("utf-16le")
    if dao_type == DB_BOOLEAN:
        return b"\x01" if value else b"\x00"
    if dao_type == DB_BYTE:
        if not isinstance(value, int) or not 0 <= value <= 255:
            raise AccessError(f"a Byte property needs 0..255, not {value!r}")
        return bytes((value,))
    if dao_type == DB_INTEGER:
        if not isinstance(value, int):
            raise AccessError(f"an Integer property needs an int, not {value!r}")
        return struct.pack("<h", value)
    if dao_type == DB_LONG:
        if not isinstance(value, int):
            raise AccessError(f"a Long property needs an int, not {value!r}")
        return struct.pack("<i", value)
    if dao_type == DB_DOUBLE:
        return struct.pack("<d", float(value))  # pyright: ignore[reportArgumentType]
    if dao_type == DB_SINGLE:
        return struct.pack("<f", float(value))  # pyright: ignore[reportArgumentType]
    if dao_type == DB_DATE:
        from pyopenvba.access._rows import encode_datetime

        if isinstance(value, float):
            return struct.pack("<d", value)
        if not isinstance(value, _dt.datetime):
            raise AccessError(f"a Date property needs a datetime, not {value!r}")
        return encode_datetime(value)
    if dao_type == DB_CURRENCY:
        if not isinstance(value, (int, Decimal)):
            raise AccessError(f"a Currency property needs an int or Decimal, not {value!r}")
        return struct.pack("<q", int(Decimal(value).scaleb(4)))
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    raise AccessError(f"property type {dao_type} takes bytes, not {value!r}")


def dao_type_for(value: object) -> int:
    """The DAO type DAO itself picks for a Python value."""
    import datetime as _dt
    from decimal import Decimal

    if isinstance(value, bool):
        return DB_BOOLEAN
    if isinstance(value, int):
        return DB_LONG if not -32768 <= value <= 32767 else DB_INTEGER
    if isinstance(value, float):
        return DB_DOUBLE
    if isinstance(value, str):
        return DB_TEXT
    if isinstance(value, _dt.datetime):
        return DB_DATE
    if isinstance(value, Decimal):
        return DB_CURRENCY
    raise AccessError(f"no property type for {value!r}")


def parse_property_blob(blob: bytes) -> PropertyBlob:
    if len(blob) < 4 or blob[:4] != SIGNATURE:
        raise AccessError(f"not an MR2 property blob (starts {blob[:4]!r})")
    out = PropertyBlob()
    pos = 4
    while pos < len(blob):
        if pos + 6 > len(blob):
            raise AccessError("property blob ends inside a block header")
        length = struct.unpack_from("<I", blob, pos)[0]
        kind = struct.unpack_from("<H", blob, pos + 4)[0]
        if length < 6 or pos + length > len(blob):
            raise AccessError(f"property block at {pos} claims {length} bytes")
        body = blob[pos + 6 : pos + length]
        if kind == BLOCK_NAMES:
            q = 0
            while q + 2 <= len(body):
                n = struct.unpack_from("<H", body, q)[0]
                out.names.append(body[q + 2 : q + 2 + n].decode("utf-16le"))
                q += 2 + n
        elif kind in (BLOCK_OBJECT, BLOCK_COLUMN):
            name_part = struct.unpack_from("<H", body, 0)[0]
            name_length = struct.unpack_from("<H", body, 4)[0]
            target = body[6 : 6 + name_length].decode("utf-16le")
            records: dict[str, PropertyValue] = {}
            q = name_part
            while q + 8 <= len(body):
                record_length, flags, dao_type, name_index, value_length = struct.unpack_from("<HBBHH", body, q)
                if record_length < 8:
                    raise AccessError(f"property record at {pos + 6 + q} claims {record_length} bytes")
                if name_index >= len(out.names):
                    raise AccessError(f"property record names index {name_index} of {len(out.names)} names")
                records[out.names[name_index]] = PropertyValue(dao_type, flags, body[q + 8 : q + 8 + value_length])
                q += record_length
            if kind == BLOCK_OBJECT:
                out.object_properties.update(records)
            else:
                out.column_properties.setdefault(target, {}).update(records)
            out.block_order.append((kind, target))
        else:
            raise AccessError(f"property block of unknown kind {kind:#x}")
        pos += length
    return out


def serialize_property_blob(blob: PropertyBlob) -> bytes:
    """The bytes DAO would write for this blob: the names block, then one
    block per object in ``block_order`` (any object not listed there is
    appended, the table's own block first)."""
    names = list(blob.names)

    def name_index(name: str) -> int:
        if name not in names:
            names.append(name)
        return names.index(name)

    blocks: list[bytes] = []
    order = list(blob.block_order)
    if blob.object_properties and (BLOCK_OBJECT, "") not in order:
        order.append((BLOCK_OBJECT, ""))
    for column in blob.column_properties:
        if (BLOCK_COLUMN, column) not in order:
            order.append((BLOCK_COLUMN, column))
    for kind, target in order:
        records = blob.object_properties if kind == BLOCK_OBJECT else blob.column_properties.get(target, {})
        name_bytes = target.encode("utf-16le")
        body = bytearray(struct.pack("<HHH", 6 + len(name_bytes), 0, len(name_bytes)) + name_bytes)
        for name, value in records.items():
            body += struct.pack("<HBBHH", 8 + len(value.raw), value.flags, value.type, name_index(name), len(value.raw)) + value.raw
        blocks.append(struct.pack("<IH", 6 + len(body), kind) + bytes(body))
    names_body = b"".join(struct.pack("<H", len(n.encode("utf-16le"))) + n.encode("utf-16le") for n in names)
    names_block = struct.pack("<IH", 6 + len(names_body), BLOCK_NAMES) + names_body
    return SIGNATURE + names_block + b"".join(blocks)
