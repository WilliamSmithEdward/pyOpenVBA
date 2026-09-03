"""Table definitions: the type-2 page (or chain of pages) that describes a
table's columns, indexes and usage maps.

Layout, as measured on Access-written files (every offset below was
confirmed by parsing a definition and checking that the bytes consumed
equal the definition length the page declares):

    0x00  u8   page type 2
    0x01  u8   1
    0x02  u16  free space on the page
    0x04  u32  next definition page (0 when the definition fits)
    0x08  u32  definition length, counted from page offset 0
    0x0C  u32  a per-table value repeated in every column header
    0x10  u32  row count
    0x14  u32  next AutoNumber value
    0x18  i32  next complex-type AutoNumber value (ACE), -1 in Jet 4
    0x28  u8   table type: 0x53 engine system table, 0x4E any other
    0x29  u16  highest column number ever used
    0x2B  u16  variable-length column count
    0x2D  u16  column count
    0x2F  u32  logical index count
    0x33  u32  real (B-tree) index count
    0x37  u32  owned-pages usage map reference
    0x3B  u32  free-space-pages usage map reference
    0x3F  real index headers, 12 bytes each
          column headers, 25 bytes each
          column names, u16 length + UTF-16LE each
          real index definitions, 52 bytes each
          logical index definitions, 28 bytes each
          logical index names
          long-value column usage maps, 10 bytes each, 0xFFFF terminated

A definition longer than one page continues on the page named at 0x04,
whose own 8-byte header is skipped.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from pyopenvba.access_read import AccessError
from pyopenvba.access._pages import PAGE_TDEF, PageStore

OFFSET_NEXT_PAGE = 0x04
OFFSET_DEFINITION_LENGTH = 0x08
OFFSET_TABLE_TAG = 0x0C
OFFSET_ROW_COUNT = 0x10
OFFSET_NEXT_AUTONUMBER = 0x14
OFFSET_NEXT_COMPLEX_AUTONUMBER = 0x18
OFFSET_TABLE_TYPE = 0x28
OFFSET_MAX_COLUMNS = 0x29
OFFSET_VAR_COLUMN_COUNT = 0x2B
OFFSET_COLUMN_COUNT = 0x2D
OFFSET_LOGICAL_INDEX_COUNT = 0x2F
OFFSET_REAL_INDEX_COUNT = 0x33
OFFSET_OWNED_PAGES = 0x37
OFFSET_FREE_SPACE_PAGES = 0x3B
OFFSET_INDEX_HEADERS = 0x3F

SIZE_REAL_INDEX_HEADER = 12
SIZE_COLUMN_HEADER = 25
SIZE_REAL_INDEX = 52
SIZE_LOGICAL_INDEX = 28
SIZE_COLUMN_USAGE_MAPS = 10
MAX_INDEX_COLUMNS = 10
INDEX_COLUMN_UNUSED = 0xFFFF

# 'S' marks the engine's own tables (MSysObjects, MSysACEs, MSysQueries,
# MSysRelationships); 'N' marks everything else, including the MSys* tables
# the Access application layer keeps, which the catalog flags as system.
TABLE_TYPE_SYSTEM = 0x53
TABLE_TYPE_USER = 0x4E

# Column data types.
TYPE_BOOLEAN = 0x01
TYPE_BYTE = 0x02
TYPE_INT = 0x03
TYPE_LONG = 0x04
TYPE_MONEY = 0x05
TYPE_FLOAT = 0x06
TYPE_DOUBLE = 0x07
TYPE_DATETIME = 0x08
TYPE_BINARY = 0x09
TYPE_TEXT = 0x0A
TYPE_OLE = 0x0B
TYPE_MEMO = 0x0C
TYPE_GUID = 0x0F
TYPE_NUMERIC = 0x10
TYPE_COMPLEX = 0x12
TYPE_BIGINT = 0x13
TYPE_EXTENDED_DATETIME = 0x14

LONG_VALUE_TYPES = frozenset((TYPE_OLE, TYPE_MEMO))
FIXED_LENGTH_TYPES = frozenset(
    (
        TYPE_BOOLEAN,
        TYPE_BYTE,
        TYPE_INT,
        TYPE_LONG,
        TYPE_MONEY,
        TYPE_FLOAT,
        TYPE_DOUBLE,
        TYPE_DATETIME,
        TYPE_GUID,
        TYPE_NUMERIC,
        TYPE_COMPLEX,
        TYPE_BIGINT,
        TYPE_EXTENDED_DATETIME,
    )
)
TYPE_NAMES = {
    TYPE_BOOLEAN: "Boolean",
    TYPE_BYTE: "Byte",
    TYPE_INT: "Integer",
    TYPE_LONG: "Long",
    TYPE_MONEY: "Currency",
    TYPE_FLOAT: "Single",
    TYPE_DOUBLE: "Double",
    TYPE_DATETIME: "DateTime",
    TYPE_BINARY: "Binary",
    TYPE_TEXT: "Text",
    TYPE_OLE: "OLE",
    TYPE_MEMO: "Memo",
    TYPE_GUID: "GUID",
    TYPE_NUMERIC: "Decimal",
    TYPE_COMPLEX: "Complex",
    TYPE_BIGINT: "BigInt",
    TYPE_EXTENDED_DATETIME: "DateTimeExtended",
}

# Column flags (byte 15 of the header).
COLUMN_FIXED = 0x01
COLUMN_NULLABLE = 0x02
COLUMN_AUTONUMBER = 0x04
# Set on every column of the engine's own catalog tables; meaning unknown.
COLUMN_SYSTEM_UNKNOWN = 0x10
COLUMN_AUTONUMBER_GUID = 0x40
COLUMN_HYPERLINK = 0x80
# Byte 16: Unicode compression for Text and Memo.  Access sets it on the
# columns it creates; SQL DDL leaves it clear unless WITH COMPRESSION is
# given.  The engine compresses its own catalog's text regardless.
COLUMN_COMPRESSED_UNICODE = 0x01

# Real index flags.
INDEX_UNIQUE = 0x01
INDEX_IGNORE_NULLS = 0x02
INDEX_REQUIRED = 0x08
INDEX_ALWAYS_SET = 0x80

# Logical index kinds.
INDEX_KIND_NORMAL = 0x00
INDEX_KIND_PRIMARY = 0x01
INDEX_KIND_FOREIGN = 0x02


@dataclass
class ColumnDef:
    """One column header plus its name.  ``raw`` is the 25-byte header as
    read, so an unedited definition serializes back byte for byte."""

    name: str
    type_code: int
    number: int
    var_index: int
    fixed_offset: int
    length: int
    flags: int
    misc_flags: int
    sort_order: int
    sort_version: int
    raw: bytes

    @property
    def is_fixed(self) -> bool:
        return bool(self.flags & COLUMN_FIXED)

    @property
    def is_long_value(self) -> bool:
        return self.type_code in LONG_VALUE_TYPES

    @property
    def nullable(self) -> bool:
        return bool(self.flags & COLUMN_NULLABLE)

    @property
    def auto_number(self) -> bool:
        return bool(self.flags & COLUMN_AUTONUMBER)

    @property
    def compressed_unicode(self) -> bool:
        return bool(self.misc_flags & COLUMN_COMPRESSED_UNICODE)

    @property
    def precision(self) -> int:
        """Decimal precision; shares bytes with the text sort order."""
        return self.raw[11]

    @property
    def scale(self) -> int:
        return self.raw[12]

    @property
    def type_name(self) -> str:
        return TYPE_NAMES.get(self.type_code, f"type 0x{self.type_code:02X}")


@dataclass
class IndexColumn:
    number: int
    ascending: bool


@dataclass
class RealIndex:
    """A B-tree: the physical index one or more logical indexes share."""

    header_raw: bytes
    entry_count: int  # distinct keys, header +4
    row_count: int  # rows the index counts, header +0; 0 unless built over rows
    columns: list[IndexColumn]
    usage_map_ref: int
    root_page: int
    unknown: int
    flags: int
    raw: bytes

    @property
    def unique(self) -> bool:
        return bool(self.flags & INDEX_UNIQUE)


@dataclass
class LogicalIndex:
    """A named index as Access shows it, pointing at a real index."""

    name: str
    number: int
    real_index: int
    relationship_kind: int
    relationship_index: int
    relationship_table_page: int
    cascade_updates: bool
    cascade_deletes: bool
    kind: int
    raw: bytes

    @property
    def is_primary_key(self) -> bool:
        return self.kind == INDEX_KIND_PRIMARY


@dataclass
class TableDefinition:
    page: int
    tag: int
    row_count: int
    next_autonumber: int
    next_complex_autonumber: int
    table_type: int
    max_columns: int
    var_column_count: int
    logical_index_count: int
    real_index_count: int
    owned_pages_ref: int
    free_space_pages_ref: int
    columns: list[ColumnDef]
    real_indexes: list[RealIndex]
    logical_indexes: list[LogicalIndex]
    column_usage_maps: dict[int, tuple[int, int]]
    definition_length: int
    pages: list[int] = field(default_factory=lambda: [])
    header_raw: bytes = b""

    @property
    def is_system(self) -> bool:
        return self.table_type == TABLE_TYPE_SYSTEM

    def columns_by_number(self) -> list[ColumnDef]:
        return sorted(self.columns, key=lambda c: c.number)

    def column(self, name: str) -> ColumnDef:
        for column in self.columns:
            if column.name.lower() == name.lower():
                return column
        raise AccessError(f"no column named {name!r}")

    def column_by_number(self, number: int) -> ColumnDef:
        for column in self.columns:
            if column.number == number:
                return column
        raise AccessError(f"no column numbered {number}")

    def primary_key(self) -> LogicalIndex | None:
        for index in self.logical_indexes:
            if index.is_primary_key:
                return index
        return None


def read_definition_bytes(store: PageStore, page: int) -> tuple[bytes, list[int]]:
    """Concatenate a definition that may span pages.  Continuation pages
    contribute everything after their 8-byte header."""
    first = store.read(page)
    if first[0] != PAGE_TDEF:
        raise AccessError(
            f"page {page} is type {first[0]:#04x}, not a table definition"
        )
    pages = [page]
    out = bytearray(first)
    next_page = struct.unpack_from("<I", first, OFFSET_NEXT_PAGE)[0]
    while next_page:
        if next_page in pages:
            raise AccessError(f"table definition at page {page} loops")
        cont = store.read(next_page)
        if cont[0] != PAGE_TDEF:
            raise AccessError(
                f"continuation page {next_page} is type {cont[0]:#04x}"
            )
        pages.append(next_page)
        out.extend(cont[8:])
        next_page = struct.unpack_from("<I", cont, OFFSET_NEXT_PAGE)[0]
    return bytes(out), pages


def _read_name(buf: bytes, pos: int) -> tuple[str, int]:
    length = struct.unpack_from("<H", buf, pos)[0]
    raw = buf[pos + 2 : pos + 2 + length]
    if len(raw) != length:
        raise AccessError("name runs past the end of the definition")
    return raw.decode("utf-16-le"), pos + 2 + length


def parse_column_header(raw: bytes, name: str) -> ColumnDef:
    if len(raw) != SIZE_COLUMN_HEADER:
        raise AccessError(f"column header is {len(raw)} bytes, not {SIZE_COLUMN_HEADER}")
    return ColumnDef(
        name=name,
        type_code=raw[0],
        number=struct.unpack_from("<H", raw, 5)[0],
        var_index=struct.unpack_from("<H", raw, 7)[0],
        sort_order=struct.unpack_from("<H", raw, 11)[0],
        sort_version=struct.unpack_from("<H", raw, 13)[0],
        flags=raw[15],
        misc_flags=raw[16],
        fixed_offset=struct.unpack_from("<H", raw, 21)[0],
        length=struct.unpack_from("<H", raw, 23)[0],
        raw=bytes(raw),
    )


def parse_real_index(header_raw: bytes, raw: bytes) -> RealIndex:
    if len(raw) != SIZE_REAL_INDEX:
        raise AccessError(f"index definition is {len(raw)} bytes, not {SIZE_REAL_INDEX}")
    columns: list[IndexColumn] = []
    for i in range(MAX_INDEX_COLUMNS):
        number = struct.unpack_from("<H", raw, 4 + 3 * i)[0]
        if number == INDEX_COLUMN_UNUSED:
            continue
        columns.append(IndexColumn(number, bool(raw[6 + 3 * i] & 0x01)))
    return RealIndex(
        header_raw=bytes(header_raw),
        entry_count=struct.unpack_from("<I", header_raw, 4)[0],
        row_count=struct.unpack_from("<I", header_raw, 0)[0],
        columns=columns,
        usage_map_ref=struct.unpack_from("<I", raw, 34)[0],
        root_page=struct.unpack_from("<I", raw, 38)[0],
        unknown=struct.unpack_from("<I", raw, 42)[0],
        flags=raw[46],
        raw=bytes(raw),
    )


def parse_logical_index(raw: bytes, name: str) -> LogicalIndex:
    if len(raw) != SIZE_LOGICAL_INDEX:
        raise AccessError(f"logical index is {len(raw)} bytes, not {SIZE_LOGICAL_INDEX}")
    return LogicalIndex(
        name=name,
        number=struct.unpack_from("<I", raw, 4)[0],
        real_index=struct.unpack_from("<I", raw, 8)[0],
        relationship_kind=raw[12],
        relationship_index=struct.unpack_from("<I", raw, 13)[0],
        relationship_table_page=struct.unpack_from("<I", raw, 17)[0],
        cascade_updates=bool(raw[21]),
        cascade_deletes=bool(raw[22]),
        kind=raw[23],
        raw=bytes(raw),
    )


def parse_table_definition(store: PageStore, page: int) -> TableDefinition:
    buf, pages = read_definition_bytes(store, page)

    def u16(off: int) -> int:
        return struct.unpack_from("<H", buf, off)[0]

    def u32(off: int) -> int:
        return struct.unpack_from("<I", buf, off)[0]

    definition_length = u32(OFFSET_DEFINITION_LENGTH)
    column_count = u16(OFFSET_COLUMN_COUNT)
    logical_count = u32(OFFSET_LOGICAL_INDEX_COUNT)
    real_count = u32(OFFSET_REAL_INDEX_COUNT)
    if definition_length > len(buf):
        raise AccessError(
            f"table definition at page {page} declares {definition_length} "
            f"bytes but only {len(buf)} were read"
        )

    pos = OFFSET_INDEX_HEADERS
    index_headers = [
        buf[pos + i * SIZE_REAL_INDEX_HEADER : pos + (i + 1) * SIZE_REAL_INDEX_HEADER]
        for i in range(real_count)
    ]
    pos += real_count * SIZE_REAL_INDEX_HEADER

    column_raws = [
        buf[pos + i * SIZE_COLUMN_HEADER : pos + (i + 1) * SIZE_COLUMN_HEADER]
        for i in range(column_count)
    ]
    pos += column_count * SIZE_COLUMN_HEADER
    columns: list[ColumnDef] = []
    for raw in column_raws:
        name, pos = _read_name(buf, pos)
        columns.append(parse_column_header(raw, name))

    real_indexes: list[RealIndex] = []
    for i in range(real_count):
        raw = buf[pos : pos + SIZE_REAL_INDEX]
        pos += SIZE_REAL_INDEX
        real_indexes.append(parse_real_index(index_headers[i], raw))

    logical_raws: list[bytes] = []
    for _ in range(logical_count):
        logical_raws.append(buf[pos : pos + SIZE_LOGICAL_INDEX])
        pos += SIZE_LOGICAL_INDEX
    logical_indexes: list[LogicalIndex] = []
    for raw in logical_raws:
        name, pos = _read_name(buf, pos)
        logical_indexes.append(parse_logical_index(raw, name))

    column_usage_maps: dict[int, tuple[int, int]] = {}
    while True:
        if pos + 2 > len(buf):
            raise AccessError(
                f"table definition at page {page} ends inside its usage-map block"
            )
        number = u16(pos)
        pos += 2
        if number == INDEX_COLUMN_UNUSED:
            break
        column_usage_maps[number] = (u32(pos), u32(pos + 4))
        pos += 8

    if pos != definition_length:
        raise AccessError(
            f"table definition at page {page}: parsed {pos} bytes, the page "
            f"declares {definition_length}"
        )

    return TableDefinition(
        page=page,
        tag=u32(OFFSET_TABLE_TAG),
        row_count=u32(OFFSET_ROW_COUNT),
        next_autonumber=u32(OFFSET_NEXT_AUTONUMBER),
        next_complex_autonumber=struct.unpack_from(
            "<i", buf, OFFSET_NEXT_COMPLEX_AUTONUMBER
        )[0],
        table_type=buf[OFFSET_TABLE_TYPE],
        max_columns=u16(OFFSET_MAX_COLUMNS),
        var_column_count=u16(OFFSET_VAR_COLUMN_COUNT),
        logical_index_count=logical_count,
        real_index_count=real_count,
        owned_pages_ref=u32(OFFSET_OWNED_PAGES),
        free_space_pages_ref=u32(OFFSET_FREE_SPACE_PAGES),
        columns=columns,
        real_indexes=real_indexes,
        logical_indexes=logical_indexes,
        column_usage_maps=column_usage_maps,
        definition_length=definition_length,
        pages=pages,
        header_raw=bytes(buf[:OFFSET_INDEX_HEADERS]),
    )

