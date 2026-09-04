"""Creating and dropping tables: the table definition page, the usage-map
page that serves it, the empty index roots, and the catalog rows.

Everything here reproduces what the engine wrote for ``CREATE TABLE``,
``CREATE INDEX`` and ``DROP TABLE`` statements, diffed page by page
(docs/access_engine.md):

* A new table takes two pages: its definition and a data-shaped page
  (owner 0) holding its usage maps as 69-byte inline rows -- owned pages,
  free-space pages, one per index, then two per long-value column.
* Each index gets an empty leaf as its root.
* A definition takes ``ceil(length / 4088)`` pages, chained through the
  word at 0x04; the last page's free word is ``4088 * pages - length``
  and every other page's is 0 (see :func:`definition_pages`).  The
  per-table tag at 0x0C and in every column header is the database's
  (0x659 in every file seen); a real index definition starts ``83 07 00
  00``; a logical one carries ``04 04`` before its kind byte.
* Fixed-length columns are laid out in column order; Boolean columns are
  "fixed" of length 1 but take no row space; GUIDs are variable-length.
* The catalog gets an MSysObjects row (Id = definition page, parent the
  Tables container, Type 1) and three MSysACEs rows with the database's
  user, group and admin SIDs at ACM 0xFFEFF.
* Dropping releases every page the table held, tombstones its usage-map
  rows, marks the definition page type 8 and deletes the catalog rows.
"""

from __future__ import annotations

from collections.abc import Sequence

import struct
from dataclasses import dataclass, field

from pyopenvba.access_read import AccessError
from pyopenvba.access._btree import serialize_index_page
from pyopenvba.access._pages import (
    OFFSET_PAGE_FREE_SPACE,
    PAGE_DATA,
    PAGE_SIZE,
    PAGE_TDEF,
)
from pyopenvba.access._tdef import (
    IndexColumn,
    LogicalIndex,
    RealIndex,
    TableDefinition,
)
from pyopenvba.access._tdef import (
    COLUMN_AUTONUMBER,
    COLUMN_COMPRESSED_UNICODE,
    COLUMN_FIXED,
    COLUMN_NULLABLE,
    INDEX_ALWAYS_SET,
    INDEX_IGNORE_NULLS,
    INDEX_KIND_FOREIGN,
    INDEX_KIND_NORMAL,
    INDEX_KIND_PRIMARY,
    INDEX_REQUIRED,
    INDEX_UNIQUE,
    LONG_VALUE_TYPES,
    MAX_INDEX_COLUMNS,
    OFFSET_INDEX_HEADERS,
    SIZE_COLUMN_HEADER,
    SIZE_LOGICAL_INDEX,
    SIZE_REAL_INDEX,
    SIZE_REAL_INDEX_HEADER,
    TABLE_TYPE_USER,
    TYPE_BIGINT,
    TYPE_BINARY,
    TYPE_BOOLEAN,
    TYPE_COMPLEX,
    TYPE_BYTE,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_FLOAT,
    TYPE_GUID,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MEMO,
    TYPE_MONEY,
    TYPE_NUMERIC,
    TYPE_OLE,
    TYPE_TEXT,
)

TYPE_BY_NAME = {
    "boolean": TYPE_BOOLEAN,
    "yesno": TYPE_BOOLEAN,
    "byte": TYPE_BYTE,
    "integer": TYPE_INT,
    "int": TYPE_INT,
    "long": TYPE_LONG,
    "currency": TYPE_MONEY,
    "money": TYPE_MONEY,
    "single": TYPE_FLOAT,
    "float": TYPE_FLOAT,
    "double": TYPE_DOUBLE,
    "datetime": TYPE_DATETIME,
    "date": TYPE_DATETIME,
    "binary": TYPE_BINARY,
    "text": TYPE_TEXT,
    "ole": TYPE_OLE,
    "memo": TYPE_MEMO,
    "guid": TYPE_GUID,
    "decimal": TYPE_NUMERIC,
    "numeric": TYPE_NUMERIC,
    "bigint": TYPE_BIGINT,
    # An attachment or multi-valued column: the row holds only a Long,
    # and `pyopenvba.access._complex` explains where the values live.
    "complex": TYPE_COMPLEX,
}
#: Types the engine stores in a row's fixed block, and their widths.  A
#: GUID and a BigInt are eight and sixteen bytes wide but the engine keeps
#: them among the variable columns (measured on tables it wrote), so they
#: are not here.
FIXED_SIZES = {
    TYPE_BYTE: 1,
    TYPE_INT: 2,
    TYPE_LONG: 4,
    TYPE_MONEY: 8,
    TYPE_FLOAT: 4,
    TYPE_DOUBLE: 8,
    TYPE_DATETIME: 8,
    TYPE_NUMERIC: 17,
    TYPE_COMPLEX: 4,
}
#: What a value of each type occupies wherever it is stored.
VALUE_SIZES = {**FIXED_SIZES, TYPE_BIGINT: 8, TYPE_GUID: 16}
DEFAULT_TEXT_CHARS = 255
MAX_TEXT_CHARS = 255
MAX_BINARY_BYTES = 510
SORT_ORDER = 0x409
REAL_INDEX_LEAD = 0x0783
LOGICAL_INDEX_FLAGS = bytes((0x04, 0x04))
NO_RELATIONSHIP = 0xFFFFFFFF
USAGE_MAP_ROW = bytes(69)
DEFINITION_FREE_RESERVE = 8
DEFINITION_PAGE_SHARE = PAGE_SIZE - DEFINITION_FREE_RESERVE  # the engine's per-page accounting unit
FREED_TABLE_DEFINITION = 0x08
DEFAULT_ACM = 0xFFEFF


@dataclass(frozen=True)
class ColumnSpec:
    """A column to create.  ``size`` is characters for Text, bytes for
    Binary, ``(precision, scale)`` for Decimal; ignored elsewhere.

    The last five are not in the column header at all: the engine keeps
    them as properties on the column, so they are written after the table
    exists, one blob write per column, which is what the engine does for
    ``CREATE TABLE ... NOT NULL`` (measured).  ``default`` and the two
    validation fields hold Jet expressions as text; see
    :mod:`pyopenvba.access._validate` for what they mean to a row."""

    name: str
    type: str
    size: int | tuple[int, int] | None = None
    autonumber: bool = False
    compressed: bool = True
    #: Keep this column among the variable-length ones even though its
    #: type would normally sit in the fixed block.  A complex column's
    #: flat table stores its two Long bookkeeping columns this way.
    variable: bool = False
    #: The two words at bytes 11-14 of the header, when not the usual
    #: collation (1033, 0): a complex column's two Long bookkeeping
    #: columns carry (0, 0), and in a table a query makes every column
    #: after a Decimal carries that Decimal's (precision, scale).  A
    #: Decimal's own precision and scale live there and ignore this.
    collation: tuple[int, int] | None = None
    #: What bytes 9-10 of the header hold, when not the column number:
    #: a table made by a query numbers them from one there (measured).
    ordinal: int | None = None
    required: bool = False
    default: str | None = None
    allow_zero_length: bool | None = None
    validation_rule: str | None = None
    validation_text: str | None = None

    def properties(self) -> dict[str, object]:
        """The column properties this spec asks for, in the order they are
        written; empty when it asks for none."""
        out: dict[str, object] = {}
        if self.required:
            out["Required"] = True
        if self.allow_zero_length is not None:
            out["AllowZeroLength"] = self.allow_zero_length
        if self.default is not None:
            out["DefaultValue"] = self.default
        if self.validation_rule is not None:
            out["ValidationRule"] = self.validation_rule
        if self.validation_text is not None:
            out["ValidationText"] = self.validation_text
        return out

    @property
    def type_code(self) -> int:
        try:
            return TYPE_BY_NAME[self.type.lower()]
        except KeyError:
            raise AccessError(f"column {self.name!r}: unknown type {self.type!r}") from None


@dataclass(frozen=True)
class IndexSpec:
    """An index to create.  ``columns`` are names, or ``(name, ascending)``."""

    name: str
    columns: tuple[str | tuple[str, bool], ...]
    unique: bool = False
    primary: bool = False
    ignore_nulls: bool = False
    required: bool = False

    def resolved(self) -> list[tuple[str, bool]]:
        out: list[tuple[str, bool]] = []
        for item in self.columns:
            if isinstance(item, str):
                out.append((item, True))
            else:
                out.append((item[0], item[1]))
        return out


@dataclass
class DefinitionLayout:
    """Everything the definition page needs that is decided by the caller:
    which pages and usage-map rows were allocated for it."""

    page: int
    tag: int
    owned_ref: int
    free_ref: int
    index_umap_refs: list[int]
    index_roots: list[int]
    column_map_refs: dict[int, tuple[int, int]] = field(default_factory=lambda: {})


#: The width of a Decimal a make-table query computes: the 96-bit value
#: and its sign, without the four bytes a declared one carries besides.
COMPUTED_DECIMAL_BYTES = 13


def column_header(spec: ColumnSpec, number: int, var_index: int, fixed_offset: int, tag: int) -> bytes:
    code = spec.type_code
    # A complex column is flagged AutoNumber too, but takes its id from
    # the complex counter at 0x1C rather than the ordinary one.
    if spec.autonumber and code not in (TYPE_LONG, TYPE_COMPLEX):
        raise AccessError(
            f"column {spec.name!r}: only a Long or a complex column can be an AutoNumber"
        )
    raw = bytearray(SIZE_COLUMN_HEADER)
    raw[0] = code
    struct.pack_into("<I", raw, 1, tag)
    struct.pack_into("<H", raw, 5, number)
    struct.pack_into("<H", raw, 7, var_index)
    struct.pack_into("<H", raw, 9, number if spec.ordinal is None else spec.ordinal)
    if code == TYPE_NUMERIC:
        precision, scale = (18, 0)
        if isinstance(spec.size, tuple):
            precision, scale = spec.size
        raw[11] = precision
        raw[12] = scale
    elif spec.collation is not None:
        struct.pack_into("<HH", raw, 11, *spec.collation)
    else:
        struct.pack_into("<H", raw, 11, SORT_ORDER)
    flags = COLUMN_NULLABLE
    length = 0
    if code == TYPE_BOOLEAN:
        flags |= COLUMN_FIXED
        length = 1
    elif code in FIXED_SIZES:
        # A column kept among the variable ones has no fixed bit.  On a
        # complex column's flat table the two Long bookkeeping columns
        # also carry no collation (flags 2 and 6 with sort order 0, against
        # 3 and 7 with 1033 for an ordinary Long), which their spec says
        # with ``collation=(0, 0)``; a variable Long a make-table query
        # builds keeps its 1033 (measured).
        if not spec.variable:
            flags |= COLUMN_FIXED
        length = FIXED_SIZES[code]
        if code == TYPE_NUMERIC and spec.variable:
            # A declared Decimal is 17 bytes wide; one a query computes,
            # kept among the variable columns, is 13 (measured).
            length = COMPUTED_DECIMAL_BYTES
        if spec.autonumber:
            if code not in (TYPE_LONG, TYPE_COMPLEX):
                raise AccessError(
                    f"column {spec.name!r}: only a Long or a complex column can be an AutoNumber"
                )
            flags |= COLUMN_AUTONUMBER
    elif code == TYPE_BINARY:
        if not spec.variable:
            flags |= COLUMN_FIXED
        length = spec.size if isinstance(spec.size, int) else MAX_BINARY_BYTES
        if not 1 <= length <= MAX_BINARY_BYTES:
            raise AccessError(f"column {spec.name!r}: Binary size must be 1..{MAX_BINARY_BYTES}")
    elif code == TYPE_TEXT:
        chars = spec.size if isinstance(spec.size, int) else DEFAULT_TEXT_CHARS
        if not 1 <= chars <= MAX_TEXT_CHARS:
            raise AccessError(f"column {spec.name!r}: Text size must be 1..{MAX_TEXT_CHARS}")
        length = 2 * chars
        if spec.compressed:
            raw[16] = COLUMN_COMPRESSED_UNICODE
    elif code == TYPE_GUID:
        length = 16
    elif code == TYPE_BIGINT:
        # Eight bytes, but among the variable columns (measured).
        length = 8
    elif code in LONG_VALUE_TYPES:
        length = 0
        if code == TYPE_MEMO and spec.compressed:
            raw[16] = COLUMN_COMPRESSED_UNICODE
    else:
        raise AccessError(f"column {spec.name!r}: type {spec.type!r} cannot be created yet")
    raw[15] = flags
    struct.pack_into("<H", raw, 21, fixed_offset)
    struct.pack_into("<H", raw, 23, length)
    return bytes(raw)


def _name(text: str) -> bytes:
    data = text.encode("utf-16-le")
    return struct.pack("<H", len(data)) + data


def build_definition(columns: list[ColumnSpec], indexes: list[IndexSpec], layout: DefinitionLayout) -> bytes:
    """The definition stream (header and body, not yet laid over pages; see
    :func:`definition_pages`) for these columns and indexes."""
    if not columns:
        raise AccessError("a table needs at least one column")
    names_lower = [c.name.lower() for c in columns]
    if len(set(names_lower)) != len(names_lower):
        raise AccessError("column names must be unique")
    if len(indexes) != len(layout.index_roots) or len(indexes) != len(layout.index_umap_refs):
        raise AccessError("one root page and one usage map per index are needed")
    headers: list[bytes] = []
    var_index = 0
    fixed_offset = 0
    for number, spec in enumerate(columns):
        code = spec.type_code
        is_fixed = code == TYPE_BOOLEAN or code in FIXED_SIZES or code == TYPE_BINARY
        if is_fixed and not spec.variable:
            offset = 0 if code == TYPE_BOOLEAN else fixed_offset
            # Bytes 7-8 count the variable columns declared before this one,
            # which for a variable column is its own index (measured: a
            # Currency column after two Text columns carries 2, and a fixed
            # column added by ALTER TABLE carries the table's whole count).
            headers.append(column_header(spec, number, var_index, offset, layout.tag))
            if code != TYPE_BOOLEAN:
                fixed_offset += FIXED_SIZES.get(code, spec.size if isinstance(spec.size, int) else MAX_BINARY_BYTES)
        else:
            headers.append(column_header(spec, number, var_index, 0, layout.tag))
            var_index += 1
    var_count = var_index

    real_defs: list[bytes] = []
    logical_defs: list[bytes] = []
    # Real indexes keep creation order; the logical list (and its names)
    # is stored sorted by name, each entry naming its own index number.
    logical_order = sorted(range(len(indexes)), key=lambda i: indexes[i].name.lower())
    for i, spec in enumerate(indexes):
        raw = bytearray(SIZE_REAL_INDEX)
        struct.pack_into("<I", raw, 0, REAL_INDEX_LEAD)
        resolved = spec.resolved()
        if not 1 <= len(resolved) <= MAX_INDEX_COLUMNS:
            raise AccessError(f"index {spec.name!r}: 1..{MAX_INDEX_COLUMNS} columns")
        for slot in range(MAX_INDEX_COLUMNS):
            pos = 4 + 3 * slot
            if slot < len(resolved):
                col_name, ascending = resolved[slot]
                if col_name.lower() not in names_lower:
                    raise AccessError(f"index {spec.name!r}: no column {col_name!r}")
                struct.pack_into("<H", raw, pos, names_lower.index(col_name.lower()))
                raw[pos + 2] = 0x01 if ascending else 0x00
            else:
                struct.pack_into("<H", raw, pos, 0xFFFF)
        struct.pack_into("<I", raw, 34, layout.index_umap_refs[i])
        struct.pack_into("<I", raw, 38, layout.index_roots[i])
        flags = INDEX_ALWAYS_SET
        if spec.unique or spec.primary:
            flags |= INDEX_UNIQUE
        if spec.ignore_nulls:
            flags |= INDEX_IGNORE_NULLS
        if spec.required or spec.primary:
            flags |= INDEX_REQUIRED
        raw[46] = flags
        real_defs.append(bytes(raw))
        logical = bytearray(SIZE_LOGICAL_INDEX)
        struct.pack_into("<I", logical, 0, layout.tag)
        struct.pack_into("<I", logical, 4, i)
        struct.pack_into("<I", logical, 8, i)
        struct.pack_into("<I", logical, 13, NO_RELATIONSHIP)
        logical[21:23] = LOGICAL_INDEX_FLAGS
        logical[23] = INDEX_KIND_PRIMARY if spec.primary else INDEX_KIND_NORMAL
        logical_defs.append(bytes(logical))
    if sum(1 for s in indexes if s.primary) > 1:
        raise AccessError("a table has at most one primary key")

    body = bytearray()
    body += bytes(SIZE_REAL_INDEX_HEADER) * len(indexes)
    body += b"".join(headers)
    body += b"".join(_name(c.name) for c in columns)
    body += b"".join(real_defs)
    body += b"".join(logical_defs[i] for i in logical_order)
    body += b"".join(_name(indexes[i].name) for i in logical_order)
    for number, spec in enumerate(columns):
        if spec.type_code in LONG_VALUE_TYPES:
            owned, free = layout.column_map_refs.get(number, (0, 0))
            if not owned or not free:
                raise AccessError(f"column {spec.name!r} needs two usage maps")
            body += struct.pack("<HII", number, owned, free)
    body += b"\xff\xff"

    raw = bytearray(OFFSET_INDEX_HEADERS)
    raw[0] = PAGE_TDEF
    raw[1] = 0x01
    struct.pack_into("<I", raw, 8, OFFSET_INDEX_HEADERS + len(body))
    struct.pack_into("<I", raw, 0x0C, layout.tag)
    struct.pack_into("<i", raw, 0x18, 1)  # next complex-type AutoNumber
    raw[0x28] = TABLE_TYPE_USER
    struct.pack_into("<H", raw, 0x29, len(columns))
    struct.pack_into("<H", raw, 0x2B, var_count)
    struct.pack_into("<H", raw, 0x2D, len(columns))
    struct.pack_into("<I", raw, 0x2F, len(indexes))
    struct.pack_into("<I", raw, 0x33, len(indexes))
    struct.pack_into("<I", raw, 0x37, layout.owned_ref)
    struct.pack_into("<I", raw, 0x3B, layout.free_ref)
    return bytes(raw) + bytes(body)


def definition_page_count(length: int) -> int:
    """How many pages a definition of ``length`` bytes takes.  The engine
    counts ``PAGE_SIZE - 8`` bytes per page, so 4088 bytes fit one page and
    4089 already take two, the second holding nothing but its header."""
    return max(1, -(-length // DEFINITION_PAGE_SHARE))


def definition_pages(stream: bytes, chain: Sequence[int]) -> list[bytes]:
    """Lay a definition stream (header and body, as :func:`build_definition`
    and :func:`serialize_definition` return it) over its pages.  ``chain``
    names the continuation pages in chain order; the caller allocates them
    in ascending order and chains them in reverse, as the engine does.

    The first page physically holds the first 4096 bytes and every
    continuation the next 4088 after its 8-byte header, while the free
    words follow the engine's 4088-per-page accounting: 0 on every page
    but the last, ``4088 * pages - length`` on the last."""
    length = len(stream)
    count = definition_page_count(length)
    if len(chain) != count - 1:
        raise AccessError(f"a {length}-byte definition needs {count - 1} continuation pages, {len(chain)} given")
    images: list[bytes] = []
    for index in range(count):
        page = bytearray(PAGE_SIZE)
        page[0] = PAGE_TDEF
        page[1] = 0x01
        free = DEFINITION_PAGE_SHARE * count - length if index == count - 1 else 0
        struct.pack_into("<H", page, 2, free)
        struct.pack_into("<I", page, 4, chain[index] if index < count - 1 else 0)
        if index == 0:
            chunk = stream[:PAGE_SIZE]
            page[8 : len(chunk)] = chunk[8:]
        else:
            start = PAGE_SIZE + DEFINITION_PAGE_SHARE * (index - 1)
            chunk = stream[start : start + DEFINITION_PAGE_SHARE]
            page[8 : 8 + len(chunk)] = chunk
        images.append(bytes(page))
    return images


def serialize_definition(definition: TableDefinition) -> bytes:
    """The definition stream rebuilt from a parsed definition: fixed header
    fields from the object, every column and index from its raw bytes,
    logical indexes in name order.  Laid over the definition's pages with
    :func:`definition_pages`, an unchanged definition reproduces them byte
    for byte, and parsing gives the object back."""
    columns = definition.columns
    body = bytearray()
    for real in definition.real_indexes:
        # The header's two counters move as rows come and go; the raw bytes
        # were read before that, so the live values go in.
        header = bytearray(real.header_raw)
        struct.pack_into("<II", header, 0, real.row_count, real.entry_count)
        body += header
    body += b"".join(c.raw for c in columns)
    body += b"".join(_name(c.name) for c in columns)
    body += b"".join(r.raw for r in definition.real_indexes)
    order = sorted(range(len(definition.logical_indexes)), key=lambda i: definition.logical_indexes[i].name.lower())
    body += b"".join(definition.logical_indexes[i].raw for i in order)
    body += b"".join(_name(definition.logical_indexes[i].name) for i in order)
    for number, (owned, free) in definition.column_usage_maps.items():
        body += struct.pack("<HII", number, owned, free)
    body += b"\xff\xff"
    raw = bytearray(OFFSET_INDEX_HEADERS)
    raw[0] = PAGE_TDEF
    raw[1] = 0x01
    struct.pack_into("<I", raw, 8, OFFSET_INDEX_HEADERS + len(body))
    struct.pack_into("<I", raw, 0x0C, definition.tag)
    struct.pack_into("<I", raw, 0x10, definition.row_count)
    struct.pack_into("<I", raw, 0x14, definition.next_autonumber & 0xFFFFFFFF)
    struct.pack_into("<i", raw, 0x18, definition.complex_marker)
    struct.pack_into("<I", raw, 0x1C, definition.last_complex_id)
    raw[0x1C:0x28] = definition.header_raw[0x1C:0x28]
    raw[0x28] = definition.table_type
    struct.pack_into("<H", raw, 0x29, definition.max_columns)
    struct.pack_into("<H", raw, 0x2B, definition.var_column_count)
    struct.pack_into("<H", raw, 0x2D, len(columns))
    struct.pack_into("<I", raw, 0x2F, len(definition.logical_indexes))
    struct.pack_into("<I", raw, 0x33, len(definition.real_indexes))
    struct.pack_into("<I", raw, 0x37, definition.owned_pages_ref)
    struct.pack_into("<I", raw, 0x3B, definition.free_space_pages_ref)
    return bytes(raw) + bytes(body)


def new_index_parts(
    spec: IndexSpec,
    definition: TableDefinition,
    number: int,
    umap_ref: int,
    root_page: int,
) -> tuple[RealIndex, LogicalIndex]:
    """The real and logical index objects for an index added to an
    existing definition, with the raw bytes the page will carry."""
    names_lower = [c.name.lower() for c in definition.columns_by_number()]
    numbers = {c.name.lower(): c.number for c in definition.columns}
    raw = bytearray(SIZE_REAL_INDEX)
    struct.pack_into("<I", raw, 0, REAL_INDEX_LEAD)
    resolved = spec.resolved()
    if not 1 <= len(resolved) <= MAX_INDEX_COLUMNS:
        raise AccessError(f"index {spec.name!r}: 1..{MAX_INDEX_COLUMNS} columns")
    index_columns: list[IndexColumn] = []
    for slot in range(MAX_INDEX_COLUMNS):
        pos = 4 + 3 * slot
        if slot < len(resolved):
            col_name, ascending = resolved[slot]
            if col_name.lower() not in names_lower:
                raise AccessError(f"index {spec.name!r}: no column {col_name!r}")
            struct.pack_into("<H", raw, pos, numbers[col_name.lower()])
            raw[pos + 2] = 0x01 if ascending else 0x00
            index_columns.append(IndexColumn(numbers[col_name.lower()], ascending))
        else:
            struct.pack_into("<H", raw, pos, 0xFFFF)
    struct.pack_into("<I", raw, 34, umap_ref)
    struct.pack_into("<I", raw, 38, root_page)
    flags = INDEX_ALWAYS_SET
    if spec.unique or spec.primary:
        flags |= INDEX_UNIQUE
    if spec.ignore_nulls:
        flags |= INDEX_IGNORE_NULLS
    if spec.required or spec.primary:
        flags |= INDEX_REQUIRED
    raw[46] = flags
    real = RealIndex(
        header_raw=bytes(SIZE_REAL_INDEX_HEADER),
        entry_count=0,
        row_count=0,
        columns=index_columns,
        usage_map_ref=umap_ref,
        root_page=root_page,
        unknown=0,
        flags=flags,
        raw=bytes(raw),
    )
    logical = bytearray(SIZE_LOGICAL_INDEX)
    struct.pack_into("<I", logical, 0, definition.tag)
    struct.pack_into("<I", logical, 4, number)
    struct.pack_into("<I", logical, 8, number)
    struct.pack_into("<I", logical, 13, NO_RELATIONSHIP)
    logical[21:23] = LOGICAL_INDEX_FLAGS
    logical[23] = INDEX_KIND_PRIMARY if spec.primary else INDEX_KIND_NORMAL
    return real, LogicalIndex(
        name=spec.name,
        number=number,
        real_index=number,
        relationship_kind=0,
        relationship_index=NO_RELATIONSHIP,
        relationship_table_page=0,
        cascade_updates=bool(LOGICAL_INDEX_FLAGS[0]),
        cascade_deletes=bool(LOGICAL_INDEX_FLAGS[1]),
        kind=logical[23],
        raw=bytes(logical),
    )


RELATIONSHIP_REFERENCED = 0x01  # this table's index is the one referred to
RELATIONSHIP_REFERENCING = 0x02  # this table's index is the foreign key


def foreign_key_logical(
    tag: int,
    name: str,
    number: int,
    real_index: int,
    *,
    referencing: bool,
    other_logical: int,
    other_page: int,
    cascade_updates: bool,
    cascade_deletes: bool,
) -> LogicalIndex:
    """The 28-byte logical index entry a relationship adds to a table: the
    foreign key on the referencing side, a ``.r<letter>`` entry sharing the
    referenced unique index on the other.  Bytes 13..20 name the other
    table's logical index number and definition page, bytes 21 and 22
    carry the cascade flags (normal indexes hold ``04 04`` there), byte 23
    is the kind, 2."""
    raw = bytearray(SIZE_LOGICAL_INDEX)
    struct.pack_into("<I", raw, 0, tag)
    struct.pack_into("<I", raw, 4, number)
    struct.pack_into("<I", raw, 8, real_index)
    raw[12] = RELATIONSHIP_REFERENCING if referencing else RELATIONSHIP_REFERENCED
    struct.pack_into("<I", raw, 13, other_logical)
    struct.pack_into("<I", raw, 17, other_page)
    raw[21] = 1 if cascade_updates else 0
    raw[22] = 1 if cascade_deletes else 0
    raw[23] = INDEX_KIND_FOREIGN
    return LogicalIndex(
        name=name,
        number=number,
        real_index=real_index,
        relationship_kind=raw[12],
        relationship_index=other_logical,
        relationship_table_page=other_page,
        cascade_updates=cascade_updates,
        cascade_deletes=cascade_deletes,
        kind=INDEX_KIND_FOREIGN,
        raw=bytes(raw),
    )


def usage_map_page(row_count: int) -> bytes:
    """A fresh data-shaped page (owner 0) holding ``row_count`` empty inline
    usage maps."""
    from pyopenvba.access._datapage import DataPage

    page = DataPage.new(0)
    for _ in range(row_count):
        page.add_row(USAGE_MAP_ROW)
    return page.to_bytes()


def empty_index_root(owner: int) -> bytes:
    return serialize_index_page([], is_leaf=True, owner=owner, prev=0, next=0, tail=0, level=0, prefix=0)


def mark_definition_freed(raw: bytes) -> bytes:
    out = bytearray(raw)
    out[0] = FREED_TABLE_DEFINITION
    return bytes(out)


def data_page_free_space(raw: bytes) -> int:
    if raw[0] != PAGE_DATA:
        raise AccessError("not a data page")
    return struct.unpack_from("<H", raw, OFFSET_PAGE_FREE_SPACE)[0]
