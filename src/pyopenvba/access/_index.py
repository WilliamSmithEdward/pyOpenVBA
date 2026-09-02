"""B-tree indexes: node and leaf pages, their entries, and the key codec.

An index page (type 3 node, type 4 leaf) is laid out as::

    0x02  u16  free bytes in the entry area
    0x04  u32  owning table definition page
    0x0C  u32  previous leaf
    0x10  u32  next leaf
    0x14  u32  tail child (node pages): the child holding every key
               greater than the last entry's
    0x18  u16  prefix length: every entry after the first is stored
               without its first that-many bytes, which equal the first
               entry's
    0x1B  453-byte bit mask over the entry area; a set bit marks the END
               of an entry, the first entry starting at 0
    0x1E0 entry area

An entry is the encoded key columns, then the home slot of the row as a
big-endian three-byte page and a one-byte row, then on node pages the
big-endian child page.  A node entry carries the LAST key of its child.

Each key column starts with a flag byte: 0x7F for a value in an
ascending column, 0x80 descending, 0x00 for null ascending, 0xFF null
descending.  A descending value is the ascending encoding with every
byte inverted.  Ascending encodings, all big-endian:

    Boolean   one byte, 0x00 for True, 0xFF for False (True sorts first)
    Byte      the byte
    Integer, Long, BigInt, Currency (scaled by 10 000)
              the two's-complement value with the sign bit flipped
    Single, Double, DateTime
              the IEEE bits with the sign bit flipped when positive and
              every bit inverted when negative
    Decimal   0xFF then the 16-byte magnitude for a positive value
    GUID      the 16 bytes in textual order, run through the binary scheme
    Binary    eight-byte chunks, each followed by 0x09 when another chunk
              follows and by the count of real bytes in the last chunk
    Text      one or two collation bytes per character, 0x01, extra
              weight bytes, 0x00 -- see ``_collation``

Every rule above was read off ACE-written indexes, one column type at a
time, by pairing leaf entries with the rows they point at.
"""

from __future__ import annotations

import struct
import uuid
from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from decimal import Decimal

from pyopenvba.access_read import AccessError
from pyopenvba.access._pages import (
    PAGE_INDEX_LEAF,
    PAGE_INDEX_NODE,
    PAGE_SIZE,
    PageStore,
)
from pyopenvba.access._rows import decode_datetime
from pyopenvba.access._tdef import (
    TYPE_BIGINT,
    TYPE_BINARY,
    TYPE_BOOLEAN,
    TYPE_BYTE,
    TYPE_COMPLEX,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_FLOAT,
    TYPE_GUID,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MONEY,
    TYPE_NUMERIC,
    TYPE_TEXT,
    ColumnDef,
)

OFFSET_FREE_SPACE = 0x02
OFFSET_OWNER = 0x04
OFFSET_PREV = 0x0C
OFFSET_NEXT = 0x10
OFFSET_TAIL = 0x14
OFFSET_PREFIX_LENGTH = 0x18
OFFSET_ENTRY_MASK = 0x1B
SIZE_ENTRY_MASK = 453
OFFSET_ENTRIES = OFFSET_ENTRY_MASK + SIZE_ENTRY_MASK
ENTRY_AREA = PAGE_SIZE - OFFSET_ENTRIES

FLAG_ASCENDING = 0x7F
FLAG_DESCENDING = 0x80
FLAG_NULL_ASCENDING = 0x00
FLAG_NULL_DESCENDING = 0xFF
TEXT_END = 0x01
EXTRA_END = 0x00
BINARY_CHUNK = 8
BINARY_MORE = 0x09

FIXED_KEY_SIZES = {
    TYPE_BOOLEAN: 1,
    TYPE_BYTE: 1,
    TYPE_INT: 2,
    TYPE_LONG: 4,
    TYPE_COMPLEX: 4,
    TYPE_MONEY: 8,
    TYPE_FLOAT: 4,
    TYPE_DOUBLE: 8,
    TYPE_DATETIME: 8,
    TYPE_BIGINT: 8,
    TYPE_NUMERIC: 17,
}


@dataclass(frozen=True)
class TextKey:
    """A text column's key as stored: collation bytes, then the extra
    weight bytes that follow the 0x01 separator.  Case is not stored, so
    a key cannot be turned back into its text; two keys compare exactly
    as the engine compares the strings they came from."""

    primary: bytes
    extra: bytes


@dataclass
class IndexEntry:
    key: bytes
    page: int
    row: int
    child: int | None

    @property
    def is_node(self) -> bool:
        return self.child is not None


@dataclass
class IndexPage:
    number: int
    is_leaf: bool
    owner: int
    free_space: int
    prev: int
    next: int
    tail: int
    prefix_length: int
    entries: list[IndexEntry]


def parse_index_page(store: PageStore, number: int) -> IndexPage:
    raw = store.read(number)
    kind = raw[0]
    if kind not in (PAGE_INDEX_NODE, PAGE_INDEX_LEAF):
        raise AccessError(f"page {number} is type {kind:#04x}, not an index page")
    is_leaf = kind == PAGE_INDEX_LEAF
    prefix_length = struct.unpack_from("<H", raw, OFFSET_PREFIX_LENGTH)[0]
    mask = raw[OFFSET_ENTRY_MASK:OFFSET_ENTRIES]
    entries: list[IndexEntry] = []
    start = 0
    first: bytes | None = None
    for byte_index, byte in enumerate(mask):
        if not byte:
            continue
        for bit in range(8):
            if not byte & (1 << bit):
                continue
            end = byte_index * 8 + bit
            if end <= start or end > ENTRY_AREA:
                raise AccessError(
                    f"index page {number}: entry boundary {end} after {start} is impossible"
                )
            stored = raw[OFFSET_ENTRIES + start : OFFSET_ENTRIES + end]
            if first is None:
                first = stored
                full = stored
            else:
                if prefix_length > len(first):
                    raise AccessError(
                        f"index page {number}: prefix {prefix_length} longer than its first entry"
                    )
                full = first[:prefix_length] + stored
            entries.append(_split_entry(full, is_leaf, number))
            start = end
    return IndexPage(
        number=number,
        is_leaf=is_leaf,
        owner=struct.unpack_from("<I", raw, OFFSET_OWNER)[0],
        free_space=struct.unpack_from("<H", raw, OFFSET_FREE_SPACE)[0],
        prev=struct.unpack_from("<I", raw, OFFSET_PREV)[0],
        next=struct.unpack_from("<I", raw, OFFSET_NEXT)[0],
        tail=struct.unpack_from("<I", raw, OFFSET_TAIL)[0],
        prefix_length=prefix_length,
        entries=entries,
    )


def _split_entry(full: bytes, is_leaf: bool, number: int) -> IndexEntry:
    trailer = 4 if is_leaf else 8
    if len(full) < trailer:
        raise AccessError(f"index page {number}: entry of {len(full)} bytes has no row pointer")
    key = full[: len(full) - trailer]
    pointer = full[len(full) - trailer : len(full) - trailer + 4]
    page = int.from_bytes(pointer[:3], "big")
    row = pointer[3]
    child = None if is_leaf else int.from_bytes(full[-4:], "big")
    return IndexEntry(key=key, page=page, row=row, child=child)


def first_leaf(store: PageStore, root: int) -> IndexPage:
    """Descend from the root along first children to the leftmost leaf."""
    seen: set[int] = set()
    page = parse_index_page(store, root)
    while not page.is_leaf:
        if page.number in seen:
            raise AccessError(f"index rooted at {root} loops through page {page.number}")
        seen.add(page.number)
        child = page.entries[0].child if page.entries else page.tail
        if not child:
            raise AccessError(f"index node {page.number} has no children")
        page = parse_index_page(store, child)
    return page


def leaf_pages(store: PageStore, root: int) -> Iterator[IndexPage]:
    """Every leaf page in key order, following the next pointers from the
    leftmost leaf.  On a single-page index the root is the leaf."""
    page = first_leaf(store, root)
    seen: set[int] = set()
    while True:
        if page.number in seen:
            raise AccessError(f"index rooted at {root}: leaf chain loops at {page.number}")
        seen.add(page.number)
        yield page
        if not page.next:
            return
        page = parse_index_page(store, page.next)
        if not page.is_leaf:
            raise AccessError(f"index rooted at {root}: page {page.number} in the leaf chain is a node")


def leaf_entries(store: PageStore, root: int) -> Iterator[IndexEntry]:
    for page in leaf_pages(store, root):
        yield from page.entries


def node_pages(store: PageStore, root: int) -> Iterator[IndexPage]:
    """Every node page, root first, breadth-first."""
    queue = [root]
    seen: set[int] = set()
    while queue:
        number = queue.pop(0)
        if number in seen:
            raise AccessError(f"index rooted at {root} loops through page {number}")
        seen.add(number)
        page = parse_index_page(store, number)
        if page.is_leaf:
            continue
        yield page
        queue.extend(e.child for e in page.entries if e.child)
        if page.tail:
            queue.append(page.tail)


# --- key codec ---------------------------------------------------------------


def _invert(raw: bytes) -> bytes:
    return bytes(b ^ 0xFF for b in raw)


def _from_sign_flipped(raw: bytes) -> int:
    value = int.from_bytes(raw, "big")
    value ^= 1 << (8 * len(raw) - 1)
    if value >= 1 << (8 * len(raw) - 1):
        value -= 1 << (8 * len(raw))
    return value


def _float_bits(raw: bytes) -> bytes:
    """Undo the float ordering trick: sign bit set means positive with the
    bit flipped; sign bit clear means negative with every bit inverted."""
    if raw[0] & 0x80:
        return bytes((raw[0] ^ 0x80,)) + raw[1:]
    return _invert(raw)


def _read_binary_key(key: bytes, pos: int) -> tuple[bytes, int]:
    out = bytearray()
    while True:
        chunk = key[pos : pos + BINARY_CHUNK]
        marker_pos = pos + BINARY_CHUNK
        if len(chunk) < BINARY_CHUNK or marker_pos >= len(key):
            raise AccessError("binary key runs off the end of the entry")
        marker = key[marker_pos]
        pos = marker_pos + 1
        if marker == BINARY_MORE:
            out.extend(chunk)
            continue
        if not 1 <= marker <= BINARY_CHUNK:
            raise AccessError(f"binary key has an impossible chunk marker {marker:#04x}")
        out.extend(chunk[:marker])
        return bytes(out), pos


def _read_text_key(key: bytes, pos: int) -> tuple[TextKey, int]:
    end = key.find(bytes((TEXT_END,)), pos)
    if end < 0:
        raise AccessError("text key has no 0x01 separator")
    extra_end = key.find(bytes((EXTRA_END,)), end + 1)
    if extra_end < 0:
        raise AccessError("text key has no 0x00 terminator")
    return TextKey(key[pos:end], key[end + 1 : extra_end]), extra_end + 1


def decode_key(key: bytes, columns: Sequence[tuple[ColumnDef, bool]]) -> list[object]:
    """Decode the key columns of one entry to Python values, ``None`` for
    null and :class:`TextKey` for text.  ``columns`` pairs each column
    definition with its ascending flag."""
    values: list[object] = []
    pos = 0
    for column, ascending in columns:
        if pos >= len(key):
            raise AccessError(f"key ends before column {column.name!r}")
        flag = key[pos]
        pos += 1
        if flag in (FLAG_NULL_ASCENDING, FLAG_NULL_DESCENDING):
            values.append(None)
            continue
        if flag not in (FLAG_ASCENDING, FLAG_DESCENDING):
            raise AccessError(f"column {column.name!r}: unknown key flag {flag:#04x}")
        descending = flag == FLAG_DESCENDING
        if descending != (not ascending):
            raise AccessError(
                f"column {column.name!r}: key flag says "
                f"{'descending' if descending else 'ascending'}, the index says otherwise"
            )
        code = column.type_code
        if code in FIXED_KEY_SIZES:
            size = FIXED_KEY_SIZES[code]
            raw = key[pos : pos + size]
            if len(raw) != size:
                raise AccessError(f"column {column.name!r}: key truncated")
            pos += size
            if descending:
                raw = _invert(raw)
            values.append(_decode_fixed(column, raw))
            continue
        if code == TYPE_GUID:
            # GUIDs go through the binary scheme: 16 bytes become 8 + 0x09 + 8 + 0x08.
            region = key[pos : pos + 18]
            if descending:
                region = _invert(region)
            data, consumed = _read_binary_key(region, 0)
            pos += consumed
            if len(data) != 16:
                raise AccessError(f"column {column.name!r}: GUID key holds {len(data)} bytes")
            values.append(uuid.UUID(bytes=data))
            continue
        if code == TYPE_BINARY:
            region = key[pos:] if not descending else _invert(key[pos:])
            data, consumed = _read_binary_key(region, 0)
            pos += consumed
            values.append(data)
            continue
        if code == TYPE_TEXT:
            region = key[pos:] if not descending else _invert(key[pos:])
            text, consumed = _read_text_key(region, 0)
            pos += consumed
            values.append(text)
            continue
        raise AccessError(f"column {column.name!r}: type {column.type_name} is not indexable")
    if pos != len(key):
        raise AccessError(f"{len(key) - pos} unexplained bytes after the last key column")
    return values


def _decode_fixed(column: ColumnDef, raw: bytes) -> object:
    code = column.type_code
    if code == TYPE_BOOLEAN:
        if raw[0] == 0x00:
            return True
        if raw[0] == 0xFF:
            return False
        raise AccessError(f"column {column.name!r}: Boolean key byte {raw[0]:#04x}")
    if code == TYPE_BYTE:
        return raw[0]
    if code in (TYPE_INT, TYPE_LONG, TYPE_BIGINT, TYPE_COMPLEX):
        return _from_sign_flipped(raw)
    if code == TYPE_MONEY:
        return Decimal(_from_sign_flipped(raw)).scaleb(-4)
    if code == TYPE_FLOAT:
        return struct.unpack(">f", _float_bits(raw))[0]
    if code == TYPE_DOUBLE:
        return struct.unpack(">d", _float_bits(raw))[0]
    if code == TYPE_DATETIME:
        return decode_datetime(struct.pack("<d", struct.unpack(">d", _float_bits(raw))[0]))
    if code == TYPE_NUMERIC:
        sign = raw[0]
        magnitude = raw[1:]
        if sign == 0xFF:
            value = Decimal(int.from_bytes(magnitude, "big"))
        elif sign == 0x00:
            value = -Decimal(int.from_bytes(_invert(magnitude), "big"))
        else:
            raise AccessError(f"column {column.name!r}: Decimal key sign byte {sign:#04x}")
        return value.scaleb(-column.scale)
    raise AccessError(f"column {column.name!r}: no fixed key codec for {column.type_name}")
