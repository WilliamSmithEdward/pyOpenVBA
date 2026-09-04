"""The page layer of a Jet 4 / ACE database file.

A database is an array of 4 KiB pages.  Page 0 is the database definition
page, part of which is masked with a fixed RC4 keystream; page 1 holds the
global usage map; every other page carries a one-byte type tag.  Rows on
data-shaped pages (data, LVAL, usage-map pages) share one slot layout, read
here so the table, long-value and usage-map code do not each reinvent it.

Nothing in this module knows what a table is.  Everything here was checked
against Access-written files and, where a rule is not visible in a file,
against Access itself.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from pyopenvba.access_read import AccessError
from pyopenvba.exceptions import UnsupportedFormatError

PAGE_SIZE = 4096

# Page type tags (first byte of every page except page 0, whose tag is 0).
PAGE_DB_DEF = 0x00
PAGE_DATA = 0x01
PAGE_TDEF = 0x02
PAGE_RETIRED = 0x09  # a data-shaped page emptied of rows and given back
PAGE_INDEX_NODE = 0x03
PAGE_INDEX_LEAF = 0x04
PAGE_USAGE_BITMAP = 0x05

# A long-value page is a data page whose owner field spells "LVAL".
LVAL_OWNER_TAG = 0x4C41564C

ACE_SIGNATURE = b"Standard ACE DB\x00"
JET4_SIGNATURE = b"Standard Jet DB\x00"

# Format version, u32 at 0x14 of page 0.
VERSION_JET3 = 0
VERSION_JET4 = 1
VERSION_ACE_2007 = 2
VERSION_ACE_2010 = 3
VERSION_ACE_2016 = 5

# Bytes 0x18..0x96 of page 0 are XORed with the RC4 keystream of a fixed
# key.  Jet 3 masks 114 bytes; Jet 4 and every ACE version mask 126.
MASK_OFFSET = 0x18
MASK_LENGTH = 126
_RC4_KEY = bytes((0xC7, 0xDA, 0x39, 0x6B))

# Fields inside the masked region.
OFFSET_VERSION = 0x14
OFFSET_CODE_PAGE = 0x3C
OFFSET_ENCODING_KEY = 0x3E
OFFSET_PASSWORD = 0x42
PASSWORD_LENGTH = 40
OFFSET_SORT_ORDER = 0x6E
OFFSET_SORT_VERSION = 0x70
OFFSET_CREATION_DATE = 0x72

# Row slot flags in the u16 offset table of a data-shaped page.  Measured
# on Access-written pages: a live row has neither bit; a row that was
# moved to another page keeps a 4-byte pointer here under 0x4000 alone,
# and its old bytes may still trail the pointer; the moved row's own slot
# on the target page carries 0x8000, so 0x8000 means "deleted" only on a
# page reached directly; a dead slot has both bits (0xC000), often with a
# stale offset shared with a neighbour.
ROW_OFFSET_MASK = 0x1FFF
ROW_DELETED = 0x8000
ROW_OVERFLOW = 0x4000

# Data-shaped page header.
OFFSET_PAGE_FREE_SPACE = 0x02
OFFSET_PAGE_OWNER = 0x04
OFFSET_PAGE_ROW_COUNT = 0x0C
OFFSET_PAGE_ROW_TABLE = 0x0E


def _rc4_keystream(key: bytes, length: int) -> bytes:
    box = list(range(256))
    j = 0
    for i in range(256):
        j = (j + box[i] + key[i % len(key)]) & 0xFF
        box[i], box[j] = box[j], box[i]
    out = bytearray()
    i = j = 0
    for _ in range(length):
        i = (i + 1) & 0xFF
        j = (j + box[i]) & 0xFF
        box[i], box[j] = box[j], box[i]
        out.append(box[(box[i] + box[j]) & 0xFF])
    return bytes(out)


HEADER_MASK = _rc4_keystream(_RC4_KEY, MASK_LENGTH)


def toggle_definition_mask(page: bytes) -> bytearray:
    """Apply (or remove -- XOR is its own inverse) the page-0 header mask."""
    out = bytearray(page)
    for i, m in enumerate(HEADER_MASK):
        out[MASK_OFFSET + i] ^= m
    return out


@dataclass(frozen=True)
class DatabaseHeader:
    """The fields of page 0 a reader needs, decoded from the unmasked page."""

    version: int
    code_page: int
    encoding_key: int
    sort_order: int
    sort_version: int
    creation_date: float
    password: bytes

    @property
    def is_ace(self) -> bool:
        return self.version >= VERSION_ACE_2007

    @classmethod
    def from_page(cls, raw_page0: bytes) -> DatabaseHeader:
        signature = raw_page0[4 : 4 + len(ACE_SIGNATURE)]
        if signature not in (ACE_SIGNATURE, JET4_SIGNATURE):
            raise UnsupportedFormatError(
                f"not a Jet 4 / ACE database (signature {bytes(signature)!r})"
            )
        version = struct.unpack_from("<I", raw_page0, OFFSET_VERSION)[0]
        if version == VERSION_JET3:
            raise UnsupportedFormatError(
                "Jet 3 (Access 97) databases use 2 KiB pages and are not supported"
            )
        plain = toggle_definition_mask(raw_page0)
        creation = struct.unpack_from("<d", plain, OFFSET_CREATION_DATE)[0]
        # The password is XORed with the whole-day part of the creation
        # date as a 4-byte integer, repeated; a database without a
        # password therefore shows that integer over and over.
        date_mask = struct.pack("<i", int(creation))
        masked = plain[OFFSET_PASSWORD : OFFSET_PASSWORD + PASSWORD_LENGTH]
        password = bytes(b ^ date_mask[i % 4] for i, b in enumerate(masked))
        if not any(password):
            password = b""
        return cls(
            version=version,
            code_page=struct.unpack_from("<H", plain, OFFSET_CODE_PAGE)[0],
            encoding_key=struct.unpack_from("<I", plain, OFFSET_ENCODING_KEY)[0],
            sort_order=struct.unpack_from("<H", plain, OFFSET_SORT_ORDER)[0],
            sort_version=struct.unpack_from("<H", plain, OFFSET_SORT_VERSION)[0],
            creation_date=creation,
            password=password,
        )


class PageStore:
    """A database file held in memory as whole pages."""

    def __init__(self, data: bytes) -> None:
        if len(data) < 2 * PAGE_SIZE:
            raise AccessError(
                f"file too small to be a database ({len(data)} bytes)"
            )
        if len(data) % PAGE_SIZE:
            raise AccessError(
                f"file length {len(data)} is not a multiple of the "
                f"{PAGE_SIZE}-byte page size"
            )
        self._data = bytearray(data)
        #: Pages released while this store has been open.  The engine hands
        #: a released page out again only after the database is reopened,
        #: so within a session the allocator skips these.
        self.released: set[int] = set()
        #: Pages allocated while this store has been open.
        self.allocated: set[int] = set()
        #: Pages given back by a definition rewrite, or by a freed value that
        #: this session had allocated: they come back into use together once
        #: five are waiting (measured on CREATE INDEX runs).
        self.pending: list[int] = []
        #: Per long-value column (keyed by its free-space map reference),
        #: the LVAL page most recently written this session: the engine
        #: tries it first for the next single-row value.
        self.lval_cursor: dict[int, int] = {}

    def truncate(self, page_count: int) -> int:
        """Drop pages from the end of the file, and say how many went.

        Nothing checks here that they are free -- the caller does that --
        and the first two pages are never dropped.
        """
        if page_count < 2 or page_count >= self.page_count:
            return 0
        dropped = self.page_count - page_count
        del self._data[page_count * PAGE_SIZE :]
        self.released = {p for p in self.released if p < page_count}
        self.allocated = {p for p in self.allocated if p < page_count}
        self.pending = [p for p in self.pending if p < page_count]
        self.lval_cursor = {k: v for k, v in self.lval_cursor.items() if v < page_count}
        return dropped

    def snapshot(self) -> tuple[bytes, set[int], set[int], list[int], dict[int, int]]:
        """Everything a rollback has to put back: the pages and the state
        the session keeps about them."""
        return (bytes(self._data), set(self.released), set(self.allocated), list(self.pending), dict(self.lval_cursor))

    def restore(self, state: tuple[bytes, set[int], set[int], list[int], dict[int, int]]) -> None:
        data, released, allocated, pending, cursor = state
        self._data = bytearray(data)
        self.released = released
        self.allocated = allocated
        self.pending = pending
        self.lval_cursor = cursor

    @property
    def page_count(self) -> int:
        return len(self._data) // PAGE_SIZE

    def read(self, page: int) -> bytes:
        if not 0 <= page < self.page_count:
            raise AccessError(
                f"page {page} out of range (0..{self.page_count - 1})"
            )
        start = page * PAGE_SIZE
        return bytes(self._data[start : start + PAGE_SIZE])

    def write(self, page: int, content: bytes) -> None:
        if len(content) != PAGE_SIZE:
            raise AccessError(
                f"page content must be {PAGE_SIZE} bytes, got {len(content)}"
            )
        if not 0 <= page < self.page_count:
            raise AccessError(
                f"page {page} out of range (0..{self.page_count - 1})"
            )
        start = page * PAGE_SIZE
        self._data[start : start + PAGE_SIZE] = content

    def append(self) -> int:
        """Grow the file by one zeroed page and return its number."""
        self._data.extend(bytes(PAGE_SIZE))
        return self.page_count - 1

    def page_type(self, page: int) -> int:
        return self._data[page * PAGE_SIZE]

    def to_bytes(self) -> bytes:
        return bytes(self._data)


# --- rows on data-shaped pages ---------------------------------------------


def row_slots(page: bytes) -> list[int]:
    """The raw u16 slot table of a data-shaped page (flags included)."""
    count = struct.unpack_from("<H", page, OFFSET_PAGE_ROW_COUNT)[0]
    if OFFSET_PAGE_ROW_TABLE + 2 * count > PAGE_SIZE:
        raise AccessError(f"row count {count} does not fit on the page")
    return list(
        struct.unpack_from(f"<{count}H", page, OFFSET_PAGE_ROW_TABLE)
    )


def row_span(slots: list[int], slot: int) -> tuple[int, int]:
    """Byte range of a row: rows are laid down from the page end, so a row
    ends where the previous slot's row starts.  A deleted slot still
    bounds its neighbour."""
    if not 0 <= slot < len(slots):
        raise AccessError(f"slot {slot} out of range (0..{len(slots) - 1})")
    start = slots[slot] & ROW_OFFSET_MASK
    end = PAGE_SIZE if slot == 0 else slots[slot - 1] & ROW_OFFSET_MASK
    if end < start:
        raise AccessError(f"slot {slot} spans {start}..{end}, which is inverted")
    return start, end


def row_bytes(page: bytes, slot: int, *, overflow_target: bool = False) -> bytes | None:
    """The bytes of one row, or ``None`` when the slot is dead.  An
    overflow slot returns its 4-byte pointer (any stale bytes after it are
    dropped); callers follow it with :func:`row_pointer` and read the
    target with ``overflow_target=True``, where the 0x8000 bit marks the
    moved row rather than a deletion."""
    slots = row_slots(page)
    entry = slots[slot]
    if entry & ROW_DELETED and not overflow_target:
        return None
    if overflow_target and entry & ROW_OVERFLOW:
        raise AccessError(f"overflow target slot {slot} is itself flagged as overflow")
    start, end = row_span(slots, slot)
    if entry & ROW_OVERFLOW:
        return page[start : start + 4]
    return page[start:end]


def row_pointer(pointer: bytes) -> tuple[int, int]:
    """Decode the 4-byte ``(row, page)`` reference Jet uses everywhere: one
    byte of row number, then a three-byte little-endian page number."""
    if len(pointer) < 4:
        raise AccessError("row pointer shorter than 4 bytes")
    return pointer[0], int.from_bytes(pointer[1:4], "little")


def encode_row_pointer(page: int, row: int) -> bytes:
    if not 0 <= row <= 0xFF or not 0 <= page <= 0xFFFFFF:
        raise AccessError(f"row pointer ({page}, {row}) out of range")
    return bytes((row,)) + page.to_bytes(3, "little")


def page_owner(page: bytes) -> int:
    return struct.unpack_from("<I", page, OFFSET_PAGE_OWNER)[0]


def is_lval_page(page: bytes) -> bool:
    return page[0] == PAGE_DATA and page_owner(page) == LVAL_OWNER_TAG


# --- usage maps --------------------------------------------------------------
#
# A usage map is a row on a data-shaped page.  Its first byte is the kind:
# 0 is an inline bitmap preceded by a u32 start page, 1 is a list of u32
# page numbers, each naming a type-5 page whose bytes from offset 4 form
# one 32 736-page bitmap chunk (0 meaning no chunk allocated yet).  A set
# bit means "this page is in the map".  For a table's owned-pages and
# free-space maps that is membership; for the global map on page 1 it is
# "this page is free", and pages past the end of the file count as free.

USAGE_MAP_INLINE = 0
USAGE_MAP_REFERENCE = 1
USAGE_BITMAP_PAGE_DATA = 4
PAGES_PER_BITMAP_PAGE = (PAGE_SIZE - USAGE_BITMAP_PAGE_DATA) * 8

GLOBAL_USAGE_MAP_PAGE = 1
GLOBAL_USAGE_MAP_ROW = 0


def usage_map_location(reference: int) -> tuple[int, int]:
    """Split a u32 usage-map reference into ``(page, row)``."""
    return reference >> 8, reference & 0xFF


@dataclass
class UsageMap:
    """A decoded usage map plus where it lives, so it can be rewritten."""

    page: int
    row: int
    kind: int
    start_page: int
    bitmap: bytearray
    reference_pages: list[int]

    def pages(self) -> list[int]:
        out: list[int] = []
        if self.kind == USAGE_MAP_INLINE:
            for byte_index, byte in enumerate(self.bitmap):
                if not byte:
                    continue
                for bit in range(8):
                    if byte & (1 << bit):
                        out.append(self.start_page + byte_index * 8 + bit)
            return out
        for chunk, _ in enumerate(self.reference_pages):
            base = chunk * PAGES_PER_BITMAP_PAGE
            chunk_bytes = self.bitmap[
                chunk * (PAGE_SIZE - USAGE_BITMAP_PAGE_DATA) :
                (chunk + 1) * (PAGE_SIZE - USAGE_BITMAP_PAGE_DATA)
            ]
            for byte_index, byte in enumerate(chunk_bytes):
                if not byte:
                    continue
                for bit in range(8):
                    if byte & (1 << bit):
                        out.append(base + byte_index * 8 + bit)
        return out

    def contains(self, page: int) -> bool:
        if self.kind == USAGE_MAP_INLINE:
            index = page - self.start_page
            if index < 0 or index // 8 >= len(self.bitmap):
                return False
        else:
            index = page
            if index // 8 >= len(self.bitmap):
                return False
        return bool(self.bitmap[index // 8] & (1 << (index % 8)))


def read_usage_map(store: PageStore, page: int, row: int) -> UsageMap:
    raw = row_bytes(store.read(page), row)
    if raw is None:
        raise AccessError(f"usage map row ({page}, {row}) is deleted")
    if not raw:
        raise AccessError(f"usage map row ({page}, {row}) is empty")
    kind = raw[0]
    if kind == USAGE_MAP_INLINE:
        if len(raw) < 5:
            raise AccessError(f"inline usage map ({page}, {row}) is truncated")
        start = struct.unpack_from("<I", raw, 1)[0]
        return UsageMap(page, row, kind, start, bytearray(raw[5:]), [])
    if kind == USAGE_MAP_REFERENCE:
        count = (len(raw) - 1) // 4
        refs = list(struct.unpack_from(f"<{count}I", raw, 1))
        bitmap = bytearray()
        for ref in refs:
            if ref == 0:
                bitmap.extend(bytes(PAGE_SIZE - USAGE_BITMAP_PAGE_DATA))
                continue
            chunk = store.read(ref)
            if chunk[0] != PAGE_USAGE_BITMAP:
                raise AccessError(
                    f"usage map ({page}, {row}) references page {ref}, "
                    f"which is type {chunk[0]:#04x}, not a usage bitmap"
                )
            bitmap.extend(chunk[USAGE_BITMAP_PAGE_DATA:])
        return UsageMap(page, row, kind, 0, bitmap, refs)
    raise AccessError(f"usage map ({page}, {row}) has unknown kind {kind}")


def read_usage_map_ref(store: PageStore, reference: int) -> UsageMap:
    page, row = usage_map_location(reference)
    return read_usage_map(store, page, row)
