"""Long values: Memo and OLE data that does not live in the row.

The row holds a 12-byte definition (see :class:`LongValueRef`).  Its kind
byte says where the bytes are:

    0x80  right after the definition, in the row itself
    0x40  one row on an LVAL page, the whole value
    0x00  a chain of LVAL rows; each begins with a 4-byte pointer to the
          next (row byte, three-byte page), (0, 0) on the last

Where the engine puts a value, measured on values it wrote
(docs/access_engine.md): up to 64 bytes inline; up to
``LVAL_SINGLE_MAX`` bytes as one row on an LVAL page shared with other
values of the same column; beyond that as a chain of 4072-byte payloads,
one chunk per page, the first page carrying the same 4-byte stamp as the
row's definition.  Memo text is compressed inline when that is shorter
and stored uncompressed outside the row.
"""

from __future__ import annotations

import struct

from pyopenvba.access_read import AccessError
from pyopenvba.access._alloc import (
    add_to_map,
    allocate_page,
    release_page,
    remove_from_map,
)
from pyopenvba.access._datapage import DataPage
from pyopenvba.access._pages import (
    LVAL_OWNER_TAG,
    OFFSET_PAGE_FREE_SPACE,
    OFFSET_PAGE_OWNER,
    PAGE_DATA,
    PAGE_SIZE,
    PageStore,
    encode_row_pointer,
    is_lval_page,
    read_usage_map_ref,
    row_bytes,
    row_pointer,
)
from pyopenvba.access._rows import LongValueRef, encode_text

LVAL_INLINE_MAX = 64
LVAL_SINGLE_MAX = 3816
LVAL_CHUNK_PAYLOAD = 4072
LVAL_PAGE_MIN_FREE = 257  # listed in the free-space map while free space exceeds 256 bytes
LVAL_SHARED_MAX = 256  # a value this size or smaller fits any page the map lists
OFFSET_LVAL_STAMP = 0x08


# --- reading -------------------------------------------------------------------


def read_long_value(store: PageStore, ref: LongValueRef) -> bytes:
    if ref.kind == LongValueRef.KIND_INLINE:
        return ref.inline
    if ref.kind == LongValueRef.KIND_SINGLE_PAGE:
        row = _lval_row(store, ref.page, ref.row)
        if len(row) != ref.length:
            raise AccessError(
                f"long value at ({ref.page}, {ref.row}) holds {len(row)} bytes, "
                f"its definition says {ref.length}"
            )
        return row
    if ref.kind == LongValueRef.KIND_CHAINED:
        out = bytearray()
        for _page, _row, chunk in _chain(store, ref):
            out.extend(chunk[4:])
        if len(out) != ref.length:
            raise AccessError(
                f"long value chain holds {len(out)} bytes, its definition says {ref.length}"
            )
        return bytes(out)
    raise AccessError(f"unknown long-value kind {ref.kind:#04x}")


def _chain(store: PageStore, ref: LongValueRef) -> list[tuple[int, int, bytes]]:
    out: list[tuple[int, int, bytes]] = []
    seen: set[tuple[int, int]] = set()
    page, row = ref.page, ref.row
    while page:
        if (page, row) in seen:
            raise AccessError(f"long value chain loops at ({page}, {row})")
        seen.add((page, row))
        chunk = _lval_row(store, page, row)
        if len(chunk) < 4:
            raise AccessError(f"long value chunk at ({page}, {row}) is too short for its pointer")
        out.append((page, row, chunk))
        row, page = row_pointer(chunk)
    return out


def _lval_row(store: PageStore, page: int, row: int) -> bytes:
    raw_page = store.read(page)
    if not is_lval_page(raw_page):
        raise AccessError(f"page {page} is not a long-value page")
    data = row_bytes(raw_page, row)
    if data is None:
        raise AccessError(f"long value row ({page}, {row}) is deleted")
    return data


# --- writing -------------------------------------------------------------------


def memo_bytes(text: str) -> bytes:
    """Bytes to store for a Memo: compressed when the value will live
    inline and compression shortens it, otherwise plain UTF-16LE."""
    plain = text.encode("utf-16-le")
    if len(plain) <= LVAL_INLINE_MAX:
        compressed = encode_text(text)
        if len(compressed) < len(plain):
            return compressed
    return plain


def _definition(length: int, kind: int, row: int, page: int, stamp: int) -> bytes:
    if length >= 1 << 24:
        raise AccessError(f"a long value of {length} bytes exceeds the format's limit")
    return length.to_bytes(3, "little") + bytes((kind, row)) + page.to_bytes(3, "little") + struct.pack("<I", stamp)


def new_lval_page(store: PageStore, stamp: int = 0) -> int:
    page = allocate_page(store)
    raw = bytearray(PAGE_SIZE)
    raw[0] = PAGE_DATA
    raw[1] = 0x01
    struct.pack_into("<H", raw, OFFSET_PAGE_FREE_SPACE, PAGE_SIZE - 14)
    struct.pack_into("<I", raw, OFFSET_PAGE_OWNER, LVAL_OWNER_TAG)
    struct.pack_into("<I", raw, OFFSET_LVAL_STAMP, stamp)
    store.write(page, bytes(raw))
    return page


def write_long_value(store: PageStore, maps: tuple[int, int], data: bytes, stamp: int) -> bytes:
    """Store ``data`` for a Memo/OLE column whose (owned, free-space) usage
    map references are ``maps``; returns the column's row bytes -- the
    12-byte definition, followed by the data itself when inline."""
    owned_ref, free_ref = maps
    if len(data) <= LVAL_INLINE_MAX:
        return _definition(len(data), LongValueRef.KIND_INLINE, 0, 0, 0) + data
    if len(data) <= LVAL_SINGLE_MAX:
        page = _single_row_page(store, maps, len(data))
        lval = DataPage(store.read(page))
        slot = lval.add_row(data)
        store.write(page, lval.to_bytes())
        if len(data) > LVAL_SHARED_MAX:
            # Only a large value moves the column's cursor: a small one
            # takes the first listed page and leaves the run alone
            # (measured on a compaction's 200 memos of wrapping sizes).
            store.lval_cursor[free_ref] = page
        if lval.free_space < LVAL_PAGE_MIN_FREE and free_ref:
            remove_from_map(store, free_ref, page)
        return _definition(len(data), LongValueRef.KIND_SINGLE_PAGE, slot, page, 0)
    # A chain: one fresh page per chunk, linked front to back, so the
    # pages are made first and the pointers filled in from the last one.
    chunks = [data[i : i + LVAL_CHUNK_PAYLOAD] for i in range(0, len(data), LVAL_CHUNK_PAYLOAD)]
    pages = [new_lval_page(store, stamp if i == 0 else 0) for i in range(len(chunks))]
    for page in pages:
        add_to_map(store, owned_ref, page)
    next_pointer = encode_row_pointer(0, 0)
    for page, chunk in zip(reversed(pages), reversed(chunks)):
        lval = DataPage(store.read(page))
        slot = lval.add_row(next_pointer + chunk)
        store.write(page, lval.to_bytes())
        next_pointer = encode_row_pointer(page, slot)
    return _definition(len(data), LongValueRef.KIND_CHAINED, 0, pages[0], stamp)


def _single_row_page(store: PageStore, maps: tuple[int, int], length: int) -> int:
    """An LVAL page of this column with room for a row of ``length``
    bytes, chosen as the engine chooses.

    The free-space map lists a page while more than 256 bytes are free,
    so any listed page has room for a value of 256 bytes or fewer and the
    engine takes the first of them.  A larger value needs the page
    checked, and the engine checks one: the last page the map lists,
    which is the one it grew the column with.  When that page cannot take
    the value it starts another rather than looking further back
    (measured: with the last page holding 1676 free, a 258-byte value
    shared it, a 1706-byte value went to a new page, and the first page,
    3827 bytes free, took neither).

    For a larger value the page the last large value went to comes first,
    which is what keeps a run of values together (measured: with page A
    holding 1080 free and the last write on page B, a 900-byte value
    went to B; the next one, B full, went to A).  A small value never
    looks there: it takes the first listed page (measured on a
    compaction, where an 80-byte memo went back to the first page while
    the run continued on the third)."""
    owned_ref, free_ref = maps
    cursor = store.lval_cursor.get(free_ref)
    if length > LVAL_SHARED_MAX and cursor is not None and cursor < store.page_count:
        raw = store.read(cursor)
        if is_lval_page(raw) and DataPage(raw).fits(length):
            return cursor
    # A column the engine gave no free-space map (Access's own MSysNameMap
    # and MSysAccessXML) has no listed pages to look through.
    listed = [
        candidate
        for candidate in (read_usage_map_ref(store, free_ref).pages() if free_ref else [])
        if candidate < store.page_count and is_lval_page(store.read(candidate))
    ]
    for candidate in listed if length <= LVAL_SHARED_MAX else listed[-1:]:
        if DataPage(store.read(candidate)).fits(length):
            return candidate
    page = new_lval_page(store)
    add_to_map(store, owned_ref, page)
    if free_ref:
        add_to_map(store, free_ref, page)
    return page


def free_long_value(store: PageStore, maps: tuple[int, int], ref: LongValueRef) -> None:
    """Give back what a long value occupied: nothing for inline, the row
    for a single-page value, every page for a chain (released to the
    global map and dropped from the column's owned map, content left in
    place, as the engine does)."""
    owned_ref, free_ref = maps
    if ref.kind == LongValueRef.KIND_INLINE:
        return
    if ref.kind == LongValueRef.KIND_SINGLE_PAGE:
        lval = DataPage(store.read(ref.page))
        lval.remove_row(ref.row)
        if lval.live_rows and lval.free_space >= LVAL_PAGE_MIN_FREE:
            # Room came back: the page is listed again (measured: a page
            # at 176 free was unlisted, at 1476 after a delete listed).
            store.write(ref.page, lval.to_bytes())
            if free_ref:
                add_to_map(store, free_ref, ref.page)
            return
        if lval.live_rows == 0:
            # An LVAL page that lost its last value is retired: type 0x09,
            # released, and out of both of the column's maps (measured on
            # a column's only LVAL page as well as one of several).
            lval.retire()
            store.write(ref.page, lval.to_bytes())
            release_page(store, ref.page)
            remove_from_map(store, owned_ref, ref.page)
            if free_ref:
                remove_from_map(store, free_ref, ref.page)
            if store.lval_cursor.get(free_ref) == ref.page:
                del store.lval_cursor[free_ref]
            return
        store.write(ref.page, lval.to_bytes())
        return
    if ref.kind == LongValueRef.KIND_CHAINED:
        for page, _row, _chunk in _chain(store, ref):
            release_page(store, page, kind="value")
            remove_from_map(store, owned_ref, page)
        return
    raise AccessError(f"unknown long-value kind {ref.kind:#04x}")
