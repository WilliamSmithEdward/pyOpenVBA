"""Page allocation and usage-map maintenance.

The global usage map (page 1, row 0) marks free pages, pages past the end
of the file included; the engine takes the lowest free page and grows the
file to reach it.  A table's owned-pages and free-space maps, and an
index's owned-pages map, are the same structure and are edited in place.
"""

from __future__ import annotations

import struct

from pyopenvba.access_read import AccessError
from pyopenvba.access._pages import (
    GLOBAL_USAGE_MAP_PAGE,
    GLOBAL_USAGE_MAP_ROW,
    PAGE_SIZE,
    PAGE_USAGE_BITMAP,
    PAGES_PER_BITMAP_PAGE,
    USAGE_BITMAP_PAGE_DATA,
    USAGE_MAP_INLINE,
    USAGE_MAP_REFERENCE,
    PageStore,
    UsageMap,
    read_usage_map,
    read_usage_map_ref,
    row_slots,
    row_span,
)


#: The global map is extended a whole byte-step at a time, 64 pages.
INLINE_BITMAP_STEP = 8
#: A map's bitmap is sized up to a multiple of this to cover a new page
#: (measured on maps the engine grew: 3756 bytes for 30 048 pages, which
#: is a multiple of four and not of eight).
INLINE_BITMAP_ROUND = 4


def _rewrite_inline_row(store: PageStore, umap: UsageMap) -> None:
    """Write an inline map's row back, resizing the row when the bitmap
    grew (the other rows on the page shift, as for any row)."""
    from pyopenvba.access._datapage import DataPage

    row = bytes((USAGE_MAP_INLINE,)) + struct.pack("<I", umap.start_page) + bytes(umap.bitmap)
    page = DataPage(store.read(umap.page))
    page.replace_row(umap.row, row)
    store.write(umap.page, page.to_bytes())


def _reach(store: PageStore, umap: UsageMap, page: int) -> None:
    """Make an inline map able to hold ``page``, the way the engine does:
    an empty map is re-based to the page's 8-aligned start; a map that
    holds pages grows to cover two pages past the one being added,
    rounded up to four bytes.  The two spare pages are what a page just
    taken from the end of the file leaves ahead of it, and they show:
    taking page 542 grew a map to 72 bytes where covering 542 alone
    wanted 68.  Thirty growths across five scenarios, from 68 bytes to
    3756, agree on it."""
    if page < umap.start_page or (page - umap.start_page) // 8 >= len(umap.bitmap):
        if not any(umap.bitmap):
            umap.start_page = page & ~7
            return
        if page < umap.start_page:
            raise AccessError(
                f"page {page} lies below the start ({umap.start_page}) of the usage map at "
                f"({umap.page}, {umap.row}); the engine's answer to that has not been measured"
            )
        needed_bytes = (page + 2 - umap.start_page) // 8 + 1
        rounded = -(-needed_bytes // INLINE_BITMAP_ROUND) * INLINE_BITMAP_ROUND
        umap.bitmap.extend(bytes(max(0, rounded - len(umap.bitmap))))


def set_usage_bit(store: PageStore, umap: UsageMap, page: int, present: bool) -> None:
    """Set or clear ``page`` in the map and write the change through."""
    if umap.kind == USAGE_MAP_INLINE:
        if present:
            _reach(store, umap, page)
            if not _inline_row_fits(store, umap):
                _to_reference(store, umap)
                set_usage_bit(store, umap, page, True)
                return
        index = page - umap.start_page
        if index < 0 or index // 8 >= len(umap.bitmap):
            if not present:
                return  # clearing a page the map cannot hold changes nothing
            raise AccessError(f"page {page} is outside the usage map at ({umap.page}, {umap.row})")
        _flip(umap.bitmap, index, present)
        _rewrite_inline_row(store, umap)
        return
    if umap.kind == USAGE_MAP_REFERENCE:
        chunk, within = divmod(page, PAGES_PER_BITMAP_PAGE)
        if chunk >= len(umap.reference_pages) and present:
            raise AccessError(
                f"page {page} is beyond the reference usage map at ({umap.page}, {umap.row})"
            )
        if chunk >= len(umap.reference_pages):
            raise AccessError(
                f"page {page} is beyond the reference usage map at ({umap.page}, {umap.row})"
            )
        bitmap_page = umap.reference_pages[chunk]
        if bitmap_page == 0:
            if not present:
                return
            bitmap_page = _new_bitmap_page(store)
            umap.reference_pages[chunk] = bitmap_page
            raw_page = bytearray(store.read(umap.page))
            slots = row_slots(bytes(raw_page), store.layout)
            start, _end = row_span(slots, umap.row, store.layout)
            struct.pack_into("<I", raw_page, start + 1 + 4 * chunk, bitmap_page)
            store.write(umap.page, bytes(raw_page))
        raw_bitmap = bytearray(store.read(bitmap_page))
        byte_index = USAGE_BITMAP_PAGE_DATA + within // 8
        if present:
            raw_bitmap[byte_index] |= 1 << (within % 8)
        else:
            raw_bitmap[byte_index] &= ~(1 << (within % 8)) & 0xFF
        store.write(bitmap_page, bytes(raw_bitmap))
        _flip(umap.bitmap, page, present)
        return
    raise AccessError(f"usage map kind {umap.kind} cannot be edited")


#: A reference map's row: the kind byte and seventeen chunk pointers,
#: which reach further than a database is allowed to grow.
REFERENCE_CHUNKS = 17
REFERENCE_ROW_SIZE = 1 + 4 * REFERENCE_CHUNKS


def _inline_row_fits(store: PageStore, umap: UsageMap) -> bool:
    """Whether the inline row, at its current bitmap size, still fits the
    page it lives on."""
    from pyopenvba.access._datapage import DataPage

    page = DataPage(store.read(umap.page))
    start, end = page.span(umap.row)
    return 5 + len(umap.bitmap) - (end - start) <= page.free_space


def _to_reference(store: PageStore, umap: UsageMap, *, global_map: bool = False) -> None:
    """Turn an inline map into the reference form: the pages it holds move
    onto bitmap pages of their own and the row becomes a list of those
    pages.  The engine does this when growing the inline bitmap would push
    its row off the page (measured: a map row reached 3761 bytes and
    converted at the step that would have needed 3797, where 3796 was
    left)."""
    from pyopenvba.access._datapage import DataPage

    held = umap.pages()
    umap.kind = USAGE_MAP_REFERENCE
    umap.reference_pages = [0] * REFERENCE_CHUNKS
    umap.bitmap = bytearray()
    umap.start_page = 0
    page = DataPage(store.read(umap.page))
    page.replace_row(umap.row, bytes((USAGE_MAP_REFERENCE,)) + bytes(4 * REFERENCE_CHUNKS))
    store.write(umap.page, page.to_bytes())
    _refresh(store, umap)
    if global_map:
        # The global map lists free pages, so its first chunk takes over
        # what was free and everything the file has not reached yet.
        _global_chunk(store, umap)
    for number in held:
        set_usage_bit(store, umap, number, True)


def _global_chunk(store: PageStore, umap: UsageMap) -> None:
    """Give the global free map its next bitmap page and mark every page of
    that chunk which the file has not reached as free.  The bitmap page is
    the one just past the end of the file, which is where the engine put
    both of the ones measured (32 000 and 32 736)."""
    chunk = next((i for i, page in enumerate(umap.reference_pages) if not page), None)
    if chunk is None:
        raise AccessError("the global usage map has no room for another chunk")
    bitmap_page = store.page_count
    store.append()
    raw = bytearray(PAGE_SIZE)
    raw[0] = PAGE_USAGE_BITMAP
    raw[1] = 0x01
    base = chunk * PAGES_PER_BITMAP_PAGE
    for page in range(max(store.page_count, base), base + PAGES_PER_BITMAP_PAGE):
        index = page - base
        raw[USAGE_BITMAP_PAGE_DATA + index // 8] |= 1 << (index % 8)
    store.write(bitmap_page, bytes(raw))
    _write_reference_slot(store, umap, chunk, bitmap_page)
    _refresh(store, umap)


def _refresh(store: PageStore, umap: UsageMap) -> None:
    """Read the map back after its shape changed."""
    fresh = read_usage_map(store, umap.page, umap.row)
    umap.kind = fresh.kind
    umap.start_page = fresh.start_page
    umap.bitmap = fresh.bitmap
    umap.reference_pages = fresh.reference_pages


def _write_reference_slot(store: PageStore, umap: UsageMap, chunk: int, bitmap_page: int) -> None:
    raw_page = bytearray(store.read(umap.page))
    slots = row_slots(bytes(raw_page), store.layout)
    start, _end = row_span(slots, umap.row, store.layout)
    struct.pack_into("<I", raw_page, start + 1 + 4 * chunk, bitmap_page)
    store.write(umap.page, bytes(raw_page))


def _flip(bitmap: bytearray, index: int, present: bool) -> None:
    if index // 8 >= len(bitmap):
        return
    if present:
        bitmap[index // 8] |= 1 << (index % 8)
    else:
        bitmap[index // 8] &= ~(1 << (index % 8)) & 0xFF


def _new_bitmap_page(store: PageStore) -> int:
    page = allocate_page(store)
    raw = bytearray(PAGE_SIZE)
    raw[0] = PAGE_USAGE_BITMAP
    raw[1] = 0x01
    store.write(page, bytes(raw))
    return page


def allocate_page(store: PageStore) -> int:
    """Take the lowest free page from the global map, growing the file if
    that page lies past its end, and mark it used.  Pages released during
    this session (``store.released``) are passed over, as the engine
    passes them over until the database is reopened.  When the map lists
    no usable free page it is extended by one 8-byte step, the 64 new
    pages counting as free, which is how the engine grows it."""
    free = read_usage_map(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    held = store.released | set(store.pending)
    candidates = [p for p in free.pages() if p not in held]
    if not candidates:
        if free.kind == USAGE_MAP_INLINE:
            grown = UsageMap(free.page, free.row, free.kind, free.start_page, free.bitmap + b"\xff" * INLINE_BITMAP_STEP, [])
            if _inline_row_fits(store, grown):
                free.bitmap.extend(b"\xff" * INLINE_BITMAP_STEP)
                _rewrite_inline_row(store, free)
            else:
                _to_reference(store, free, global_map=True)
        else:
            _global_chunk(store, free)
        candidates = [p for p in free.pages() if p not in held]
    page = candidates[0]
    while store.page_count <= page:
        store.append()
    set_usage_bit(store, free, page, False)
    store.allocated.add(page)
    return page


PENDING_FLUSH = 5


def release_page(store: PageStore, page: int, *, kind: str = "object") -> None:
    """Return ``page`` to the global free map.  When the session may take it
    again depends on what gave it back (all measured with DAO):

    * ``"object"`` -- DROP TABLE, truncation, the retirement of an emptied
      page, a dropped index or column: never in this session, however
      many pile up (seven tables dropped and recreated in one session all
      took fresh pages).
    * ``"value"`` -- a freed long-value chain: at once when the page
      predates the session (the next 10 KB value, and a new table's
      pages, took them in order), else it waits like a rewrite's page.
    * ``"rewrite"`` -- the continuation page a definition rewrite replaced:
      it waits with the other pending pages, and they all come back into
      use once five are waiting (a run of CREATE INDEX statements reused
      its released continuation pages in a batch, lowest first).
    """
    free = read_usage_map(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    set_usage_bit(store, free, page, True)
    if kind == "object":
        store.released.add(page)
        return
    if kind == "value" and page not in store.allocated:
        store.released.discard(page)
        return
    store.pending.append(page)
    if len(store.pending) >= PENDING_FLUSH:
        store.pending.clear()


def add_to_map(store: PageStore, reference: int, page: int) -> None:
    set_usage_bit(store, read_usage_map_ref(store, reference), page, True)


def remove_from_map(store: PageStore, reference: int, page: int) -> None:
    set_usage_bit(store, read_usage_map_ref(store, reference), page, False)
