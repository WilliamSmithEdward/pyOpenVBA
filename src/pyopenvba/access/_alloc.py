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


INLINE_BITMAP_STEP = 8


def _rewrite_inline_row(store: PageStore, umap: UsageMap) -> None:
    """Write an inline map's row back, resizing the row when the bitmap
    grew (the other rows on the page shift, as for any row)."""
    from pyopenvba.access._datapage import DataPage

    row = bytes((USAGE_MAP_INLINE,)) + struct.pack("<I", umap.start_page) + bytes(umap.bitmap)
    page = DataPage(store.read(umap.page))
    page.replace_row(umap.row, row)
    store.write(umap.page, page.to_bytes())


def _reach(umap: UsageMap, page: int) -> None:
    """Make an inline map able to hold ``page``, the way the engine does:
    an empty map is re-based to the page's 8-aligned start; a map with
    pages grows its bitmap in 8-byte steps to the least size covering the
    page."""
    if page < umap.start_page or (page - umap.start_page) // 8 >= len(umap.bitmap):
        if not any(umap.bitmap):
            umap.start_page = page & ~7
            return
        if page < umap.start_page:
            raise AccessError(
                f"page {page} lies below the start ({umap.start_page}) of the usage map at "
                f"({umap.page}, {umap.row}); the engine's answer to that has not been measured"
            )
        needed_bytes = (page - umap.start_page) // 8 + 1
        rounded = -(-needed_bytes // INLINE_BITMAP_STEP) * INLINE_BITMAP_STEP
        umap.bitmap.extend(bytes(rounded - len(umap.bitmap)))


def set_usage_bit(store: PageStore, umap: UsageMap, page: int, present: bool) -> None:
    """Set or clear ``page`` in the map and write the change through."""
    if umap.kind == USAGE_MAP_INLINE:
        if present:
            _reach(umap, page)
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
            slots = row_slots(bytes(raw_page))
            start, _end = row_span(slots, umap.row)
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
        if free.kind != USAGE_MAP_INLINE:
            raise AccessError("the global usage map lists no free page and is not inline; not supported yet")
        free.bitmap.extend(b"\xff" * INLINE_BITMAP_STEP)
        _rewrite_inline_row(store, free)
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
