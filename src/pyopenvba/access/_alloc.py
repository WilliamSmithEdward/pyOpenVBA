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


def set_usage_bit(store: PageStore, umap: UsageMap, page: int, present: bool) -> None:
    """Set or clear ``page`` in the map and write the change through."""
    if umap.kind == USAGE_MAP_INLINE:
        index = page - umap.start_page
        if index < 0 or index // 8 >= len(umap.bitmap):
            raise AccessError(
                f"page {page} is outside the inline usage map at ({umap.page}, {umap.row}), "
                f"which covers {umap.start_page}..{umap.start_page + 8 * len(umap.bitmap) - 1}"
            )
        _flip(umap.bitmap, index, present)
        raw_page = bytearray(store.read(umap.page))
        slots = row_slots(bytes(raw_page))
        start, end = row_span(slots, umap.row)
        if end - start != 5 + len(umap.bitmap):
            raise AccessError(f"usage map row ({umap.page}, {umap.row}) changed size")
        raw_page[start + 5 : end] = umap.bitmap
        store.write(umap.page, bytes(raw_page))
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
    that page lies past its end, and mark it used."""
    free = read_usage_map(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    candidates = free.pages()
    if not candidates:
        raise AccessError(
            "the global usage map lists no free page; growing the map itself is not supported yet"
        )
    page = candidates[0]
    while store.page_count <= page:
        store.append()
    set_usage_bit(store, free, page, False)
    return page


def release_page(store: PageStore, page: int) -> None:
    free = read_usage_map(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    set_usage_bit(store, free, page, True)


def add_to_map(store: PageStore, reference: int, page: int) -> None:
    set_usage_bit(store, read_usage_map_ref(store, reference), page, True)


def remove_from_map(store: PageStore, reference: int, page: int) -> None:
    set_usage_bit(store, read_usage_map_ref(store, reference), page, False)
