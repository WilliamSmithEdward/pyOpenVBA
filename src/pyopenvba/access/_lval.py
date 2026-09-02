"""Long values: Memo and OLE data that does not live in the row.

The row holds a 12-byte definition (see :class:`LongValueRef`).  Its kind
byte says where the bytes are:

    0x80  right after the definition, in the row itself
    0x40  one row on an LVAL page, the whole value
    0x00  a chain of LVAL rows; each begins with a 4-byte pointer to the
          next (row byte, three-byte page), (0, 0) on the last
"""

from __future__ import annotations

from pyopenvba.access_read import AccessError
from pyopenvba.access._pages import PageStore, is_lval_page, row_bytes, row_pointer
from pyopenvba.access._rows import LongValueRef


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
        seen: set[tuple[int, int]] = set()
        page, row = ref.page, ref.row
        while page:
            if (page, row) in seen:
                raise AccessError(f"long value chain loops at ({page}, {row})")
            seen.add((page, row))
            chunk = _lval_row(store, page, row)
            if len(chunk) < 4:
                raise AccessError(
                    f"long value chunk at ({page}, {row}) is too short for its pointer"
                )
            row, page = row_pointer(chunk)
            out.extend(chunk[4:])
        if len(out) != ref.length:
            raise AccessError(
                f"long value chain holds {len(out)} bytes, its definition says {ref.length}"
            )
        return bytes(out)
    raise AccessError(f"unknown long-value kind {ref.kind:#04x}")


def _lval_row(store: PageStore, page: int, row: int) -> bytes:
    raw_page = store.read(page)
    if not is_lval_page(raw_page):
        raise AccessError(f"page {page} is not a long-value page")
    data = row_bytes(raw_page, row)
    if data is None:
        raise AccessError(f"long value row ({page}, {row}) is deleted")
    return data
