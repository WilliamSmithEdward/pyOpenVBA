"""In-place editing of a data-shaped page: the slot table and the rows.

Measured on pages the ACE engine wrote (docs/access_engine.md):

* Rows are laid down from the end of the page toward the slot table, in
  slot order, with no gaps: row k occupies [offset(k), offset(k-1)).
* Free space (u16 at 2) is exactly 4096 - 14 - 2 * slots - bytes below
  the lowest row.
* Deleting a row shifts every row below it up over the hole, updates
  their offsets, and leaves the slot in place flagged 0xC000 with its
  offset at the boundary it now sits on (zero length).  The bytes freed
  at the bottom are not cleared.
* Replacing a row shifts the rows below it by the size difference.
* Inserting appends a slot and places the row below the lowest one.
"""

from __future__ import annotations

import struct

from pyopenvba.access_read import AccessError
from pyopenvba.access._pages import (
    OFFSET_PAGE_FREE_SPACE,
    OFFSET_PAGE_OWNER,
    OFFSET_PAGE_ROW_COUNT,
    OFFSET_PAGE_ROW_TABLE,
    PAGE_DATA,
    PAGE_SIZE,
    ROW_DELETED,
    ROW_OFFSET_MASK,
    ROW_OVERFLOW,
)

INITIAL_FREE_SPACE = PAGE_SIZE - OFFSET_PAGE_ROW_TABLE
DEAD_SLOT = ROW_DELETED | ROW_OVERFLOW


class DataPage:
    """A 4 KiB data page held as a mutable byte array."""

    def __init__(self, raw: bytes) -> None:
        if len(raw) != PAGE_SIZE:
            raise AccessError(f"a page is {PAGE_SIZE} bytes, got {len(raw)}")
        if raw[0] != PAGE_DATA:
            raise AccessError(f"page type {raw[0]:#04x} is not a data page")
        self.raw = bytearray(raw)

    @classmethod
    def new(cls, owner: int) -> DataPage:
        raw = bytearray(PAGE_SIZE)
        raw[0] = PAGE_DATA
        raw[1] = 0x01
        struct.pack_into("<H", raw, OFFSET_PAGE_FREE_SPACE, INITIAL_FREE_SPACE)
        struct.pack_into("<I", raw, OFFSET_PAGE_OWNER, owner)
        return cls(bytes(raw))

    # -- reading -----------------------------------------------------------

    @property
    def owner(self) -> int:
        return struct.unpack_from("<I", self.raw, OFFSET_PAGE_OWNER)[0]

    @property
    def free_space(self) -> int:
        return struct.unpack_from("<H", self.raw, OFFSET_PAGE_FREE_SPACE)[0]

    @property
    def slot_count(self) -> int:
        return struct.unpack_from("<H", self.raw, OFFSET_PAGE_ROW_COUNT)[0]

    @property
    def slots(self) -> list[int]:
        count = self.slot_count
        return list(struct.unpack_from(f"<{count}H", self.raw, OFFSET_PAGE_ROW_TABLE))

    def span(self, slot: int) -> tuple[int, int]:
        slots = self.slots
        start = slots[slot] & ROW_OFFSET_MASK
        end = PAGE_SIZE if slot == 0 else slots[slot - 1] & ROW_OFFSET_MASK
        return start, end

    def lowest_offset(self) -> int:
        slots = self.slots
        if not slots:
            return PAGE_SIZE
        return min(entry & ROW_OFFSET_MASK for entry in slots)

    def fits(self, row_length: int) -> bool:
        return self.free_space >= row_length + 2

    def row(self, slot: int) -> bytes | None:
        entry = self.slots[slot]
        if entry & ROW_DELETED:
            return None
        start, end = self.span(slot)
        if entry & ROW_OVERFLOW:
            return bytes(self.raw[start : start + 4])
        return bytes(self.raw[start:end])

    # -- writing -----------------------------------------------------------

    def _set_slot(self, slot: int, value: int) -> None:
        struct.pack_into("<H", self.raw, OFFSET_PAGE_ROW_TABLE + 2 * slot, value)

    def _set_free_space(self, value: int) -> None:
        if not 0 <= value <= INITIAL_FREE_SPACE:
            raise AccessError(f"free space {value} is impossible")
        struct.pack_into("<H", self.raw, OFFSET_PAGE_FREE_SPACE, value)

    def add_row(self, row: bytes) -> int:
        """Append a slot for ``row`` and return the slot number."""
        if not self.fits(len(row)):
            raise AccessError(
                f"row of {len(row)} bytes does not fit; {self.free_space} bytes free"
            )
        slot = self.slot_count
        start = self.lowest_offset() - len(row)
        if start < OFFSET_PAGE_ROW_TABLE + 2 * (slot + 1):
            raise AccessError("row would overlap the slot table")
        self.raw[start : start + len(row)] = row
        struct.pack_into("<H", self.raw, OFFSET_PAGE_ROW_COUNT, slot + 1)
        self._set_slot(slot, start)
        self._set_free_space(self.free_space - len(row) - 2)
        return slot

    def _shift_below(self, boundary: int, delta: int) -> None:
        """Move every row that starts below ``boundary`` by ``delta`` bytes
        (positive moves toward the page end) and fix its slot offset."""
        lowest = self.lowest_offset()
        if lowest >= boundary:
            return
        block = bytes(self.raw[lowest:boundary])
        self.raw[lowest + delta : boundary + delta] = block
        for slot, entry in enumerate(self.slots):
            offset = entry & ROW_OFFSET_MASK
            if offset < boundary:
                self._set_slot(slot, (entry & ~ROW_OFFSET_MASK) | (offset + delta))

    def remove_row(self, slot: int) -> None:
        """Delete a row the way the engine does: close the hole and leave a
        dead slot at the boundary."""
        entry = self.slots[slot]
        if entry & ROW_DELETED:
            raise AccessError(f"slot {slot} is already dead")
        start, end = self.span(slot)
        length = end - start
        self._shift_below(start, length)
        self._set_slot(slot, DEAD_SLOT | end)
        self._set_free_space(self.free_space + length)

    def replace_row(self, slot: int, row: bytes) -> None:
        entry = self.slots[slot]
        if entry & DEAD_SLOT:
            raise AccessError(f"slot {slot} does not hold a plain row")
        start, end = self.span(slot)
        delta = len(row) - (end - start)
        if delta > self.free_space:
            raise AccessError(
                f"row grows by {delta} bytes but only {self.free_space} are free"
            )
        if delta:
            self._shift_below(start, -delta)
            start -= delta
            self._set_slot(slot, start)
        self.raw[start : start + len(row)] = row
        self._set_free_space(self.free_space - delta)

    def to_bytes(self) -> bytes:
        return bytes(self.raw)
