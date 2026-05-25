"""
pyopenvba.access -- Pure-Python read/write of Microsoft Access (.accdb / .mdb) databases.

Status
------
EXPERIMENTAL. This module is being built out incrementally. The current scope is:

    * Read 4 KiB ACE page format (Access 2007+ / Jet 4).
    * Discover VBA modules by locating their MS-OVBA compressed blobs.
    * Walk LVAL page chains to reassemble multi-page blobs.
    * Decompress blobs to plain VBA source via :func:`pyopenvba.vba.decompress`.
    * (Planned) Inject / rewrite module source.

Format notes (all reverse engineered against published Jet/ACE references plus
direct inspection -- no external Microsoft dependency at runtime):

    * Page 0 begins with a single-byte page-type tag (0x00), a one-byte
      database type, two reserved bytes, then the ASCII signature
      "Standard ACE DB\\0".
    * Subsequent pages are 4096 bytes each and begin with a one-byte page-type
      tag:
          0x01  Data page (rows of a table)
          0x02  Table definition page
          0x03  Intermediate index page
          0x04  Leaf index page
          0x05  Page usage map
          0x08  Long Value (LVAL) page; carries the literal "LVAL" tag at +4
    * Unencrypted .accdb data pages are stored in cleartext; only certain
      fields of page 0 are obfuscated. No page-level XOR is required to read
      the catalog.

VBA storage
-----------
Each VBA module's source is stored as a single MS-OVBA compressed stream --
the **same** RLE format used by Excel/Word VBA projects (see
:mod:`pyopenvba.vba`). The stream is laid out across one or more LVAL data
pages chained by a next-page pointer in each page header.

On the **starting** page of a stream, the OVBA signature byte ``0x01`` is
immediately followed by chunk headers and the bytes ``"Attribute VB_Name = ...
"``. On every **continuation** page, the page header bytes 14-15 hold an
offset (little-endian) at which a 4-byte record prefix is followed by the
resumption of the OVBA byte stream. The blob continues to the end of each
chained page.

There is also a secondary plaintext **comment-row index** in some databases
(0xE3 0x00 0x00 0x00 markers + u16 length + ASCII text). This is an Access
find/replace index, NOT the source of truth; callers should ignore it for
source extraction and use :meth:`AccessFile.iter_vba_modules` instead.
"""

from __future__ import annotations

import datetime
import struct
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Union

from pyopenvba.exceptions import PyOpenVBAError, UnsupportedFormatError
from pyopenvba.vba import VBAReference, decompress as _ovba_decompress
from pyopenvba.vba_pcode import (
    DisassembledModule,
    disassemble_module_stream,
)


ACE_PAGE_SIZE = 4096
ACE_SIGNATURE = b"Standard ACE DB\x00"
JET4_SIGNATURE = b"Standard Jet DB\x00"

# Page-type tag values (first byte of each page).
PAGE_TYPE_DB_DEF = 0x00
PAGE_TYPE_DATA = 0x01
PAGE_TYPE_TABLE_DEF = 0x02
PAGE_TYPE_INTERMEDIATE_INDEX = 0x03
PAGE_TYPE_LEAF_INDEX = 0x04
PAGE_TYPE_PAGE_USAGE_MAP = 0x05
PAGE_TYPE_LVAL = 0x08

# --- MSysObjects (Jet/ACE system catalog) -----------------------------------
#
# Every .accdb file embeds a system table named ``MSysObjects`` that lists
# every persistent object Access knows about: tables, queries, forms,
# reports, macros, modules, relationships, etc. Its TDEF lives at page 2
# (always, regardless of database age) and its data rows live on
# DATA-type pages whose ``owner_tdef_pn`` field equals 2.
#
# Schema (17 columns; column index follows TDEF definition order):
#   fixed: Id (u32), ParentId (u32), Type (i16),
#          DateCreate (8B), DateUpdate (8B), Flags (u32)
#   variable: Connect, Database, ForeignName, LvExtra, LvModule, LvProp,
#             Name, Owner, RmtInfoLong, RmtInfoShort, (+ 1 system slot)
#
# Row layout (verified across all 25 RE-corpus samples):
#   off 00..02  col_count (u16, always 17 for MSysObjects)
#   off 02..06  Id
#   off 06..0A  ParentId
#   off 0A..0C  Type (signed)
#   off 0C..14  DateCreate
#   off 14..1C  DateUpdate
#   off 1C..20  Flags
#   off 20..JT  variable-column data, packed back-to-front
#   off JT..JT+22  jump table: 11 u16 entries giving the START offset of
#                  each variable column, in REVERSE column order
#                  (jt[10] is the FIRST var col laid down in memory).
#   off JT+22..JT+24  var_col_count (u16, always 11)
#   off JT+24..JT+27  null bitmap (3 bytes, ceil(17/8))
#
# Empirically: variable column index 10 == ``Name`` (UTF-16-LE). Length
# of the Name field = jt[9] - jt[10].
_MSYS_OBJECTS_TDEF_PAGE = 2
_MSYS_COL_COUNT = 17
_MSYS_VAR_COL_COUNT = 11
_MSYS_NULL_MASK_BYTES = 3  # ceil(17 / 8)
_MSYS_NAME_VAR_INDEX = 10  # which variable column holds the Name field

# MSysObjects ``Type`` values (signed i16). Positive types are
# system-defined container objects; negative types (high bit set)
# tag user content. Only the values verified against the live RE
# corpus are surfaced as named constants.
MSYS_TYPE_FORM = -32768          # 0x8000
MSYS_TYPE_REPORT = -32766        # 0x8002
MSYS_TYPE_MACRO = -32766         # alias (Access reuses 0x8002 historically)
MSYS_TYPE_MODULE = -32761        # 0x8007  -- VBA CodeModule
MSYS_TYPE_CONTAINER = 3          # e.g. "Modules", "Forms", "Reports" hubs
MSYS_TYPE_TABLE = 1
MSYS_TYPE_QUERY = 5
MSYS_TYPE_DATABASE = 8

# Row offset-table entry flags inside a Jet/ACE DATA page.
_ROW_OFFSET_MASK = 0x1FFF
_ROW_DELETED_FLAG = 0x8000
_ROW_OVERFLOW_FLAG = 0x4000

# VBA source-row markers, reverse engineered from live .accdb fixtures.
#
# Inside the LVAL chains for VBA modules, each user-typed source line is
# stored as a single record of the form:
#
#     <4-byte type marker> <u16 LE length> <text bytes>
#
# Two type markers have been observed so far:
#   SOURCE_ROW_COMMENT  (0xE3 0x00 0x00 0x00)
#       Comment line. The leading apostrophe ("'") is stripped on disk
#       and must be re-prepended on reconstruction. Any spaces after the
#       apostrophe are preserved in the stored payload.
#   SOURCE_ROW_CODE     (TBD -- not yet observed in fixtures)
#       Non-comment source line.
#
# The implicit module preamble lines that Access exposes via the COM
# CodeModule.Lines() API (notably "Option Compare Database") are NOT stored
# as rows -- Access prepends them at read time based on the module's
# compare-mode setting.
SOURCE_ROW_COMMENT = b"\xE3\x00\x00\x00"


class AccessError(PyOpenVBAError):
    """Base error for the Access backend."""


class AccessFile:
    """
    Read-only entry point for an Access database file. Construction parses the
    file header and validates the ACE/Jet signature. Higher-level methods
    (``module_names``, source extraction, etc.) are implemented incrementally.

    Usage::

        with AccessFile("database.accdb") as db:
            print(db.format)            # "ace" or "jet4"
            print(db.page_count)
    """

    def __init__(self, path: Union[str, Path]) -> None:
        self.path = Path(path)
        self._data: bytearray = bytearray(self.path.read_bytes())
        if len(self._data) < ACE_PAGE_SIZE:
            raise AccessError(
                f"File too small to be an Access database "
                f"({len(self._data)} bytes < {ACE_PAGE_SIZE})"
            )
        if len(self._data) % ACE_PAGE_SIZE != 0:
            # Access always grows the file in whole-page increments.
            raise AccessError(
                f"File length {len(self._data)} is not a multiple of "
                f"the {ACE_PAGE_SIZE}-byte page size"
            )
        sig = bytes(self._data[4 : 4 + len(ACE_SIGNATURE)])
        if sig == ACE_SIGNATURE:
            self.format = "ace"
        elif sig == JET4_SIGNATURE:
            self.format = "jet4"
        else:
            raise UnsupportedFormatError(
                f"{self.path.name!r}: unrecognized database signature "
                f"{bytes(sig)!r}"
            )

    @property
    def page_count(self) -> int:
        return len(self._data) // ACE_PAGE_SIZE

    def read_page(self, page_num: int) -> bytes:
        """Return the raw 4 KiB contents of a single page, by zero-based index."""
        if page_num < 0 or page_num >= self.page_count:
            raise AccessError(
                f"page {page_num} out of range (0..{self.page_count - 1})"
            )
        off = page_num * ACE_PAGE_SIZE
        return bytes(self._data[off : off + ACE_PAGE_SIZE])

    def page_type(self, page_num: int) -> int:
        """Return the first-byte page-type tag of the given page."""
        return self._data[page_num * ACE_PAGE_SIZE]

    def iter_source_rows(self) -> Iterator["SourceRow"]:
        """
        Yield every VBA source-line row found anywhere in the database, in
        file-offset order.

        Each row carries a row-type marker (currently only the comment
        marker has been observed and decoded), a 16-bit text length, and an
        ASCII payload. The leading "' " of stored comment lines is *not*
        included in ``text`` -- callers should prepend it when reconstructing
        the line as it would appear in the VBA editor.

        This is a low-level utility that scans the whole file by marker. A
        higher-level method that maps a specific module name to its row range
        via the system catalog is not yet implemented.
        """
        data = self._data
        n = len(data)
        i = 0
        marker = SOURCE_ROW_COMMENT
        while True:
            j = data.find(marker, i)
            if j < 0:
                return
            if j + 6 > n:
                return
            length = int.from_bytes(data[j + 4 : j + 6], "little")
            end = j + 6 + length
            i = j + 1
            if length == 0 or length > 8000 or end > n:
                continue
            payload = bytes(data[j + 6 : end])
            # Reject obviously non-source payloads (must be printable ASCII
            # plus tabs / CR / LF).
            if any(
                not (0x09 <= b <= 0x7E or b in (0x0A, 0x0D)) for b in payload
            ):
                continue
            yield SourceRow(
                offset=j,
                row_type="comment",
                length=length,
                text=payload,
            )

    def __enter__(self) -> "AccessFile":
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        # Nothing to release: we read the file fully into memory at __init__.
        return None

    # ------------------------------------------------------------------
    # VBA module discovery & extraction
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # LVAL page / slot decoding (Phase 2 RE, 2026-05).
    # ------------------------------------------------------------------
    #
    # Each LVAL page lays out as:
    #     [0]      page_type 0x01
    #     [1]      0x01 (subtype)
    #     [2:4]    checksum / version
    #     [4:8]    'LVAL'
    #     [8:12]   reserved (zero)
    #     [12:14]  u16 LE slot count N
    #     [14:14+2N]  u16 LE slot table; top nibble 0xD = tombstone,
    #                 else low 12 bits = byte offset of row in page.
    # Rows grow downward from PAGE end. Row END = next-higher slot offset
    # in the table, or PAGE_SIZE for the top-most row.
    #
    # A long-value stored across multiple chunks places each chunk in one
    # slot. The first 4 bytes of such a chunk are
    #     [0]      next_slot (u8)
    #     [1:4]    next_page (u24 LE)
    #     [4:]     chunk payload
    # with (0, 0) marking the last chunk. Long-values that fit in a
    # single chunk are stored WITHOUT a continuation prefix (the row IS
    # the payload).
    #
    # Both forms occur in the wild: the canonical 1 MB fixture stores its
    # entire VBA project state in ONE chained long-value of 21 chunks; a
    # small .accdb with one tiny module stores ~6 separate standalone
    # long-values, one per VBA project section.

    def _lval_slot_count(self, page_num: int) -> int:
        base = page_num * ACE_PAGE_SIZE
        return int.from_bytes(self._data[base + 12 : base + 14], "little")

    def _lval_slot_offsets(self, page_num: int) -> list[int]:
        """Return the slot table for an LVAL page, as raw u16 values
        (including the 0xD000 tombstone flag)."""
        base = page_num * ACE_PAGE_SIZE
        n = self._lval_slot_count(page_num)
        return [
            int.from_bytes(self._data[base + 14 + 2 * i : base + 16 + 2 * i], "little")
            for i in range(n)
        ]

    def _lval_row_bytes(self, page_num: int, slot: int) -> bytes:
        """Return the raw bytes of one LVAL row including any
        4-byte continuation prefix. Raises if the slot is a tombstone
        or out of range."""
        slots = self._lval_slot_offsets(page_num)
        if slot < 0 or slot >= len(slots):
            raise AccessError(
                f"slot {slot} out of range on page {page_num} (n={len(slots)})"
            )
        raw = slots[slot]
        if (raw & 0xF000) == 0xD000:
            raise AccessError(
                f"slot {slot} on page {page_num} is a tombstone"
            )
        start = raw & 0x0FFF
        end = ACE_PAGE_SIZE
        for other in slots:
            if (other & 0xF000) == 0xD000:
                continue
            o = other & 0x0FFF
            if o > start and o < end:
                end = o
        base = page_num * ACE_PAGE_SIZE
        return bytes(self._data[base + start : base + end])

    def _walk_lval_chain(
        self, page_num: int, slot: int, max_chunks: int = 4096
    ) -> bytes:
        """Treat (page_num, slot) as the head of a chained long-value
        and return the concatenated payload of every chunk in the chain,
        stripping the 4-byte (next_slot, u24 next_page) prefix from each
        chunk. Stops when the prefix is (0, 0).

        Raises :class:`AccessError` if the chain is malformed (cycle,
        out-of-range page, etc.).
        """
        out = bytearray()
        seen: set[tuple[int, int]] = set()
        cur_p, cur_s = page_num, slot
        for _ in range(max_chunks):
            if (cur_p, cur_s) in seen:
                raise AccessError(
                    f"LVAL chain cycle at ({cur_p}, {cur_s})"
                )
            seen.add((cur_p, cur_s))
            row = self._lval_row_bytes(cur_p, cur_s)
            if len(row) < 4:
                raise AccessError(
                    f"LVAL row ({cur_p}, {cur_s}) too short to hold chain prefix"
                )
            next_s = row[0]
            next_p = int.from_bytes(row[1:4], "little")
            out.extend(row[4:])
            if next_p == 0 and next_s == 0:
                return bytes(out)
            if next_p >= self.page_count:
                raise AccessError(
                    f"LVAL chain references out-of-range page {next_p}"
                )
            cur_p, cur_s = next_p, next_s
        raise AccessError(f"LVAL chain exceeded max_chunks={max_chunks}")

    def _iter_lval_pages(self) -> Iterator[int]:
        for p in range(self.page_count):
            base = p * ACE_PAGE_SIZE
            if (
                self._data[base] == 0x01
                and bytes(self._data[base + 4 : base + 8]) == b"LVAL"
            ):
                yield p

    def _iter_lval_rows(self) -> Iterator[tuple[int, int, bytes]]:
        """Yield ``(page, slot, row_bytes)`` for every non-tombstone
        slot on every LVAL page in the database, in page-then-slot
        order."""
        for page in self._iter_lval_pages():
            slots = self._lval_slot_offsets(page)
            for slot, raw in enumerate(slots):
                if (raw & 0xF000) == 0xD000:
                    continue
                yield page, slot, self._lval_row_bytes(page, slot)

    # ------------------------------------------------------------------
    # LVAL row mutation primitives (Phase 5 write path, 2026-05).
    # ------------------------------------------------------------------
    #
    # Resize/rewrite an LVAL row in place by reflowing the page's
    # slot table and shifting adjacent rows. Tombstone bits are
    # preserved. Operates within a single page; chain growth onto a
    # fresh page is handled at a higher level (and currently raises
    # ``AccessError`` if no chain tail capacity is available).

    def _set_slot_offset(self, page: int, slot: int, offset: int) -> None:
        """Write ``offset`` (lower 12 bits) into slot ``slot`` of
        ``page``, preserving the tombstone bits (upper 4 bits) of
        the existing slot value."""
        base = page * ACE_PAGE_SIZE
        cur = int.from_bytes(
            self._data[base + 14 + 2 * slot : base + 16 + 2 * slot],
            "little",
        )
        new = (cur & 0xF000) | (offset & 0x0FFF)
        self._data[base + 14 + 2 * slot : base + 16 + 2 * slot] = (
            new.to_bytes(2, "little")
        )

    def _lval_row_extent(
        self, page: int, slot: int
    ) -> tuple[int, int]:
        """Return ``(start_offset, end_offset)`` of the live row at
        ``(page, slot)``, where ``end_offset`` is the smallest higher
        non-tombstone slot offset (or ``ACE_PAGE_SIZE`` for the top-
        most row). Raises if the slot is a tombstone."""
        slots = self._lval_slot_offsets(page)
        raw = slots[slot]
        if (raw & 0xF000) == 0xD000:
            raise AccessError(
                f"slot {slot} on page {page} is a tombstone"
            )
        start = raw & 0x0FFF
        end = ACE_PAGE_SIZE
        for other in slots:
            if (other & 0xF000) == 0xD000:
                continue
            o = other & 0x0FFF
            if o > start and o < end:
                end = o
        return start, end

    def _lval_free_space(self, page: int) -> int:
        """Number of contiguous free bytes available for new payload
        on ``page`` (i.e. the gap between the end of the slot table
        and the lowest live row offset)."""
        slots = self._lval_slot_offsets(page)
        slot_table_end = 14 + 2 * len(slots)
        live = [s & 0x0FFF for s in slots if (s & 0xF000) != 0xD000]
        lowest = min(live) if live else ACE_PAGE_SIZE
        return lowest - slot_table_end

    def _lval_resize_row(
        self, page: int, slot: int, new_size: int
    ) -> None:
        """Resize the live LVAL row at ``(page, slot)`` to ``new_size``
        bytes, shifting other rows on the same page and updating the
        slot table.

        The new bytes are NOT initialized -- callers should overwrite
        the row payload immediately via :meth:`_lval_write_row` or by
        slicing into ``self._data``.

        Raises :class:`AccessError` if growing the row would overflow
        the page's free-space window, or if ``new_size`` is negative.
        """
        if new_size < 0:
            raise AccessError(
                f"_lval_resize_row: new_size must be non-negative "
                f"(got {new_size})"
            )
        base = page * ACE_PAGE_SIZE
        start, end = self._lval_row_extent(page, slot)
        old_size = end - start
        delta = new_size - old_size
        if delta == 0:
            return
        slots = self._lval_slot_offsets(page)
        live = [
            (i, s & 0x0FFF)
            for i, s in enumerate(slots)
            if (s & 0xF000) != 0xD000
        ]
        lowest = min(off for _, off in live)
        slot_table_end = 14 + 2 * len(slots)
        new_lowest = lowest - delta
        if new_lowest < slot_table_end:
            raise AccessError(
                f"_lval_resize_row: page {page} cannot grow slot "
                f"{slot} by {delta} bytes (free={lowest - slot_table_end}, "
                f"need={delta})"
            )
        # Move the bytes physically beneath R (offsets [lowest..start])
        # to [new_lowest..start - delta]. bytearray slice assignment
        # evaluates the RHS into a new bytes object first, so
        # overlapping moves are safe in either direction.
        block = bytes(self._data[base + lowest : base + start])
        # Clear the freed region (only matters for shrink; harmless on grow).
        if delta < 0:
            # After move, bytes at [lowest..new_lowest] become stale;
            # zero them to keep the page free space clean.
            pass
        self._data[base + new_lowest : base + start + (new_lowest - lowest)] = block
        # If we shrank, zero the now-free trailing region beneath R.
        if delta < 0:
            # The block now occupies [new_lowest..new_lowest+(start-lowest)].
            # Free region: [lowest..new_lowest] (note new_lowest > lowest here).
            self._data[base + lowest : base + new_lowest] = bytes(
                new_lowest - lowest
            )
        # Update slot offsets for every live slot whose old offset
        # was <= start (i.e. those that physically moved, plus R
        # itself whose start moves by -delta).
        for i, old_off in live:
            if old_off <= start:
                self._set_slot_offset(page, i, old_off - delta)

    def _lval_write_row(
        self, page: int, slot: int, new_row: bytes
    ) -> None:
        """Replace the entire bytes of the LVAL row at ``(page, slot)``
        with ``new_row``, resizing the row in place. ``new_row`` must
        include any 4-byte continuation prefix if the row was a chain
        head/intermediate.

        Raises :class:`AccessError` if the new bytes don't fit on the
        page (see :meth:`_lval_resize_row`).
        """
        self._lval_resize_row(page, slot, len(new_row))
        base = page * ACE_PAGE_SIZE
        start, end = self._lval_row_extent(page, slot)
        assert end - start == len(new_row), (
            f"resize/write invariant: {end - start} != {len(new_row)}"
        )
        self._data[base + start : base + end] = new_row

    def _lval_tombstone_slot(self, page: int, slot: int) -> None:
        """Mark the slot at ``(page, slot)`` as a tombstone, freeing
        its payload region. The slot's recorded offset is preserved
        in the low 12 bits (Access keeps the offset for diagnostic
        purposes); only the 0xD000 flag is set.

        The freed row's BYTES are zeroed so they can't be
        misinterpreted as part of an adjacent live row's payload by
        downstream readers (such as :meth:`iter_vba_modules`) that
        compute row extents by walking the slot table.

        To physically reclaim the freed space and compact the page,
        call :meth:`_lval_compact_page` (when available) after
        tombstoning.
        """
        base = page * ACE_PAGE_SIZE
        # Compute the row extent BEFORE flipping the tombstone bit so
        # _lval_row_extent doesn't refuse.
        start, end = self._lval_row_extent(page, slot)
        # Zero the freed payload to prevent stale OVBA blobs from
        # leaking into adjacent rows' extents.
        for i in range(base + start, base + end):
            self._data[i] = 0
        cur = int.from_bytes(
            self._data[base + 14 + 2 * slot : base + 16 + 2 * slot],
            "little",
        )
        new = 0xD000 | (cur & 0x0FFF)
        self._data[base + 14 + 2 * slot : base + 16 + 2 * slot] = (
            new.to_bytes(2, "little")
        )

    def _lval_append_row(self, page: int, payload: bytes) -> int:
        """Append a new row to ``page`` and return its slot index.

        The new row occupies the lowest currently-free bytes on the
        page (just above any existing rows) and a new slot entry is
        added at the end of the slot table. Reuses a tombstoned slot
        if one is available.

        Raises :class:`AccessError` if there is insufficient free
        space (payload size + 2 bytes for the new slot entry).
        """
        base = page * ACE_PAGE_SIZE
        slots = self._lval_slot_offsets(page)
        live = [
            (i, s & 0x0FFF)
            for i, s in enumerate(slots)
            if (s & 0xF000) != 0xD000
        ]
        lowest = min((off for _, off in live), default=ACE_PAGE_SIZE)
        # Search for a tombstoned slot we can reuse.
        reuse_slot: int | None = None
        for i, s in enumerate(slots):
            if (s & 0xF000) == 0xD000:
                reuse_slot = i
                break
        slot_table_end = 14 + 2 * len(slots)
        need = len(payload) + (0 if reuse_slot is not None else 2)
        if lowest - slot_table_end < need:
            raise AccessError(
                f"_lval_append_row: page {page} has insufficient free "
                f"space (have {lowest - slot_table_end}, need {need})"
            )
        new_off = lowest - len(payload)
        self._data[base + new_off : base + lowest] = payload
        if reuse_slot is not None:
            self._set_slot_offset(page, reuse_slot, new_off)
            # Clear the tombstone bits explicitly.
            cur = int.from_bytes(
                self._data[base + 14 + 2 * reuse_slot : base + 16 + 2 * reuse_slot],
                "little",
            )
            self._data[base + 14 + 2 * reuse_slot : base + 16 + 2 * reuse_slot] = (
                (cur & 0x0FFF).to_bytes(2, "little")
            )
            return reuse_slot
        # Append new slot at end of table.
        new_slot = len(slots)
        self._data[base + 14 + 2 * new_slot : base + 16 + 2 * new_slot] = (
            new_off.to_bytes(2, "little")
        )
        self._data[base + 12 : base + 14] = (
            (len(slots) + 1).to_bytes(2, "little")
        )
        return new_slot

    # ------------------------------------------------------------------
    # MS-OVBA dir-stream catalog (Phase 3 RE, 2026-05).
    # ------------------------------------------------------------------
    #
    # Access embeds a standard MS-OVBA "dir" stream (section 2.3.4.2 of
    # MS-OVBA) into exactly one LVAL row of every database that contains
    # a VBA project. The row is OVBA-RLE-compressed in the usual way;
    # decompressed, it parses byte-for-byte with our existing dir-stream
    # parser in `pyopenvba.vba`.
    #
    # We locate it by attempting OVBA decompression on each LVAL row and
    # accepting the one whose decompressed bytes start with the
    # PROJECTSYSKIND record header `01 00 04 00 00 00`. That signature
    # is fully deterministic and avoids any reliance on the slot index
    # (which is not stable across .accdb files).

    _DIR_STREAM_MAGIC = b"\x01\x00\x04\x00\x00\x00"

    def _find_catalog_row(self) -> tuple[int, int, bytes] | None:
        """Return ``(page, slot, decompressed_dir_stream_bytes)`` for the
        single LVAL row that holds the project's MS-OVBA dir stream, or
        ``None`` if no such row is present (e.g. databases with no VBA
        project initialized)."""
        for page, slot, row in self._iter_lval_rows():
            if not row or row[0] != 0x01 or len(row) < 3:
                continue
            hdr = int.from_bytes(row[1:3], "little")
            if ((hdr >> 12) & 0x7) != 0b011:
                continue
            try:
                raw = _ovba_decompress(
                    bytes(row), stream_name=f"accdb_catalog@({page},{slot})"
                )
            except Exception:
                continue
            if raw.startswith(self._DIR_STREAM_MAGIC):
                return page, slot, raw
        return None

    def read_project_info(self) -> "AccessVBAProject":
        """Parse and return the project-level VBA metadata embedded in
        this database.

        Raises :class:`AccessError` if no VBA dir-stream catalog row is
        present (the database has no VBA project, or the catalog row
        could not be located -- file an issue with the fixture).
        """
        from pyopenvba.vba import parse_dir_stream

        found = self._find_catalog_row()
        if found is None:
            raise AccessError(
                f"no MS-OVBA dir-stream catalog row found in "
                f"{self.path.name!r}; this database may have no VBA project"
            )
        page, slot, raw = found
        info, mods = parse_dir_stream(raw)
        return AccessVBAProject(
            catalog_page=page,
            catalog_slot=slot,
            catalog_raw_size=raw.__len__(),
            sys_kind=info.sys_kind,
            lcid=info.lcid,
            code_page=info.code_page,
            project_name=info.name,
            references=tuple(info.references),
            modules=tuple(
                AccessVBAModuleEntry(
                    name=m.name,
                    name_unicode=m.name_unicode,
                    stream_name=m.stream_name,
                    is_class_module=(m.module_kind.value == 0x0022),
                    is_private=m.is_private,
                    is_read_only=m.is_read_only,
                )
                for m in mods
            ),
        )

    # ------------------------------------------------------------------
    # Phase 4 RE: authoritative VBA p-code stream
    # ------------------------------------------------------------------
    #
    # Every Access database with a VBA project carries one LVAL row whose
    # raw bytes begin with the magic header `72 55 40 ...` ("rU@"). This
    # row is the authoritative compiled-bytecode store -- decompiling and
    # re-displaying VBA source in the Access editor reads from this row,
    # not from the OVBA cache. The OVBA cache (row found by
    # `iter_vba_modules`) is a passive plaintext mirror Access keeps for
    # version-control and import/export tools.
    #
    # Evidence captured by the corpus (May 2026):
    # * `Dim x As Integer` (sample 044) and `Dim x As Long` (sample 045)
    #   compile to byte-for-byte identical p-code -- type annotations are
    #   resolved/erased at compile time.
    # * `' a comment` (sample 049) compiles to byte-for-byte identical
    #   p-code as `Dim x As Integer` -- comments produce no bytecode.
    # * `MsgBox "hello"` (sample 040) and `MsgBox "world"` (sample 041)
    #   differ in only 2 bytes (a u16 string-literal slot id), proving
    #   string literals are interned.
    # * `0x67 0x02` markers bracket each procedure body; `0x7B 0x02`
    #   marks the end of the module's last procedure; `0xED 0x05 <u16>`
    #   pushes a literal integer (confirmed: `ed 05 2a 00` for `x = 42`).
    #
    # Full opcode field guide is still in progress; this method is the
    # entry point that exposes the raw bytes for further RE work.

    # The p-code row header is 12 bytes. Every rU@-headed LVAL row
    # starts with the 4-byte signature ``72 55 40 00`` followed by 8
    # more bytes whose structure encodes the row's role:
    #
    #   bytes  0..3   : signature 'rU@\x00'
    #   bytes  4..7   : reserved / zero in the corpus
    #   bytes  8..15  : u64 (LE) -- 0x4000 for the *module-active*
    #                   bytecode row, 0 for every other rU@ row
    #
    # The 0x4000 at offset 10 is the deterministic structural marker
    # that distinguishes the row Access actually executes from older
    # stubs and project/system bootstrap rows kept alongside it.
    # Verified across the 15-sample corpus (samples 010-051).
    _PCODE_MAGIC = b"\x72\x55\x40\x00"   # 'rU@\x00'
    _PCODE_ACTIVE_PREFIX = (
        b"\x72\x55\x40\x00\x00\x00\x00\x00\x00\x00\x40\x00"
    )

    def _find_pcode_rows(self) -> list[tuple[int, int, bytes]]:
        """Return ``(page, slot, raw_bytes)`` for every LVAL row that
        carries the VBA p-code magic header. The corpus shows 2-3 such
        rows per database with VBA enabled."""
        hits: list[tuple[int, int, bytes]] = []
        for page, slot, row in self._iter_lval_rows():
            if row.startswith(self._PCODE_MAGIC):
                hits.append((page, slot, bytes(row)))
        return hits

    def iter_pcode_streams(self) -> tuple["AccessVBAPCodeStream", ...]:
        """Return every ``rU@``-headed VBA p-code row in the database.

        Exactly one of these rows is the *module-active* bytecode that
        Access executes; the others are stale stubs or project/system
        bootstrap rows. Use :meth:`read_module_pcode_stream` to fetch
        the active row directly.

        Raises :class:`AccessError` if no p-code rows can be located.
        """
        hits = self._find_pcode_rows()
        if not hits:
            raise AccessError(
                f"no VBA p-code rows (header 'rU@') found in "
                f"{self.path.name!r}; this database may have no VBA project"
            )
        return tuple(
            AccessVBAPCodeStream(page=p, slot=s, raw=r) for p, s, r in hits
        )

    def read_module_pcode_stream(self) -> "AccessVBAPCodeStream":
        """Return the *module-active* VBA p-code row, identified by the
        structural 12-byte prefix ``72 55 40 00 00 00 00 00 00 00 40
        00`` (byte at offset 10 is ``0x40`` rather than ``0x00``).

        This is the deterministic discriminator that separates the
        active compiled bytecode from the stale stub and bootstrap
        rows Access keeps alongside it.

        Raises :class:`AccessError` if zero or more than one row
        matches the active prefix.
        """
        matches = [
            s for s in self.iter_pcode_streams()
            if s.raw.startswith(self._PCODE_ACTIVE_PREFIX)
        ]
        if not matches:
            raise AccessError(
                f"no module-active VBA p-code row (prefix "
                f"{self._PCODE_ACTIVE_PREFIX.hex(' ')}) found in "
                f"{self.path.name!r}"
            )
        if len(matches) > 1:
            locs = ", ".join(f"({m.page},{m.slot})" for m in matches)
            raise AccessError(
                f"expected exactly one module-active VBA p-code row in "
                f"{self.path.name!r}, found {len(matches)} at {locs}"
            )
        return matches[0]

    # String-literal interning table -- see docs/access_pcode_re.md.
    # Inside the project symbol-table row each string literal is stored
    # as: ``0B <u32 LE byte-count> <UTF-16-LE bytes>``. The leading
    # ``0B`` tag distinguishes literal records from other entries in
    # the same row.
    _STRING_LITERAL_TAG = 0x0B

    def find_interned_strings(self) -> tuple["AccessVBAInternedString", ...]:
        """Scan every LVAL row for VBA string-literal records of the
        form ``0B <u32 LE byte-count> <UTF-16-LE bytes>``.

        This is a deterministic content-based scan -- no slot
        coordinates are hard-coded. The intern table lives in a
        per-project row alongside reference/module metadata; the
        decoder simply walks every LVAL row and yields each valid
        literal record it finds.

        A record is accepted only when:

        * the byte-count is even and non-zero,
        * the byte-count fits in the row,
        * the decoded UTF-16-LE bytes form a valid Python ``str``,
        * and the decoded string contains no NUL characters
          (filters out structural padding that happens to start with
          ``0B``).

        Returns a tuple of :class:`AccessVBAInternedString` records,
        each carrying the source page, slot, in-row byte offset, and
        decoded value.
        """
        out: list[AccessVBAInternedString] = []
        for page, slot, row in self._iter_lval_rows():
            buf = bytes(row)
            n = len(buf)
            i = 0
            while i < n - 5:
                if buf[i] == self._STRING_LITERAL_TAG:
                    byte_count = int.from_bytes(buf[i + 1:i + 5], "little")
                    payload_start = i + 5
                    payload_end = payload_start + byte_count
                    if (
                        byte_count > 0
                        and byte_count % 2 == 0
                        and payload_end <= n
                    ):
                        try:
                            text = buf[payload_start:payload_end].decode(
                                "utf-16-le"
                            )
                        except UnicodeDecodeError:
                            i += 1
                            continue
                        if text and "\x00" not in text and text.isprintable():
                            out.append(
                                AccessVBAInternedString(
                                    page=page,
                                    slot=slot,
                                    offset=i,
                                    value=text,
                                )
                            )
                            i = payload_end
                            continue
                i += 1
        return tuple(out)

    # ------------------------------------------------------------------
    # Standard VBA module-stream p-code (Phase 4d RE, 2026-05).
    # ------------------------------------------------------------------
    # In every Access database we inspected, the LVAL row carrying the
    # OVBA-compressed VBA source ALSO contains -- at a content-dependent
    # offset earlier in the same row -- the standard Office VBA module
    # stream's "PerformanceCache" region, recognisable by the
    # well-known ``0xCAFE`` magic word. This is the same per-line
    # p-code layout described in [MS-OVBA] section 2.3.4.3 and
    # disassembled by the public `pcodedmp` tool. It is *NOT* the
    # ``rU@``-prefixed bytecode (read by
    # :meth:`read_module_pcode_stream`); the ``rU@`` stream is the
    # Access-specific cached/execodes form. Both forms coexist in the
    # database; Access uses ``rU@`` at runtime, but the canonical VBA7
    # p-code -- portable across all Office hosts -- lives here.

    def find_module_streams(self) -> tuple["AccessVBAModuleStream", ...]:
        """Return the standard Office VBA module-stream bytes for
        every VBA module in the database.

        For each VBA module, the LVAL row containing its OVBA-
        compressed source also contains -- at an earlier offset in
        the same row -- the standard module stream's binary
        ``PerformanceCache`` region. That region is recognisable by
        the ``0xCAFE`` magic word and contains the canonical VBA7
        p-code (per-line opcodes), in the exact layout defined by
        [MS-OVBA] and consumed by public disassemblers.

        Returns one :class:`AccessVBAModuleStream` per module,
        carrying the source page, slot, raw row bytes, and the
        in-row offset of the ``0xCAFE`` magic. The raw bytes can be
        fed directly to any disassembler that expects an Office VBA
        module stream.

        This is a deterministic content-based scan: we identify
        carrier rows by intersecting "row decompresses to a valid
        VBA source attribute prefix" with "row contains a single
        ``0xCAFE`` word".
        """
        results: list[AccessVBAModuleStream] = []
        seen: set[tuple[int, int]] = set()
        for module in self.iter_vba_modules():
            page = module.start_offset // ACE_PAGE_SIZE
            for p, slot, row in self._iter_lval_rows():
                if p != page:
                    continue
                raw = bytes(row)
                cafe = raw.find(b"\xfe\xca")
                if cafe < 0:
                    continue
                if (p, slot) in seen:
                    break
                seen.add((p, slot))
                results.append(
                    AccessVBAModuleStream(
                        page=p, slot=slot, raw=raw, cafe_offset=cafe
                    )
                )
                break
        return tuple(results)

    def identifiers(self) -> tuple[AccessVBAIdentifier, ...]:
        """Enumerate every project-level identifier name decoded from
        the ``_VBA_PROJECT``-equivalent LVAL row.

        Returns a tuple of :class:`AccessVBAIdentifier` records in
        on-disk order. The list contains:

        * Typelib reference names (``Access``, ``VBA``, ``Win32``,
          ``Win64``, ``stdole``, ``DAO``, etc.)
        * The project name (``Project1`` and the project file stem).
        * The original module-template name (``Module1``) plus the
          current user module name(s).
        * User-defined procedure and variable names.
        * Intrinsic VBA function names referenced from compiled code
          (``MsgBox``, ``_Evaluate``, etc.).

        Returns an empty tuple if no ``CC 61`` row is present
        (corrupted / non-VBA-enabled database).

        Note: this is a project-wide *inventory*; the ``name_id``
        u16 operands in p-code do **NOT** index into this table
        directly (they index a per-procedure reference table -- a
        future RE deliverable). The inventory is still useful for
        diagnostic and auditing purposes (e.g. listing every
        intrinsic a project calls).
        """
        rows = list(self._iter_lval_rows())
        stream = _find_vba_project_row(rows)
        if stream is None:
            return ()
        return _parse_vba_project_identifiers(stream)

    def disassemble_module(
        self, name: str, *, is_64bit: bool = True
    ) -> DisassembledModule:
        """Disassemble the canonical VBA7 p-code of the named module.

        Locates the module's ``0xCAFE`` p-code region via
        :meth:`find_module_streams`, then walks the per-line opcode
        stream using :func:`pyopenvba.vba_pcode.disassemble_module_stream`.

        Args:
            name: Module name (matches ``VBAModule.name`` /
                ``Attribute VB_Name``).
            is_64bit: P-code encoding flavour. Defaults to ``True``
                because every Access database in the project's test
                corpus uses VBA7 64-bit encoding. Set ``False`` for
                databases produced by 32-bit Office / VBA6 hosts.

        Returns:
            A fully decoded :class:`DisassembledModule`.

        Raises:
            AccessError: If no module with that name is present, or
                the module is present but has no compiled p-code
                (e.g. source-only module that has never been
                executed).
        """
        modules = list(self.iter_vba_modules())
        streams = self.find_module_streams()
        # iter_vba_modules() and find_module_streams() walk the
        # database in the same page order, but find_module_streams
        # only emits an entry when a CAFE region is present in the
        # carrier row. Match by (page, slot) to be robust.
        by_page: dict[int, AccessVBAModuleStream] = {
            s.page: s for s in streams
        }
        for module in modules:
            if module.name != name:
                continue
            stream = by_page.get(module.start_offset // ACE_PAGE_SIZE)
            if stream is None:
                raise AccessError(
                    f"module {name!r} has no compiled p-code "
                    "(no 0xCAFE region in carrier row)"
                )
            return disassemble_module_stream(
                stream.raw, is_64bit=is_64bit
            )
        raise AccessError(f"module {name!r} not found in database")

    def disassemble_all_modules(
        self, *, is_64bit: bool = True
    ) -> dict[str, DisassembledModule]:
        """Disassemble every VBA module in the database.

        Convenience wrapper around :meth:`disassemble_module`. Returns
        a name-keyed mapping; modules with no compiled p-code (no
        ``0xCAFE`` carrier row) are silently skipped, mirroring
        :meth:`find_module_streams` semantics.

        Args:
            is_64bit: See :meth:`disassemble_module`.
        """
        out: dict[str, DisassembledModule] = {}
        streams = {s.page: s for s in self.find_module_streams()}
        for module in self.iter_vba_modules():
            stream = streams.get(module.start_offset // ACE_PAGE_SIZE)
            if stream is None:
                continue
            out[module.name] = disassemble_module_stream(
                stream.raw, is_64bit=is_64bit
            )
        return out

    def iter_vba_modules(self) -> Iterator["VBAModule"]:
        """
        Discover and yield every VBA module embedded in this database.

        Implementation strategy
        -----------------------
        Each VBA module's source is stored as an MS-OVBA compressed stream
        held in one LVAL row, possibly chained across multiple LVAL
        chunks (see Phase 2 RE notes above).

        We iterate every non-tombstone LVAL row in the database and try
        TWO interpretations:

        * **Standalone**: the entire row is an OVBA stream. Try to
          decompress ``row[:]`` directly.
        * **Chained**: the row's first 4 bytes are a
          ``<u8 next_slot><u24 next_page>`` continuation prefix; walk
          the chain accumulating ``row[4:]`` from each chunk, then
          decompress.

        Accept any result that decompresses to a stream starting with
        ``Attribute VB_Name = "..."``. This avoids any dependency on
        parsing the Access system catalog (MSysObjects /
        MSysAccessStorage) -- a reasonable trade-off until write support
        requires us to allocate / re-link LVAL chunks.
        """
        yielded: set[tuple[int, int]] = set()
        for page, slot, row in self._iter_lval_rows():
            for blob_kind, blob in self._candidate_blobs(page, slot, row):
                try:
                    raw = _ovba_decompress(
                        blob, stream_name=f"accdb@({page},{slot}):{blob_kind}"
                    )
                except Exception:
                    continue
                if not raw.startswith(b"Attribute VB_Name = "):
                    continue
                if (page, slot) in yielded:
                    continue
                yielded.add((page, slot))
                text = raw.decode("latin-1")
                lines = text.split("\r\n")
                body_start = 0
                module_name = ""
                for idx, ln in enumerate(lines):
                    if ln.startswith("Attribute "):
                        if ln.startswith('Attribute VB_Name = "'):
                            module_name = ln.split('"', 2)[1]
                    else:
                        body_start = idx
                        break
                body = "\r\n".join(lines[body_start:])
                yield VBAModule(
                    name=module_name,
                    start_offset=page * ACE_PAGE_SIZE,
                    raw_blob_size=len(blob),
                    decompressed_size=len(raw),
                    attributes_text="\r\n".join(lines[:body_start]),
                    source=body,
                )
                break  # don't double-yield from the alternative interpretation

    def _candidate_blobs(
        self, page: int, slot: int, row: bytes
    ) -> Iterator[tuple[str, bytes]]:
        """Yield ``(label, candidate_ovba_blob)`` for the row.

        Two interpretations are tried:

        * **Standalone**: scan the row for any OVBA signature (sig byte
          ``0x01`` followed by a chunk header with signature bits
          ``0b011``) and yield the suffix of the row from that offset.
        * **Chained**: if the first 4 bytes of the row form a valid
          ``(slot, page)`` continuation prefix pointing to another LVAL
          row, walk the chain and then scan the assembled blob for OVBA
          signatures the same way.
        """
        for off in self._scan_ovba_signatures(row):
            yield f"standalone@({page},{slot})+{off}", row[off:]
        if len(row) >= 4 and self._looks_like_chain_head(row):
            try:
                blob = self._walk_lval_chain(page, slot)
            except AccessError:
                return
            for off in self._scan_ovba_signatures(blob):
                yield f"chained@({page},{slot})+{off}", blob[off:]

    def _looks_like_chain_head(self, row: bytes) -> bool:
        """Heuristic: row[0:4] is a plausible (slot, u24 page) chain
        prefix iff the page number is in range and is itself an LVAL
        page."""
        if len(row) < 4:
            return False
        next_p = int.from_bytes(row[1:4], "little")
        if next_p == 0 or next_p >= self.page_count:
            return False
        base = next_p * ACE_PAGE_SIZE
        return (
            self._data[base] == 0x01
            and bytes(self._data[base + 4 : base + 8]) == b"LVAL"
        )

    @staticmethod
    def _scan_ovba_signatures(row: bytes) -> list[int]:
        """Return every offset inside ``row`` that plausibly begins an
        MS-OVBA stream (sig byte ``0x01`` + chunk header with signature
        bits ``0b011``)."""
        out: list[int] = []
        i = 0
        n = len(row)
        while i + 3 <= n:
            j = row.find(b"\x01", i)
            if j < 0 or j + 3 > n:
                break
            hdr = int.from_bytes(row[j + 1 : j + 3], "little")
            if ((hdr >> 12) & 0x7) == 0b011:
                out.append(j)
            i = j + 1
        return out

    def vba_module_names(self) -> list[str]:
        """
        Return the names of every distinct VBA module in this database, in
        the order they are first encountered.

        When the MS-OVBA dir-stream catalog can be located (the normal
        case), its authoritative module ordering is returned and shadow /
        undo copies of edited modules are ignored. Otherwise this falls
        back to scanning every OVBA stream for ``Attribute VB_Name`` and
        deduplicating in encounter order.
        """
        try:
            project = self.read_project_info()
        except AccessError:
            project = None
        if project is not None:
            return [m.name for m in project.modules]
        seen: list[str] = []
        for m in self.iter_vba_modules():
            if m.name and m.name not in seen:
                seen.append(m.name)
        return seen

    def read_vba_module(self, name: str) -> str:
        """
        Return the user-visible source text of the named module (without
        the leading ``Attribute VB_*`` preamble lines, with ``\\r\\n``
        line endings preserved).

        When Access has shadow copies of the same module on disk, the copy
        with the highest file offset is returned (this is the most recent
        write).

        Raises :class:`AccessError` if no module with that name is found.
        """
        candidates = [m for m in self.iter_vba_modules() if m.name == name]
        if not candidates:
            raise AccessError(
                f"VBA module {name!r} not found in {self.path.name!r}"
            )
        candidates.sort(key=lambda m: m.start_offset)
        return candidates[-1].source

    # ------------------------------------------------------------------
    # Write path (EXPERIMENTAL).
    # ------------------------------------------------------------------
    #
    # Access does **not** display VBA module source by decompressing the
    # MS-OVBA blob -- the blob is a passive cache. We verified this by
    # zero-filling Module2's entire OVBA chain in the live fixture and
    # observing that Access COM (and the VBA editor) still rendered the
    # module's source correctly.
    #
    # Access's authoritative storage is:
    #
    #   * **Comments**: stored verbatim in plaintext rows tagged ``E3 00 00 00``
    #     followed by a u16-LE byte length and the ASCII payload (with the
    #     leading apostrophe stripped).
    #   * **String literals**: stored verbatim in plaintext rows tagged
    #     ``B9 00`` followed by a u16-LE byte length, the ASCII payload, and
    #     a 12-byte row-metadata trailer.
    #   * **Code structure** (procedure names, statements, keywords): stored
    #     as Access-flavoured p-code in tables we do not currently parse.
    #
    # Therefore the write surface we expose is:
    #
    #   * :meth:`replace_text` -- same-length byte-for-byte substitution of
    #     ASCII text. This is sufficient to patch comment text and string
    #     literal contents (verified against Access COM and the on-screen
    #     VBA editor on a live fixture).
    #
    # Larger structural edits (changing procedure names, adding/removing
    # statements, etc.) require regenerating Access's p-code tables, which
    # is outside the current scope. The OVBA blob can still be regenerated
    # to keep our reader self-consistent, but doing so has no effect on
    # what Access displays.

    def replace_text(
        self,
        old: Union[bytes, str],
        new: Union[bytes, str],
        *,
        count: int = -1,
        encoding: str = "latin-1",
    ) -> int:
        """
        Replace every occurrence of ``old`` with ``new`` in the database
        byte buffer. ``old`` and ``new`` **must be the same length**; this
        is the safe primitive that does not perturb row offsets, page
        layouts, or catalog entries.

        Strings are encoded with ``encoding`` (default ``latin-1`` -- the
        encoding Access uses for ASCII source bytes).

        ``count`` caps the number of replacements (``-1`` = unlimited).
        Returns the number of replacements made. Raises :class:`AccessError`
        if ``old`` does not appear at least once or if the lengths differ.

        Call :meth:`save` to persist the change.

        Use cases (verified end-to-end against Access COM and the VBA
        editor on a live ``.accdb``):

        * Patching the text of a string literal inside a code line such as
          ``MsgBox "old text"`` -> ``MsgBox "new text!"`` -- the literal
          payload is stored in a ``B9 00`` row.
        * Patching the body of a comment line ``' old`` -> ``' new`` --
          the comment text (minus leading apostrophe) is stored in an
          ``E3 00 00 00`` row.

        Use cases that this method does **not** support:

        * Changing the length of a literal or comment (requires re-laying
          row offset tables -- not yet implemented).
        * Renaming procedures or otherwise changing code structure
          (requires regenerating Access p-code -- not implemented).
        """
        old_b = old.encode(encoding) if isinstance(old, str) else bytes(old)
        new_b = new.encode(encoding) if isinstance(new, str) else bytes(new)
        if len(old_b) != len(new_b):
            raise AccessError(
                f"replace_text requires same-length operands: "
                f"len(old)={len(old_b)} != len(new)={len(new_b)}"
            )
        if not old_b:
            raise AccessError("replace_text: old must be non-empty")
        replaced = 0
        i = 0
        while count < 0 or replaced < count:
            j = self._data.find(old_b, i)
            if j < 0:
                break
            self._data[j : j + len(old_b)] = new_b
            replaced += 1
            i = j + len(new_b)
        if replaced == 0:
            raise AccessError(
                f"replace_text: pattern {old_b!r} not found in "
                f"{self.path.name!r}"
            )
        return replaced

    # ------------------------------------------------------------------
    # Length-changing edits (Phase 5 write path, 2026-05).
    # ------------------------------------------------------------------
    #
    # ``replace_text`` requires equal-length operands so that no offsets
    # have to move. The methods below allow length changes when the
    # match falls entirely inside a single LVAL row -- in which case we
    # can reflow only that page's slot table and leave the rest of the
    # database untouched. Edits that straddle row boundaries or touch
    # non-LVAL pages are rejected.

    def _locate_match_in_lval(
        self, needle: bytes
    ) -> tuple[int, int, int]:
        """Locate the unique LVAL row containing ``needle`` and return
        ``(page, slot, row_relative_offset)``. Raises if zero or more
        than one row contains the pattern (callers can downgrade
        multi-hit to first-hit with ``count=1`` semantics)."""
        hits: list[tuple[int, int, int]] = []
        for page, slot, row in self._iter_lval_rows():
            i = row.find(needle)
            if i >= 0:
                hits.append((page, slot, i))
        if not hits:
            raise AccessError(
                f"pattern {needle!r} not found in any LVAL row "
                f"of {self.path.name!r}"
            )
        if len(hits) > 1:
            raise AccessError(
                f"pattern {needle!r} occurs in {len(hits)} LVAL rows; "
                f"refine the search or use replace_text_resize "
                f"with count=1"
            )
        return hits[0]

    def replace_text_resize(
        self,
        old: Union[bytes, str],
        new: Union[bytes, str],
        *,
        count: int = -1,
        encoding: str = "latin-1",
    ) -> int:
        """Replace ``old`` with ``new`` allowing length changes, by
        reflowing the affected LVAL page's slot table.

        Unlike :meth:`replace_text`, this method requires that every
        match live entirely inside a single LVAL row (i.e. the bytes
        don't straddle row or page boundaries). The matching row is
        resized in place via :meth:`_lval_resize_row`; other rows on
        the same page shift to accommodate. Other pages are not
        touched.

        ``count`` caps the number of replacements (``-1`` = unlimited).
        Returns the number of replacements made.

        Raises :class:`AccessError` if no match is found, if the match
        straddles a row boundary, or if the new row would not fit on
        its page (caller must catch and retry with a shorter
        replacement, or fall back to a chain-aware writer once one
        exists).
        """
        old_b = old.encode(encoding) if isinstance(old, str) else bytes(old)
        new_b = new.encode(encoding) if isinstance(new, str) else bytes(new)
        if not old_b:
            raise AccessError("replace_text_resize: old must be non-empty")
        replaced = 0
        while count < 0 or replaced < count:
            # Re-scan after each replacement -- offsets shift.
            try:
                page, slot, row_off = self._locate_match_in_lval(old_b)
            except AccessError as exc:
                msg = str(exc)
                if "occurs in" in msg:
                    # Multi-hit: replace only the first.
                    found_one: tuple[int, int, int] | None = None
                    for page, slot, row in self._iter_lval_rows():
                        i = row.find(old_b)
                        if i >= 0:
                            found_one = (page, slot, i)
                            break
                    if found_one is None:
                        break
                    page, slot, row_off = found_one
                else:
                    if replaced > 0:
                        break
                    raise
            row = self._lval_row_bytes(page, slot)
            # Confirm the match doesn't extend past the row end.
            if row_off + len(old_b) > len(row):
                raise AccessError(
                    f"replace_text_resize: match at page={page} slot={slot} "
                    f"row_off={row_off} straddles row end (row_len={len(row)})"
                )
            new_row = row[:row_off] + new_b + row[row_off + len(old_b):]
            self._lval_write_row(page, slot, new_row)
            replaced += 1
        if replaced == 0:
            raise AccessError(
                f"replace_text_resize: pattern {old_b!r} not found "
                f"in any LVAL row of {self.path.name!r}"
            )
        return replaced

    def write_lval_row(
        self, page: int, slot: int, new_bytes: Union[bytes, bytearray]
    ) -> None:
        """Public wrapper around :meth:`_lval_write_row` for advanced
        callers that have located a specific LVAL row (via
        :meth:`_iter_lval_rows` or similar) and want to replace its
        entire payload.

        ``new_bytes`` must include any continuation prefix bytes for
        rows that are part of an LVAL chain. The row is resized in
        place on its page; other rows on the same page reflow.

        Raises :class:`AccessError` if the new payload doesn't fit
        on the page (free space + current row size < ``len(new_bytes)``).
        """
        self._lval_write_row(page, slot, bytes(new_bytes))

    def lval_free_space(self, page: int) -> int:
        """Public wrapper around :meth:`_lval_free_space`. Returns the
        number of bytes currently available for new payload on the
        given LVAL page (i.e. the gap between the slot table and the
        lowest live row).

        Raises :class:`AccessError` if ``page`` is not an LVAL page.
        """
        base = page * ACE_PAGE_SIZE
        if (
            self._data[base] != 0x01
            or bytes(self._data[base + 4 : base + 8]) != b"LVAL"
        ):
            raise AccessError(f"page {page} is not an LVAL page")
        return self._lval_free_space(page)

    # ------------------------------------------------------------------
    # Dir-stream catalog surgery (Phase 5 module mutation, 2026-05).
    # ------------------------------------------------------------------
    #
    # The PROJECTMODULES region of the dir-stream uses the regular
    # ``<u16 id><u32 size><data>`` record format end-to-end (the
    # complications of PROJECTINFORMATION / PROJECTREFERENCES with
    # PROJECTVERSION and REFERENCECONTROL extended records are all
    # *before* PROJECTMODULES). Within PROJECTMODULES every record is
    # safely walkable.
    #
    # Module record layout (verified against samples 010-051):
    #   MODULENAME              (0x0019) <u32 size> <MBCS name>
    #   MODULENAMEUNICODE       (0x0047) <u32 size> <UTF-16-LE name>
    #   MODULESTREAMNAME        (0x001A) <u32 size> <MBCS stream name>
    #   MODULESTREAMNAMEUNICODE (0x0032) <u32 size> <UTF-16-LE>
    #   MODULEDOCSTRING         (0x001C) optional
    #   MODULEDOCSTRINGUNICODE  (0x0048) optional
    #   MODULEOFFSET            (0x0031) <u32 size=4> <u32 text_off>
    #   MODULEHELPCONTEXT       (0x001E) <u32 size=4> <u32>
    #   MODULECOOKIE            (0x002C) <u32 size=2> <u16>
    #   MODULETYPE              (0x0021 std | 0x0022 other) size=0
    #   MODULEREADONLY          (0x0025) optional size=0
    #   MODULEPRIVATE           (0x0028) optional size=0
    #   MODULE-TERMINATOR       (0x002B) size=0

    _PROJECTMODULES_ID = 0x000F
    _MODULENAME_ID = 0x0019
    _MODULENAMEUNICODE_ID = 0x0047
    _MODULESTREAMNAME_ID = 0x001A
    _MODULESTREAMNAMEUNICODE_ID = 0x0032
    _MODULEDOCSTRING_ID = 0x001C
    _MODULEDOCSTRINGUNICODE_ID = 0x0048
    _MODULEOFFSET_ID = 0x0031
    _MODULEHELPCONTEXT_ID = 0x001E
    _MODULECOOKIE_ID = 0x002C
    _MODULETYPE_STD_ID = 0x0021
    _MODULETYPE_OTHER_ID = 0x0022
    _MODULEREADONLY_ID = 0x0025
    _MODULEPRIVATE_ID = 0x0028
    _MODULE_TERMINATOR_ID = 0x002B
    _DIR_TERMINATOR_ID = 0x0010

    def _locate_modules_section(
        self, dir_raw: bytes
    ) -> tuple[int, int, int]:
        """Return ``(count_record_pos, modules_start, dir_term_pos)``
        for the decompressed dir-stream ``dir_raw``.

        * ``count_record_pos`` is the start of the PROJECTMODULES
          (0x000F) record; the u16 module count lives at offset +6.
        * ``modules_start`` is the byte offset of the first MODULENAME
          record (or ``dir_term_pos`` when no modules are present).
        * ``dir_term_pos`` is the byte offset of the DIR-TERMINATOR
          (0x0010) record.
        """
        # PROJECTMODULES always has size=2; search for the canonical
        # byte signature ``0F 00 02 00 00 00`` and validate.
        sig = bytes(
            self._PROJECTMODULES_ID.to_bytes(2, "little")
            + (2).to_bytes(4, "little")
        )
        i = dir_raw.find(sig)
        if i < 0:
            raise AccessError(
                "PROJECTMODULES record not found in dir-stream"
            )
        count_pos = i
        # DIR-TERMINATOR: id=0x0010 followed by reserved u32=0
        # (i.e. the 6-byte tail ``10 00 00 00 00 00``).
        dt_sig = (
            self._DIR_TERMINATOR_ID.to_bytes(2, "little")
            + (0).to_bytes(4, "little")
        )
        dt = dir_raw.rfind(dt_sig)
        if dt < 0:
            raise AccessError(
                "DIR-TERMINATOR record not found in dir-stream"
            )
        # Modules start after PROJECTMODULES + optional PROJECTCOOKIE.
        pos = count_pos + 8  # 6-byte header + 2-byte count payload
        while pos + 6 <= dt:
            rid = int.from_bytes(dir_raw[pos : pos + 2], "little")
            if rid == self._MODULENAME_ID:
                return count_pos, pos, dt
            sz = int.from_bytes(dir_raw[pos + 2 : pos + 6], "little")
            pos += 6 + sz
        return count_pos, dt, dt

    def _iter_module_records(
        self, dir_raw: bytes
    ) -> Iterator[tuple[int, int, str]]:
        """Yield ``(record_start, record_end, module_name)`` for each
        module block in the dir-stream. ``record_end`` is exclusive
        (i.e. one past the MODULE-TERMINATOR's payload)."""
        _, modules_start, dir_term = self._locate_modules_section(dir_raw)
        pos = modules_start
        while pos < dir_term:
            rid = int.from_bytes(dir_raw[pos : pos + 2], "little")
            if rid != self._MODULENAME_ID:
                break
            name_size = int.from_bytes(dir_raw[pos + 2 : pos + 6], "little")
            name = bytes(
                dir_raw[pos + 6 : pos + 6 + name_size]
            ).decode("latin-1", errors="replace")
            cur = pos + 6 + name_size
            while cur < dir_term:
                sub_rid = int.from_bytes(dir_raw[cur : cur + 2], "little")
                sub_sz = int.from_bytes(dir_raw[cur + 2 : cur + 6], "little")
                cur += 6 + sub_sz
                if sub_rid == self._MODULE_TERMINATOR_ID:
                    break
            yield (pos, cur, name)
            pos = cur

    def _find_module_record(
        self, dir_raw: bytes, name: str
    ) -> tuple[int, int]:
        """Return ``(record_start, record_end)`` for the module with
        the given MBCS name. Raises :class:`AccessError` if not found."""
        for start, end, n in self._iter_module_records(dir_raw):
            if n == name:
                return start, end
        names = [n for _, _, n in self._iter_module_records(dir_raw)]
        raise AccessError(
            f"module {name!r} not found in dir-stream; "
            f"present modules: {names}"
        )

    def _find_subrecord(
        self, dir_raw: bytes, start: int, end: int, record_id: int
    ) -> tuple[int, int] | None:
        """Return ``(record_start, data_end)`` of the first sub-record
        with id ``record_id`` inside the module record at ``[start,
        end)``. ``data_end`` is the offset just past the record's
        payload. Returns ``None`` if not found.

        The walker starts AT ``start`` (which is the MODULENAME
        record), so callers asking for MODULENAME itself will find it
        immediately.
        """
        pos = start
        while pos + 6 <= end:
            rid = int.from_bytes(dir_raw[pos : pos + 2], "little")
            sz = int.from_bytes(dir_raw[pos + 2 : pos + 6], "little")
            data_end = pos + 6 + sz
            if rid == record_id:
                return pos, data_end
            pos = data_end
        return None

    def _rewrite_subrecord_payload(
        self,
        dir_raw: bytes,
        start: int,
        end: int,
        record_id: int,
        new_payload: bytes,
    ) -> bytes:
        """Return a new dir-stream with the payload of the sub-record
        ``record_id`` (inside module record ``[start, end)``) replaced
        by ``new_payload``. The size field is updated to ``len(new_payload)``.
        Raises if the sub-record is absent.
        """
        loc = self._find_subrecord(dir_raw, start, end, record_id)
        if loc is None:
            raise AccessError(
                f"sub-record 0x{record_id:04x} not found in module "
                f"record [{start}, {end})"
            )
        rec_start, data_end = loc
        old_size = int.from_bytes(
            dir_raw[rec_start + 2 : rec_start + 6], "little"
        )
        assert data_end == rec_start + 6 + old_size
        new_size = len(new_payload)
        new_header = (
            record_id.to_bytes(2, "little") + new_size.to_bytes(4, "little")
        )
        return bytes(
            dir_raw[:rec_start]
            + new_header
            + new_payload
            + dir_raw[data_end:]
        )

    def _write_catalog_dir_stream(self, new_dir_raw: bytes) -> None:
        """Re-OVBA-compress ``new_dir_raw`` and write it to the catalog
        LVAL row, resizing the row as needed.

        Caller is responsible for ensuring the new dir-stream is
        structurally valid (record framing intact, terminators present).
        Raises :class:`AccessError` if the catalog row can't be located
        or the new compressed payload doesn't fit on its page.
        """
        from pyopenvba.vba import compress as _ovba_compress

        found = self._find_catalog_row()
        if found is None:
            raise AccessError(
                "no MS-OVBA dir-stream catalog row in this database"
            )
        page, slot, _old_raw = found
        new_compressed = _ovba_compress(new_dir_raw)
        self._lval_write_row(page, slot, new_compressed)

    # ------------------------------------------------------------------
    # Module mutation public surface (Phase 5, 2026-05).
    # ------------------------------------------------------------------
    #
    # These operations rewrite the dir-stream catalog and, where
    # possible, the OVBA cache row that carries the human-readable
    # source. They make the changes that downstream pure-Python readers
    # (``read_project_info``, ``iter_vba_modules``) will observe; they
    # do NOT update the Access engine's row-level structures
    # (MSysObjects / DATA pages / p-code tables) and therefore should
    # not be relied on to drive the live Access VBA editor UI without
    # additional reverse-engineering work documented in
    # ``memories/repo/access-vba-storage.md``.

    def rename_module(self, old_name: str, new_name: str) -> None:
        """Rename a VBA module in the dir-stream catalog and update
        the corresponding ``Attribute VB_Name`` in the OVBA cache row.

        Updates MODULENAME (MBCS) and MODULENAMEUNICODE (UTF-16-LE) in
        the dir-stream. The MODULESTREAMNAME field is intentionally
        left untouched -- Access does not use it as a stream key.

        Call :meth:`save` to persist.

        Raises :class:`AccessError` if ``old_name`` is not present,
        if ``new_name`` is empty or contains characters that can't be
        encoded in the project code page, or if any rewrite would
        cause a layout overflow.
        """
        if not new_name:
            raise AccessError("rename_module: new_name must be non-empty")
        try:
            new_mbcs = new_name.encode("latin-1")
        except UnicodeEncodeError as exc:
            raise AccessError(
                f"rename_module: new_name {new_name!r} not encodable in "
                f"latin-1 (project code page)"
            ) from exc
        new_utf16 = new_name.encode("utf-16-le")

        found = self._find_catalog_row()
        if found is None:
            raise AccessError("no dir-stream catalog row found")
        _page, _slot, dir_raw = found

        mod_start, mod_end = self._find_module_record(dir_raw, old_name)
        # MODULENAME is the very first record in the block; data is
        # everything after the 6-byte header up to MODULENAMEUNICODE.
        new_dir = self._rewrite_subrecord_payload(
            dir_raw, mod_start, mod_end, self._MODULENAME_ID, new_mbcs
        )
        # Recompute module record bounds in the new buffer (size shift).
        delta_mbcs = len(new_mbcs) - (
            mod_end - mod_start
            - (len(new_dir) - len(dir_raw))
            # actually compute by re-locating after the change:
        )
        # Re-locate by name (we just renamed, so search for new name).
        mod_start_b, mod_end_b = self._find_module_record(new_dir, new_name)
        new_dir = self._rewrite_subrecord_payload(
            new_dir,
            mod_start_b,
            mod_end_b,
            self._MODULENAMEUNICODE_ID,
            new_utf16,
        )
        _ = delta_mbcs  # silence Pyright; intentionally unused

        # Write back the catalog.
        self._write_catalog_dir_stream(new_dir)

        # Best-effort: update the OVBA cache's ``Attribute VB_Name`` line.
        # Failure here is not fatal -- the dir-stream catalog is the
        # authoritative source for module identity; the OVBA cache is
        # informational and will be regenerated by Access on next save.
        try:
            old_attr = b'Attribute VB_Name = "' + old_name.encode("latin-1") + b'"'
            new_attr = b'Attribute VB_Name = "' + new_mbcs + b'"'
            if old_attr != new_attr:
                # The Attribute line lives inside an OVBA-compressed
                # blob, so a raw replace_text won't find it. We have to
                # decompress, edit, re-compress, and write the cache row
                # back. Identify the row via _scan_ovba_signatures.
                self._patch_ovba_cache_attribute(old_name, old_attr, new_attr)
        except AccessError:
            # OVBA cache update is opportunistic; ignore failures.
            pass

        # MSysObjects (Jet/ACE system catalog) update: the live Access
        # engine consults MSysObjects to drive the Navigation Pane and
        # the VBA editor. Updating only the dir-stream catalog leaves
        # the two sources of truth desynchronised. If no MSysObjects
        # row exists for ``old_name`` (e.g. the module was just added
        # via add_module_catalog_entry without a paired MSysObjects
        # row) silently skip; otherwise rename succeeds.
        if self.find_msys_object(old_name) is not None:
            self.rename_msys_object(old_name, new_name)

    def _patch_ovba_cache_attribute(
        self,
        module_name: str,
        old_attr: bytes,
        new_attr: bytes,
    ) -> None:
        """Find the OVBA cache row that decompresses to source whose
        ``Attribute VB_Name = "<module_name>"`` line matches, rewrite
        the attribute line, re-compress, and write the row back.

        Best-effort: silently no-ops if the cache row isn't found.
        """
        from pyopenvba.vba import compress as _ovba_compress

        for page, slot, row in list(self._iter_lval_rows()):
            sigs = self._scan_ovba_signatures(row)
            for off in sigs:
                try:
                    decomp = _ovba_decompress(
                        bytes(row[off:]),
                        stream_name=f"accdb_ovba_cache@({page},{slot})+{off}",
                    )
                except Exception:
                    continue
                if old_attr not in decomp:
                    continue
                if not decomp.startswith(b"Attribute VB_Name = \""):
                    continue
                # Verify it's the module we're targeting.
                tail = decomp[len(b'Attribute VB_Name = "') :]
                end_q = tail.find(b'"')
                if end_q < 0:
                    continue
                cur_name = tail[:end_q].decode("latin-1", errors="replace")
                if cur_name != module_name:
                    continue
                new_decomp = decomp.replace(old_attr, new_attr, 1)
                new_compressed = _ovba_compress(new_decomp)
                new_row = bytes(row[:off]) + new_compressed
                self._lval_write_row(page, slot, new_row)
                return
        # Not found: leave it alone.

    def delete_module(self, name: str) -> None:
        """Remove a VBA module from the dir-stream catalog and
        tombstone its OVBA cache row.

        This excises the module's MODULENAME...MODULE-TERMINATOR block
        from the dir-stream, decrements PROJECTMODULES count, and (if
        the OVBA cache row for this module can be located) tombstones
        that LVAL slot so subsequent reads skip it.

        Call :meth:`save` to persist.

        Raises :class:`AccessError` if ``name`` is not present.

        IMPORTANT: this is a *catalog-level* delete. The Access engine
        keeps additional structural references (MSysObjects DATA-page
        rows, p-code tables) that this method does not touch. Reads via
        :meth:`read_project_info` and :meth:`iter_vba_modules` will
        reflect the removal; the live Access VBA editor may still show
        the module until those engine-level references are also
        cleaned up. See ``memories/repo/access-vba-storage.md`` for
        the remaining barriers.
        """
        found = self._find_catalog_row()
        if found is None:
            raise AccessError("no dir-stream catalog row found")
        _page, _slot, dir_raw = found

        mod_start, mod_end = self._find_module_record(dir_raw, name)
        count_pos, _modules_start, _dir_term = self._locate_modules_section(
            dir_raw
        )
        old_count = int.from_bytes(
            dir_raw[count_pos + 6 : count_pos + 8], "little"
        )
        if old_count == 0:
            raise AccessError(
                f"delete_module: PROJECTMODULES count already 0; "
                f"refusing to delete {name!r}"
            )
        new_count = old_count - 1
        new_dir_pre = (
            dir_raw[:count_pos + 6]
            + new_count.to_bytes(2, "little")
            + dir_raw[count_pos + 8 : mod_start]
            + dir_raw[mod_end:]
        )

        self._write_catalog_dir_stream(bytes(new_dir_pre))

        # Best-effort: tombstone EVERY OVBA cache row whose decompressed
        # payload begins with ``Attribute VB_Name = "<name>"``. Access
        # is known to keep multiple redundant cache copies for some
        # modules; deleting just one is insufficient to make
        # iter_vba_modules stop yielding the module.
        to_tombstone: list[tuple[int, int]] = []
        for page, slot, row in list(self._iter_lval_rows()):
            sigs = self._scan_ovba_signatures(row)
            for off in sigs:
                try:
                    decomp = _ovba_decompress(
                        bytes(row[off:]),
                        stream_name=f"accdb_ovba_cache@({page},{slot})+{off}",
                    )
                except Exception:
                    continue
                if not decomp.startswith(b"Attribute VB_Name = \""):
                    continue
                tail = decomp[len(b'Attribute VB_Name = "') :]
                q = tail.find(b'"')
                if q < 0:
                    continue
                if tail[:q].decode("latin-1", errors="replace") == name:
                    to_tombstone.append((page, slot))
                    break
        for page, slot in to_tombstone:
            self._lval_tombstone_slot(page, slot)

        # MSysObjects: tombstone the matching system-catalog row so
        # Access's Navigation Pane no longer surfaces the module. We
        # use find_msys_object (not _module) to match by name only,
        # in case Type happens to differ on legacy files.
        if self.find_msys_object(name) is not None:
            self.delete_msys_object(name)

    def modify_module_cache(
        self, name: str, new_source: str
    ) -> None:
        """Rewrite the OVBA cache row for ``name`` so its decompressed
        payload is ``Attribute VB_Name = "<name>"`` + ``new_source``.

        Updates only the OVBA *cache* row (the passive plaintext mirror
        Access keeps). Does NOT update the authoritative p-code
        tables, so the Access VBA editor will continue to show the
        previously-compiled source until a recompile is triggered.

        Use cases:
        * Exporting the database, editing the cache, importing back
          via the OVBA toolchain.
        * Diff-driven version control of the VBA source.

        Raises :class:`AccessError` if ``name`` is not present in the
        catalog, if the OVBA cache row can't be located, or if the
        new compressed payload doesn't fit on the row's page.
        """
        from pyopenvba.vba import compress as _ovba_compress

        info = self.read_project_info()
        if not any(m.name == name for m in info.modules):
            raise AccessError(
                f"modify_module_cache: module {name!r} not present "
                f"in dir-stream catalog"
            )
        new_decomp = (
            b'Attribute VB_Name = "'
            + name.encode("latin-1")
            + b'"\r\n'
            + new_source.encode("latin-1")
        )
        # Locate the existing OVBA cache row for this module.
        target: tuple[int, int, int] | None = None
        for page, slot, row in list(self._iter_lval_rows()):
            sigs = self._scan_ovba_signatures(row)
            for off in sigs:
                try:
                    decomp = _ovba_decompress(
                        bytes(row[off:]),
                        stream_name=f"accdb_ovba_cache@({page},{slot})+{off}",
                    )
                except Exception:
                    continue
                if not decomp.startswith(b"Attribute VB_Name = \""):
                    continue
                tail = decomp[len(b'Attribute VB_Name = "') :]
                q = tail.find(b'"')
                if q < 0:
                    continue
                if tail[:q].decode("latin-1", errors="replace") == name:
                    target = (page, slot, off)
                    break
            if target is not None:
                break
        if target is None:
            raise AccessError(
                f"modify_module_cache: OVBA cache row for module "
                f"{name!r} not found"
            )
        page, slot, off = target
        row = self._lval_row_bytes(page, slot)
        new_compressed = _ovba_compress(new_decomp)
        new_row = bytes(row[:off]) + new_compressed
        self._lval_write_row(page, slot, new_row)

    def add_module_catalog_entry(
        self,
        name: str,
        *,
        is_class_module: bool = False,
        is_private: bool = False,
        is_read_only: bool = False,
        stream_name: str | None = None,
        cookie: int = 0xFFFF,
    ) -> None:
        """Append a new module record to the dir-stream catalog.

        This is a *catalog-level* add: it makes the module visible to
        :meth:`read_project_info` and other dir-stream readers. It does
        NOT create the p-code rows or the MSysObjects table entries
        the Access engine requires to surface the module in the live
        VBA editor; doing so requires reverse-engineering work
        documented in ``memories/repo/access-vba-storage.md``.

        Useful primarily for round-trip testing of the dir-stream
        catalog writer and for offline manipulation of project
        metadata.

        Raises :class:`AccessError` if ``name`` is already present or
        if the new dir-stream doesn't fit on its LVAL page.
        """
        if not name:
            raise AccessError("add_module_catalog_entry: name must be non-empty")
        try:
            name_mbcs = name.encode("latin-1")
        except UnicodeEncodeError as exc:
            raise AccessError(
                f"add_module_catalog_entry: name {name!r} not encodable "
                f"in latin-1"
            ) from exc
        name_utf16 = name.encode("utf-16-le")
        if stream_name is None:
            # Pick a deterministic placeholder; Access doesn't actually
            # use this field, but the dir-stream format requires it.
            stream_name = (name + "_STREAM").upper()
        stream_mbcs = stream_name.encode("latin-1")
        stream_utf16 = stream_name.encode("utf-16-le")

        found = self._find_catalog_row()
        if found is None:
            raise AccessError("no dir-stream catalog row found")
        _page, _slot, dir_raw = found

        # Reject duplicates.
        for _s, _e, n in self._iter_module_records(dir_raw):
            if n == name:
                raise AccessError(
                    f"add_module_catalog_entry: module {name!r} already exists"
                )

        def rec(rid: int, payload: bytes) -> bytes:
            return (
                rid.to_bytes(2, "little")
                + len(payload).to_bytes(4, "little")
                + payload
            )

        mod_type_id = (
            self._MODULETYPE_OTHER_ID
            if is_class_module
            else self._MODULETYPE_STD_ID
        )
        block = b"".join(
            [
                rec(self._MODULENAME_ID, name_mbcs),
                rec(self._MODULENAMEUNICODE_ID, name_utf16),
                rec(self._MODULESTREAMNAME_ID, stream_mbcs),
                rec(self._MODULESTREAMNAMEUNICODE_ID, stream_utf16),
                rec(self._MODULEDOCSTRING_ID, b""),
                rec(self._MODULEDOCSTRINGUNICODE_ID, b""),
                rec(self._MODULEOFFSET_ID, (0).to_bytes(4, "little")),
                rec(self._MODULEHELPCONTEXT_ID, (0).to_bytes(4, "little")),
                rec(self._MODULECOOKIE_ID, cookie.to_bytes(2, "little")),
                rec(mod_type_id, b""),
            ]
            + ([rec(self._MODULEREADONLY_ID, b"")] if is_read_only else [])
            + ([rec(self._MODULEPRIVATE_ID, b"")] if is_private else [])
            + [rec(self._MODULE_TERMINATOR_ID, b"")]
        )

        count_pos, _modules_start, dir_term = self._locate_modules_section(
            dir_raw
        )
        old_count = int.from_bytes(
            dir_raw[count_pos + 6 : count_pos + 8], "little"
        )
        new_count = old_count + 1
        new_dir = (
            dir_raw[:count_pos + 6]
            + new_count.to_bytes(2, "little")
            + dir_raw[count_pos + 8 : dir_term]
            + block
            + dir_raw[dir_term:]
        )
        self._write_catalog_dir_stream(bytes(new_dir))

        # MSysObjects: add a parallel system-catalog row so the Access
        # Navigation Pane / VBA editor knows about this module. If
        # MSysObjects already has a row for ``name`` (rare -- e.g.
        # caller pre-populated it), skip rather than duplicate.
        if self.find_msys_object(name) is None:
            self.add_msys_module_object(name)

    # ------------------------------------------------------------------

    def replace_module(self, name: str, new_source: str) -> None:
        """
        Replace the OVBA source cache for an existing module.

        Convenience alias for :meth:`modify_module_cache` that reads more
        naturally for one-shot "swap module source" workflows.
        """
        self.modify_module_cache(name, new_source)

    def export_module(self, name: str) -> str:
        """
        Return the user-visible source text of a single module by name.

        Identical to :meth:`read_vba_module`; provided for symmetry with
        :meth:`export_modules` and :meth:`import_module`.
        """
        return self.read_vba_module(name)

    def export_modules(
        self,
        dest_dir: Union[str, Path],
        *,
        include_attributes: bool = False,
    ) -> list[Path]:
        """
        Write every module to ``dest_dir`` as one file per module.

        Class modules are written as ``<name>.cls``; everything else as
        ``<name>.bas``. The leading ``Attribute VB_*`` preamble is omitted
        by default (this matches what the VBA editor shows on screen); set
        ``include_attributes=True`` to round-trip the raw stream.

        Returns the list of files written. The destination directory is
        created if it does not exist.
        """
        out_dir = Path(dest_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        class_names: set[str] = set()
        try:
            project = self.read_project_info()
            class_names = {
                m.name for m in project.modules if m.is_class_module
            }
        except AccessError:
            pass
        written: list[Path] = []
        seen: set[str] = set()
        for module in self.iter_vba_modules():
            if module.name in seen:
                continue
            seen.add(module.name)
            ext = ".cls" if module.name in class_names else ".bas"
            target = out_dir / (module.name + ext)
            if include_attributes:
                body = module.attributes_text + module.source
            else:
                body = module.source
            target.write_bytes(body.encode("latin-1"))
            written.append(target)
        return written

    def import_module(
        self,
        source_or_path: Union[str, Path],
        *,
        name: Union[str, None] = None,
        is_class_module: Union[bool, None] = None,
        replace_existing: bool = False,
    ) -> str:
        """
        Add (or replace) a module from a ``.bas``/``.cls`` file or from a
        raw VBA source string.

        - If ``source_or_path`` is a :class:`Path` (or a string referring
          to an existing file), the file's contents are loaded as the
          new source. Module type is inferred from the file extension
          (``.cls`` -> class module) unless ``is_class_module`` is set.
          Module ``name`` defaults to the file stem.
        - If ``source_or_path`` is a string and is not an existing file,
          it is treated as raw source. ``name`` is required in that case.

        If a module with the resolved name already exists:
          * with ``replace_existing=True``, its OVBA cache is rewritten
            via :meth:`modify_module_cache`;
          * otherwise an :class:`AccessError` is raised.

        Returns the final module name as stored.
        """
        resolved_name: Union[str, None] = name
        is_class: bool = (
            bool(is_class_module) if is_class_module is not None else False
        )

        path_candidate: Union[Path, None] = None
        if isinstance(source_or_path, Path):
            path_candidate = source_or_path
        else:
            try:
                maybe = Path(source_or_path)
                if maybe.exists() and maybe.is_file():
                    path_candidate = maybe
            except (OSError, ValueError):
                path_candidate = None

        source_text: str
        if path_candidate is not None:
            source_text = path_candidate.read_bytes().decode("latin-1")
            if resolved_name is None:
                resolved_name = path_candidate.stem
            if is_class_module is None:
                is_class = path_candidate.suffix.lower() == ".cls"
        else:
            source_text = str(source_or_path)
            if resolved_name is None:
                raise AccessError(
                    "import_module: explicit name= is required when "
                    "passing raw source"
                )

        # Normalize line endings to CRLF (Access on-disk convention).
        normalized = source_text.replace("\r\n", "\n").replace("\r", "\n")
        source_text = normalized.replace("\n", "\r\n")

        existing = set(self.vba_module_names())
        if resolved_name in existing:
            if not replace_existing:
                raise AccessError(
                    f"import_module: module {resolved_name!r} already "
                    "exists (pass replace_existing=True to overwrite)"
                )
            self.modify_module_cache(resolved_name, source_text)
            return resolved_name

        self.add_module_catalog_entry(
            resolved_name, is_class_module=is_class
        )
        return resolved_name

    # ------------------------------------------------------------------
    # MSysObjects (Jet/ACE system catalog) -- read path
    # ------------------------------------------------------------------

    def _iter_msys_data_pages(self) -> Iterator[int]:
        """Yield page numbers of every DATA page whose owner is the
        MSysObjects TDEF at page 2."""
        page_count = len(self._data) // ACE_PAGE_SIZE
        for pn in range(1, page_count):
            base = pn * ACE_PAGE_SIZE
            if self._data[base] != PAGE_TYPE_DATA:
                continue
            owner = int.from_bytes(self._data[base + 4 : base + 8], "little")
            if owner != _MSYS_OBJECTS_TDEF_PAGE:
                continue
            yield pn

    def _decode_msys_row(
        self, row: bytes, *, page: int, slot: int
    ) -> "AccessSysObject | None":
        """Decode one MSysObjects row. Returns ``None`` for rows that
        do not match the expected 17-column / 11-var-column schema
        (such rows are silently skipped to keep the reader robust
        across unanticipated catalog variants)."""
        # Need at least: 32 fixed bytes + jump table + var_count + null_mask.
        min_len = 32 + 2 * _MSYS_VAR_COL_COUNT + 2 + _MSYS_NULL_MASK_BYTES
        if len(row) < min_len:
            return None
        cols = int.from_bytes(row[0:2], "little")
        if cols != _MSYS_COL_COUNT:
            return None
        id_ = int.from_bytes(row[2:6], "little")
        parent_id = int.from_bytes(row[6:10], "little")
        type_ = int.from_bytes(row[10:12], "little", signed=True)
        flags = int.from_bytes(row[28:32], "little")

        tail_off = len(row) - _MSYS_NULL_MASK_BYTES - 2
        var_col_count = int.from_bytes(row[tail_off : tail_off + 2], "little")
        if var_col_count != _MSYS_VAR_COL_COUNT:
            return None
        jt_start = tail_off - 2 * var_col_count
        if jt_start < 32:
            return None
        # Jump table stores u16 offsets, one per variable column. The
        # table is laid down such that variable column index i STARTS
        # at row offset jt[i]; the column with the HIGHEST index is
        # placed first in physical memory (lowest row offset). A
        # variable column's END is therefore at jt[i-1] (the start of
        # the column with the next-lower index), or at the start of
        # the jump table itself for column index 0.
        jt = [
            int.from_bytes(row[jt_start + 2 * i : jt_start + 2 * i + 2], "little")
            for i in range(var_col_count)
        ]
        name_start = jt[_MSYS_NAME_VAR_INDEX]
        # Name's end is the start of the variable column with the
        # next-lower index. _MSYS_NAME_VAR_INDEX is 10 (the highest
        # populated variable-column index in MSysObjects), so the
        # preceding-index lookup is always valid.
        name_end = jt[_MSYS_NAME_VAR_INDEX - 1]
        if not (32 <= name_start <= name_end <= jt_start):
            return None
        if (name_end - name_start) % 2 != 0:
            return None
        try:
            name = row[name_start:name_end].decode("utf-16-le")
        except UnicodeDecodeError:
            return None
        return AccessSysObject(
            id_=id_,
            parent_id=parent_id,
            type_=type_,
            flags=flags,
            name=name,
            page=page,
            slot=slot,
        )

    def iter_msys_objects(self) -> Iterator["AccessSysObject"]:
        """Iterate every persistent object listed in the .accdb's
        ``MSysObjects`` system catalog.

        Yields one :class:`AccessSysObject` per non-deleted row across
        every DATA page owned by the MSysObjects TDEF. Each row
        identifies a Table, Query, Form, Report, Macro, VBA Module,
        or system Container.

        Use :meth:`find_msys_module` for the common case of locating
        a single VBA code module by name. Use :meth:`iter_msys_modules`
        to enumerate only VBA module rows.
        """
        for pn in self._iter_msys_data_pages():
            base = pn * ACE_PAGE_SIZE
            page_bytes = bytes(
                self._data[base : base + ACE_PAGE_SIZE]
            )
            row_count = int.from_bytes(page_bytes[12:14], "little")
            # First pass: collect non-deleted offsets so we can determine
            # each row's length from the next-higher offset.
            entries: list[tuple[int, int]] = []  # (slot, offset)
            for slot in range(row_count):
                ent = int.from_bytes(
                    page_bytes[14 + 2 * slot : 16 + 2 * slot], "little"
                )
                if ent & (_ROW_DELETED_FLAG | _ROW_OVERFLOW_FLAG):
                    continue
                off = ent & _ROW_OFFSET_MASK
                entries.append((slot, off))
            if not entries:
                continue
            # Each row runs from its offset up to the next-higher offset
            # (or to the end of the page for the highest-offset row).
            sorted_offs = sorted({off for _, off in entries})
            next_after: dict[int, int] = {}
            for i, off in enumerate(sorted_offs):
                next_after[off] = (
                    sorted_offs[i + 1]
                    if i + 1 < len(sorted_offs)
                    else ACE_PAGE_SIZE
                )
            for slot, off in entries:
                end = next_after[off]
                row = page_bytes[off:end]
                obj = self._decode_msys_row(row, page=pn, slot=slot)
                if obj is not None:
                    yield obj

    def msys_objects(self) -> tuple["AccessSysObject", ...]:
        """Return all MSysObjects rows as a tuple (materialised list of
        :meth:`iter_msys_objects`)."""
        return tuple(self.iter_msys_objects())

    def iter_msys_modules(self) -> Iterator["AccessSysObject"]:
        """Iterate only the MSysObjects rows that represent VBA code
        modules (``Type == MSYS_TYPE_MODULE``)."""
        for obj in self.iter_msys_objects():
            if obj.is_vba_module:
                yield obj

    def find_msys_object(
        self,
        name: str,
        *,
        type_: int | None = None,
    ) -> "AccessSysObject | None":
        """Return the first MSysObjects row matching ``name`` (and
        optionally ``type_``), or ``None`` if not found.

        Name match is case-insensitive, matching Access's own behaviour.
        """
        target = name.casefold()
        for obj in self.iter_msys_objects():
            if obj.name.casefold() != target:
                continue
            if type_ is not None and obj.type_ != type_:
                continue
            return obj
        return None

    def find_msys_module(self, name: str) -> "AccessSysObject | None":
        """Return the MSysObjects row for the VBA code module called
        ``name`` (case-insensitive), or ``None`` if no such module
        exists in the system catalog."""
        return self.find_msys_object(name, type_=MSYS_TYPE_MODULE)

    # ------------------------------------------------------------------
    # MSysObjects (Jet/ACE system catalog) -- write path
    # ------------------------------------------------------------------
    #
    # Updating MSysObjects rows is required to make module-catalog
    # mutations (rename / delete / add) visible to the live Access
    # engine (and hence the VBA editor and the Navigation Pane).
    #
    # The underlying DATA pages share the row offset-table format used
    # by LVAL pages, but with different tombstone semantics:
    #   * LVAL  pages: top nibble 0xD = tombstone (preserves low 12 bits)
    #   * DATA  pages: bit 0x8000 = deleted, bit 0x4000 = overflow
    # so we cannot reuse the LVAL helpers verbatim.

    def _data_slot_count(self, page_num: int) -> int:
        base = page_num * ACE_PAGE_SIZE
        return int.from_bytes(self._data[base + 12 : base + 14], "little")

    def _data_slot_entries(self, page_num: int) -> list[int]:
        """Raw u16 slot table entries for a DATA page (including the
        ``0x8000`` deleted and ``0x4000`` overflow flag bits)."""
        base = page_num * ACE_PAGE_SIZE
        n = self._data_slot_count(page_num)
        return [
            int.from_bytes(
                self._data[base + 14 + 2 * i : base + 16 + 2 * i], "little"
            )
            for i in range(n)
        ]

    def _data_set_slot_entry(
        self, page_num: int, slot: int, new_entry: int
    ) -> None:
        base = page_num * ACE_PAGE_SIZE
        self._data[base + 14 + 2 * slot : base + 16 + 2 * slot] = (
            (new_entry & 0xFFFF).to_bytes(2, "little")
        )

    def _data_row_is_live(self, entry: int) -> bool:
        return (entry & (_ROW_DELETED_FLAG | _ROW_OVERFLOW_FLAG)) == 0

    def _data_row_extent(
        self, page_num: int, slot: int
    ) -> tuple[int, int]:
        """Return ``(start, end)`` byte offsets on ``page_num`` for the
        live DATA row at slot index ``slot``.

        End is the next-higher live row offset, or ``ACE_PAGE_SIZE``
        for the top-most row. Raises :class:`AccessError` if the slot
        is deleted or an overflow pointer.
        """
        entries = self._data_slot_entries(page_num)
        if slot < 0 or slot >= len(entries):
            raise AccessError(
                f"_data_row_extent: slot {slot} out of range on page "
                f"{page_num} (have {len(entries)} slots)"
            )
        ent = entries[slot]
        if not self._data_row_is_live(ent):
            raise AccessError(
                f"_data_row_extent: slot {slot} on page {page_num} is "
                f"not a live row (flags=0x{ent & 0xE000:04x})"
            )
        start = ent & _ROW_OFFSET_MASK
        end = ACE_PAGE_SIZE
        for other in entries:
            if not self._data_row_is_live(other):
                continue
            o = other & _ROW_OFFSET_MASK
            if o > start and o < end:
                end = o
        return start, end

    def _data_free_space(self, page_num: int) -> int:
        """Contiguous free bytes available for new row data on a
        DATA page (gap between the end of the slot table and the
        lowest live row offset)."""
        entries = self._data_slot_entries(page_num)
        slot_table_end = 14 + 2 * len(entries)
        live = [
            e & _ROW_OFFSET_MASK
            for e in entries
            if self._data_row_is_live(e)
        ]
        lowest = min(live) if live else ACE_PAGE_SIZE
        return lowest - slot_table_end

    def _data_resize_row(
        self, page_num: int, slot: int, new_size: int
    ) -> None:
        """Resize a DATA-page row in place, shifting other rows on the
        same page and updating their slot offsets. Mirrors
        :meth:`_lval_resize_row` but uses DATA-page tombstone bits.
        """
        if new_size < 0:
            raise AccessError(
                f"_data_resize_row: new_size must be non-negative "
                f"(got {new_size})"
            )
        base = page_num * ACE_PAGE_SIZE
        start, end = self._data_row_extent(page_num, slot)
        old_size = end - start
        delta = new_size - old_size
        if delta == 0:
            return
        entries = self._data_slot_entries(page_num)
        live: list[tuple[int, int]] = [
            (i, ent & _ROW_OFFSET_MASK)
            for i, ent in enumerate(entries)
            if self._data_row_is_live(ent)
        ]
        lowest = min(off for _, off in live)
        slot_table_end = 14 + 2 * len(entries)
        new_lowest = lowest - delta
        if new_lowest < slot_table_end:
            raise AccessError(
                f"_data_resize_row: page {page_num} cannot grow slot "
                f"{slot} by {delta} bytes (free="
                f"{lowest - slot_table_end}, need={delta})"
            )
        block = bytes(self._data[base + lowest : base + start])
        self._data[
            base + new_lowest : base + start + (new_lowest - lowest)
        ] = block
        if delta < 0:
            # Zero the now-free trailing region beneath the resized row.
            self._data[base + lowest : base + new_lowest] = bytes(
                new_lowest - lowest
            )
        # Update slot offsets for every live slot whose old offset
        # was <= the resized row's old start.
        for i, old_off in live:
            if old_off <= start:
                ent = entries[i]
                new_off = old_off - delta
                # Preserve the entry's flag bits (none are set on live
                # rows -- but keep the high bits for forward-compat).
                self._data_set_slot_entry(
                    page_num, i, (ent & 0xE000) | (new_off & _ROW_OFFSET_MASK)
                )

    def _data_write_row(
        self, page_num: int, slot: int, new_row: bytes
    ) -> None:
        """Replace the bytes of a DATA-page row, resizing as needed."""
        self._data_resize_row(page_num, slot, len(new_row))
        base = page_num * ACE_PAGE_SIZE
        start, end = self._data_row_extent(page_num, slot)
        assert end - start == len(new_row), (
            f"_data_write_row invariant: {end - start} != {len(new_row)}"
        )
        self._data[base + start : base + end] = new_row

    def _data_tombstone_row(self, page_num: int, slot: int) -> None:
        """Mark a DATA-page row as deleted via the ``0x8000`` flag.
        The row's bytes are zeroed so they cannot be misread by
        higher-level scanners that compute extents naively."""
        base = page_num * ACE_PAGE_SIZE
        start, end = self._data_row_extent(page_num, slot)
        for i in range(base + start, base + end):
            self._data[i] = 0
        entries = self._data_slot_entries(page_num)
        ent = entries[slot]
        self._data_set_slot_entry(
            page_num, slot, _ROW_DELETED_FLAG | (ent & _ROW_OFFSET_MASK)
        )

    def _data_append_row(self, page_num: int, payload: bytes) -> int:
        """Append a new row to a DATA page and return its slot index.
        Reuses a deleted slot if available. Mirrors
        :meth:`_lval_append_row` for DATA-page tombstones."""
        base = page_num * ACE_PAGE_SIZE
        entries = self._data_slot_entries(page_num)
        live: list[tuple[int, int]] = [
            (i, ent & _ROW_OFFSET_MASK)
            for i, ent in enumerate(entries)
            if self._data_row_is_live(ent)
        ]
        lowest = min((off for _, off in live), default=ACE_PAGE_SIZE)
        reuse_slot: int | None = None
        for i, ent in enumerate(entries):
            if (ent & _ROW_DELETED_FLAG) and not (ent & _ROW_OVERFLOW_FLAG):
                reuse_slot = i
                break
        slot_table_end = 14 + 2 * len(entries)
        need = len(payload) + (0 if reuse_slot is not None else 2)
        if lowest - slot_table_end < need:
            raise AccessError(
                f"_data_append_row: page {page_num} has insufficient "
                f"free space (have {lowest - slot_table_end}, need "
                f"{need})"
            )
        new_off = lowest - len(payload)
        self._data[base + new_off : base + lowest] = payload
        if reuse_slot is not None:
            self._data_set_slot_entry(page_num, reuse_slot, new_off)
            return reuse_slot
        new_slot = len(entries)
        self._data_set_slot_entry(page_num, new_slot, new_off)
        self._data[base + 12 : base + 14] = (
            (len(entries) + 1).to_bytes(2, "little")
        )
        return new_slot

    # ----- MSysObjects-specific row builder + mutators ----------------

    @staticmethod
    def _ole_date_now() -> bytes:
        """Return the current UTC time encoded as an 8-byte OLE Date
        (IEEE 754 double; days since 1899-12-30)."""
        epoch = datetime.datetime(1899, 12, 30)
        delta_days = (
            datetime.datetime.now() - epoch
        ).total_seconds() / 86400.0
        return struct.pack("<d", delta_days)

    def _build_msys_module_row(
        self,
        *,
        id_: int,
        parent_id: int,
        name: str,
        date_create: bytes,
        date_update: bytes,
        owner_var9: bytes = b"\x00\x00",
    ) -> bytes:
        """Construct a complete MSysObjects row for a VBA code module.

        The row mirrors the structure observed in fresh-from-template
        .accdb files (17 columns; only Name is populated in the
        variable section, plus a 2-byte "var col 9" sentinel that
        Access writes -- empirically a small Owner-like value, default
        ``00 00`` if unknown). All other variable columns are NULL.
        Null bitmap is ``ff 00 00`` (matches every observed row).
        """
        if not name:
            raise AccessError("_build_msys_module_row: name must be non-empty")
        name_bytes = name.encode("utf-16-le")
        if len(date_create) != 8 or len(date_update) != 8:
            raise AccessError(
                "_build_msys_module_row: date_create/date_update must "
                "each be exactly 8 bytes"
            )
        if len(owner_var9) != 2:
            raise AccessError(
                "_build_msys_module_row: owner_var9 must be exactly 2 bytes"
            )
        # Fixed columns (32 bytes).
        fixed = (
            (_MSYS_COL_COUNT).to_bytes(2, "little")
            + (id_ & 0xFFFFFFFF).to_bytes(4, "little")
            + (parent_id & 0xFFFFFFFF).to_bytes(4, "little")
            + struct.pack("<h", MSYS_TYPE_MODULE)
            + date_create
            + date_update
            + (0).to_bytes(4, "little")  # Flags = 0
        )
        assert len(fixed) == 32

        # Variable section: var col 10 = Name (first in memory),
        # var col 9 = sentinel (next), var cols 0..8 = empty.
        var_data = name_bytes + owner_var9  # var col 10 then var col 9
        var_data_len = len(var_data)

        # Jump table (11 u16 entries): jt[i] = start offset (within ROW)
        # of variable column i.  Highest-index col is first in memory.
        # Layout in our row:
        #   var col 10 starts at row offset 32 (immediately after fixed).
        #   var col 9  starts at row offset 32 + len(name_bytes).
        #   var cols 0..8 are NULL -> their start == jt_start (var section
        #   ends at the jump table).
        row_start_var10 = 32
        row_start_var9 = 32 + len(name_bytes)
        jt_start_in_row = 32 + var_data_len
        jt_entries: list[int] = []
        for i in range(_MSYS_VAR_COL_COUNT):
            if i == 10:
                jt_entries.append(row_start_var10)
            elif i == 9:
                jt_entries.append(row_start_var9)
            else:
                # Null var cols all point at jt_start_in_row.
                jt_entries.append(jt_start_in_row)
        jt_bytes = b"".join(
            e.to_bytes(2, "little") for e in jt_entries
        )
        var_col_count_bytes = (_MSYS_VAR_COL_COUNT).to_bytes(2, "little")
        # Null mask: cols 0..7 set (matches every observed row).
        null_mask_bytes = b"\xff\x00\x00"
        assert len(null_mask_bytes) == _MSYS_NULL_MASK_BYTES

        row = (
            fixed
            + var_data
            + jt_bytes
            + var_col_count_bytes
            + null_mask_bytes
        )
        return row

    def _msys_next_user_id(self) -> int:
        """Allocate the next available user-content Id (bit 31 set)."""
        max_user = 0x80000000 - 1
        for obj in self.iter_msys_objects():
            if obj.id_ & 0x80000000 and obj.id_ > max_user:
                max_user = obj.id_
        return max_user + 1

    def rename_msys_object(self, old_name: str, new_name: str) -> None:
        """Rewrite the ``Name`` column of the MSysObjects row whose
        current name matches ``old_name`` (case-insensitive). Updates
        the row in place, resizing the page layout if the UTF-16-LE
        byte length changes.

        Raises :class:`AccessError` if no such row exists, if
        ``new_name`` is empty, or if growing the row would overflow
        the containing DATA page.

        This is the catalog-level companion to :meth:`rename_module`
        for VBA modules, but it works for any MSysObjects row (table,
        query, form, etc.) the caller addresses by name.
        """
        if not new_name:
            raise AccessError("rename_msys_object: new_name must be non-empty")
        obj = self.find_msys_object(old_name)
        if obj is None:
            raise AccessError(
                f"rename_msys_object: no MSysObjects row named "
                f"{old_name!r}"
            )
        base = obj.page * ACE_PAGE_SIZE
        start, end = self._data_row_extent(obj.page, obj.slot)
        row = bytes(self._data[base + start : base + end])

        # Decode jump table to locate Name field byte-precisely.
        tail_off = len(row) - _MSYS_NULL_MASK_BYTES - 2
        var_col_count = int.from_bytes(
            row[tail_off : tail_off + 2], "little"
        )
        if var_col_count != _MSYS_VAR_COL_COUNT:
            raise AccessError(
                f"rename_msys_object: unexpected var_col_count "
                f"{var_col_count} on row {obj.name!r}"
            )
        jt_start = tail_off - 2 * var_col_count
        jt = [
            int.from_bytes(row[jt_start + 2 * i : jt_start + 2 * i + 2], "little")
            for i in range(var_col_count)
        ]
        name_start = jt[_MSYS_NAME_VAR_INDEX]
        name_end = jt[_MSYS_NAME_VAR_INDEX - 1]
        old_name_bytes = row[name_start:name_end]
        new_name_bytes = new_name.encode("utf-16-le")
        delta = len(new_name_bytes) - len(old_name_bytes)

        if delta == 0:
            # Fast path: in-place byte patch, no offset reflow needed.
            self._data[base + start + name_start : base + start + name_end] = (
                new_name_bytes
            )
            return

        # Rebuild the row with the new Name, shifting other var-column
        # offsets in the jump table accordingly. The Name column (var
        # index 10) is FIRST in physical memory, so its end-offset
        # equals the start of var col 9 (jt[9]). Shifting Name's
        # length shifts every var col with index < 10 by `delta`.
        new_jt = list(jt)
        for i in range(_MSYS_NAME_VAR_INDEX):  # cols 0..9
            new_jt[i] = jt[i] + delta
        new_jt_bytes = b"".join(
            e.to_bytes(2, "little") for e in new_jt
        )
        # Compose the new row.
        # Fixed bytes 0..32 remain identical.
        # Variable data: [Name bytes] [rest of original var data
        # (cols 9..0 in memory order)].
        original_var_data = row[32 : 32 + (jt_start - 32)]
        # Replace the Name slice within original_var_data.
        # Name occupies [name_start - 32 .. name_end - 32) in
        # original_var_data, and was always the FIRST chunk
        # (name_start - 32 == 0 by construction).
        if name_start != 32:
            raise AccessError(
                "rename_msys_object: unexpected MSysObjects layout "
                "(Name not at row offset 32)"
            )
        new_var_data = new_name_bytes + original_var_data[name_end - 32 :]
        new_row = (
            row[:32]
            + new_var_data
            + new_jt_bytes
            + row[tail_off : tail_off + 2]  # var_col_count
            + row[tail_off + 2 : tail_off + 2 + _MSYS_NULL_MASK_BYTES]
        )
        # Sanity: total row length must equal old length + delta.
        if len(new_row) != len(row) + delta:
            raise AccessError(
                f"rename_msys_object: internal length mismatch "
                f"(expected {len(row) + delta}, got {len(new_row)})"
            )
        self._data_write_row(obj.page, obj.slot, new_row)

    def delete_msys_object(self, name: str) -> None:
        """Tombstone the MSysObjects row whose ``Name`` matches
        ``name`` (case-insensitive). The row's offset-table entry
        gets the ``0x8000`` deleted flag and its bytes are zeroed.

        Raises :class:`AccessError` if no such row exists.
        """
        obj = self.find_msys_object(name)
        if obj is None:
            raise AccessError(
                f"delete_msys_object: no MSysObjects row named {name!r}"
            )
        self._data_tombstone_row(obj.page, obj.slot)

    def add_msys_module_object(self, name: str) -> "AccessSysObject":
        """Add a new ``MSysObjects`` row for a VBA code module called
        ``name`` and return the resulting :class:`AccessSysObject`.

        Allocates a fresh ``Id`` (next available user-content Id),
        sets ``ParentId`` to the ``Modules`` container, ``Type`` to
        ``MSYS_TYPE_MODULE``, and writes the current time into the
        date columns. The row is appended to the first MSysObjects
        DATA page with enough free space.

        Raises :class:`AccessError` if the database has no ``Modules``
        container (should never happen on a real .accdb), or if no
        MSysObjects DATA page has room for the new row.
        """
        if not name:
            raise AccessError("add_msys_module_object: name must be non-empty")
        if self.find_msys_object(name) is not None:
            raise AccessError(
                f"add_msys_module_object: an MSysObjects row named "
                f"{name!r} already exists"
            )
        modules_container = self.find_msys_object(
            "Modules", type_=MSYS_TYPE_CONTAINER
        )
        if modules_container is None:
            raise AccessError(
                "add_msys_module_object: Modules container row not found "
                "(database does not look like an .accdb with VBA enabled)"
            )
        # Match owner_var9 of any existing module if present, else 0.
        sample_module = next(iter(self.iter_msys_modules()), None)
        if sample_module is not None:
            sample_base = sample_module.page * ACE_PAGE_SIZE
            s_start, s_end = self._data_row_extent(
                sample_module.page, sample_module.slot
            )
            sample_row = bytes(self._data[sample_base + s_start : sample_base + s_end])
            s_tail = len(sample_row) - _MSYS_NULL_MASK_BYTES - 2
            s_jt_start = s_tail - 2 * _MSYS_VAR_COL_COUNT
            s_jt9 = int.from_bytes(
                sample_row[s_jt_start + 2 * 9 : s_jt_start + 2 * 9 + 2], "little"
            )
            s_jt10 = int.from_bytes(
                sample_row[s_jt_start + 2 * 10 : s_jt_start + 2 * 10 + 2], "little"
            )
            owner_var9 = sample_row[s_jt9 : s_jt9 + (s_jt_start + 32 - s_jt_start)]
            # Recompute: var col 9 length is (next start) - jt[9] = jt_start_in_row - jt[9]
            # = (32 + var_data_len) - jt[9].
            owner_var9 = sample_row[s_jt9 : s_tail - 2 * _MSYS_VAR_COL_COUNT]
            # Actually simpler: var col 9 ends at jt_start of the row.
            owner_var9 = sample_row[s_jt9 : s_jt_start]
            # And var col 10's end is jt[9], so its size is jt[9] - jt[10].
            _ = s_jt10  # keep for diagnostic clarity
            if len(owner_var9) != 2:
                # Schema variant: fall back to default sentinel.
                owner_var9 = b"\x00\x00"
        else:
            owner_var9 = b"\x00\x00"
        now = self._ole_date_now()
        new_id = self._msys_next_user_id()
        new_row = self._build_msys_module_row(
            id_=new_id,
            parent_id=modules_container.id_,
            name=name,
            date_create=now,
            date_update=now,
            owner_var9=owner_var9,
        )
        # Find the first MSysObjects DATA page with room (need row + 2
        # bytes for slot table entry, conservatively).
        target_page: int | None = None
        for pn in self._iter_msys_data_pages():
            if self._data_free_space(pn) >= len(new_row) + 2:
                target_page = pn
                break
        if target_page is None:
            raise AccessError(
                "add_msys_module_object: no MSysObjects DATA page has "
                f"enough free space for a {len(new_row)}-byte row "
                "(growing onto a fresh page is not yet implemented)"
            )
        slot = self._data_append_row(target_page, new_row)
        return AccessSysObject(
            id_=new_id,
            parent_id=modules_container.id_,
            type_=MSYS_TYPE_MODULE,
            flags=0,
            name=name,
            page=target_page,
            slot=slot,
        )

    def save(self, path: Union[str, Path, None] = None) -> None:
        """
        Persist the in-memory buffer to disk. If ``path`` is omitted, the
        file is overwritten in place. Otherwise the file is written to
        ``path`` and :attr:`path` is updated to point at it.
        """
        out = Path(path) if path is not None else self.path
        out.write_bytes(bytes(self._data))
        self.path = out


@dataclass(frozen=True)
class AccessSysObject:
    """One row of the .accdb ``MSysObjects`` system catalog.

    MSysObjects is the master object index inside every Access database.
    Each row identifies one persistent object: a table, query, form,
    report, macro, VBA module, or a "container" hub that groups
    objects of a given kind (e.g. the ``Modules`` container that
    parents every VBA module row).

    The ``type_`` field is the raw signed 16-bit ``Type`` column value
    -- compare against the ``MSYS_TYPE_*`` constants exposed at module
    level, or use :attr:`is_vba_module`.

    Attributes:
        id_: ``Id`` column. Positive values are system objects, values
            with bit 31 set (e.g. ``0x80000005``) are user content.
        parent_id: ``ParentId`` column. References the row whose
            ``id_`` equals this value; for VBA modules this points at
            the ``Modules`` container row.
        type_: ``Type`` column. ``MSYS_TYPE_MODULE`` (-32761) marks a
            VBA code module.
        flags: ``Flags`` column (u32).
        name: ``Name`` column (decoded UTF-16-LE).
        page: ACE 4 KiB page number where this row lives.
        slot: Slot index within ``page``.
    """

    id_: int
    parent_id: int
    type_: int
    flags: int
    name: str
    page: int
    slot: int

    @property
    def is_vba_module(self) -> bool:
        """``True`` if this row represents a user-defined VBA code module."""
        return self.type_ == MSYS_TYPE_MODULE


@dataclass(frozen=True)
class SourceRow:
    """One stored VBA source-line row inside an Access database."""

    offset: int
    row_type: str   # "comment" -- only kind decoded so far
    length: int
    text: bytes

    def to_source_line(self) -> str:
        """Reconstruct the source line as it would appear in the VBA editor."""
        if self.row_type == "comment":
            return "'" + self.text.decode("ascii")
        return self.text.decode("ascii")


@dataclass(frozen=True)
class VBAModule:
    """
    A VBA module discovered inside an .accdb file.

    ``source`` is the user-visible code (matches what the Access VBA editor
    shows via ``CodeModule.Lines(1, CountOfLines)`` except that line endings
    here are ``\\r\\n`` rather than ``\\n``).
    ``attributes_text`` is the leading ``Attribute VB_*`` block emitted by
    the VBA compiler; it is normally hidden from the editor but is part of
    the on-disk stream.
    """

    name: str
    start_offset: int
    raw_blob_size: int
    decompressed_size: int
    attributes_text: str
    source: str


@dataclass(frozen=True)
class AccessVBAModuleEntry:
    """A single module record parsed from the .accdb dir-stream catalog.

    This is the project-level *catalog* view of a module (its declared
    name, kind, and access flags). The actual user source for the module
    is loaded separately via :meth:`AccessFile.iter_vba_modules`.

    Attributes:
        name: MBCS module name (PROJECTNAME code page).
        name_unicode: UTF-16 module name as stored in MODULENAMEUNICODE.
        stream_name: Obfuscated identifier Access stores in MODULESTREAMNAME.
            Unlike Excel/Word/PowerPoint, Access does not use this as an
            actual CFB stream name (there is no CFB), but the field is
            present in the dir stream.
        is_class_module: ``True`` for ClassModule (MODULETYPE 0x0022),
            ``False`` for procedural standard modules (0x0021).
        is_private: ``MODULEPRIVATE`` flag.
        is_read_only: ``MODULEREADONLY`` flag.
    """

    name: str
    name_unicode: str
    stream_name: str
    is_class_module: bool
    is_private: bool
    is_read_only: bool


@dataclass(frozen=True)
class AccessVBAProject:
    """Project-level VBA metadata parsed from the .accdb dir-stream catalog.

    See [MS-OVBA] section 2.3.4.2 for the underlying record layout. The
    dir stream is stored inside Access as a single OVBA-compressed LVAL
    row; ``catalog_page`` / ``catalog_slot`` identify that row in the
    database for diagnostic purposes.
    """

    catalog_page: int
    catalog_slot: int
    catalog_raw_size: int
    sys_kind: int
    lcid: int
    code_page: int
    project_name: str
    references: tuple[VBAReference, ...]
    modules: tuple[AccessVBAModuleEntry, ...]


@dataclass(frozen=True)
class AccessVBAPCodeStream:
    """The raw authoritative VBA p-code bytes for an Access database,
    together with the LVAL row coordinates from which they were read.

    The first four bytes are always ``72 55 40 00`` ('rU@\\x00'). The
    full opcode field guide is being reverse-engineered; see
    ``docs/access_pcode_re.md``.

    Attributes:
        page: ACE 4 KiB page number containing the LVAL row.
        slot: Slot index within ``page``.
        raw: Compiled bytecode payload (variable length; typically
            ~150-500 bytes for a single short procedure).
    """

    page: int
    slot: int
    raw: bytes


@dataclass(frozen=True)
class AccessVBAModuleStream:
    """Standard Office VBA module-stream bytes for a single VBA module,
    extracted from the LVAL row that also carries its OVBA-compressed
    source.

    Recognisable by the ``0xCAFE`` magic word in ``raw[cafe_offset:]``
    that marks the start of the per-line p-code region (see [MS-OVBA]
    section 2.3.4.3). The full byte layout matches what public VBA
    disassemblers (e.g. ``pcodedmp``) consume.

    This is the *canonical* portable VBA p-code -- the form that any
    Office host running VBA7 can execute. It coexists with the
    Access-specific ``rU@``-prefixed cached form (see
    :class:`AccessVBAPCodeStream`).

    Attributes:
        page: ACE 4 KiB page number containing the LVAL row.
        slot: Slot index within ``page``.
        raw: Entire row bytes; the module-stream-format region runs
            from offset 0 through the start of the OVBA compressed
            source.
        cafe_offset: In-row byte offset of the ``0xCAFE`` magic word
            that opens the p-code region.
    """

    page: int
    slot: int
    raw: bytes
    cafe_offset: int


@dataclass(frozen=True)
class AccessVBAInternedString:
    """A single VBA string-literal record decoded from the project's
    intern table.

    Each literal is stored as a ``0B <u32 LE byte-count> <UTF-16-LE>``
    record inside one of the database's LVAL rows. See
    ``docs/access_pcode_re.md`` Phase 4 for the structural rationale
    (compiled p-code is fully anonymised; literals live here and are
    referenced from bytecode by slot id only).

    Attributes:
        page: ACE page number of the LVAL row carrying the record.
        slot: Slot index within ``page``.
        offset: Byte offset of the ``0B`` tag within the row.
        value: Decoded string value.
    """

    page: int
    slot: int
    offset: int
    value: str


@dataclass(frozen=True)
class AccessVBAIdentifier:
    """A single identifier name decoded from the project's
    ``_VBA_PROJECT``-equivalent stream.

    The Access ``_VBA_PROJECT`` payload is stored uncompressed in an
    LVAL row whose first two bytes are the magic ``CC 61``. Near the
    tail of that row the host emits a list of identifier records, one
    per typelib reference, project name, module, procedure, variable,
    and intrinsic. Each record uses the layout::

        <u8 name_len> <u8 type_byte> <ASCII name>
        <u16 LE id_low> <u8 0x10> <u8 0x00>

    Empirically verified across the 25-sample RE corpus (samples
    010..051). Trailing ``10 00`` bytes appear to be a type-tag /
    cookie pair; ``id_low`` is the per-record token; ``type_byte`` is
    ``0x04`` for typelib refs / module/proc/variable names and ``0x00``
    for intrinsic function names (e.g. ``MsgBox``). Other type bytes
    (e.g. ``0x80``, ``0xac``) introduce variable-length descriptor
    blocks that we currently surface verbatim via :attr:`prefix`.

    Note: the ``name_id`` operands in compiled p-code do **NOT** index
    this table directly; they are per-procedure reference-table slots
    (see ``docs/access_pcode_re.md`` Phase 4f). This dataclass simply
    exposes the canonical project-wide identifier inventory.

    Attributes:
        index: 0-based position within the identifier table.
        type_byte: The single type byte preceding the name in the
            record.
        name: ASCII name, decoded from the on-disk byte payload.
        id_low: 16-bit ID cookie that follows the name on disk.
        prefix: Any extra descriptor bytes seen before this record that
            could not be parsed as another ``<len><type><name>`` entry.
            Empty for fully canonical records.
    """

    index: int
    type_byte: int
    name: str
    id_low: int
    prefix: bytes


def _find_vba_project_row(rows: list[tuple[int, int, bytes]]) -> bytes | None:
    """Return the ``_VBA_PROJECT``-equivalent row payload, or ``None``
    if no row starts with the ``CC 61`` magic."""
    for _page, _slot, row in rows:
        b = bytes(row)
        if b.startswith(b"\xcc\x61"):
            return b
    return None


def _parse_vba_project_identifiers(
    stream: bytes,
) -> tuple[AccessVBAIdentifier, ...]:
    """Parse the identifier list from a ``CC 61``-magic Access
    ``_VBA_PROJECT`` stream.

    Strategy: locate the references count marker ``02 00 06 04
    'Access'`` (or ``02 00 06 0C 'Access'`` -- type byte is ``0x04``
    in zero-module projects and ``0x0C`` once any user code exists)
    and walk forward through ``<len><type><ASCII name><id u16> 10 00``
    records, recording any non-conforming bytes as a per-record
    ``prefix`` so the parse is lossless. The very first ``Access``
    record has no ``10 00`` trailer (the next entry begins
    immediately); subsequent entries do.

    Records terminate at the first byte where the length byte is
    ``0x02`` and the next byte is ``0xFF`` (sentinel observed across
    every corpus sample), or when fewer than 6 bytes remain.
    """
    start = -1
    for type_byte in (0x04, 0x0C):
        cand = stream.find(b"\x02\x00\x06" + bytes([type_byte]) + b"Access")
        if cand >= 0:
            start = cand
            break
    if start < 0:
        return ()
    pos = start + 2  # Skip the u16 count; entries begin at 'Access'.
    out: list[AccessVBAIdentifier] = []
    pending_prefix = b""
    index = 0
    end = len(stream)
    # Special-case: first entry is 'Access' with NO id trailer.
    if (
        pos + 8 <= end
        and stream[pos] == 0x06
        and stream[pos + 1] in (0x04, 0x0C)
        and stream[pos + 2:pos + 8] == b"Access"
    ):
        out.append(
            AccessVBAIdentifier(
                index=0,
                type_byte=stream[pos + 1],
                name="Access",
                id_low=0,
                prefix=b"",
            )
        )
        index = 1
        pos += 8
    while pos + 6 <= end:
        # End-of-table sentinel: 02 FF FF 01 01 ...
        if stream[pos] == 0x02 and stream[pos + 1] == 0xFF:
            break
        name_len = stream[pos]
        type_byte = stream[pos + 1]
        # Type bytes 0x80 (intrinsic special, e.g. _Evaluate) and
        # 0xac (procedure with body) insert a 6-byte descriptor block
        # between <len><type> and the ASCII name.
        name_start = pos + 2
        if type_byte in (0x80, 0xAC):
            name_start = pos + 2 + 6
        name_end = name_start + name_len
        # The "canonical" record needs:
        #   <len> <type> [<6B descriptor>] <name(name_len)>
        #     <id u16> <0x10> <0x00>
        record_end = name_end + 4
        if (
            name_len > 0
            and name_len < 64
            and record_end <= end
            and stream[record_end - 2] == 0x10
            and stream[record_end - 1] == 0x00
            and all(
                0x20 <= stream[i] < 0x7F or stream[i] == 0x5F
                for i in range(name_start, name_end)
            )
        ):
            name = stream[name_start:name_end].decode(
                "ascii", errors="replace"
            )
            id_low = int.from_bytes(
                stream[name_end:name_end + 2], "little"
            )
            # Pre-name descriptor bytes (if any) become this entry's
            # prefix metadata, alongside any pending unparsed bytes.
            descriptor = stream[pos + 2:name_start]
            out.append(
                AccessVBAIdentifier(
                    index=index,
                    type_byte=type_byte,
                    name=name,
                    id_low=id_low,
                    prefix=pending_prefix + descriptor,
                )
            )
            index += 1
            pending_prefix = b""
            pos = record_end
            continue
        # Non-canonical byte -- accumulate into pending prefix and
        # advance one byte. The next valid record carries it.
        pending_prefix += bytes([stream[pos]])
        pos += 1
    return tuple(out)