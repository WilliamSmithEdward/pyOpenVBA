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

from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Union

from pyopenvba.exceptions import PyOpenVBAError, UnsupportedFormatError
from pyopenvba.vba import decompress as _ovba_decompress


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

        Access keeps shadow / undo copies of edited modules, so the same
        module name may appear multiple times in
        :meth:`iter_vba_modules`. This helper deduplicates.
        """
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
