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