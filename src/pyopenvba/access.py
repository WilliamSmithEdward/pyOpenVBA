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

    def _lval_next_page(self, page_num: int) -> int:
        """
        Return the next page in this LVAL chain, or 0 if this is the last page.

        For ACE LVAL pages, the chain pointer is a 3-byte little-endian
        value at offsets 21..23 of the page header (byte 24 is part of the
        following field). A value of 0 marks end-of-chain.
        """
        base = page_num * ACE_PAGE_SIZE
        return int.from_bytes(self._data[base + 21 : base + 24], "little")

    def _lval_continuation_offset(self, page_num: int) -> int:
        """
        For a continuation LVAL page, return the absolute file offset at
        which this page's contribution to the long-value byte stream begins.

        Page header bytes 14..15 (little-endian) give the page-relative
        offset of the record header for this page's long-value chunk; the
        first 4 bytes at that offset are a record prefix and the stream
        bytes themselves start 4 bytes later.
        """
        base = page_num * ACE_PAGE_SIZE
        rec_off = int.from_bytes(self._data[base + 14 : base + 16], "little")
        return base + rec_off + 4

    def _read_lval_chain(self, start_offset: int) -> bytes:
        """
        Walk the LVAL page chain starting at ``start_offset`` (the absolute
        file offset where the long value begins on its first page) and
        return the reassembled byte stream.

        The first page contributes bytes from ``start_offset`` to the end
        of that page. Each subsequent chained page contributes from
        :meth:`_lval_continuation_offset` to the end of that page.
        """
        return b"".join(
            bytes(self._data[off : off + length])
            for off, length in self._lval_segments(start_offset)
        )

    def _lval_segments(self, start_offset: int) -> list[tuple[int, int]]:
        """
        Return the list of ``(file_offset, length)`` slots in the LVAL
        chain anchored at ``start_offset``, in stream order. Concatenating
        the bytes at these slots reproduces the long-value byte stream.

        The first slot starts at ``start_offset`` and runs to the end of
        its page. Each subsequent slot runs from
        :meth:`_lval_continuation_offset` to the end of its page.
        """
        page_num = start_offset // ACE_PAGE_SIZE
        page_end = (page_num + 1) * ACE_PAGE_SIZE
        segs: list[tuple[int, int]] = [(start_offset, page_end - start_offset)]
        next_page = self._lval_next_page(page_num)
        while next_page:
            start = self._lval_continuation_offset(next_page)
            end = (next_page + 1) * ACE_PAGE_SIZE
            segs.append((start, end - start))
            next_page = self._lval_next_page(next_page)
        return segs

    def _find_ovba_signature_offsets(self) -> list[int]:
        """
        Return every file offset that plausibly begins an MS-OVBA stream:
        a 0x01 signature byte followed by a 2-byte little-endian chunk
        header whose top nibble is 0xB (signature bits = 0b011, flag = 1
        for a compressed chunk).
        """
        data = self._data
        out: list[int] = []
        i = 0
        while True:
            j = data.find(b"\x01", i)
            if j < 0 or j + 3 > len(data):
                return out
            i = j + 1
            if (data[j + 2] & 0xF0) == 0xB0:
                out.append(j)

    def iter_vba_modules(self) -> Iterator["VBAModule"]:
        """
        Discover and yield every VBA module embedded in this database.

        Implementation strategy
        -----------------------
        Each VBA module is stored as a single MS-OVBA compressed stream on
        one or more chained LVAL pages (see module docstring). We scan the
        file for plausible OVBA signature bytes, walk the LVAL chain at
        each candidate, attempt MS-OVBA decompression, and accept any
        result that begins with ``Attribute VB_Name = "..."`` (every Office
        VBA module starts with this attribute line).

        This avoids any dependency on parsing the Access system catalog
        (MSysObjects / MSysAccessStorage) -- a reasonable trade-off until
        write support requires us to allocate / re-link LVAL chains.
        """
        for off in self._find_ovba_signature_offsets():
            try:
                blob = self._read_lval_chain(off)
                raw = _ovba_decompress(blob, stream_name=f"accdb@0x{off:X}")
            except Exception:
                continue
            if not raw.startswith(b"Attribute VB_Name = "):
                continue
            text = raw.decode("latin-1")
            # Split off the leading Attribute VB_* preamble lines from the
            # user-visible source body.
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
                start_offset=off,
                raw_blob_size=len(blob),
                decompressed_size=len(raw),
                attributes_text="\r\n".join(lines[:body_start]),
                source=body,
            )

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
